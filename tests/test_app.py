"""Smoke tests for the Streamlit UI.

Drives the real app through Streamlit's AppTest harness with the synthetic
samples, so a change that breaks the page - not just the counting - fails here.
"""

import re

import pytest

from conftest import DETAILS, EXPECTED, OLDER_VERSION, PORTAL, POWERBI, ROOT, SAMPLES

pytest.importorskip("streamlit.testing.v1")
from streamlit.testing.v1 import AppTest  # noqa: E402

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

AREAS = {k: DETAILS[k]["practice_area"].split(" > ")[1] for k in SAMPLES}
PAGE = "\n".join(AREAS[k] for k in ("alpha", "beta", "gamma"))


def _upload(path, mime):
    return (path.name, path.read_bytes(), mime)


def _load(at, documents=None, with_portal=True):
    uploaders = at.get("file_uploader")
    uploaders[0].set_value(
        [_upload(p, DOCX_MIME) for p in (documents or list(SAMPLES.values()))]
    )
    uploaders[1].set_value(_upload(POWERBI, XLSX_MIME))
    if with_portal:
        uploaders[2].set_value(_upload(PORTAL, XLSX_MIME))
    return at.run()


@pytest.fixture(scope="module")
def app():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    return _load(at)


def _row_order(at):
    return next(b for b in at.get("selectbox") if b.label == "Row order")


def _with_pasted_page(order_text, documents=None):
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, documents)
    at.get("text_area")[0].set_value(order_text)
    return at.run()


# --- it runs ----------------------------------------------------------------


def test_app_starts_and_asks_for_files():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120).run()
    assert not at.exception
    assert len(at.get("file_uploader")) == 3
    assert at.info, "should prompt for files before any are supplied"


def test_app_runs_the_full_pipeline_without_error(app):
    assert not app.exception


def test_opening_memo_is_a_single_sentence():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120).run()
    assert len(at.info) == 1
    assert at.info[0].value == (
        "Add at least one submission document and the Power BI export to begin."
    )


def test_staff_portal_uploader_is_labelled_optional_and_has_a_hint():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120).run()
    uploader = at.get("file_uploader")[2]
    assert uploader.label == "Staff Portal export - optional (.xlsx or .csv)"
    assert uploader.help.startswith("This helps with populating the file name column")
    for point in ["file name", "practice area", "blank", "missing"]:
        assert point in uploader.help


def test_a_note_says_to_upload_the_whole_firm_each_time():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120).run()
    captions = " ".join(c.value for c in at.caption)
    assert "not just the latest batch" in captions
    assert "nothing is saved between runs" in captions


def test_document_count_is_shown(app):
    captions = [c.value for c in app.caption]
    assert any(f"**{len(SAMPLES)}** documents added" in c for c in captions)


def test_document_count_reads_zero_before_anything_is_added():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120).run()
    assert any("No documents added yet" in c.value for c in at.caption)


def test_single_document_is_not_pluralised():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, documents=[SAMPLES["alpha"]])
    assert any("**1** document added" in c.value for c in at.caption)


def test_privacy_wording(app):
    assert any(
        "Everything runs on your machine (offline) for data privacy." in c.value
        for c in app.caption
    )


# --- the summary ------------------------------------------------------------


def test_summary_counts_are_shown(app):
    summary = {m.label: m.value for m in app.metric}
    assert summary["Files read"] == f"{len(SAMPLES)} / {len(SAMPLES)}"
    assert summary["Duplicates"] == "0"
    assert int(summary["Corrections"]) > 0


def test_summary_has_no_unexplained_tiles(app):
    labels = {m.label for m in app.metric}
    assert "Metrics compared" not in labels
    assert "Matched export" not in labels
    assert "Missing links" in labels


def test_missing_links_reads_as_unknown_without_a_pasted_page(app):
    assert {m.label: m.value for m in app.metric}["Missing links"] == "—"


def test_files_read_tile_explains_a_failure(tmp_path):
    bad = tmp_path / "not-a-document.docx"
    bad.write_bytes(b"this is not a docx")

    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, documents=[SAMPLES["alpha"], bad], with_portal=False)

    tile = next(m for m in at.metric if m.label == "Files read")
    assert tile.value == "1 / 2"
    assert "could not" in tile.help
    assert any("could not be read" in m.value for m in at.markdown)


def test_unreadable_files_are_named_and_listed(tmp_path):
    bad = tmp_path / "not-a-document.docx"
    bad.write_bytes(b"this is not a docx")
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, documents=[SAMPLES["alpha"], bad], with_portal=False)

    problem = next(m.value for m in at.markdown if "could not be read" in m.value)
    assert "\n    - `not-a-document.docx`" in problem


# --- the results table ------------------------------------------------------


def test_results_table_has_a_row_per_document_and_the_agreed_columns(app):
    frame = app.dataframe[0].value
    assert frame.shape[0] == len(SAMPLES)
    assert list(frame.columns) == [
        "FIRMREF", "FIRM", "REGION", "PRACTICE AREA", "FILE NAME",
        "MATTERS", "CLIENTS", "NEW", "LPs", "NGs", "LAs",
        "NEEDS A LOOK", "DOUBLE-CHECKED",
    ]


def test_the_tick_column_starts_unticked(app):
    frame = app.dataframe[0].value
    assert "DOUBLE-CHECKED" in frame.columns
    assert frame["DOUBLE-CHECKED"].tolist() == [False] * len(SAMPLES)
    assert any("row(s) ticked" in c.value for c in app.caption)


def test_the_table_is_editable_by_default_and_colourable_on_demand(app):
    """Streamlit colours a table or lets you edit it, not both."""
    boxes = [c.label for c in app.checkbox]
    assert "Colour the cells (read-only)" in boxes
    frame = app.dataframe[0].value
    # editable by default, so the figures come through as numbers
    assert str(frame["MATTERS"].dtype) == "Int64"


def test_the_needs_a_look_column_says_why():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at)
    frame = at.dataframe[0].value
    notes = " ".join(str(v) for v in frame["NEEDS A LOOK"])
    assert "numbering does not add up" in notes


def test_no_practice_area_cell_is_blank_for_the_samples(app):
    assert (app.dataframe[0].value["PRACTICE AREA"].str.len() > 0).all()


def test_row_order_selector_is_offered(app):
    selects = app.get("selectbox")
    assert selects
    assert _row_order(app).label == "Row order"
    assert "Power BI export order" in _row_order(app).options


def test_pasted_order_option_is_hidden_until_a_sheet_is_supplied(app):
    assert "Staff Portal order (pasted)" not in _row_order(app).options


def test_paste_box_is_offered_and_empty_by_default(app):
    boxes = app.get("text_area")
    assert boxes
    assert boxes[0].value == ""


def test_download_button_is_offered(app):
    assert [b.label for b in app.get("download_button")] == [
        "Download .xlsx with highlighting"
    ]


def test_toggle_reveals_original_export_values(app):
    app.get("toggle")[0].set_value(True).run()
    assert not app.exception
    frame = app.dataframe[0].value
    assert "was NGs" in frame.columns
    row = frame[frame["FILE NAME"].str.contains("alpha")].iloc[0]
    assert row["NGs"] == EXPECTED["alpha"]["next_gen"]
    assert row["was NGs"] == 0


def test_the_table_row_count_is_explained(app):
    caption = next(c.value for c in app.caption if "row(s) in the table" in c.value)
    assert f"**{len(SAMPLES)}** row(s) in the table" in caption
    assert f"**{len(SAMPLES)}** submission document file(s) uploaded here" in caption
    assert "One row per firm, region and practice area." in caption


def test_the_row_count_shows_what_was_taken_off():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, documents=[SAMPLES["alpha"], OLDER_VERSION])

    caption = next(c.value for c in at.caption if "row(s) in the table" in c.value)
    assert "**1** row(s) in the table" in caption
    assert "**2** submission document file(s) uploaded here" in caption
    assert "**1** duplicate version(s) removed to keep the newer version" in caption


def test_older_submission_for_the_same_slot_is_left_out_of_the_table():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, documents=[OLDER_VERSION, SAMPLES["alpha"]])

    assert not at.exception
    frame = at.dataframe[0].value
    assert frame.shape[0] == 1, "the portal shows one row, so the table must too"
    assert frame.iloc[0]["FILE NAME"] == SAMPLES["alpha"].name
    assert any("older submission" in m.value for m in at.markdown)


def test_duplicate_upload_is_reported_and_left_out_of_the_table():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    files = [_upload(p, DOCX_MIME) for p in SAMPLES.values()]
    files.append(
        (SAMPLES["alpha"].name.replace(".docx", " (1).docx"),
         SAMPLES["alpha"].read_bytes(), DOCX_MIME)
    )
    uploaders = at.get("file_uploader")
    uploaders[0].set_value(files)
    uploaders[1].set_value(_upload(POWERBI, XLSX_MIME))
    uploaders[2].set_value(_upload(PORTAL, XLSX_MIME))
    at.run()

    assert not at.exception
    assert at.dataframe[0].value.shape[0] == len(SAMPLES)
    assert any("second copy" in m.value for m in at.markdown)


# --- the pasted page --------------------------------------------------------


def test_pasted_order_drives_the_table():
    at = _with_pasted_page("\n".join(AREAS[k] for k in ("gamma", "alpha", "beta")))
    assert not at.exception
    selector = _row_order(at)
    assert selector.options[0] == "Staff Portal order (pasted)"
    assert selector.value == "Staff Portal order (pasted)"
    assert any("3 entries, all recognised" in s.value for s in at.success)

    names = list(at.dataframe[0].value["FILE NAME"])
    assert names == [SAMPLES[k].name for k in ("gamma", "alpha", "beta")]


def test_wrong_column_is_refused_at_paste_time():
    at = _with_pasted_page(
        f"A Firm\t{AREAS['alpha']}\nA Firm\t{AREAS['beta']}"
    )
    assert not at.exception
    picker = next(
        b for b in at.get("selectbox")
        if b.label == "Which column holds the practice area?"
    )
    assert "0 of 2 recognised" in picker.options[0]
    assert "2 of 2 recognised" in picker.options[1]
    assert picker.value == picker.options[1]

    picker.set_value(picker.options[0]).run()
    assert not at.exception
    assert any("cannot order the table" in e.value for e in at.error)


def test_the_page_panel_can_be_collapsed():
    at = _with_pasted_page(PAGE)
    assert not at.exception
    assert "Against the pasted Staff Portal page" in [e.label for e in at.expander]


def test_the_page_panel_does_not_nest_expanders():
    """Streamlit refuses nested expanders; the no-document list must be plain."""
    from conftest import MANIFEST

    at = _with_pasted_page(MANIFEST["portal_rows_without_document"][0] + "\n" + PAGE)
    assert not at.exception
    assert [e.label for e in at.expander].count(
        "Against the pasted Staff Portal page"
    ) == 1
    assert not any("No document on the portal" in e.label for e in at.expander)
    assert any("No document on the portal" in m.value for m in at.markdown)


def test_row_order_dropdown_offers_every_mode_and_holds_a_choice():
    at = _with_pasted_page(PAGE)
    picker = _row_order(at)
    assert len(picker.options) == 6
    assert picker.value == "Staff Portal order (pasted)"

    picker.set_value("Region, then practice area (A-Z)").run()
    assert not at.exception
    assert _row_order(at).value == "Region, then practice area (A-Z)"


def test_checklist_is_labelled_as_its_own_list_not_table_rows():
    at = _with_pasted_page(PAGE)
    captions = [c.value for c in at.caption]
    assert any("pasted row(s), in your order" in c for c in captions)
    assert any("Pasted order from Staff Portal" in m.value for m in at.markdown)
    assert "Show only the rows in the table" in [c.label for c in at.checkbox]


def test_checklist_key_is_its_own_block_not_a_caption_line():
    at = _with_pasted_page(PAGE)
    assert any(e.label == "Legend" for e in at.expander)
    assert not any("✅ uploaded ❌ missing" in c.value for c in at.caption)


def test_every_checklist_status_is_explained_in_the_key():
    from conftest import MANIFEST

    at = _with_pasted_page(MANIFEST["portal_rows_without_document"][0] + "\n" + PAGE)
    key = next(m.value for m in at.markdown if "**uploaded**" in m.value)
    for mark, word in [("✅", "uploaded"), ("❌", "missing"),
                       ("⛓️‍💥", "no document link"),
                       ("⏳", "not started"), ("❔", "not found"),
                       ("🕓", "older version")]:
        assert mark in key and word in key
    assert "nothing to download or check" in key


def test_show_only_rows_in_the_table_matches_the_table_row_count():
    from conftest import MANIFEST

    at = _with_pasted_page(MANIFEST["portal_rows_without_document"][0] + "\n" + PAGE)
    rows_in_table = at.dataframe[0].value.shape[0]

    box = next(c for c in at.checkbox if c.label == "Show only the rows in the table")
    box.set_value(True).run()
    assert not at.exception

    checklist = next(m.value for m in at.markdown if re.match(r"^`\s*\d+`", m.value))
    assert checklist.count("✅") == rows_in_table
    assert "⛓️‍💥" not in checklist


def test_missing_links_tile_and_list_use_the_same_name():
    from conftest import MANIFEST

    at = _with_pasted_page(MANIFEST["portal_rows_without_document"][0] + "\n" + PAGE)
    tile = next(m for m in at.metric if m.label == "Missing links")
    assert tile.value == "1"
    assert "listed in the panel below" in tile.help

    heading = next(m.value for m in at.markdown if "No document on the portal" in m.value)
    assert heading.startswith("**No document on the portal (1)**")


def test_rows_covered_reads_as_rows_not_files():
    at = _with_pasted_page(PAGE)
    tile = next(m for m in at.metric if m.label == "Covered")
    assert "/" in tile.value
    assert "Rows, not files" in tile.help
    assert "Files read" in tile.help
    assert not any(m.label == "Uploaded" for m in at.metric)


def test_the_panel_figures_add_up_on_screen():
    from conftest import MANIFEST

    at = _with_pasted_page(MANIFEST["portal_rows_without_document"][0] + "\n" + PAGE)
    metrics = {m.label: m.value for m in at.metric}
    rows = int(metrics["Rows on the page"])
    covered, to_check = (int(x) for x in metrics["Covered"].split(" / "))
    assert rows == int(metrics["Missing links"]) + to_check
    assert to_check == covered + int(metrics["Missing"]) + int(
        metrics["Not started"]
    ) + int(metrics["Check on the page"])


def test_the_maths_explains_the_chain():
    at = _with_pasted_page(PAGE)
    maths = next(
        m.value for m in at.markdown if "rows pasted from the Staff Portal" in m.value
    )
    assert "separate documents to download" in maths
    assert "submission documents uploaded here" in maths
    assert "rows in the table" in maths


# --- the parse log ----------------------------------------------------------


def test_every_document_gets_a_parse_log_entry(app):
    labels = [e.label for e in app.expander]
    for path in SAMPLES.values():
        assert any(path.name in label for label in labels)


def test_parse_log_has_a_legend_for_its_icons(app):
    log = {h.value: h for h in app.subheader}["Per-document parse log"]
    assert log.help
    for mark in ["✅", "🕓", "♻️", "❓", "⚠️"]:
        assert mark in log.help
    assert "older version" in log.help


def test_an_older_version_says_it_was_not_counted():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, documents=[OLDER_VERSION, SAMPLES["alpha"]])

    labels = [e.label for e in at.expander]
    old_entry = next(l for l in labels if OLDER_VERSION.name in l)
    new_entry = next(l for l in labels if SAMPLES["alpha"].name in l)
    assert old_entry.startswith("🕓") and "not counted" in old_entry
    assert new_entry.startswith("✅") and "not counted" not in new_entry


# --- notes, problems and presentation --------------------------------------


def test_notes_and_problems_are_separate_sections(app):
    headings = [m.value for m in app.markdown]
    assert any("Needs attention" in h for h in headings) or any(
        "Nothing needs attention" in s.value for s in app.success
    )
    labels = [e.label for e in app.expander]
    assert any(label.startswith("Notes about the files read") for label in labels)
    assert not any("Notes and problems" in label for label in labels)


def test_a_note_explains_what_is_lost_without_the_portal_export():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at, with_portal=False)
    assert not at.exception
    notes = " ".join(m.value for m in at.markdown)
    assert "No Staff Portal export was added" in notes


def test_that_note_is_absent_when_the_portal_export_is_supplied(app):
    notes = " ".join(m.value for m in app.markdown)
    assert "No Staff Portal export was added" not in notes


def test_uncountable_metrics_are_labelled_as_an_instruction():
    from parsing import UNPARSED

    assert UNPARSED == "double-check manually"


def test_highlight_styles_carry_an_explicit_colour():
    """Pale fills need dark text, or they vanish in dark mode."""
    source = (ROOT / "app.py").read_text()
    assert "background-color: {fill}; color: {ON_FILL}" in source


def test_theme_is_not_pinned_to_light():
    config = (ROOT / ".streamlit" / "config.toml").read_text()
    assert 'base = "light"' not in config
    assert "Settings -> Theme" in config


def test_the_heading_is_discreet():
    """The table is the point of the page, not the name of the tool."""
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120).run()
    assert not at.title, "st.title is fixed at a size that dominates the page"
    heading = next(m.value for m in at.markdown if "Submission recount" in m.value)
    assert "opacity" in heading and "font-size" in heading


# --- starting a new batch ---------------------------------------------------


def _restart_button(at):
    return next(b for b in at.get("button") if b.label == "↺")


def test_the_restart_button_is_there_before_any_files():
    """It has to be reachable on the empty page, which stops early."""
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=120).run()
    assert not at.exception
    assert _restart_button(at)


def test_the_restart_button_asks_first():
    """Clicking it must confirm, not wipe the batch outright."""
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at)
    before = at.session_state["batch_id"]
    _restart_button(at).click().run()
    assert not at.exception
    assert at.session_state["batch_id"] == before
    assert any(b.label == "Start a new batch" for b in at.get("button"))


def test_confirming_clears_the_batch():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at)
    at.session_state["corrections"] = {("x.docx", "matters"): 3}
    at.session_state["checked_rows"] = {"x.docx"}
    at.run()

    _restart_button(at).click().run()
    next(b for b in at.get("button") if b.label == "Start a new batch").click().run()

    assert not at.exception
    assert at.session_state["batch_id"] == 1
    assert not at.session_state.get("corrections")
    assert not at.session_state.get("checked_rows")
    # Back to the empty page, asking for files again.
    assert any("Add at least one submission document" in i.value for i in at.get("info"))


def test_cancelling_keeps_the_batch():
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=300).run()
    _load(at)
    _restart_button(at).click().run()
    next(b for b in at.get("button") if b.label == "Cancel").click().run()

    assert not at.exception
    assert at.session_state["batch_id"] == 0
    assert not any("Add at least one submission document" in i.value for i in at.get("info"))
