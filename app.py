"""Streamlit UI for recounting law firm submission metrics.

Runs entirely on this machine. Nothing is uploaded anywhere: the documents are
read in memory by python-docx and the results are written back to your browser.

All parsing and matching lives in parsing.py / matching.py so it can be tested
without launching the app.
"""

from __future__ import annotations

import streamlit as st

from export_writer import CHECKED_COLUMN, build_workbook
from matching import (
    SORT_MODES,
    is_elite_area,
    build_portal_order,
    build_results,
    build_vocabulary,
    collapse_superseded,
    created_at_text,
    describe_record,
    difference_records,
    REVIEW_COLUMN,
    display_value,
    final_differs,
    final_value,
    load_portal,
    load_powerbi,
    needs_attention,
    normalise,
    read_pasted_order,
    override_key,
    reconcile,
    results_to_frame,
    review_note,
    sort_results,
    unordered_rows,
)
from parsing import METRIC_COLUMNS, METRIC_KEYS, METRIC_LABELS, UNPARSED, parse_document

# Pale fills, so the text on them must be dark whatever theme is in use - in
# dark mode the inherited text colour is light and would vanish.
GREEN = "#C6EFCE"
AMBER = "#FFE9A8"
ON_FILL = "#1A1A1A"

st.set_page_config(page_title="Submission recount", page_icon="📋", layout="wide")


# --- cached loaders ---------------------------------------------------------
# Keyed on the file's bytes, so re-running does not re-read a 12,000-row export.


@st.cache_data(show_spinner=False)
def _parse(name: str, payload: bytes):
    from io import BytesIO

    return parse_document(BytesIO(payload), filename=name)


@st.cache_data(show_spinner=False)
def _powerbi(payload: bytes):
    from io import BytesIO

    return load_powerbi(BytesIO(payload))


@st.cache_data(show_spinner=False)
def _portal(payload: bytes):
    from io import BytesIO

    return load_portal(BytesIO(payload))


def _sum(title: str, rows: list[tuple[int, str]]) -> str:
    """A titled column of figures, each line following from the one above."""
    body = "\n".join(f"{n:>5}  {label}" for n, label in rows)
    return f"**{title}**\n\n```\n{body}\n```"


def _arithmetic(results, table_results, duplicates, superseded, report) -> str:
    """Every figure on the page, in the order one turns into the next.

    Lines that would read 0 are left out: they add length, not information.
    """
    parts: list[str] = []

    if report is not None:
        rows = [(report.page_size, "rows pasted from the Staff Portal")]
        if report.without_document:
            rows.append((-len(report.without_document), "with no document on the portal"))
        if report.unrecognised:
            rows.append((-len(report.unrecognised), "not found for this firm"))
        if report.elite:
            rows.append((-len(report.elite), "Elite (not recounted)"))
        rows.append((len(report.expected), "rows to check"))
        if report.shared_rows:
            rows.append((-report.shared_rows, "share a document with another row"))
        rows.append((report.distinct_documents, "separate documents to download"))
        parts.append(_sum("Staff Portal page", rows))
        if report.repeats:
            parts.append(
                f"{len(report.repeats)} line(s) repeat the line above in the Staff Portal export. "
                "Each is counted as a document of its own."
            )

    rows = [(len(results), "submission documents uploaded here")]
    elite = [r for r in results if r.status == "elite"]
    if elite:
        rows.append((-len(elite), "Elite (not recounted)"))
    if superseded:
        rows.append((-len(superseded), "older versions removed"))
    if duplicates:
        rows.append((-len(duplicates), "second copies removed"))
    rows.append((len(table_results), "rows in the table"))
    parts.append(_sum("Your files", rows))

    if report is not None:
        # Compared by name: equal counts can hide a wrong file standing in for
        # a missing one.
        names = report.not_downloaded
        if names:
            parts.append(
                f"**Still to download ({len(names)}):** "
                + ", ".join(f"`{n}`" for n in names)
            )
        elif report.unclear:
            parts.append(
                f"Every document the Staff Portal export names is uploaded, but "
                f"{len(report.unclear)} line(s) need checking on the page."
            )
        else:
            parts.append("Every document on the page is uploaded.")

    firms = {normalise(r.firm) for r in table_results if r.firm}
    firm_rows = [r for r in powerbi.rows if normalise(r.get("FIRM_NAME")) in firms]
    export_rows = [
        r for r in firm_rows if not is_elite_area(r.get("CONCATENATED_PRACTICE_AREA_NAME"))
    ]
    if export_rows:
        # Two files can match one export row, so count the rows, not the files.
        matched = len({id(r.export_row) for r in table_results if r.export_row})
        rows = [(len(export_rows), "rows in the Power BI export for this firm")]
        if len(export_rows) > matched:
            rows.append((matched - len(export_rows), "not matched to any uploaded file"))
        rows.append((matched, "matched and compared"))
        title = "Power BI export"
    if len(export_rows) < len(firm_rows):
        title += f" ({len(firm_rows) - len(export_rows)} Elite row(s) left out)"
    parts.append(_sum(title, rows))

    return "\n\n".join(parts)


# --- inputs -----------------------------------------------------------------

# A quiet heading: the table is what matters on this page, not the name of the
# --- starting over ----------------------------------------------------------

# Every input widget is keyed to this number. Bumping it hands Streamlit a set
# of widgets it has never seen, which is the only way to empty a file uploader
# from code: there is no clear() on the widget itself.
batch = st.session_state.setdefault("batch_id", 0)


def _start_new_batch() -> None:
    """Drop everything belonging to the batch on screen."""
    for key in [k for k in st.session_state if k != "batch_id"]:
        del st.session_state[key]
    st.session_state["batch_id"] = batch + 1


@st.dialog("Start a new batch?")
def _confirm_new_batch() -> None:
    st.write(
        "This clears the documents you have added, both exports, the pasted "
        "Staff Portal order, and every figure you have typed in or ticked off."
    )
    st.caption(
        "Nothing is kept between batches, so download the results first if you "
        "still need them."
    )
    confirm, cancel = st.columns(2)
    if confirm.button("Start a new batch", type="primary", width="stretch"):
        _start_new_batch()
        st.rerun()
    if cancel.button("Cancel", width="stretch"):
        st.session_state.pop("confirming_new_batch", None)
        st.rerun()


# tool. Sized and dimmed by hand because st.title is fixed at a size that
# dominates everything under it.
heading, restart = st.columns([12, 1])
with heading:
    st.markdown(
        """
        <div style="font-size:1.15rem; font-weight:600; opacity:0.55;
                    letter-spacing:0.01em; margin:0 0 0.15rem 0;">
          Submission recount
        </div>
        """,
        unsafe_allow_html=True,
    )
with restart:
    if st.button(
        "↺",
        help="Start a new batch - clears everything on screen",
        width="stretch",
    ):
        st.session_state["confirming_new_batch"] = True
# Held open by a flag of its own. A dialog opened straight from the button
# closes again on the next rerun - which is the rerun its own buttons cause -
# so the confirmation would never be read.
if st.session_state.get("confirming_new_batch"):
    _confirm_new_batch()
st.caption(
    "Recounts the metrics in submission documents and compares them against the "
    "Power BI export. Everything runs on your machine (offline) for data privacy."
)

left, right = st.columns([2, 1])
with left:
    doc_files = st.file_uploader(
        "Submission documents (.docx)",
        type=["docx"],
        accept_multiple_files=True,
        key=f"docs_{batch}",
        help="Drag in as many as you like. A full firm of ~114 documents takes "
        "about 15 seconds, and uploading the lot each time keeps one downloaded results file "
        "covering everything you have downloaded so far.",
    )
    count = len(doc_files or [])
    st.caption(
        (
            f"**{count}** document{'' if count == 1 else 's'} added"
            if count
            else "No documents added yet"
        )
        + "  \nUpload all of a firm's documents each time, not just the latest "
        "batch — nothing is saved between runs."
    )
with right:
    powerbi_file = st.file_uploader(
        "Power BI export (.xlsx)", type=["xlsx"], key=f"powerbi_{batch}"
    )
    portal_file = st.file_uploader(
        "Staff Portal export - optional (.xlsx or .csv)",
        type=["xlsx", "csv"],
        key=f"portal_{batch}",
        help=(
            "This helps with populating the file name column - nothing else can "
            "supply it, since the Power BI export does not hold document file "
            "names.\n\nIt also fills the firm, region and practice area, and "
            "cross-checks them against the document. A document whose "
            "practice-area dropdown was never used has no region or practice area "
            "of its own, so without this file those cells come out blank.\n\n"
            "It is also what lets the app tell you which documents are missing "
            "from a batch."
        ),
    )

if not doc_files or not powerbi_file:
    st.info("Add at least one submission document and the Power BI export to begin.")
    st.stop()

# --- run --------------------------------------------------------------------

try:
    powerbi = _powerbi(powerbi_file.getvalue())
except ValueError as exc:
    st.error(f"Could not read the Power BI export: {exc}")
    st.stop()

portal = None
if portal_file is not None:
    try:
        portal = _portal(portal_file.getvalue())
    except ValueError as exc:
        st.error(f"Could not read the Staff Portal export: {exc}")
        st.stop()

# --- Staff Portal row order (pasted straight in) ----------------------------

portal_order = None
with st.expander("Staff Portal row order (optional)", expanded=False):
    st.caption(
        "Select the rows on the Staff Portal page, copy, and paste here to order "
        "the table the way the portal shows it. A whole pasted table is fine - the "
        "practice area column is found for you and checked against both exports. "
        "Nothing is uploaded."
    )
    order_text = st.text_area(
        "Paste the Staff Portal rows",
        height=160,
        placeholder="Professional negligence\nPensions\nProperty finance\n...",
        label_visibility="collapsed",
        key=f"order_{batch}",
    )

    if order_text.strip():
        vocabulary = build_vocabulary(powerbi, portal)
        pasted = read_pasted_order(order_text, vocabulary)

        chosen = pasted.suggested
        if pasted.width > 1:
            chosen = pasted.column_labels.index(
                st.selectbox(
                    "Which column holds the practice area?",
                    pasted.column_labels,
                    index=pasted.suggested,
                    help="Each column shows how many of its values the Power BI or Staff "
                    "Portal export recognises as practice areas or file names.",
                )
            )

        try:
            portal_order = build_portal_order(pasted, chosen)
        except ValueError as exc:
            st.error(str(exc))
        else:
            recognised = pasted.recognised(chosen)
            total = pasted.entry_count
            if recognised == total:
                st.success(f"{total} entries, all recognised")
            else:
                st.warning(
                    f"{total} entries, {recognised} recognised. "
                    f"{total - recognised} value(s) are not practice areas or file "
                    "names in either export."
                )
                # A paste that is mostly unrecognised is usually the wrong
                # shape rather than the wrong firm: copying the portal page
                # cell by cell puts each field on its own line, so dates,
                # counts and "DOCX" arrive as if they were practice areas.
                if total and recognised / total < 0.5:
                    st.info(
                        "Most of this is not practice areas, which usually means "
                        "each cell landed on its own line rather than each row. "
                        "Select the rows on the portal page and copy them as a "
                        "block, or paste just the practice area column."
                    )
                unknown = pasted.unrecognised(chosen)
                if unknown:
                    st.caption(
                        "Not recognised: "
                        + ", ".join(repr(u) for u in unknown)
                        + (" ..." if total - recognised > len(unknown) else "")
                    )


with st.spinner("Reading documents..."):
    documents = [_parse(f.name, f.getvalue()) for f in doc_files]
results = build_results(documents, powerbi, portal)
collapse_superseded(results)

# --- summary ----------------------------------------------------------------

# A second copy of a document already read must not become a second row in a
# table that is meant to have one row per firm and practice area.
table_results = [
    r for r in results if r.status not in ("duplicate", "superseded", "elite")
]

differences = difference_records(table_results)
parsed_ok = sum(1 for r in results if r.document.ok)
failed = len(results) - parsed_ok
unparsed = sum(len(r.document.unparsed_metrics) for r in table_results)
unmatched = [r for r in results if r.status in ("unmatched", "ambiguous")]
duplicates = [r for r in results if r.status == "duplicate"]
superseded = [r for r in results if r.status == "superseded"]
elite = [r for r in results if r.status == "elite"]
undecided = [r for r in table_results if r.undecided_supersession]

report = None
if portal_order is not None and portal is not None:
    report = reconcile(results, portal, portal_order)

st.subheader("Summary")
cols = st.columns(5)
cols[0].metric(
    "Files read",
    f"{parsed_ok} / {len(results)}",
    help=f"You added {len(results)} file(s) and {parsed_ok} could be read."
    + (
        f" {failed} could not — see Needs attention."
        if failed
        else " None failed."
    ),
)
cols[1].metric(
    "Duplicates",
    len(duplicates),
    help="Files that are a second copy of a submission already read. Each one "
    "usually means a different document was never downloaded.",
)
cols[2].metric(
    "Missing links",
    len(report.without_document) if report else "—",
    help="Rows on the pasted page with no document on the portal, so nothing "
    "to download. They're listed in the panel below.",
)
cols[3].metric(
    "Corrections",
    len(differences),
    help="Counts where this recount disagrees with the Power BI export. These "
    "are the green cells.",
)
cols[4].metric(
    "Double-check manually",
    unparsed,
    help="Counts that could not be established from the document and need "
    "counting by hand. These are the amber cells.",
)

# --- what the pasted portal page expected ------------------------------------

if report is not None:
    # A collapsible panel: it is long when a page is worked in parts, and it sits
    # between the summary and the table. Nothing inside it may be an expander -
    # Streamlit does not allow those to nest.
    with st.expander("Against the pasted Staff Portal page", expanded=True):
        # Everything here is counted in rows, so the tiles add up: rows to check
        # = covered + missing + not started + check on the page (+ unreadable).
        tiles = [
            ("Rows on the page", report.page_size, "Lines in what you pasted."),
            (
                "Covered",
                f"{len(report.accounted)} / {len(report.expected)}",
                "Rows with a document that one of your files covers. Rows, not "
                "files: see 'Files read' above for those.",
            ),
            ("Missing", len(report.missing), "In the part you've done, but not uploaded."),
            ("Not started", len(report.not_reached), "Further down the page than this batch."),
            (
                "Check on the page",
                len(report.unclear),
                "The Staff Portal export repeats the line above, so the app can't "
                "tell which document these are.",
            ),
        ]
        for column, (label, value, hint) in zip(st.columns(len(tiles)), tiles):
            column.metric(label, value, help=hint)

        def _lines(rows, detail=lambda r: "") -> str:
            return "\n".join(f"- line {r.position} · {r.value}{detail(r)}" for r in rows)

        if report.missing:
            st.error(f"**Not uploaded ({len(report.missing)})**")
            st.markdown(
                _lines(report.missing, lambda r: f" · `{r.candidates[0]}`" if r.candidates else "")
            )

        if report.unclear:
            st.warning(
                f"**Check {len(report.unclear)} line(s) on the portal page.** If "
                "the page shows a different person or file there, download it."
            )
            st.markdown(_lines(report.unclear, lambda r: f" · same Staff Portal row as line {r.repeat_of}"))

        if report.unreadable:
            st.warning(f"**Could not be read ({len(report.unreadable)})**")
            st.markdown(_lines(report.unreadable))

        if report.without_document:
            st.markdown(
                f"**No document on the portal ({len(report.without_document)})**, "
                "nothing to download:"
            )
            shown = report.without_document[:10]
            st.markdown(_lines(shown, lambda r: f" ({r.region})" if r.region else ""))
            if len(report.without_document) > len(shown):
                st.caption(f"... and {len(report.without_document) - len(shown)} more.")

        if report.not_on_page:
            st.info(
                f"Not on the pasted page ({len(report.not_on_page)}): "
                + ", ".join(f"`{r.source_filename or r.filename}`" for r in report.not_on_page[:5])
                + (" ..." if len(report.not_on_page) > 5 else "")
            )

        if report.unrecognised:
            st.caption(
                f"Not found for this firm ({len(report.unrecognised)}): "
                + ", ".join(repr(r.value) for r in report.unrecognised[:5])
                + (" ..." if len(report.unrecognised) > 5 else "")
            )

        if report.accounted_for and report.expected:
            st.success("Nothing outstanding. Every row with a document is uploaded.")
        elif not (report.missing or report.unclear) and report.not_reached:
            st.info(f"Nothing missing so far. {len(report.not_reached)} row(s) still to do.")

elif portal is not None:
    st.caption(
        "Paste the Staff Portal page into the row order box above to check the "
        "batch is complete - which rows have no submission document, and which "
        "documents are missing."
    )

# Two separate lists. A note is background; a problem needs a decision. Mixing
# them makes a reader either alarmed by a note or dismissive of a problem.
problems: list[str] = []
notes: list[str] = []

if failed:
    unreadable = [r for r in results if not r.document.ok]
    problems.append(
        f"- **{failed}** document(s) could not be read — nothing was counted from "
        + ("it" if failed == 1 else "them")
        + ". Open the file to check it is a real submission document.\n"
        + "\n".join(
            f"    - `{r.source_filename or r.filename}`" for r in unreadable
        )
    )
if duplicates:
    problems.append(
        f"- **{len(duplicates)}** file(s) are a second copy of a submission already "
        "read and are left out of the table. A duplicate download usually means "
        "another document was missed.\n"
        + "\n".join(f"    - `{r.source_filename}`" for r in duplicates)
    )
if unmatched:
    problems.append(
        f"- **{len(unmatched)}** document(s) could not be matched to a Power BI "
        "export row:\n"
        + "\n".join(f"    - `{r.source_filename or r.filename}`" for r in unmatched)
    )
numbering = [
    r for r in table_results
    if any(m.numbering_gap or m.needs_check for m in r.document.metrics.values())
]
if numbering:
    def second_look_reason(row) -> str:
        """A few words, not the whole note.

        The note itself is shown once, on the document's own entry in the parse
        log, next to the figure and the evidence it refers to. Repeating it here
        buried the list of affected files in several lines of explanation each.
        """
        reasons = []
        if any(m.numbering_gap for m in row.document.metrics.values()):
            reasons.append("matter numbering")
        if any(m.needs_check for m in row.document.metrics.values()):
            reasons.append("nomination description")
        return ", ".join(reasons) or "see the parse log"

    problems.append(
        f"- **{len(numbering)}** document(s) have a figure that was counted but is "
        "worth a second look (amber). The reason is in each one's parse log "
        "below:\n"
        + "\n".join(
            f"    - `{r.source_filename or r.filename}` — {second_look_reason(r)}"
            for r in numbering[:10]
        )
        + (f"\n    - ... and {len(numbering) - 10} more" if len(numbering) > 10 else "")
    )
if undecided:
    problems.append(
        f"**{len(undecided)}** practice area(s) have more than one submission, but "
        "both have the same date, so the newer one could not be identified. Check "
        "by hand: "
        + ", ".join(f"{r.practice_area} ({r.filename})" for r in undecided)
    )
if unparsed:
    problems.append(
        f"**{unparsed}** figure(s) could not be counted. Those cells are amber "
        "and show the Power BI figure as a starting point. Type your own count "
        "over it in the table."
    )
blank_area = [r for r in table_results if not r.practice_area]
if blank_area:
    problems.append(
        f"**{len(blank_area)}** row(s) have no practice area. The reference columns "
        "come from the Staff Portal export, so this means the filename was not found "
        "in it: "
        + ", ".join(r.filename for r in blank_area[:5])
        + (" ..." if len(blank_area) > 5 else "")
    )

if portal_order is not None:
    notes.extend(portal_order.notes)
    missing_from_order = unordered_rows(table_results, portal_order)
    if missing_from_order and len(missing_from_order) == len(table_results):
        problems.append(
            "**None** of the rows could be placed using the pasted Staff Portal "
            f"order ({portal_order.column_label}). Is that the column holding the "
            "practice areas?"
        )
    elif missing_from_order:
        problems.append(
            f"**{len(missing_from_order)}** row(s) are not in the pasted Staff Portal "
            "order and are listed last: "
            + ", ".join(r.practice_area or r.filename for r in missing_from_order[:5])
            + (" ..." if len(missing_from_order) > 5 else "")
        )

if superseded:
    notes.append(
        f"- **{len(superseded)}** older submission(s) were left out because a newer "
        "one exists for the same firm, region and practice area:\n"
        + "\n".join(
            f"    - `{r.source_filename}` — {created_at_text(r.portal_record)}"
            + (f" — {r.practice_area}" if r.practice_area else "")
            for r in superseded
        )
    )
if elite:
    notes.append(
        f"- **{len(elite)}** Elite submission(s) left out - Elites are not recounted:\n"
        + "\n".join(f"    - `{r.source_filename}`" for r in elite)
    )
notes.extend(powerbi.notes)
if portal is not None:
    notes.extend(portal.notes)
else:
    notes.append(
        "No Staff Portal export was added. File names fall back to the name of the "
        "file you dragged in, and the firm, region and practice area come from each "
        "document's own fields - a document whose practice-area dropdown was never "
        "used will leave those cells blank. Adding it also lets the app check the "
        "batch for missing and duplicate documents."
    )

if problems:
    st.markdown("#### Needs attention")
    for problem in problems:
        st.markdown(problem if problem.lstrip().startswith("-") else f"- {problem}")
elif not failed:
    st.success("Nothing needs attention.")

if notes:
    with st.expander(f"Notes about the files read ({len(notes)})", expanded=False):
        for note in notes:
            st.markdown(note if note.lstrip().startswith("-") else f"- {note}")

# --- results table ----------------------------------------------------------

st.subheader("Results")

sort_options = [
    label
    for label in SORT_MODES
    if SORT_MODES[label] != "portal_paste" or portal_order is not None
]

# Each status gets its own symbol. "—" was previously used for "no document",
# which read as punctuation when the key was written on one line.
MARKERS = {
    "uploaded": ("✅", ""),
    "missing": ("❌", "**"),
    "older version": ("🕓", ""),
    "could not read": ("⚠️", "**"),
    "check page": ("🔍", "**"),
    "no document": ("⛓️‍💥", ""),
    "not reached": ("⏳", ""),
    "unrecognised": ("❔", ""),
    "elite": ("🚫", ""),
}

# The columns are opened before the controls so the pasted-order panel starts
# level with them, rather than beginning where the table does and leaving a band
# of white space above it.
if report is not None and report.page_size:
    table_col, page_col = st.columns([7, 3])
else:
    table_col, page_col = st.container(), None

with table_col:
    order_col, toggle_col = st.columns([2, 3])

    with order_col:
        sort_label = st.selectbox(
            "Row order",
            sort_options,
            index=0,
            help="Paste the Staff Portal rows above to sort the table exactly as "
            "the portal shows it. Otherwise Power BI export order mirrors that file "
            "row for row, so ordering that file once orders this table too.",
        )

    with toggle_col:
        show_original = st.toggle(
            "Show the Power BI value alongside corrections",
            value=False,
            help="Off: corrected numbers only. On: shows '22 (was 20)' in every "
            "changed cell.",
        )

    results = sort_results(results, SORT_MODES[sort_label], portal_order)
    table_results = [
        r for r in results if r.status not in ("duplicate", "superseded", "elite")
    ]

    # Where the row count comes from. Without this the table's length looks
    # unrelated to the number of files added.
    removed = [
        f"**{len(superseded)}** duplicate version(s) removed to keep the newer version"
        if superseded
        else "",
        f"**{len(duplicates)}** second copy/copies of the same file removed"
        if duplicates
        else "",
    ]
    removed = [part for part in removed if part]
    st.caption(
        f"**{len(table_results)}** row(s) in the table: **{len(results)}** "
        "submission document file(s) uploaded here"
        + (", " + ", ".join(removed) if removed else "")
        + ". One row per firm, region and practice area."
    )

    with st.popover("Show the maths", icon=":material/calculate:"):
        st.markdown(_arithmetic(results, table_results, duplicates, superseded, report))
    st.caption(
        "**Green**: differs from the Power BI export. **Amber**: needs a look, "
        "either not counted (the cell shows the Power BI figure) or counted but "
        "unusual. Type over a figure and the colours follow it."
    )

    # Hand corrections and double-checked marks live in session state, keyed by
    # file name and metric, so they survive a rerun and follow a row when the
    # sort order changes.
    overrides = st.session_state.setdefault("corrections", {})
    checked = st.session_state.setdefault("checked_rows", set())

    coloured = st.checkbox(
        "Colour the cells (read-only)",
        value=False,
        help="Streamlit can colour a table or let you edit it, not both. Tick "
        "this to see the highlighting; untick it to make changes. The "
        "downloaded file always carries the colours.",
    )

    frame = results_to_frame(
        table_results,
        show_original=show_original,
        overrides=overrides,
        numeric=not coloured,
        checked=checked,
    )
    frame[REVIEW_COLUMN] = [review_note(r, overrides, checked) for r in table_results]
    frame[CHECKED_COLUMN] = [r.filename.lower() in checked for r in table_results]

    # Tall enough to show a real batch without endless scrolling, capped so the
    # page stays navigable.
    table_height = min(38 + 35 * max(len(frame), 3), 640)

    if coloured:
        def _styles(_):
            import pandas as pd

            styled = pd.DataFrame("", index=frame.index, columns=frame.columns)
            for i, row in enumerate(table_results):
                for key in METRIC_KEYS:
                    position = frame.columns.get_loc(METRIC_COLUMNS[key])
                    value, unverified = display_value(row, key, overrides, checked)
                    if unverified or needs_attention(row, key, checked):
                        fill = AMBER
                    elif final_differs(row, key, overrides):
                        fill = GREEN
                    else:
                        continue
                    styled.iloc[i, position] = (
                        f"background-color: {fill}; color: {ON_FILL}"
                    )
                if frame.iloc[i][REVIEW_COLUMN]:
                    styled.iloc[i, frame.columns.get_loc(REVIEW_COLUMN)] = (
                        f"background-color: {AMBER}; color: {ON_FILL}"
                    )
            return styled

        st.dataframe(
            frame.style.apply(_styles, axis=None),
            width="stretch",
            hide_index=True,
            height=table_height,
        )
        st.caption(
            "Read-only while the colours are on. Untick **Colour the cells** to "
            "change a figure or tick a row."
        )
    else:
        editable = [METRIC_COLUMNS[k] for k in METRIC_KEYS] + [CHECKED_COLUMN]
        edited = st.data_editor(
            frame,
            width="stretch",
            hide_index=True,
            height=table_height,
            key=f"results_table_{batch}",
            disabled=[c for c in frame.columns if c not in editable],
            column_config={
                **{
                    METRIC_COLUMNS[k]: st.column_config.NumberColumn(
                        METRIC_COLUMNS[k],
                        help="Type over a figure to overrule it. Leave it alone "
                        "and the app's own count stands.",
                        min_value=0,
                        step=1,
                    )
                    for k in METRIC_KEYS
                },
                REVIEW_COLUMN: st.column_config.TextColumn(
                    REVIEW_COLUMN,
                    help="What on this row wants checking, and why.",
                    width="medium",
                ),
                CHECKED_COLUMN: st.column_config.CheckboxColumn(
                    CHECKED_COLUMN,
                    help="Tick a row once you have checked it. Ticked rows are "
                    "marked in the exported file.",
                    default=False,
                ),
            },
        )

        # A cell counts as edited only when it differs from what was rendered,
        # so an amber cell left alone keeps its "nobody has checked this" state
        # rather than being taken as accepted.
        import pandas as pd

        changed = False
        for i, row in enumerate(table_results):
            for key in METRIC_KEYS:
                column = METRIC_COLUMNS[key]
                shown, typed = frame.iloc[i][column], edited.iloc[i][column]
                if pd.isna(shown) and pd.isna(typed):
                    continue
                if not pd.isna(shown) and not pd.isna(typed) and int(shown) == int(typed):
                    continue
                value = None if pd.isna(typed) else int(typed)
                if value == row.recount(key):
                    changed |= overrides.pop(override_key(row, key), "-") != "-"
                else:
                    changed |= overrides.get(override_key(row, key), "-") != value
                    overrides[override_key(row, key)] = value

        ticked = {
            row.filename.lower()
            for row, mark in zip(table_results, edited[CHECKED_COLUMN])
            if bool(mark)
        }
        if ticked != checked:
            st.session_state["checked_rows"] = ticked
            checked = ticked
            changed = True
        if changed:
            st.rerun()

    marks = []
    if overrides:
        marks.append(f"**{len(overrides)}** figure(s) entered by hand")
    marks.append(f"**{len(checked)}** of {len(table_results)} row(s) ticked")
    st.caption(
        ", ".join(marks)
        + ". Both are written to the exported file, and the highlighting there "
        "follows whatever is in the table."
    )

    st.download_button(
        "Download .xlsx with highlighting",
        data=build_workbook(table_results, overrides, checked),
        file_name="submission-recount.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="The fills are real cell formatting, so they survive being pasted "
        "into the master sheet.",
    )
    if overrides and st.button("Undo all hand corrections"):
        st.session_state["corrections"] = {}
        st.rerun()

# The pasted page as a checklist beside the table. With the row order set to
# the pasted page (the default once one is pasted) the two run in the same
# order, so they can be read side by side.
if page_col is not None:
    with page_col:
        st.markdown("**Pasted order from Staff Portal**")

        in_table = {id(r) for r in table_results}
        only_in_table = st.checkbox(
            "Show only the rows in the table",
            value=False,
            help="Leaves only the rows that became a table row, so the two lists "
            "match one for one when the row order is Staff Portal order (pasted).",
        )
        shown_rows = (
            [r for r in report.rows if r.uploaded is not None and id(r.uploaded) in in_table]
            if only_in_table
            else report.rows
        )

        st.caption(f"{len(shown_rows)} of {report.page_size} pasted row(s), in your order.")

        with st.container(height=max(table_height - 24, 200)):
            lines = []
            for page_row in shown_rows:
                marker, emphasis = MARKERS[page_row.status]
                label = page_row.value
                if len(label) > 40:
                    label = label[:39] + "…"
                lines.append(
                    f"`{page_row.position:>3}` {marker} {emphasis}{label}{emphasis}"
                )
            st.markdown("  \n".join(lines))

        with st.expander("Legend", expanded=False):
            st.markdown(
                "- ✅ **uploaded** — in your batch and in the table\n"
                "- ❌ **missing** — the portal has a document, but it wasn't uploaded\n"
                "- 🔍 **check on the page** — the Staff Portal export repeats the line above, so "
                "look at the page: a different person or file there means another "
                "document to download\n"
                "- ⚠️ **could not be read** — uploaded, but the file wouldn't open\n"
                "- 🕓 **older version** — a newer submission was counted instead\n"
                "- ⛓️‍💥 **no document link** — nothing on the portal, so nothing "
                "to download or check\n"
                "- ⏳ **not started** — further down than this batch reaches\n"
                "- 🚫 **Elite** — not recounted, nothing to do\n"
                "- ❔ **not found** — no portal record under this firm; check the "
                "spelling of what you pasted"
            )

# --- per-document log -------------------------------------------------------

st.subheader(
    "Per-document parse log",
    help=(
        "One entry per file you added, showing where each number was counted from.\n\n"
        "**The mark on each entry:**\n\n"
        "- ✅ **counted** — matched to a Power BI export row and included in the table\n"
        "- 🕓 **older version, not counted** — two files here are for the same "
        "firm, region and practice area. The newer one was counted (it is the ✅ "
        "entry); this older one was left out of the table. It is listed only so "
        "every file you added is accounted for.\n"
        "- ♻️ **duplicate** — the same document was already read from another file, "
        "usually a second download. Counted once.\n"
        "- 🚫 **Elite, not counted** — Elites are not recounted\n"
        "- ❓ **not matched** — could not be tied to a row in the Power BI export\n"
        "- ⚠️ **could not be read** — the file could not be opened or is not a "
        "submission document"
    ),
)
st.caption("Where each number came from, so you can check it against the document.")

for row in results:
    icon = {
        "matched": "✅",
        "unmatched": "❓",
        "ambiguous": "❓",
        "duplicate": "♻️",
        "superseded": "🕓",
        "parse_error": "⚠️",
        "elite": "🚫",
    }[row.status]
    headline = f"{icon} {row.source_filename or row.filename}"
    if row.practice_area:
        headline += f" — {row.practice_area}"
    if row.status in ("duplicate", "superseded", "elite"):
        # Its counts never reach the table, so a correction count would mislead.
        headline += " — not counted"
    else:
        changed = sum(1 for k in METRIC_KEYS if row.differs(k))
        if changed:
            headline += f" — {changed} correction(s)"

    with st.expander(headline, expanded=False):
        if not row.document.ok:
            st.error(row.document.error)

        detail = st.columns(3)
        detail[0].markdown(f"**Firm**  \n{row.firm or '—'}  \n`ref {row.firm_ref or '—'}`")
        detail[1].markdown(f"**Region**  \n{row.region or '—'}")
        detail[2].markdown(f"**Practice area**  \n{row.practice_area or '—'}")
        if row.reference_source:
            st.caption(
                f"Firm, region and practice area taken from the {row.reference_source}."
            )
        if row.document.dropdown_raw:
            st.caption(f"The document's own dropdown reads: {row.document.dropdown_raw}")
        else:
            st.caption("The document's practice-area dropdown was left unselected.")

        for key in METRIC_KEYS:
            metric = row.document.metrics.get(key)
            if metric is None:
                continue
            recount, export_value = row.recount(key), row.export_value(key)

            if metric.is_unparsed:
                verdict = f"❓ **{UNPARSED}**"
            elif row.differs(key):
                verdict = f"🟩 **{recount}**: Power BI said {export_value}"
            elif export_value is None:
                verdict = f"**{recount}**: no Power BI value to compare"
            else:
                verdict = f"**{recount}**: matches Power BI"
            if not metric.is_unparsed and (metric.numbering_gap or metric.needs_check):
                # Amber in the table, so not a confident green here: the note
                # below says why the figure itself may be wrong.
                verdict = verdict.replace("🟩 ", "") + " 🟨 **worth a look, see the note below**"

            st.markdown(f"**{METRIC_LABELS[key]}**: {verdict}")
            if metric.evidence:
                with st.container():
                    st.caption("Counted from:")
                    st.code("\n".join(metric.evidence), language=None)
            for note in metric.notes:
                st.caption(f"⚠️ {note}")

        if row.issues:
            st.markdown("**Issues**")
            for issue in row.issues:
                st.markdown(f"- {issue}")
