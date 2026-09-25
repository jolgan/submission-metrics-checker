"""Checking an uploaded batch against the portal page it came from.

Covers the ways a batch goes wrong in practice: a portal row with no submission
document to download, a document downloaded twice (which means some other
document was never downloaded), and a page worked in parts.
"""

import pytest

from conftest import DETAILS, MANIFEST, OLDER_VERSION, PORTAL, POWERBI, SAMPLES
from matching import (
    build_portal_order,
    build_results,
    build_vocabulary,
    collapse_superseded,
    load_portal,
    load_powerbi,
    read_pasted_order,
    reconcile,
    strip_download_suffix,
)
from parsing import parse_document

AREAS = {k: DETAILS[k]["practice_area"].split(" > ")[1] for k in SAMPLES}
NO_DOCUMENT = MANIFEST["portal_rows_without_document"][0]


@pytest.fixture(scope="module")
def exports():
    return load_powerbi(POWERBI), load_portal(PORTAL)


@pytest.fixture(scope="module")
def results(exports):
    powerbi, portal = exports
    return build_results([parse_document(p) for p in SAMPLES.values()], powerbi, portal)


def order_for(text, exports):
    powerbi, portal = exports
    pasted = read_pasted_order(text, build_vocabulary(powerbi, portal))
    return build_portal_order(pasted, pasted.suggested)


def page(*lines):
    return "\n".join(lines)


# --- portal rows with no submission document -------------------------------


def test_portal_rows_without_a_document_are_kept(exports):
    """They used to be discarded, which is why the app could not see them."""
    _, portal = exports
    assert portal.without_document >= 1
    assert all(not r["has_document"] for r in portal.records if not r["filename"])


def test_a_row_with_no_document_is_not_reported_as_missing(results, exports):
    _, portal = exports
    text = page(*AREAS.values(), NO_DOCUMENT)
    report = reconcile(results, portal, order_for(text, exports))

    assert report.page_size == 4
    assert len(report.expected) == 3
    assert len(report.without_document) == 1
    assert report.missing == [], "rows with no document must not count as missing"


def test_rows_without_a_document_are_named(results, exports):
    _, portal = exports
    report = reconcile(results, portal, order_for(page(NO_DOCUMENT), exports))
    assert NO_DOCUMENT.split(" > ")[1] in report.without_document[0].value


def test_a_complete_page_is_accounted_for(results, exports):
    _, portal = exports
    report = reconcile(results, portal, order_for(page(*AREAS.values()), exports))
    assert report.accounted_for
    assert len(report.expected) == 3
    assert report.missing == []


# --- documents that should be there but are not ----------------------------


def test_a_gap_inside_the_span_is_missing(exports):
    """Two of three uploaded, with the gap above the last one."""
    powerbi, portal = exports
    documents = [parse_document(SAMPLES[k]) for k in ("alpha", "beta")]
    results = build_results(documents, powerbi, portal)
    text = page(AREAS["alpha"], AREAS["gamma"], AREAS["beta"])
    report = reconcile(results, portal, order_for(text, exports))

    assert report.attempted == 2
    assert [r.value for r in report.missing] == [AREAS["gamma"]]
    assert not report.accounted_for


def test_rows_below_the_batch_are_not_reported_missing(exports):
    """One upload against a longer page: only one row is in scope."""
    powerbi, portal = exports
    results = build_results([parse_document(SAMPLES["alpha"])], powerbi, portal)
    report = reconcile(results, portal, order_for(page(*AREAS.values()), exports))

    assert report.attempted == 1
    assert report.missing == [], "rows below the batch are not missing"
    assert len(report.not_reached) == 2


def test_the_span_is_the_number_of_files_and_nothing_more(exports):
    """A stray match far down the page must not make everything above missing."""
    powerbi, portal = exports
    results = build_results([parse_document(SAMPLES["beta"])], powerbi, portal)
    text = page(AREAS["alpha"], AREAS["gamma"], AREAS["beta"])
    report = reconcile(results, portal, order_for(text, exports))

    assert report.attempted == 1
    assert [r.value for r in report.missing] == [AREAS["alpha"]]
    assert len(report.not_reached) == 1


def test_uploads_not_on_the_pasted_page_are_reported(results, exports):
    _, portal = exports
    report = reconcile(results, portal, order_for(page(AREAS["alpha"]), exports))
    assert len(report.not_on_page) == 2


def test_pasted_rows_with_no_portal_record_are_reported(results, exports):
    _, portal = exports
    text = page(*AREAS.values(), "Not A Real Practice Area")
    report = reconcile(results, portal, order_for(text, exports))
    assert [r.value for r in report.unrecognised] == ["Not A Real Practice Area"]


def test_position_in_the_order_is_recoverable_for_a_missing_row(exports):
    powerbi, portal = exports
    results = build_results([parse_document(SAMPLES[k]) for k in ("alpha", "beta")],
                            powerbi, portal)
    text = page(AREAS["alpha"], AREAS["gamma"], AREAS["beta"])
    report = reconcile(results, portal, order_for(text, exports))
    assert report.missing[0].position == 2


# --- duplicate downloads ---------------------------------------------------


def test_strip_download_suffix():
    assert strip_download_suffix("A Firm_submission (1).docx") == (
        "A Firm_submission.docx", 1,
    )
    assert strip_download_suffix("A Firm_submission (12).docx") == (
        "A Firm_submission.docx", 12,
    )
    assert strip_download_suffix("A Firm_submission.docx") == (
        "A Firm_submission.docx", None,
    )


def test_a_bracketed_name_that_is_not_a_copy_marker_is_left_alone():
    name = "Firm_submission_Intellectual property (patents).docx"
    assert strip_download_suffix(name) == (name, None)


def test_a_copy_suffixed_file_still_matches_its_portal_row(exports):
    _, portal = exports
    name = SAMPLES["alpha"].name.replace(".docx", " (1).docx")
    record = portal.lookup(name)
    assert record is not None
    assert record["practice_area"] == DETAILS["alpha"]["practice_area"].split(" > ")[1]


def test_a_copy_suffixed_upload_is_matched_and_flagged(exports):
    powerbi, portal = exports
    doc = parse_document(SAMPLES["alpha"])
    doc.filename = SAMPLES["alpha"].name.replace(".docx", " (1).docx")
    rows = build_results([doc], powerbi, portal)

    assert rows[0].status == "matched", "the suffix must not break the match"
    assert any("copy 1" in i for i in rows[0].issues)


def test_a_duplicate_upload_points_at_the_document_that_was_missed(exports):
    powerbi, portal = exports
    documents = [parse_document(SAMPLES[k]) for k in ("alpha", "beta")]
    extra = parse_document(SAMPLES["alpha"])
    extra.filename = SAMPLES["alpha"].name.replace(".docx", " (1).docx")
    documents.append(extra)

    results = build_results(documents, powerbi, portal)
    text = page(AREAS["alpha"], AREAS["beta"], AREAS["gamma"])
    report = reconcile(results, portal, order_for(text, exports))

    assert len(report.duplicates) == 1
    assert report.attempted == 3
    assert [r.value for r in report.missing] == [AREAS["gamma"]]


# --- one document listed against several rows ------------------------------


def test_shared_documents_are_counted_once(exports):
    """The portal lists beta's practice area twice against the same file."""
    powerbi, portal = exports
    shared = MANIFEST["shared_document"]
    results = build_results([parse_document(p) for p in SAMPLES.values()],
                            powerbi, portal)
    area = shared["practice_area"].split(" > ")[1]
    report = reconcile(results, portal, order_for(page(area, area), exports))

    assert len(report.expected) == 2
    assert report.shared_rows == 1
    assert report.distinct_documents == 1


def test_no_sharing_means_documents_equal_rows(results, exports):
    _, portal = exports
    report = reconcile(results, portal, order_for(page(AREAS["alpha"]), exports))
    assert report.shared_rows == 0
    assert report.distinct_documents == len(report.expected)


# --- the figures hold together ---------------------------------------------


def test_page_rows_carry_their_line_number(results, exports):
    _, portal = exports
    report = reconcile(results, portal, order_for(page(*AREAS.values()), exports))
    assert [r.position for r in report.rows] == [1, 2, 3]


def test_page_rows_carry_a_status_for_the_checklist(exports):
    powerbi, portal = exports
    results = build_results([parse_document(SAMPLES["alpha"])], powerbi, portal)
    text = page(AREAS["alpha"], NO_DOCUMENT, AREAS["beta"], "Nonsense")
    report = reconcile(results, portal, order_for(text, exports))

    statuses = [r.status for r in report.rows]
    assert statuses.count("uploaded") == 1
    assert statuses.count("no document") == 1
    assert statuses.count("unrecognised") == 1
    assert statuses.count("not reached") == 1
    assert len(statuses) == report.page_size


def test_an_unreadable_file_does_not_tick_its_row(exports, tmp_path):
    powerbi, portal = exports
    bad = tmp_path / SAMPLES["alpha"].name
    bad.write_bytes(b"not a docx")
    results = build_results([parse_document(bad)], powerbi, portal)
    report = reconcile(results, portal, order_for(page(AREAS["alpha"]), exports))

    assert len(report.unreadable) == 1
    assert report.accounted == []
    assert not report.accounted_for


def test_the_figures_add_up(exports):
    powerbi, portal = exports
    results = build_results([parse_document(SAMPLES["alpha"])], powerbi, portal)
    text = page(*AREAS.values(), NO_DOCUMENT)
    report = reconcile(results, portal, order_for(text, exports))

    assert report.page_size == (
        len(report.without_document) + len(report.expected) + len(report.unrecognised)
    )
    assert len(report.expected) == (
        len(report.accounted) + len(report.missing)
        + len(report.not_reached) + len(report.unreadable)
    )


def test_an_older_version_is_marked_not_ticked(exports):
    powerbi, portal = exports
    documents = [parse_document(SAMPLES["alpha"]), parse_document(OLDER_VERSION)]
    documents[1].filename = OLDER_VERSION.name
    results = build_results(documents, powerbi, portal)
    collapse_superseded(results)
    report = reconcile(results, portal, order_for(page(AREAS["alpha"]), exports))

    statuses = {r.status for r in report.rows}
    assert statuses <= {"uploaded", "older version"}


# --- practice areas the portal names with three levels ----------------------


def test_a_three_level_paste_finds_its_row(exports):
    """The portal screen shows 'Group > Area > Table' for some submissions.

    The export row may print two of those levels and the paste three, so the
    spellings have to be reconciled rather than compared literally.
    """
    _, portal = exports
    three = MANIFEST["three_level_row"]
    records = portal.records_for_area(three)
    assert records, f"{three!r} should find a portal row"
    assert len(records) == 1


def test_the_same_row_is_found_by_its_table_name_alone(exports):
    _, portal = exports
    table = MANIFEST["three_level_row"].split(" > ")[-1]
    assert portal.records_for_area(table)


def test_a_three_level_row_keeps_all_three_levels_in_the_table(exports):
    """Two submissions can share a group and area, differing only by table."""
    from matching import _portal_practice_area

    _, portal = exports
    three = MANIFEST["three_level_row"]
    record = portal.records_for_area(three)[0]
    assert _portal_practice_area(record) == three


def test_a_punctuated_table_name_still_finds_its_portal_record(exports):
    """The bare area is what the Staff Portal page prints, so it must resolve.

    The portal row names the same practice area twice, once bracketed and once
    not. Read literally that is a third level, and the record is then indexed
    only under spellings nobody pastes.
    """
    _, portal = exports
    area = MANIFEST["punctuated_table_row"]["practice_area"]
    records = portal.records_for_area(area)
    assert [r for r in records if r.get("firm") == MANIFEST["firms"]["alpha"]], area


def test_either_spelling_of_the_punctuated_area_resolves(exports):
    _, portal = exports
    row = MANIFEST["punctuated_table_row"]
    for spelling in (row["practice_area"], row["table_name"]):
        assert portal.records_for_area(spelling), spelling


def test_a_table_name_built_from_the_group_and_area_still_matches(exports):
    """The portal's table name is sometimes the group and area joined.

    "City focus - Brasilia - Government relations" for the area "Government
    relations" repeats nothing, so the name stays three levels deep and the
    bare area - the only thing the Staff Portal page prints - was indexed under
    no form at all, leaving the document reported as not on the pasted page.
    """
    _, portal = exports
    row = MANIFEST["derived_table_row"]
    records = portal.records_for_area(row["practice_area"])
    assert [r for r in records if r.get("firm") == MANIFEST["firms"]["alpha"]]


def test_the_full_three_level_name_still_matches(exports):
    """Widening the bare form must not cost the specific one."""
    _, portal = exports
    row = MANIFEST["derived_table_row"]
    full = f"{row['practice_group']} > {row['practice_area']} > {row['table_name']}"
    assert portal.records_for_area(full)


def test_a_bare_area_places_against_the_pasted_order():
    from matching import _order_keys, build_portal_order, read_pasted_order

    order = build_portal_order(read_pasted_order("Government relations\nTax"))
    area = (
        "City focus - Brasilia > Government relations "
        "> City focus - Brasilia - Government relations"
    )
    keys = _order_keys(None, "Example Advogados", "Brazil", None, area)
    assert any(k in order.index for k in keys)
