"""One row per firm, region and practice area: the newest submission wins.

A firm often has more than one document for the same slot - an earlier
submission and a replacement. The Staff Portal shows one row, so the output
table must too.
"""

import pytest

from conftest import MANIFEST, OLDER_VERSION, PORTAL, POWERBI, SAMPLES
from matching import (
    build_results,
    collapse_superseded,
    created_at_text,
    load_portal,
    load_powerbi,
)
from parsing import parse_document

NEWER = SAMPLES["alpha"].name
OLDER = OLDER_VERSION.name


@pytest.fixture(scope="module")
def exports():
    return load_powerbi(POWERBI), load_portal(PORTAL)


def document_named(name, key="alpha"):
    doc = parse_document(SAMPLES[key])
    doc.filename = name
    return doc


def build(names, exports):
    """Build rows and the set the table would show, as the app does."""
    powerbi, portal = exports
    rows = build_results([document_named(n) for n in names], powerbi, portal)
    collapse_superseded(rows)
    kept = [r for r in rows if r.status not in ("duplicate", "superseded")]
    return rows, kept


# --- reading the dates ------------------------------------------------------


def test_created_at_is_read_from_the_portal_export(exports):
    _, portal = exports
    assert created_at_text(portal.lookup(OLDER)) < created_at_text(portal.lookup(NEWER))


def test_a_missing_date_sorts_oldest():
    from datetime import datetime

    from matching import _created_at

    assert _created_at({"created_at": None}) == datetime.min
    assert _created_at({}) == datetime.min
    assert _created_at(None) == datetime.min
    assert _created_at({"created_at": "not a date"}) == datetime.min


def test_an_iso_string_date_is_understood():
    from datetime import datetime

    from matching import _created_at

    assert _created_at({"created_at": "2026-03-30 00:00:00"}) == datetime(2026, 3, 30)


# --- the rule ---------------------------------------------------------------


def test_the_newer_submission_is_kept(exports):
    rows, kept = build([OLDER, NEWER], exports)
    assert [r.status for r in rows] == ["superseded", "matched"]
    assert [r.filename for r in kept] == [NEWER]


def test_order_of_upload_does_not_matter(exports):
    _, kept = build([NEWER, OLDER], exports)
    assert [r.filename for r in kept] == [NEWER]


def test_the_dropped_row_explains_itself(exports):
    rows, _ = build([OLDER, NEWER], exports)
    older = next(r for r in rows if r.source_filename == OLDER)
    issue = " ".join(older.issues)
    assert NEWER in issue
    assert "newer submission is counted" in issue


def test_a_single_submission_is_untouched(exports):
    rows, kept = build([NEWER], exports)
    assert [r.status for r in rows] == ["matched"]
    assert len(kept) == 1
    assert not any("newer submission is counted" in i for i in rows[0].issues)


def test_different_practice_areas_are_not_collapsed(exports):
    powerbi, portal = exports
    rows = build_results([parse_document(p) for p in SAMPLES.values()], powerbi, portal)
    kept = collapse_superseded(rows)
    assert len(kept) == len(SAMPLES)
    assert all(r.status == "matched" for r in rows)


# --- the cases where the date cannot decide ---------------------------------


def test_identical_filenames_are_a_true_duplicate(exports):
    powerbi, portal = exports
    duplicate = NEWER.replace(".docx", " (1).docx")
    rows = build_results(
        [document_named(NEWER), document_named(duplicate)], powerbi, portal
    )
    collapse_superseded(rows)
    kept = [r for r in rows if r.status not in ("duplicate", "superseded")]
    assert [r.status for r in rows] == ["matched", "duplicate"]
    assert len(kept) == 1


def test_the_file_with_the_suffix_is_the_copy_whatever_the_upload_order(exports):
    powerbi, portal = exports
    copy = NEWER.replace(".docx", " (1).docx")
    for order in ([NEWER, copy], [copy, NEWER]):
        rows = build_results([document_named(n) for n in order], powerbi, portal)
        by_name = {r.source_filename: r.status for r in rows}
        assert by_name[NEWER] == "matched", f"order {order}"
        assert by_name[copy] == "duplicate", f"order {order}"


def test_a_tied_date_is_flagged_rather_than_guessed(exports):
    powerbi, portal = exports
    rows = build_results(
        [document_named(OLDER), document_named(NEWER)], powerbi, portal
    )
    for row in rows:
        row.portal_record = dict(row.portal_record, created_at=None)
    kept = collapse_superseded(rows)

    assert len(kept) == 1
    assert kept[0].undecided_supersession is True
    assert any("cannot be told" in i for i in kept[0].issues)


def test_a_decided_supersession_is_not_flagged_as_undecided(exports):
    _, kept = build([OLDER, NEWER], exports)
    assert kept[0].undecided_supersession is False


def test_unmatched_rows_are_never_collapsed(exports):
    powerbi, portal = exports
    documents = []
    for name in ("not-in-the-portal-a.docx", "not-in-the-portal-b.docx"):
        doc = document_named(name)
        doc.region = doc.practice_group = doc.practice_area = None
        documents.append(doc)

    rows = build_results(documents, powerbi, portal)
    assert all(r.status == "unmatched" for r in rows)
    assert len(collapse_superseded(rows)) == 2


def test_the_manifest_describes_this_pair():
    assert MANIFEST["superseded"]["older"] == OLDER
    assert MANIFEST["superseded"]["newer"] == NEWER
