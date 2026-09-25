"""Lawyer submissions the Staff Portal export lists wrongly.

A firm can nominate two lawyers for one table, each with their own document.
The LatAm export then lists the first lawyer twice - same portal ID, same file -
and never names the second. One firm's batch read "13 separate documents,
nothing left to download" against a page that held 15: the second lawyer's
file in each of two tables matched nothing, the first lawyer's file was never
downloaded, and the counts happened to agree.

Two things answer that. A file the export does not name is matched by firm
reference and table code (the EC number, which names the table, not the
document). And a page line the export merely repeats is reported as one to
check on the page, not as a second listing of the same document.

Elite tables, where this was first seen, are now left out altogether (see
test_elite.py). The tables here are ordinary lawyer tables, so the repeat
handling stays covered for any other table the export repeats.
"""

import csv

import openpyxl
import pytest

from conftest import make_submission
from matching import (
    build_portal_order,
    build_results,
    build_vocabulary,
    collapse_superseded,
    load_portal,
    load_powerbi,
    read_pasted_order,
    reconcile,
)
from parsing import parse_document

FIRM, REF = "Example Advogados", "77001"
FIRM_FILE = f"Example_Advogados_{REF}_submission_2027 Edition_EC100.docx"
ANN = f"Example_Advogados_Ann_Lee_{REF}_submission_2027 Edition_EC900.docx"
BOB = f"Example_Advogados_Bob_Ray_{REF}_submission_2027 Edition_EC900.docx"
BASE = "https://example.invalid/docs/"
HEADER = ["ID", "Firm", "Firm Ref", "Region", "Practice Group Name",
          "Practice Area Name", "Table Name", "Created At", "Submission Type",
          "Person Name", "Document", "File URI"]


def firm_row(id_, area, filename):
    return [id_, FIRM, REF, "Brazil", "", area, area, "2026-02-13 00:00:00",
            "firm", "", "Y", BASE + filename]


def lawyer_row(id_, person, filename):
    return [id_, FIRM, REF, "Brazil", "City Lawyers", "Tax", "Tax",
            "2026-02-27 00:00:00", "lawyer", person, "Y", BASE + filename]


def write_portal(path, rows):
    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADER)
        writer.writerows(rows)
    return load_portal(path)


# The shape the export really has: Ann listed twice under one ID, Bob absent.
REPEATED = [
    firm_row("id-1", "Insurance", FIRM_FILE),
    lawyer_row("id-2", "Ann Lee", ANN),
    lawyer_row("id-2", "Ann Lee", ANN),
]


@pytest.fixture
def powerbi(tmp_path):
    path = tmp_path / "powerbi.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["FIRM_NAME", "FIRM_REF", "COUNTRY_NAME",
                  "CONCATENATED_PRACTICE_AREA_NAME", "NUM_MATTERS",
                  "NUM_ACTIVE_CLIENTS", "NUM_NEW_CLIENTS", "NUM_LEAD_PARTNERS",
                  "NUM_NEXT_GEN", "NUM_ASSOCIATES"])
    sheet.append([FIRM, REF, "Brazil", "Insurance", 1, 1, 1, 0, 0, 0])
    sheet.append([FIRM, REF, "Brazil", "City Lawyers > Tax > Tax", 1, 1, 1, 0, 0, 0])
    workbook.save(path)
    return load_powerbi(path)


def documents(tmp_path, *names):
    return [parse_document(make_submission(tmp_path / name)) for name in names]


def run(tmp_path, powerbi, portal, names, page="Insurance\nTax\nTax"):
    results = build_results(documents(tmp_path, *names), powerbi, portal)
    collapse_superseded(results)
    pasted = read_pasted_order(page, build_vocabulary(powerbi, portal))
    report = reconcile(results, portal, build_portal_order(pasted, pasted.suggested))
    return results, report


# --- matching a file the export does not name -------------------------------


def test_an_unnamed_lawyer_file_is_matched_by_firm_and_table_code(tmp_path, powerbi):
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    rows = build_results(documents(tmp_path, BOB), powerbi, portal)

    row = rows[0]
    assert row.matched_by_code
    assert row.status == "matched"
    assert row.practice_area == "City Lawyers > Tax"
    assert row.export_row is not None
    assert any("firm reference and table code" in i for i in row.issues)


def test_a_file_matched_by_code_keeps_its_own_name(tmp_path, powerbi):
    """It is a different document from the one the portal lists."""
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    rows = build_results(documents(tmp_path, BOB), powerbi, portal)
    assert rows[0].filename == BOB


def test_two_known_files_for_one_table_are_not_guessed_between(tmp_path, powerbi):
    portal = write_portal(tmp_path / "portal.csv", [
        lawyer_row("id-2", "Ann Lee", ANN),
        lawyer_row("id-3", "Cid Lima", ANN.replace("Ann_Lee", "Cid_Lima")),
    ])
    rows = build_results(documents(tmp_path, BOB), powerbi, portal)
    assert not rows[0].matched_by_code
    assert rows[0].status == "unmatched"


def test_the_same_table_code_at_another_firm_is_not_a_match(tmp_path, powerbi):
    """EC codes name tables, which every firm shares."""
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    other = BOB.replace(REF, "88002")
    rows = build_results(documents(tmp_path, other), powerbi, portal)
    assert not rows[0].matched_by_code


def test_two_lawyers_in_one_table_are_both_kept(tmp_path, powerbi):
    """Not an older and a newer version, and not two copies of one file."""
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    results, _ = run(tmp_path, powerbi, portal, [FIRM_FILE, ANN, BOB])
    assert [r.status for r in results] == ["matched", "matched", "matched"]


# --- the page against a repeating export -------------------------------------


def test_a_repeated_line_is_flagged_to_check_on_the_page(tmp_path, powerbi):
    """The reported shape: the second lawyer's file is in, the first's is not."""
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    _, report = run(tmp_path, powerbi, portal, [BOB, FIRM_FILE], page="Tax\nTax\nInsurance")

    assert [r.status for r in report.rows] == ["uploaded", "check page", "uploaded"]
    assert report.rows[1].repeat_of == 1
    assert report.missing == []
    assert not report.accounted_for


def test_a_repeated_line_counts_as_a_document_of_its_own(tmp_path, powerbi):
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE, BOB])
    assert report.distinct_documents == 3
    assert report.shared_rows == 0


def test_both_lawyers_uploaded_covers_both_lines(tmp_path, powerbi):
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE, ANN, BOB])
    assert [r.status for r in report.rows] == ["uploaded"] * 3
    assert report.accounted_for
    assert report.not_on_page == []


def test_the_named_file_covers_the_first_line(tmp_path, powerbi):
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE, BOB, ANN])
    assert report.rows[1].uploaded.source_filename == ANN
    assert report.rows[2].uploaded.source_filename == BOB


def test_a_repeated_line_below_the_batch_is_not_started(tmp_path, powerbi):
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE])
    assert [r.status for r in report.rows] == ["uploaded", "not reached", "not reached"]


def test_distinct_rows_sharing_a_file_are_still_shared(tmp_path, powerbi):
    """Different portal IDs on one file is a genuine shared document."""
    portal = write_portal(tmp_path / "portal.csv", [
        firm_row("id-1", "Insurance", FIRM_FILE),
        firm_row("id-9", "Insurance", FIRM_FILE),
    ])
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE], page="Insurance\nInsurance")
    assert report.repeats == []
    assert report.shared_rows == 1


# --- what is still to download, by name -------------------------------------


def test_still_to_download_is_decided_by_name_not_count(tmp_path, powerbi):
    """Equal counts once hid a wrong file standing in for a missing one."""
    portal = write_portal(tmp_path / "portal.csv", [
        firm_row("id-1", "Insurance", FIRM_FILE),
        lawyer_row("id-2", "Ann Lee", ANN),
    ])
    other = FIRM_FILE.replace("EC100", "EC555")  # a file this page does not list
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE, other], page="Insurance\nTax")
    assert report.not_downloaded == [ANN]


def test_nothing_to_download_when_every_file_is_in(tmp_path, powerbi):
    portal = write_portal(tmp_path / "portal.csv", REPEATED)
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE, ANN, BOB])
    assert report.not_downloaded == []
