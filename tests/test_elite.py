"""Elite tables are not recounted, so an uploaded Elite file is left out.

"Rio de Janeiro Elite" tables are lawyer nominations, and the Staff Portal and
Power BI disagree on which lawyers they hold. Rather than ask the checker to
spot them, the app drops them from the table, the reconciliation and the
Power BI comparison. Only the practice group decides: a table merely named
"Elite Boutique" is an ordinary submission.
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
    is_elite,
    is_elite_area,
    load_portal,
    load_powerbi,
    read_pasted_order,
    reconcile,
)
from parsing import parse_document

FIRM, REF = "Example Advogados", "77001"
FIRM_FILE = f"Example_Advogados_{REF}_submission_2027 Edition_EC100.docx"
ELITE_NAMED = f"Example_Advogados_Ann_Lee_{REF}_submission_2027 Edition_EC900.docx"
ELITE_UNNAMED = f"Example_Advogados_Bob_Ray_{REF}_submission_2027 Edition_EC900.docx"
HEADER = ["ID", "Firm", "Firm Ref", "Region", "Practice Group Name",
          "Practice Area Name", "Table Name", "Created At", "Submission Type",
          "Person Name", "Document", "File URI"]


@pytest.fixture
def portal(tmp_path):
    path = tmp_path / "portal.csv"
    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADER)
        writer.writerow(["id-1", FIRM, REF, "Brazil", "", "Tax", "Tax",
                         "2026-02-13", "firm", "", "Y", "https://x.invalid/" + FIRM_FILE])
        # Listed twice under one ID, as the export does for two lawyers.
        for _ in range(2):
            writer.writerow(["id-2", FIRM, REF, "Brazil", "Rio de Janeiro Elite",
                             "Tax", "Tax", "2026-02-27", "lawyer", "Ann Lee", "Y",
                             "https://x.invalid/" + ELITE_NAMED])
    return load_portal(path)


@pytest.fixture
def powerbi(tmp_path):
    path = tmp_path / "powerbi.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["FIRM_NAME", "FIRM_REF", "COUNTRY_NAME",
                  "CONCATENATED_PRACTICE_AREA_NAME", "NUM_MATTERS",
                  "NUM_ACTIVE_CLIENTS", "NUM_NEW_CLIENTS", "NUM_LEAD_PARTNERS",
                  "NUM_NEXT_GEN", "NUM_ASSOCIATES"])
    sheet.append([FIRM, REF, "Brazil", "Tax", 1, 1, 1, 0, 0, 0])
    sheet.append([FIRM, REF, "Brazil", "Rio de Janeiro Elite > Tax > Tax", 1, 0, 0, 0, 0, 0])
    workbook.save(path)
    return load_powerbi(path)


def run(tmp_path, powerbi, portal, names, page="Tax\nTax\nTax"):
    docs = [parse_document(make_submission(tmp_path / n)) for n in names]
    results = build_results(docs, powerbi, portal)
    collapse_superseded(results)
    pasted = read_pasted_order(page, build_vocabulary(powerbi, portal))
    report = reconcile(results, portal, build_portal_order(pasted, pasted.suggested))
    return results, report


# --- recognising an Elite table ---------------------------------------------


@pytest.mark.parametrize("group", ["Rio de Janeiro Elite", "São Paulo Elite", "ELITE"])
def test_an_elite_practice_group_is_recognised(group):
    assert is_elite(group)


@pytest.mark.parametrize("group", ["", None, "Energy and natural resources", "Elitex"])
def test_other_groups_are_not_elite(group):
    assert not is_elite(group)


def test_an_elite_boutique_table_is_not_an_elite_group():
    """The only 'elite' in the UK export: Mining > Elite Boutique."""
    assert not is_elite_area("Energy and natural resources > Mining > Elite Boutique")
    assert is_elite_area("Rio de Janeiro Elite > Dispute Resolution > Dispute resolution")


# --- an uploaded Elite file -------------------------------------------------


def test_an_elite_file_is_set_aside(tmp_path, powerbi, portal):
    results, _ = run(tmp_path, powerbi, portal, [FIRM_FILE, ELITE_NAMED])
    statuses = {r.source_filename: r.status for r in results}
    assert statuses == {FIRM_FILE: "matched", ELITE_NAMED: "elite"}
    elite = next(r for r in results if r.status == "elite")
    assert elite.export_row is None, "never compared against Power BI"
    assert any("not recounted" in i for i in elite.issues)


def test_an_elite_file_the_export_does_not_name_is_set_aside(tmp_path, powerbi, portal):
    """The second lawyer, found only by firm reference and table code."""
    results, _ = run(tmp_path, powerbi, portal, [FIRM_FILE, ELITE_UNNAMED])
    assert next(r for r in results if r.source_filename == ELITE_UNNAMED).status == "elite"


def test_a_regular_submission_in_the_same_area_is_still_counted(tmp_path, powerbi, portal):
    results, _ = run(tmp_path, powerbi, portal, [FIRM_FILE])
    assert results[0].status == "matched"
    assert results[0].export_row["CONCATENATED_PRACTICE_AREA_NAME"] == "Tax"


# --- the pasted page --------------------------------------------------------


def test_elite_lines_are_neither_expected_nor_missing(tmp_path, powerbi, portal):
    """Nothing to download, so a page with only the firm file in is complete."""
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE])
    assert [r.status for r in report.rows] == ["uploaded", "elite", "elite"]
    assert len(report.expected) == 1
    assert report.missing == [] and report.unclear == []
    assert report.accounted_for


def test_uploading_an_elite_file_changes_nothing_on_the_page(tmp_path, powerbi, portal):
    _, with_elite = run(tmp_path, powerbi, portal, [FIRM_FILE, ELITE_NAMED, ELITE_UNNAMED])
    assert [r.status for r in with_elite.rows] == ["uploaded", "elite", "elite"]
    assert with_elite.not_on_page == [], "Elite files are not strays"
    assert with_elite.attempted == 1, "Elite files do not stretch the batch"


def test_the_page_figures_still_add_up(tmp_path, powerbi, portal):
    _, report = run(tmp_path, powerbi, portal, [FIRM_FILE])
    assert report.page_size == (
        len(report.without_document) + len(report.expected)
        + len(report.unrecognised) + len(report.elite)
    )
