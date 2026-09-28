"""A firm's national table next to a City focus table for the same area.

One firm has both "Tax" and "City focus - Porto Alegre > Tax > City focus -
Porto Alegre - Tax" in the Power BI export. The City focus row is also indexed
under its bare area, "tax", so both rows answer to "Tax". Two things went
wrong:

* the national "Tax" submission found both rows and was reported ambiguous;
* the City focus submission matched its own row on some runs and both rows on
  others, because the spellings were tried in set order, which Python shuffles
  from one run to the next.
"""

import subprocess
import sys

import openpyxl
import pytest

from conftest import ROOT
from matching import load_powerbi

FIRM = "Example Advogados"
CITY = "City focus - Porto Alegre > Tax > City focus - Porto Alegre - Tax"


@pytest.fixture
def export_path(tmp_path):
    path = tmp_path / "powerbi.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["FIRM_NAME", "FIRM_REF", "COUNTRY_NAME", "CONCATENATED_PRACTICE_AREA_NAME"])
    sheet.append([FIRM, "1", "Brazil", "Tax"])
    sheet.append([FIRM, "1", "Brazil", CITY])
    sheet.append([FIRM, "1", "Brazil", "Finance > Banking and finance > Banking and finance"])
    sheet.append([FIRM, "1", "Brazil", "Banking and finance"])
    workbook.save(path)
    return path


def names(rows):
    return [r["CONCATENATED_PRACTICE_AREA_NAME"] for r in rows]


def test_the_national_table_is_its_own_row(export_path):
    assert names(load_powerbi(export_path).lookup(FIRM, "Brazil", "Tax")) == ["Tax"]


def test_the_city_focus_table_is_its_own_row(export_path):
    assert names(load_powerbi(export_path).lookup(FIRM, "Brazil", CITY)) == [CITY]


def test_the_two_level_city_focus_spelling_finds_the_city_row(export_path):
    """The portal names it 'City focus - Porto Alegre > Tax' plus a table."""
    found = load_powerbi(export_path).lookup(FIRM, "Brazil", "City focus - Porto Alegre > Tax")
    assert names(found) == [CITY]


def test_a_repeated_level_still_counts_as_the_same_name(export_path):
    """The exact-name rule must not split a row from its repeated spelling."""
    found = load_powerbi(export_path).lookup(FIRM, "Brazil", "Finance > Banking and finance")
    assert names(found) == ["Finance > Banking and finance > Banking and finance"]


def test_a_genuine_duplicate_is_still_ambiguous(tmp_path):
    path = tmp_path / "dupes.xlsx"
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["FIRM_NAME", "FIRM_REF", "COUNTRY_NAME", "CONCATENATED_PRACTICE_AREA_NAME"])
    sheet.append([FIRM, "1", "Brazil", "Tax"])
    sheet.append([FIRM, "1", "Brazil", "Tax"])
    workbook.save(path)
    assert len(load_powerbi(path).lookup(FIRM, "Brazil", "Tax")) == 2


@pytest.mark.parametrize("seed", ["1", "2", "6", "42"])
def test_the_answer_does_not_change_between_runs(export_path, seed):
    """Seeds 2 and 6 made the City focus row ambiguous before the fix."""
    script = (
        "import sys; sys.path.insert(0, sys.argv[1])\n"
        "from matching import load_powerbi\n"
        "rows = load_powerbi(sys.argv[2]).lookup(sys.argv[3], 'Brazil', sys.argv[4])\n"
        "print(len(rows))\n"
    )
    out = subprocess.run(
        [sys.executable, "-c", script, str(ROOT), str(export_path), FIRM, CITY],
        capture_output=True, text=True, env={"PYTHONHASHSEED": seed}, check=True,
    )
    assert out.stdout.strip() == "1"
