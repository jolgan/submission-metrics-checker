"""Reading exports whose columns have been renamed or reshaped.

The Power BI dashboard was refreshed in September 2026 and every column was
relabelled except FIRM_NAME: MATCH_STATUS became "Matched?", NUM_MATTERS became
"# Matters", CONCATENATED_PRACTICE_AREA_NAME became "Practice Area". The loader
had been matching header text exactly, so the whole file became unreadable.

Column names are now treated as spellings of a field, with the six counts
falling back to the last six columns of the sheet, which is the one thing about
the layout that holds.
"""

import csv

import openpyxl
import pytest

from conftest import POWERBI
from matching import (
    METRIC_EXPORT_COLUMNS,
    load_portal,
    load_powerbi,
    resolve_powerbi_columns,
)

OLD_HEADER = [
    "MATCH_STATUS", "SUBMISSION_FILE_NAME", "FIRM_NAME", "FIRM_REF",
    "CONCATENATED_PRACTICE_AREA_NAME", "COUNTRY_NAME",
    "NUM_MATTERS", "NUM_ACTIVE_CLIENTS", "NUM_NEW_CLIENTS",
    "NUM_LEAD_PARTNERS", "NUM_NEXT_GEN", "NUM_ASSOCIATES",
]

NEW_HEADER = [
    "Matched?", "File Name", "FIRM_NAME", "Firm Ref",
    "SUBMISSION_PRACTICE_AREA_NAME", "Country", "Practice Area",
    "Ranking Type", "Tier", "Ranking Decision", "# Referees",
    "# Referees replied", "# Matters", "# Active Clients", "# New Clients",
    "# L.Partners", "# Next Gens", "# Associates",
]


def write(path, header, rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Export"
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


# --- resolving the header ---------------------------------------------------


@pytest.mark.parametrize("header", [OLD_HEADER, NEW_HEADER])
def test_both_vocabularies_resolve(header):
    found, _ = resolve_powerbi_columns(header)
    for field in ("FIRM_NAME", "FIRM_REF", "COUNTRY_NAME",
                  "CONCATENATED_PRACTICE_AREA_NAME"):
        assert field in found, f"{field} not placed in {header[:4]}"
    for metric in METRIC_EXPORT_COLUMNS:
        assert metric in found


def test_the_new_names_point_at_the_right_columns():
    found, _ = resolve_powerbi_columns(NEW_HEADER)
    assert NEW_HEADER[found["NUM_MATTERS"]] == "# Matters"
    assert NEW_HEADER[found["NUM_LEAD_PARTNERS"]] == "# L.Partners"
    assert NEW_HEADER[found["NUM_NEXT_GEN"]] == "# Next Gens"
    assert NEW_HEADER[found["FIRM_REF"]] == "Firm Ref"
    assert NEW_HEADER[found["COUNTRY_NAME"]] == "Country"


def test_the_fuller_practice_area_name_wins():
    """Both 'Practice Area' and 'SUBMISSION_PRACTICE_AREA_NAME' are present."""
    found, _ = resolve_powerbi_columns(NEW_HEADER)
    assert NEW_HEADER[found["CONCATENATED_PRACTICE_AREA_NAME"]] == "Practice Area"


def test_headers_are_matched_case_and_space_insensitively():
    header = ["  matched?  ", "firm name", "FIRM REF", "country", "practice area",
              "# MATTERS", "# active clients", "# new clients", "# l.partners",
              "# next gens", "# associates"]
    found, _ = resolve_powerbi_columns(header)
    assert "FIRM_NAME" in found and "NUM_MATTERS" in found


def test_unknown_count_names_fall_back_to_the_last_six_columns():
    """The dashboard may rename them again, but it does not move them."""
    header = ["Firm Name", "Firm Ref", "Country", "Practice Area",
              "a", "b", "c", "d", "e", "f"]
    found, notes = resolve_powerbi_columns(header)
    assert [header[found[m]] for m in METRIC_EXPORT_COLUMNS] == list("abcdef")
    assert any("last six columns" in n for n in notes)


def test_a_partial_rename_is_reported_rather_than_assumed():
    """Half recognised is worse than none: position cannot be trusted."""
    header = ["Firm Name", "Firm Ref", "Country", "Practice Area",
              "# Matters", "b", "c", "d", "e", "f"]
    found, notes = resolve_powerbi_columns(header)
    assert "NUM_MATTERS" in found
    assert "NUM_ASSOCIATES" not in found
    assert any("nothing was assumed about position" in n for n in notes)


# --- loading a whole file ---------------------------------------------------


def test_a_renamed_export_loads(tmp_path):
    path = write(tmp_path / "new.xlsx", NEW_HEADER, [
        ["match", "x.docx", "Example Firm", 51617, "Banking", "Argentina",
         "Banking and finance", "firm recommended", 3, "ranked", 18, 3,
         15, 26, 8, 3, 1, 1],
    ])
    export = load_powerbi(path)
    assert len(export.rows) == 1
    row = export.rows[0]
    assert row["FIRM_NAME"] == "Example Firm"
    assert row["COUNTRY_NAME"] == "Argentina"
    assert row["CONCATENATED_PRACTICE_AREA_NAME"] == "Banking and finance"
    assert row["NUM_MATTERS"] == 15
    assert row["NUM_ASSOCIATES"] == 1


def test_the_old_format_still_loads():
    export = load_powerbi(POWERBI)
    assert export.rows
    assert export.rows[0]["NUM_MATTERS"] is not None


def test_a_renamed_export_is_still_searchable(tmp_path):
    path = write(tmp_path / "lookup.xlsx", NEW_HEADER, [
        ["match", "x.docx", "Example Firm", 51617, "Banking", "Argentina",
         "Banking and finance", "firm recommended", 3, "ranked", 18, 3,
         15, 26, 8, 3, 1, 1],
    ])
    export = load_powerbi(path)
    assert export.lookup("Example Firm", "Argentina", "Banking and finance")


def test_the_filter_footer_is_still_dropped(tmp_path):
    path = write(tmp_path / "footer.xlsx", NEW_HEADER, [
        ["match", "x.docx", "Example Firm", 51617, "Banking", "Argentina",
         "Banking and finance", "firm recommended", 3, "ranked", 18, 3,
         15, 26, 8, 3, 1, 1],
        [None] * len(NEW_HEADER),
        ["Applied filters:\nCOUNTRY_NAME is Argentina"] + [None] * (len(NEW_HEADER) - 1),
    ])
    export = load_powerbi(path)
    assert len(export.rows) == 1
    assert export.dropped == 2


def test_a_workbook_without_the_fields_is_still_refused(tmp_path):
    path = write(tmp_path / "wrong.xlsx", ["A", "B", "C"], [[1, 2, 3]])
    with pytest.raises(ValueError, match="expected Power BI columns"):
        load_powerbi(path)


# --- the portal export arrives as .csv for some publications ---------------


def test_the_portal_reader_accepts_csv(tmp_path):
    path = tmp_path / "portal.csv"
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Firm", "Firm Ref", "Region", "Practice Group Name",
                         "Practice Area Name", "Table Name", "Created At",
                         "Document", "File URI"])
        writer.writerow(["Example Firm", "51617", "Argentina", "",
                         "Banking and finance", "Banking and finance",
                         "28/01/2026", "Y",
                         "https://example.invalid/x/example_submission.docx"])

    portal = load_portal(path)
    record = portal.lookup("example_submission.docx")
    assert record is not None
    assert record["firm"] == "Example Firm"
    assert record["region"] == "Argentina"


def test_a_csv_with_a_byte_order_mark_still_reads(tmp_path):
    """The portal writes its CSV with a BOM."""
    path = tmp_path / "bom.csv"
    path.write_text(
        "﻿Firm,Region,Practice Area Name,Document,File URI\n"
        "Example Firm,Argentina,Tax,Y,https://example.invalid/x/a_submission.docx\n",
        encoding="utf-8",
    )
    portal = load_portal(path)
    assert portal.lookup("a_submission.docx") is not None


# --- the way the app actually calls these -----------------------------------
#
# The uploader hands an in-memory buffer with no file name attached. Deciding
# the format from the extension worked when called with a path and raised
# BadZipFile from the interface, which is the only way it is really used.


def _buffer(path):
    from io import BytesIO
    from pathlib import Path

    return BytesIO(Path(path).read_bytes())


def test_a_csv_arriving_as_a_nameless_buffer_is_read(tmp_path):
    path = tmp_path / "portal.csv"
    path.write_text(
        "Firm,Region,Practice Area Name,Document,File URI\n"
        "Example Firm,Argentina,Tax,Y,https://example.invalid/x/a_submission.docx\n",
        encoding="utf-8",
    )
    portal = load_portal(_buffer(path))
    assert portal.lookup("a_submission.docx") is not None


def test_an_xlsx_arriving_as_a_nameless_buffer_is_still_read():
    from conftest import PORTAL

    portal = load_portal(_buffer(PORTAL))
    assert portal.records


def test_the_format_is_decided_by_content_not_by_name(tmp_path):
    """A CSV saved with an .xlsx name must still be read as text."""
    path = tmp_path / "misnamed.xlsx"
    path.write_text(
        "Firm,Region,Practice Area Name,Document,File URI\n"
        "Example Firm,Argentina,Tax,Y,https://example.invalid/x/b_submission.docx\n",
        encoding="utf-8",
    )
    assert load_portal(path).lookup("b_submission.docx") is not None
    assert load_portal(_buffer(path)).lookup("b_submission.docx") is not None


def test_a_buffer_is_left_where_it_was_found(tmp_path):
    """Sniffing must not consume the stream the reader is about to use."""
    path = tmp_path / "portal.csv"
    path.write_text(
        "Firm,Region,Practice Area Name,Document,File URI\n"
        "Example Firm,Argentina,Tax,Y,https://example.invalid/x/c_submission.docx\n",
        encoding="utf-8",
    )
    buffer = _buffer(path)
    assert buffer.tell() == 0
    load_portal(buffer)
