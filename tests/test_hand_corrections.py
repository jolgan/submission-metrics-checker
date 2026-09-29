"""Figures corrected by hand, and the highlighting that follows them.

The fill used to be decided once, from whatever the parser produced. Anyone
editing the exported file then had a green cell that no longer meant "differs
from the export", and an amber cell gave no clue whether the figure they wrote
in matched the original or not.

Every value now passes through final_value, so the highlighting answers one
question only: does this cell differ from the Power BI export, whoever decided
it.
"""

from io import BytesIO

import openpyxl
import pytest

from conftest import EXPECTED, PORTAL, POWERBI, SAMPLES
from export_writer import AMBER, CHECKED_COLUMN, CHECKED_TEXT, GREEN, build_workbook
from matching import (
    build_results,
    difference_records,
    final_differs,
    final_value,
    load_portal,
    load_powerbi,
    override_key,
    results_to_frame,
    review_rows,
)
from parsing import METRIC_COLUMNS, UNPARSED, parse_document


@pytest.fixture(scope="module")
def results():
    powerbi, portal = load_powerbi(POWERBI), load_portal(PORTAL)
    return build_results([parse_document(p) for p in SAMPLES.values()], powerbi, portal)


@pytest.fixture
def corrected(results):
    """alpha's next generation count, which the export has as 0."""
    row = next(r for r in results if "alpha" in r.filename)
    return row, "next_gen"


# --- the value that wins ----------------------------------------------------


def test_without_an_override_the_parsers_figure_stands(corrected):
    row, metric = corrected
    assert final_value(row, metric, {}) == EXPECTED["alpha"]["next_gen"]
    assert final_differs(row, metric, {}) is True


def test_a_hand_figure_overrules_the_parser(corrected):
    row, metric = corrected
    overrides = {override_key(row, metric): 9}
    assert final_value(row, metric, overrides) == 9
    assert final_differs(row, metric, overrides) is True


def test_correcting_back_to_the_export_stops_the_highlight(corrected):
    """The case that prompted this: green must mean 'differs', always."""
    row, metric = corrected
    original = row.export_value(metric)
    overrides = {override_key(row, metric): original}
    assert final_value(row, metric, overrides) == original
    assert final_differs(row, metric, overrides) is False


def test_a_hand_figure_matching_the_export_leaves_the_changes_list(results, corrected):
    row, metric = corrected
    before = [
        d for d in difference_records(results, {})
        if d["Filename"] == row.filename and d["Metric"] == "Next generation"
    ]
    assert before, "it starts out as a difference"

    overrides = {override_key(row, metric): row.export_value(metric)}
    after = [
        d for d in difference_records(results, overrides)
        if d["Filename"] == row.filename and d["Metric"] == "Next generation"
    ]
    assert after == [], "corrected back to the export, so no longer a change"


def test_the_changes_list_says_who_decided(results, corrected):
    row, metric = corrected
    assert all(d["Decided by"] == "the app" for d in difference_records(results, {}))

    overrides = {override_key(row, metric): 9}
    entry = next(
        d for d in difference_records(results, overrides)
        if d["Metric"] == "Next generation"
    )
    assert entry["Decided by"] == "you"
    assert entry["Recounted value"] == 9


def test_an_uncounted_cell_filled_by_hand_is_judged_like_any_other(results):
    """An amber cell given a figure becomes green only if it differs."""
    row = results[0]
    row.document.metrics["matters"].value = None
    assert final_value(row, "matters", {}) is None
    assert final_differs(row, "matters", {}) is False, "nothing to compare yet"

    same = {override_key(row, "matters"): row.export_value("matters")}
    assert final_differs(row, "matters", same) is False

    other = {override_key(row, "matters"): row.export_value("matters") + 3}
    assert final_differs(row, "matters", other) is True


def test_clearing_a_hand_figure_marks_the_cell_uncounted_again(corrected):
    row, metric = corrected
    overrides = {override_key(row, metric): None}
    assert final_value(row, metric, overrides) is None
    assert final_differs(row, metric, overrides) is False


# --- the table on screen ----------------------------------------------------


def test_the_table_shows_the_hand_figure(results, corrected):
    row, metric = corrected
    overrides = {override_key(row, metric): 9}
    frame = results_to_frame(results, overrides=overrides)
    shown = frame[frame["FILE NAME"] == row.filename].iloc[0][METRIC_COLUMNS[metric]]
    assert shown == "9"


def test_a_cleared_cell_falls_back_to_the_export_figure(results, corrected):
    """Amber carries the export's own number, so a hand count has a baseline."""
    from matching import display_value

    row, metric = corrected
    overrides = {override_key(row, metric): None}
    frame = results_to_frame(results, overrides=overrides)
    shown = frame[frame["FILE NAME"] == row.filename].iloc[0][METRIC_COLUMNS[metric]]
    assert shown == str(row.export_value(metric))

    value, unverified = display_value(row, metric, overrides)
    assert value == row.export_value(metric)
    assert unverified is True


# --- the review list --------------------------------------------------------


def test_the_review_list_holds_only_cells_worth_a_decision(results):
    entries = review_rows(results, {})
    assert entries
    assert all(entry["Why"] != "-" for entry in entries)
    assert len(entries) < len(results) * 6


def test_the_review_list_can_show_everything(results):
    assert len(review_rows(results, {}, everything=True)) == len(results) * 6


def test_the_review_list_carries_both_figures(results):
    entry = next(e for e in review_rows(results, {}) if "differs" in e["Why"])
    assert entry["Power BI"] is not None
    assert entry["App counted"] is not None
    assert entry["Your value"] == entry["App counted"]


def test_a_corrected_cell_leaves_the_review_list(results, corrected):
    row, metric = corrected
    overrides = {override_key(row, metric): row.export_value(metric)}
    remaining = [
        e for e in review_rows(results, overrides)
        if e["_key"] == override_key(row, metric)
    ]
    assert remaining == [], "it no longer differs, so there is nothing to decide"


# --- the exported workbook --------------------------------------------------


def _sheet(results, overrides=None, checked=None):
    data = build_workbook(results, overrides, checked)
    return openpyxl.load_workbook(BytesIO(data))["Recount"]


def _cell(sheet, filename, column):
    header = [c.value for c in sheet[1]]
    name_col = header.index("FILE NAME") + 1
    for r in range(2, sheet.max_row + 1):
        if sheet.cell(row=r, column=name_col).value == filename:
            return sheet.cell(row=r, column=header.index(column) + 1)
    raise AssertionError(f"{filename} not in the sheet")


def test_the_export_writes_the_hand_figure(results, corrected):
    row, metric = corrected
    sheet = _sheet(results, {override_key(row, metric): 9})
    cell = _cell(sheet, row.filename, METRIC_COLUMNS[metric])
    assert cell.value == 9
    assert cell.fill.start_color.rgb == GREEN.start_color.rgb


def test_the_export_drops_the_fill_when_it_matches_the_export(results, corrected):
    row, metric = corrected
    original = row.export_value(metric)
    sheet = _sheet(results, {override_key(row, metric): original})
    cell = _cell(sheet, row.filename, METRIC_COLUMNS[metric])
    assert cell.value == original
    assert cell.fill.start_color.rgb != GREEN.start_color.rgb
    assert cell.fill.start_color.rgb != AMBER.start_color.rgb


def test_the_export_comment_says_who_decided(results, corrected):
    row, metric = corrected
    sheet = _sheet(results, {override_key(row, metric): 9})
    cell = _cell(sheet, row.filename, METRIC_COLUMNS[metric])
    assert "decided by you" in cell.comment.text


def test_a_flagged_figure_is_amber_and_keeps_its_number(results):
    """Amber says "your eyes here"; the counted figure is still readable.

    Bold was tried and dropped: it travels with a copy into the master sheet
    and has to be cleared by hand.
    """
    row = next(r for r in results if "gamma" in r.filename)
    assert row.document.metrics["matters"].numbering_gap is True
    cell = _cell(_sheet(results), row.filename, "MATTERS")
    assert cell.value == EXPECTED["gamma"]["matters"]
    assert cell.fill.start_color.rgb == AMBER.start_color.rgb
    assert cell.font.bold in (None, False), "no bold to clear before pasting"
    assert "worth checking" in cell.comment.text


def test_nothing_in_the_sheet_is_bold(results):
    sheet = _sheet(results)
    for r in range(2, sheet.max_row + 1):
        for c in range(1, sheet.max_column + 1):
            assert sheet.cell(row=r, column=c).font.bold in (None, False)


# --- the double-checked column ----------------------------------------------


def test_the_checked_column_is_the_last_one(results):
    sheet = _sheet(results)
    assert [c.value for c in sheet[1]][-1] == CHECKED_COLUMN


def test_an_unticked_row_leaves_the_column_blank(results):
    assert _cell(_sheet(results), results[0].filename, CHECKED_COLUMN).value is None


def test_a_ticked_row_is_marked(results):
    ticked = {results[0].filename.lower()}
    sheet = _sheet(results, checked=ticked)
    assert _cell(sheet, results[0].filename, CHECKED_COLUMN).value == CHECKED_TEXT
    assert _cell(sheet, results[1].filename, CHECKED_COLUMN).value is None


def test_ticking_does_not_disturb_the_figures(results):
    plain = _sheet(results)
    ticked = _sheet(results, checked={r.filename.lower() for r in results})
    for row in results:
        for column in METRIC_COLUMNS.values():
            assert _cell(plain, row.filename, column).value == (
                _cell(ticked, row.filename, column).value
            )


# --- an uncounted cell carries the export's own figure ----------------------


def test_an_uncounted_cell_shows_the_export_figure_in_amber(results):
    """Counting 4 by hand tells you nothing unless you can see what was there."""
    row = results[0]
    original = row.export_value("matters")
    row.document.metrics["matters"].value = None

    cell = _cell(_sheet(results), row.filename, "MATTERS")
    assert cell.value == original, "the baseline, not a phrase"
    assert cell.fill.start_color.rgb == AMBER.start_color.rgb
    assert "unverified" in cell.comment.text
    row.document.metrics["matters"].value = original  # restore for other tests


def test_an_uncounted_cell_is_not_green_even_though_it_holds_a_figure(results):
    row = results[0]
    row.document.metrics["matters"].value = None
    cell = _cell(_sheet(results), row.filename, "MATTERS")
    assert cell.fill.start_color.rgb != GREEN.start_color.rgb
    row.document.metrics["matters"].value = EXPECTED["alpha"]["matters"]


def test_a_hand_count_matching_the_export_stays_unhighlighted(results):
    """The case that prompted this: count 4, the export already said 4."""
    row = results[0]
    original = row.export_value("matters")
    row.document.metrics["matters"].value = None

    overrides = {override_key(row, "matters"): original}
    cell = _cell(_sheet(results, overrides), row.filename, "MATTERS")
    assert cell.value == original
    assert cell.fill.start_color.rgb != GREEN.start_color.rgb
    assert cell.fill.start_color.rgb != AMBER.start_color.rgb, "it is settled now"
    row.document.metrics["matters"].value = EXPECTED["alpha"]["matters"]


def test_a_hand_count_differing_from_the_export_turns_green(results):
    row = results[0]
    original = row.export_value("matters")
    row.document.metrics["matters"].value = None

    overrides = {override_key(row, "matters"): original + 2}
    cell = _cell(_sheet(results, overrides), row.filename, "MATTERS")
    assert cell.value == original + 2
    assert cell.fill.start_color.rgb == GREEN.start_color.rgb
    row.document.metrics["matters"].value = EXPECTED["alpha"]["matters"]


def test_the_phrase_remains_when_there_is_no_export_figure_either(results):
    """An unmatched row has no baseline to offer."""
    row = results[0]
    row.document.metrics["matters"].value = None
    kept = row.export_row
    row.export_row = None
    try:
        cell = _cell(_sheet(results), row.filename, "MATTERS")
        assert cell.value == UNPARSED
        assert cell.fill.start_color.rgb == AMBER.start_color.rgb
    finally:
        row.export_row = kept
        row.document.metrics["matters"].value = EXPECTED["alpha"]["matters"]


# --- editing in the table itself -------------------------------------------


def test_a_cell_left_alone_is_not_taken_as_accepted(results):
    """An amber cell shows the export's figure; leaving it must not settle it.

    The table renders that figure so a hand count has something to compare
    against, so "unchanged" has to mean unchanged from what was rendered, not
    equal to the export.
    """
    from matching import results_to_frame

    row = results[0]
    original = row.export_value("matters")
    row.document.metrics["matters"].value = None
    try:
        frame = results_to_frame(results, overrides={}, numeric=True)
        shown = frame[frame["FILE NAME"] == row.filename].iloc[0]["MATTERS"]
        assert shown == original, "the baseline is rendered into the cell"

        # the app compares what came back against what it rendered; equal means
        # untouched, so no override is recorded and the cell stays unverified
        assert final_value(row, "matters", {}) is None
    finally:
        row.document.metrics["matters"].value = EXPECTED["alpha"]["matters"]


def test_the_review_note_says_what_wants_checking(results):
    from matching import review_note

    row = next(r for r in results if "gamma" in r.filename)
    note = review_note(row, {})
    assert "MATTERS" in note
    assert "numbering does not add up" in note


def test_the_review_note_carries_the_baseline_for_an_uncounted_cell(results):
    from matching import review_note

    row = results[0]
    original = row.export_value("matters")
    row.document.metrics["matters"].value = None
    try:
        note = review_note(row, {})
        assert "MATTERS: not counted" in note
        assert f"Power BI says {original}" in note
    finally:
        row.document.metrics["matters"].value = EXPECTED["alpha"]["matters"]


def test_a_settled_row_has_an_empty_review_note(results):
    from matching import review_note

    row = next(r for r in results if "alpha" in r.filename)
    assert review_note(row, {}) == ""


def test_the_was_columns_show_the_export_figures(results):
    from matching import results_to_frame

    frame = results_to_frame(results, show_original=True, numeric=True)
    row = frame[frame["FILE NAME"].str.contains("alpha")].iloc[0]
    assert row["was NGs"] == 0
    assert row["NGs"] == EXPECTED["alpha"]["next_gen"]


# --- ticking a row is the act of verification -------------------------------


def test_ticking_clears_an_amber_cell(results):
    """Amber means nobody has looked. Ticking says somebody has."""
    row = results[0]
    original = row.export_value("matters")
    row.document.metrics["matters"].value = None
    try:
        plain = _cell(_sheet(results), row.filename, "MATTERS")
        assert plain.fill.start_color.rgb == AMBER.start_color.rgb

        ticked = _cell(
            _sheet(results, checked={row.filename.lower()}), row.filename, "MATTERS"
        )
        assert ticked.value == original, "the export's figure, now confirmed"
        assert ticked.fill.start_color.rgb != AMBER.start_color.rgb
        assert ticked.fill.start_color.rgb != GREEN.start_color.rgb
    finally:
        row.document.metrics["matters"].value = EXPECTED["alpha"]["matters"]


def test_ticking_clears_a_numbering_flag(results):
    row = next(r for r in results if "gamma" in r.filename)
    assert _cell(_sheet(results), row.filename, "MATTERS").fill.start_color.rgb == (
        AMBER.start_color.rgb
    )
    ticked = _cell(
        _sheet(results, checked={row.filename.lower()}), row.filename, "MATTERS"
    )
    assert ticked.fill.start_color.rgb != AMBER.start_color.rgb
    assert ticked.value == EXPECTED["gamma"]["matters"]


def test_ticking_does_not_hide_a_genuine_difference(results):
    """A revised figure stays green after ticking: it still differs."""
    row = next(r for r in results if "alpha" in r.filename)
    ticked = _sheet(results, checked={row.filename.lower()})
    cell = _cell(ticked, row.filename, "NGs")
    assert cell.value == EXPECTED["alpha"]["next_gen"]
    assert cell.fill.start_color.rgb == GREEN.start_color.rgb


def test_a_hand_figure_on_a_ticked_row_is_still_judged_on_its_value(results):
    row = results[0]
    ticked = {row.filename.lower()}
    original = row.export_value("matters")

    same = _cell(
        _sheet(results, {override_key(row, "matters"): original}, ticked),
        row.filename, "MATTERS",
    )
    assert same.fill.start_color.rgb != GREEN.start_color.rgb

    other = _cell(
        _sheet(results, {override_key(row, "matters"): original + 5}, ticked),
        row.filename, "MATTERS",
    )
    assert other.fill.start_color.rgb == GREEN.start_color.rgb


def test_ticking_empties_the_review_note(results):
    from matching import review_note

    row = next(r for r in results if "gamma" in r.filename)
    assert review_note(row, {}) != ""
    assert review_note(row, {}, {row.filename.lower()}) == ""


def test_ticking_one_row_leaves_the_others_alone(results):
    row = next(r for r in results if "gamma" in r.filename)
    other = next(r for r in results if "alpha" in r.filename)
    sheet = _sheet(results, checked={row.filename.lower()})
    assert _cell(sheet, other.filename, "NGs").fill.start_color.rgb == (
        GREEN.start_color.rgb
    )
