"""Ordering the table from text pasted off the Staff Portal screen."""

import pytest

from conftest import DETAILS, PORTAL, POWERBI, SAMPLES
from matching import (
    build_portal_order,
    build_results,
    build_vocabulary,
    load_portal,
    load_powerbi,
    read_pasted_order,
    sort_results,
    unordered_rows,
)
from parsing import parse_document

# The three practice areas, deliberately not in upload order.
AREAS = [
    DETAILS["gamma"]["practice_area"].split(" > ")[1],
    DETAILS["alpha"]["practice_area"].split(" > ")[1],
    DETAILS["beta"]["practice_area"].split(" > ")[1],
]
EXPECTED_ORDER = ["gamma", "alpha", "beta"]


@pytest.fixture(scope="module")
def exports():
    return load_powerbi(POWERBI), load_portal(PORTAL)


@pytest.fixture(scope="module")
def known(exports):
    return build_vocabulary(*exports)


@pytest.fixture(scope="module")
def results(exports):
    powerbi, portal = exports
    documents = [parse_document(p) for p in SAMPLES.values()]
    return build_results(documents, powerbi, portal)


def order_of(rows):
    """Which sample each row is, by the key in its filename."""
    out = []
    for row in rows:
        for key, path in SAMPLES.items():
            if path.name == row.filename:
                out.append(key)
    return out


def apply(text, results, known=None, column=None):
    pasted = read_pasted_order(text, known)
    order = build_portal_order(pasted, column)
    return order_of(sort_results(results, "portal_paste", order)), order


# --- the common case: one column of practice areas -------------------------


def test_single_column_paste_drives_the_table(results):
    ordering, order = apply("\n".join(AREAS), results)
    assert ordering == EXPECTED_ORDER
    assert order.size == 3


def test_blank_lines_and_stray_whitespace_are_ignored(results):
    text = "\n".join(["", f"  {AREAS[0]}  ", "", AREAS[1], AREAS[2], "  "])
    ordering, order = apply(text, results)
    assert ordering == EXPECTED_ORDER
    assert order.size == 3


def test_windows_line_endings_are_handled(results):
    ordering, _ = apply("\r\n".join(AREAS), results)
    assert ordering == EXPECTED_ORDER


def test_full_group_and_area_spelling_also_matches(results):
    text = "\n".join(DETAILS[k]["practice_area"] for k in EXPECTED_ORDER)
    ordering, _ = apply(text, results)
    assert ordering == EXPECTED_ORDER


def test_empty_paste_is_rejected_clearly():
    with pytest.raises(ValueError, match="Nothing was pasted"):
        build_portal_order(read_pasted_order("   \n  \n"))


# --- whole tables pasted off the page --------------------------------------


def test_pasted_table_with_headings_picks_the_practice_area_column(results):
    header = "Firm\tRegion\tPractice Area Name"
    lines = [header] + [
        f"{DETAILS[k]['firm']}\t{DETAILS[k]['region']}\t"
        f"{DETAILS[k]['practice_area'].split(' > ')[1]}"
        for k in EXPECTED_ORDER
    ]
    ordering, order = apply("\n".join(lines), results)
    assert ordering == EXPECTED_ORDER
    assert "Practice Area Name" in order.column_label


def test_headings_are_matched_case_insensitively(results):
    ordering, _ = apply("\n".join(["  PRACTICE AREA  "] + AREAS), results)
    assert ordering == EXPECTED_ORDER


def test_pasted_table_without_headings_finds_the_column_by_content(results, known):
    lines = [f"2027 Edition\tA Firm\t{area}\tfirm\tN" for area in AREAS]
    ordering, order = apply("\n".join(lines), results, known=known)
    assert ordering == EXPECTED_ORDER
    assert "Column 3" in order.column_label


def test_a_data_value_of_firm_is_not_mistaken_for_a_heading(known):
    """Portal rows carry a submission type literally spelled 'firm'."""
    lines = [f"2027 Edition\tA Firm\t{area}\tfirm\tN" for area in AREAS]
    pasted = read_pasted_order("\n".join(lines), known)
    assert pasted.headers is None
    assert pasted.entry_count == 3, "the first data row must not be eaten"


def test_the_column_can_be_overridden_by_hand(results, known):
    lines = [f"A Firm\t{area}" for area in AREAS]
    pasted = read_pasted_order("\n".join(lines), known)
    assert pasted.suggested == 1

    ordering, _ = apply("\n".join(lines), results, known=known, column=1)
    assert ordering == EXPECTED_ORDER

    with pytest.raises(ValueError) as caught:
        apply("\n".join(lines), results, known=known, column=0)
    assert "cannot order the table" in str(caught.value)


def test_a_filename_column_is_used_when_named(results):
    lines = ["File Name"] + [SAMPLES[k].name for k in EXPECTED_ORDER]
    ordering, _ = apply("\n".join(lines), results)
    assert ordering == EXPECTED_ORDER


def test_column_labels_show_a_sample_so_the_right_one_can_be_picked(known):
    lines = [f"2027 Edition\tA Firm\t{area}" for area in AREAS]
    pasted = read_pasted_order("\n".join(lines), known)
    assert pasted.width == 3
    assert pasted.column_labels[0].startswith("Column 1 — 2027 Edition")
    assert pasted.entry_count == 3


def test_ragged_rows_do_not_break_the_split(known):
    lines = [f"A Firm\t{area}" for area in AREAS] + ["A Firm"]
    pasted = read_pasted_order("\n".join(lines), known)
    assert all(len(row) == pasted.width for row in pasted.rows)
    assert build_portal_order(pasted, 1).size == 3


def test_out_of_range_column_is_rejected(known):
    pasted = read_pasted_order("\n".join(AREAS), known)
    with pytest.raises(ValueError, match="not in the pasted text"):
        build_portal_order(pasted, 5)


# --- checking the paste against the exports --------------------------------


def test_every_column_reports_how_much_is_recognised(known):
    lines = [f"A Firm\t{area}" for area in AREAS]
    pasted = read_pasted_order("\n".join(lines), known)
    assert pasted.recognised(0) == 0
    assert pasted.recognised(1) == 3
    assert "0 of 3 recognised" in pasted.column_labels[0]
    assert "3 of 3 recognised" in pasted.column_labels[1]


def test_content_beats_a_misleading_heading(known):
    lines = ["Practice Area Name\tFirm"] + [f"A Firm\t{area}" for area in AREAS]
    pasted = read_pasted_order("\n".join(lines), known)
    assert pasted.named["practice_area"] == 0
    assert pasted.suggested == 1, "column 1 holds the practice areas"


def test_unrecognised_values_are_listed_back(known):
    text = "\n".join(AREAS + ["Not A Practice Area", "Also Not One"])
    pasted = read_pasted_order(text, known)
    assert pasted.recognised(0) == 3
    assert pasted.unrecognised(0) == ["Not A Practice Area", "Also Not One"]


def test_the_portal_vocabulary_adds_to_the_powerbi_one(exports):
    powerbi, portal = exports
    powerbi_only = build_vocabulary(powerbi)
    both = build_vocabulary(powerbi, portal)
    assert len(both.areas) >= len(powerbi_only.areas)
    assert both.filenames and not powerbi_only.filenames


def test_without_a_vocabulary_nothing_is_claimed():
    pasted = read_pasted_order("\n".join(AREAS))
    assert not pasted.checked
    assert pasted.recognised(0) == 0
    assert "recognised" not in pasted.column_labels[0]
    assert build_portal_order(pasted, 0).size == 3


# --- behaviour of the resulting order --------------------------------------


def test_rows_missing_from_the_paste_go_last_not_dropped(results):
    ordering, order = apply("\n".join(AREAS[:2]), results)
    assert len(ordering) == 3
    assert ordering[:2] == EXPECTED_ORDER[:2]
    assert len(unordered_rows(results, order)) == 1


def test_extra_pasted_rows_are_harmless(results):
    text = "\n".join(AREAS + ["Pensions"])
    ordering, order = apply(text, results)
    assert ordering == EXPECTED_ORDER
    assert order.size == 4


def test_requesting_portal_order_without_a_paste_is_an_error(results):
    with pytest.raises(ValueError, match="no order sheet"):
        sort_results(results, "portal_paste", None)


def test_a_mostly_unrecognised_paste_is_recognisable_as_such(known):
    """Copying a portal page cell by cell puts each field on its own line.

    Dates, counts and the word DOCX then arrive as if they were practice areas,
    and almost nothing matches. The proportion recognised is what tells the two
    cases apart: a paste for the wrong firm matches nothing at all, a paste of
    the wrong shape matches a little.
    """
    lines = []
    for area in AREAS:
        lines += ["Example Firm", area, "28/01/2026", "DOCX", "19"]
    pasted = read_pasted_order("\n".join(lines), known)

    assert pasted.entry_count == len(AREAS) * 5
    assert pasted.recognised(0) == len(AREAS)
    assert pasted.recognised(0) / pasted.entry_count < 0.5


def test_a_well_shaped_paste_is_fully_recognised(known):
    pasted = read_pasted_order("\n".join(AREAS), known)
    assert pasted.recognised(0) == pasted.entry_count


def test_a_three_level_row_places_against_a_pasted_bare_area():
    """The Staff Portal screen prints the area; the row may carry three levels.

    Where the portal's table name differs from its area name only by
    punctuation the row reads as three levels deep, and the pasted line then
    placed nothing, pushing a correctly parsed document to the bottom of the
    table as "not in the pasted order".
    """
    from matching import _order_keys, build_portal_order, read_pasted_order

    order = build_portal_order(read_pasted_order(
        "Electricity (and renewable energy)\nCorporate and M&A"
    ))
    area = (
        "Energy and natural resources > Electricity (and renewable energy) "
        "> Electricity and renewable energy"
    )
    keys = _order_keys(None, "Example Advogados", "Brazil", None, area)
    assert any(k in order.index for k in keys)
    assert order.index[next(k for k in keys if k in order.index)] == 0
