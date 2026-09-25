"""Compare two exported result tables and list only what changed.

    .venv/bin/python compare_exports.py old.xlsx new.xlsx

Use this after re-exporting, so a fresh export does not cost you the
spot-checking you already did: rows that are identical need no second look,
and only the cells listed here do.

Rows are matched on FILE NAME, never on position, because the row order
depends on which "Row order" was selected when each file was saved.
"""

from __future__ import annotations

import sys

import openpyxl

# The current green, and the paler one used before, so a comparison against an
# older export does not report every corrected cell as a change of shade.
GREENS = {"0092D050", "00C6EFCE"}
AMBER = "00FFE9A8"
METRICS = ["MATTERS", "CLIENTS", "NEW", "LPs", "NGs", "LAs"]


def read(path: str) -> tuple[dict[str, dict], list[str]]:
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["Recount"] if "Recount" in workbook.sheetnames else workbook.worksheets[0]
    header = [str(c.value) if c.value is not None else "" for c in sheet[1]]
    if "FILE NAME" not in header:
        raise SystemExit(f"{path}: no FILE NAME column - is this a result table export?")

    key_at = header.index("FILE NAME")
    rows: dict[str, dict] = {}
    for raw in sheet.iter_rows(min_row=2):
        name = raw[key_at].value
        if not name:
            continue
        record = {}
        for i, column in enumerate(header):
            if i >= len(raw):
                continue
            cell = raw[i]
            rgb = cell.fill.start_color.rgb
            shade = "green" if rgb in GREENS else "amber" if rgb == AMBER else ""
            record[column] = (cell.value, shade)
        rows[str(name)] = record
    return rows, header


def describe(value, shade) -> str:
    text = "(blank)" if value is None else str(value)
    return f"{text} [{shade}]" if shade else text


def main(old_path: str, new_path: str) -> None:
    old, header = read(old_path)
    new, _ = read(new_path)

    only_old = sorted(set(old) - set(new))
    only_new = sorted(set(new) - set(old))
    shared = [name for name in new if name in old]

    print(f"old: {len(old)} row(s)   new: {len(new)} row(s)")
    if only_old:
        print(f"\n{len(only_old)} row(s) only in the old export:")
        for name in only_old:
            print(f"   {name}")
    if only_new:
        print(f"\n{len(only_new)} row(s) only in the new export:")
        for name in only_new:
            print(f"   {name}")

    changed_rows = 0
    changed_cells = 0
    for name in shared:
        differences = []
        for column in header:
            before = old[name].get(column, (None, ""))
            after = new[name].get(column, (None, ""))
            if before != after:
                differences.append((column, before, after))
        if not differences:
            continue
        changed_rows += 1
        changed_cells += len(differences)
        print(f"\n{name}")
        for column, before, after in differences:
            mark = "  <-- recheck" if column in METRICS else ""
            print(f"   {column:<14} {describe(*before):<28} -> {describe(*after)}{mark}")

    print(f"\n{len(shared)} row(s) in both. {changed_rows} changed, "
          f"{len(shared) - changed_rows} identical.")
    print(f"{changed_cells} cell(s) differ - those are the only ones worth "
          "checking again.")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__)
        raise SystemExit(1)
    main(sys.argv[1], sys.argv[2])
