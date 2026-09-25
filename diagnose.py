"""Show exactly what the parser sees in one submission document.

    .venv/bin/python diagnose.py "/path/to/a submission.docx"

Prints every table the counters look at, every row, and whether it was counted.
Use it when a spot-check disagrees with the app.
"""

from __future__ import annotations

import sys

from docx import Document

from parsing import (
    NOMINATION_PATTERNS,
    RE_AFFIRMATIVE,
    RE_CLIENT_BOILERPLATE,
    RE_CLIENT_HEADER_ROW,
    RE_CLIENT_TABLE,
    RE_EXISTING_WORD,
    RE_MATTER_TABLE,
    RE_NEGATIVE,
    RE_NEW_WORD,
    RE_NOT_APPLICABLE,
    clean,
    iter_blocks,
    nomination_name,
    parse_document,
    row_text,
    table_label,
)


def verdict(answer: str) -> str:
    if RE_NOT_APPLICABLE.match(answer):
        return "flagged (not yes or no)"
    if RE_AFFIRMATIVE.match(answer) or RE_NEW_WORD.match(answer):
        return "NEW"
    if RE_NEGATIVE.match(answer) or RE_EXISTING_WORD.match(answer):
        return "not new"
    if answer == "":
        return "blank -> not new"
    return "flagged (not yes or no)"


def main(path: str) -> None:
    document = Document(path)
    blocks = list(iter_blocks(document))

    print("=" * 78)
    print(path.rsplit("/", 1)[-1])
    print("=" * 78)

    client_tables = 0
    counted = 0

    for index, (kind, item) in enumerate(blocks):
        if kind != "tbl":
            continue
        label = table_label(item)
        if not RE_CLIENT_TABLE.match(label):
            continue

        client_tables += 1
        # the nearest paragraph above, which says which list this is
        heading = ""
        for j in range(index - 1, -1, -1):
            if blocks[j][0] == "p" and clean(blocks[j][1].text):
                heading = clean(blocks[j][1].text)
                break

        print(f"\nCLIENT TABLE {client_tables}  (block {index}, "
              f"{len(item.rows)} rows x {len(item.columns)} cols)")
        print(f"  under heading: {heading[:70]!r}")
        print(f"  label row    : {label[:70]!r}")

        in_this_table = 0
        for r, row in enumerate(item.rows[1:], start=1):
            cells = row_text(row)
            name = cells[0] if cells else ""
            answer = cells[1] if len(cells) > 1 else ""

            if not name:
                note = "skipped - no client name"
            elif RE_CLIENT_BOILERPLATE.search(name):
                note = "skipped - boilerplate row"
            elif RE_CLIENT_HEADER_ROW.match(name):
                note = "skipped - repeated header row"
            else:
                in_this_table += 1
                counted += 1
                note = f"COUNTED #{counted:<3} answer={answer!r} -> {verdict(answer)}"
            print(f"    r{r:<3} {name[:44]!r:<46} {note}")

        print(f"  -> {in_this_table} client(s) counted in this table")

    print(f"\n{client_tables} client table(s) found, {counted} client(s) counted in total")

    print("\n--- other tables the counters look at ---")
    for index, (kind, item) in enumerate(blocks):
        if kind != "tbl":
            continue
        label = table_label(item)
        if RE_MATTER_TABLE.match(label):
            print(f"  block {index:<4} MATTER      {label[:60]!r}")
        for key, pattern in NOMINATION_PATTERNS.items():
            if pattern.match(label):
                print(f"  block {index:<4} {key:<11} {label[:44]!r} "
                      f"name={nomination_name(item)[:28]!r}")

    # Tables that look like they should have been counted but were not. A label
    # that fails to match is otherwise silent - the table is simply absent.
    print("\n--- near misses: tables that look countable but did not match ---")
    near = 0
    for index, (kind, item) in enumerate(blocks):
        if kind != "tbl":
            continue
        label = table_label(item)
        low = label.lower()
        matched = (
            RE_CLIENT_TABLE.match(label)
            or RE_MATTER_TABLE.match(label)
            or any(p.match(label) for p in NOMINATION_PATTERNS.values())
        )
        if matched:
            continue
        # Real labels are short. Narrative cells mention these words in prose.
        looks_like = len(label) <= 80 and (
            ("matter" in low and "summary" not in low)
            or low.startswith(("partner", "associate"))
            or "active key client" in low
        )
        if looks_like:
            near += 1
            print(f"  block {index:<4} label={label[:64]!r}")
            print(f"            first rows: "
                  f"{[row_text(r)[:2] for r in item.rows[:3]]}")
    if not near:
        print("  (none - every table that looks countable was counted)")

    # Anything at all that mentions a matter number, so a renumbered or
    # oddly-labelled table shows up here even if it is not table-shaped.
    print("\n--- every table label in the document, in order ---")
    for index, (kind, item) in enumerate(blocks):
        if kind == "tbl":
            print(f"  {index:>4}  {table_label(item)[:76]!r}")

    print("\n--- what the app reports for this document ---")
    parsed = parse_document(path)
    for key, metric in parsed.metrics.items():
        print(f"  {key:<15} {metric.display}")
        for note in metric.notes:
            print(f"      note: {note}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        raise SystemExit(1)
    main(sys.argv[1])
