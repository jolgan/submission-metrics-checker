"""Check a downloaded batch against the Staff Portal, before you start checking.

    .venv/bin/python check_batch.py "Example Firm"

Names the submissions the portal lists for that firm which are not in your
downloads folder. Run it straight after downloading: the browser reports how
many downloads it started, not how many arrived, so the two disagree more often
than you would think.

With no firm name it lists the folders it can see, so you never have to guess
what to type.

It looks for your downloads under the Windows Downloads folder, and for a Staff
Portal export (.csv or .xlsx) in samples/real. Both can be given explicitly:

    .venv/bin/python check_batch.py "Example Firm" --folder "/path/to/docs" \
        --portal samples/real/latam-portal.csv
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path
from urllib.parse import unquote

HERE = Path(__file__).resolve().parent
# Under WSL the Windows Downloads folder sits in /mnt/c/Users/<name>/Downloads.
DOWNLOADS = sorted(Path("/mnt/c/Users").glob("*/Downloads")) + [Path.home() / "Downloads"]
PORTAL_DIRS = [HERE / "samples" / "real", HERE / "samples", HERE]


def basename(uri: str) -> str:
    return unquote(str(uri).split("?")[0].replace("\\", "/").rsplit("/", 1)[-1]).strip()


def created_at(row: dict) -> str:
    """The portal's Created At, as a string that sorts chronologically.

    Dates arrive as text or as a datetime depending on whether the export was
    saved as .csv or .xlsx. A row with no date sorts oldest, so it never wins a
    comparison and never silently sets another row aside.
    """
    value = row.get("Created At") or row.get("Created at") or row.get("created_at")
    if value is None or str(value).strip() == "":
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    text = str(value).strip()
    match = re.match(r"(\d{2})/(\d{2})/(\d{4})", text)  # dd/mm/yyyy
    if match:
        return f"{match.group(3)}-{match.group(2)}-{match.group(1)}"
    return text


def table_code(name: str) -> tuple[str, str] | None:
    """(firm reference, EC table code) from a submission file name."""
    match = re.search(r"_(\d+)_submission_.*_(EC\d+)\.docx$", name, re.I)
    return (match.group(1), match.group(2).upper()) if match else None


def portal_rows(path: Path) -> list[dict]:
    """Every row of a portal export, whether .csv or .xlsx."""
    if path.suffix.lower() == ".csv":
        with open(path, encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))

    import openpyxl

    workbook = openpyxl.load_workbook(path, data_only=True)
    for sheet in workbook.worksheets:
        header = [str(c.value).strip() if c.value is not None else "" for c in sheet[1]]
        if "File URI" not in header:
            continue
        rows = []
        for raw in sheet.iter_rows(min_row=2):
            row = {header[i]: c.value for i, c in enumerate(raw) if i < len(header)}
            uri_cell = raw[header.index("File URI")]
            if uri_cell.hyperlink is not None and uri_cell.hyperlink.target:
                row["File URI"] = uri_cell.hyperlink.target
            rows.append(row)
        return rows
    return []


# Downloads are filed one folder per firm, but those are often grouped by
# country ("Downloads/mexico/Sainz Abogados"), so a firm is looked for a level
# or two down rather than only at the top.
FOLDER_DEPTH = 3


def find_folders() -> list[Path]:
    out: list[Path] = []
    for root in DOWNLOADS:
        if not root.is_dir():
            continue
        for depth in range(1, FOLDER_DEPTH + 1):
            out += [p for p in sorted(root.glob("/".join(["*"] * depth)))
                    if p.is_dir() and any(p.glob("*.docx"))]
    return out


def folder_label(folder: Path) -> str:
    """The folder's name, with its parent when that is what tells firms apart."""
    for root in DOWNLOADS:
        try:
            relative = folder.relative_to(root)
        except ValueError:
            continue
        return str(relative)
    return folder.name


def nearby_portals(folder: Path) -> list[Path]:
    """Portal exports filed with the documents themselves."""
    out: list[Path] = []
    for place in (folder / "spreadsheets", folder, folder.parent / "spreadsheets"):
        if place.is_dir():
            out += sorted(p for p in place.iterdir()
                          if p.suffix.lower() in (".csv", ".xlsx")
                          and not p.name.startswith("~$"))
    return out


def find_portals() -> list[Path]:
    out = []
    for folder in PORTAL_DIRS:
        if folder.is_dir():
            out += sorted(p for p in folder.iterdir()
                          if p.suffix.lower() in (".csv", ".xlsx"))
    return out


def pick_folder(firm: str) -> Path | None:
    wanted = firm.lower().replace(" ", "")
    folders = find_folders()
    for folder in folders:  # the firm's own folder name first
        if wanted in folder.name.lower().replace(" ", ""):
            return folder
    for folder in folders:  # then anywhere in the path, so "mexico" works too
        if wanted in folder_label(folder).lower().replace(" ", ""):
            return folder
    return None


def main() -> int:
    parser = argparse.ArgumentParser(add_help=True, description=__doc__)
    parser.add_argument("firm", nargs="?", help='the firm, e.g. "Example Firm"')
    parser.add_argument("--folder", help="where the documents are, if not in Downloads")
    parser.add_argument("--portal", help="the Staff Portal export to check against")
    args = parser.parse_args()

    if not args.firm and not args.folder:
        print("Give a firm name, for example:\n")
        print('    .venv/bin/python check_batch.py "Example Firm"\n')
        folders = find_folders()
        if folders:
            print("Folders with documents in them:")
            for folder in folders:
                print(f"    {folder_label(folder)}   ({len(list(folder.glob('*.docx')))} files)")
        return 1

    folder = Path(args.folder) if args.folder else pick_folder(args.firm)
    if folder is None or not folder.is_dir():
        print(f"No downloads folder found for {args.firm!r}.")
        print("Folders I can see:")
        for candidate in find_folders():
            print(f"    {folder_label(candidate)}")
        print("\nOr point at it directly with --folder \"/path/to/folder\"")
        return 1

    # With --folder and no name, the folder is named after the firm.
    if not args.firm:
        args.firm = folder.name

    on_disk = {p.name.lower(): p.name for p in folder.glob("*.docx")}

    # The firm's own export is the one to trust. A publication-wide export
    # kept in samples/ often covers the firm too, which makes the right answer
    # arrive for the wrong reason - and the wrong answer, silently, when it
    # does not cover them.
    portals = (
        [Path(args.portal)] if args.portal
        else nearby_portals(folder) + find_portals()
    )
    rows, used = [], None
    for candidate in portals:
        if not candidate.is_file():
            continue
        try:
            found = [r for r in portal_rows(candidate)
                     if args.firm.lower() in str(r.get("Firm") or "").lower()]
        except Exception:
            continue
        if found:
            rows, used = found, candidate
            break

    if not rows:
        print(f"{args.firm} does not appear in any Staff Portal export I can find.")
        print("Exports I looked at:")
        for candidate in portals:
            print(f"    {candidate}")
        print("\nExport the firm's publication from the Staff Portal into "
              "samples/real, or pass --portal \"/path/to/export.csv\"")
        return 1

    # Elite tables ("Rio de Janeiro Elite") are not recounted, so their
    # documents are not expected in the folder. A table merely named "Elite
    # Boutique" is ordinary, so only the practice group counts.
    elite = [r for r in rows if re.search(r"\belite\b", str(r.get("Practice Group Name") or ""), re.I)]
    rows = [r for r in rows if r not in elite]

    wanted, no_document = {}, []
    for row in rows:
        uri = str(row.get("File URI") or "").strip()
        if uri:
            wanted.setdefault(basename(uri), row)
        else:
            no_document.append(row)

    # A firm that resubmits leaves both rows on the portal. Only the newer is
    # counted, so only the newer is worth fetching - listing the older one as
    # missing sends the checker after a file that would be set aside the moment
    # it arrived. Decided on Created At; a tie is left alone rather than
    # guessed at.
    superseded = {}
    slots: dict[tuple, list[tuple[str, dict]]] = {}
    for name, row in wanted.items():
        key = (
            str(row.get("Firm Name") or "").strip().lower(),
            str(row.get("Country") or row.get("Region") or "").strip().lower(),
            str(row.get("Practice Area Name") or "").strip().lower(),
        )
        slots.setdefault(key, []).append((name, row))
    for entries in slots.values():
        if len(entries) < 2:
            continue
        dated = [(n, r, created_at(r)) for n, r in entries]
        newest = max(d for _, _, d in dated)
        if sum(1 for _, _, d in dated if d == newest) > 1:
            continue
        for name, row, date in dated:
            if date != newest:
                superseded[name] = row
    wanted = {n: r for n, r in wanted.items() if n not in superseded}

    missing = {n: r for n, r in wanted.items() if n.lower() not in on_disk}
    extra = [on_disk[k] for k in on_disk if k not in {n.lower() for n in wanted}]

    # An Elite file the portal does not name - a second lawyer in the table -
    # still carries the table's firm reference and EC code.
    elite_codes = {table_code(basename(r.get("File URI") or "")) for r in elite}
    elite_codes.discard(None)
    elite_files = [n for n in extra if table_code(n) in elite_codes]
    extra = [n for n in extra if n not in elite_files]

    print(f"{args.firm}")
    print(f"  documents in {folder.name}/ : {len(on_disk)}")
    print(f"  the portal lists            : {len(rows)} row(s), "
          f"{len(wanted)} with a document")
    if no_document:
        print(f"  no document to download     : {len(no_document)}")
        for row in no_document:
            print(f"      {row.get('Region')} | {row.get('Practice Area Name')}")
    if elite:
        print(f"  Elite rows skipped          : {len(elite)} (not recounted)")
    print(f"  checked against             : {used.name}")

    if missing:
        print(f"\nMISSING {len(missing)} file(s):")
        for name, row in missing.items():
            print(f"    {name}")
            print(f"        {row.get('Region')} | {row.get('Practice Area Name')}")
    else:
        print("\nNothing missing. Every document the portal lists is in the folder.")
    if superseded:
        print(f"\nolder versions, already replaced ({len(superseded)}), "
              "nothing to download:")
        for name, row in superseded.items():
            print(f"    {name}")
            print(f"        {row.get('Region')} | {row.get('Practice Area Name')}"
                  f" | {created_at(row)[:10]}")

    if elite_files:
        print(f"\n{len(elite_files)} Elite file(s) in the folder - not recounted, "
              "the app leaves them out:")
        for name in sorted(elite_files):
            print(f"    {name}")

    if extra:
        print(f"\n{len(extra)} file(s) in the folder but not on the portal "
              "(another firm's, or renamed):")
        for name in sorted(extra):
            print(f"    {name}")

    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
