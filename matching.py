"""Load the two exports and join recounted documents to their export rows.

Pure logic: no Streamlit. Everything takes a path or a file-like object so the
tests can drive it without launching the app.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any
from urllib.parse import unquote

import openpyxl
import pandas as pd

from parsing import EXPORT_COLUMNS, METRIC_COLUMNS, METRIC_KEYS, ParsedDocument, UNPARSED

POWERBI_REQUIRED = [
    "FIRM_NAME",
    "FIRM_REF",
    "COUNTRY_NAME",
    "CONCATENATED_PRACTICE_AREA_NAME",
]

# The dashboard renames its columns from time to time - a refresh in September
# 2026 relabelled every one of them except FIRM_NAME. The names are therefore
# treated as spellings of a field rather than as the field itself. Each entry is
# tried in order, so the most specific name wins where several are present.
POWERBI_ALIASES = {
    "MATCH_STATUS": ("match_status", "matched?", "matched", "match status"),
    "FIRM_NAME": ("firm_name", "firm name", "firm"),
    "FIRM_REF": ("firm_ref", "firm ref", "firm reference"),
    "COUNTRY_NAME": ("country_name", "country", "region"),
    "CONCATENATED_PRACTICE_AREA_NAME": (
        "concatenated_practice_area_name",
        "practice area",
        "practice area name",
        "submission_practice_area_name",
    ),
    "NUM_MATTERS": ("num_matters", "# matters", "matters"),
    "NUM_ACTIVE_CLIENTS": (
        "num_active_clients", "# active clients", "active clients",
    ),
    "NUM_NEW_CLIENTS": ("num_new_clients", "# new clients", "new clients"),
    "NUM_LEAD_PARTNERS": (
        "num_lead_partners", "# l.partners", "# lead partners", "lead partners",
    ),
    "NUM_NEXT_GEN": (
        "num_next_gen", "# next gens", "# next gen", "next gens", "next gen",
    ),
    "NUM_ASSOCIATES": ("num_associates", "# associates", "associates"),
}

METRIC_EXPORT_COLUMNS = [
    "NUM_MATTERS",
    "NUM_ACTIVE_CLIENTS",
    "NUM_NEW_CLIENTS",
    "NUM_LEAD_PARTNERS",
    "NUM_NEXT_GEN",
    "NUM_ASSOCIATES",
]


def _header_key(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def resolve_powerbi_columns(header: list) -> tuple[dict[str, int], list[str]]:
    """Work out which column holds which field, and say what was assumed.

    Returns the column index for each field it could place. The six counts are
    the last six columns of the sheet, so where none of their names is
    recognised that position is used instead - the dashboard may rename them
    again, but it does not move them.
    """
    names = [_header_key(h) for h in header]
    found: dict[str, int] = {}
    notes: list[str] = []

    for canonical, aliases in POWERBI_ALIASES.items():
        for alias in aliases:
            if alias in names:
                found[canonical] = names.index(alias)
                break

    placed = [m for m in METRIC_EXPORT_COLUMNS if m in found]
    if not placed and len(header) >= 6:
        for offset, canonical in enumerate(METRIC_EXPORT_COLUMNS):
            found[canonical] = len(header) - 6 + offset
        notes.append(
            "None of the six count columns were recognised by name, so the last "
            "six columns were used: "
            + ", ".join(str(header[found[m]]) for m in METRIC_EXPORT_COLUMNS)
        )
    elif len(placed) < len(METRIC_EXPORT_COLUMNS):
        missing = [m for m in METRIC_EXPORT_COLUMNS if m not in found]
        notes.append(
            "Some count columns were recognised and others were not, so nothing "
            "was assumed about position. Missing: " + ", ".join(missing)
        )

    return found, notes


def normalise(value: Any) -> str:
    """Lowercase, trim, collapse whitespace, standardise spacing around '>'."""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    text = re.sub(r"\s*>\s*", " > ", text)
    return text.lower()


FIRM_SUFFIXES = re.compile(
    r"\b(llp|ltd|limited|plc|lp|inc|incorporated|solicitors)\b\.?", re.I
)


def firm_key(name: Any) -> str:
    """A firm name reduced to its distinctive part, for diagnostics only."""
    text = normalise(name)
    text = FIRM_SUFFIXES.sub(" ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _level_words(text: str) -> str:
    """Two levels that differ only in punctuation name the same thing.

    The portal stores the practice area and the table name in separate columns
    and staff do not always punctuate them alike, so "Electricity (and
    renewable energy)" can be filed under a table named "Electricity and
    renewable energy". Comparing the words alone keeps the pair recognisable
    as one level rather than two.
    """
    return re.sub(r"[^0-9a-z]+", " ", text).strip()


def practice_area_variants(value: Any) -> set[str]:
    """Every spelling one practice area is written in across the two systems.

    The portal names a submission by practice group, practice area and table,
    and the two exports do not agree on how many of those three they print. The
    Power BI export sometimes repeats the last level
    ("Finance > Banking and finance > Banking and finance"), and the portal
    screen shows all three ("Transport > Travel > Travel: personal injury")
    where the export row is only two. Comparing the printed strings alone leaves
    real submissions unmatched, so each is reduced to the handful of spellings
    it could legitimately take.
    """
    text = normalise(value)
    if not text:
        return set()

    parts = [p.strip() for p in text.split(" > ") if p.strip()]
    forms = {text}

    # A trailing level repeating the one before it carries no information.
    # Both spellings are kept as forms of their own so that a pasted line
    # naming either one still finds the record.
    while len(parts) >= 2 and _level_words(parts[-1]) == _level_words(parts[-2]):
        forms.add(parts[-1])
        parts = parts[:-1]
        forms.add(" > ".join(parts))

    if len(parts) >= 3:
        forms.add(" > ".join(parts[-2:]))
        forms.add(" > ".join(parts[:2]))

    # Every level below the practice group, on its own. The Staff Portal screen
    # prints the practice area by itself, and where the portal's table name is
    # neither a repeat of the area nor derived from it - "City focus - Brasilia
    # - Government relations" for the area "Government relations" - the area is
    # the middle of three levels and is named by no other form.
    #
    # The practice group is deliberately left out. A group covers many areas,
    # so on its own it identifies no single submission, and indexing a record
    # under it would let a paste naming the group match an arbitrary row.
    forms.update(parts[1:])
    if parts:
        forms.add(parts[-1])
    return forms


def to_int(value: Any) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


# --- Power BI export --------------------------------------------------------


@dataclass
class PowerBIExport:
    rows: list[dict] = field(default_factory=list)
    index: dict[tuple[str, str, str], list[dict]] = field(default_factory=dict)
    dropped: int = 0
    notes: list[str] = field(default_factory=list)

    def lookup(self, firm: str, region: str, practice_area: str) -> list[dict]:
        firm_key, region_key = normalise(firm), normalise(region)
        found: list[dict] = []
        for form in practice_area_variants(practice_area):
            for row in self.index.get((firm_key, region_key, form), []):
                if not any(row is seen for seen in found):
                    found.append(row)
            if found:
                break
        return found

    def lookup_ignoring_region(self, firm: str, practice_area: str) -> list[dict]:
        fn = normalise(firm)
        forms = practice_area_variants(practice_area)
        found: list[dict] = []
        for key, rows in self.index.items():
            if key[0] != fn or key[2] not in forms:
                continue
            for row in rows:
                if not any(row is seen for seen in found):
                    found.append(row)
        return found

    def lookup_relaxed_firm(self, firm: str, practice_area: str) -> list[dict]:
        """Rows matching on practice area whose firm name differs only by suffix.

        Used to explain a failed match, never to make one: a document saying
        "Example Partners" will not silently be matched to "Example Partners LLP".
        """
        fk = firm_key(firm)
        forms = practice_area_variants(practice_area)
        found: list[dict] = []
        for key, rows in self.index.items():
            if key[2] not in forms or firm_key(key[0]) != fk:
                continue
            if key[0] == normalise(firm):
                continue
            for row in rows:
                if not any(row is seen for seen in found):
                    found.append(row)
        return found


def load_powerbi(source: Any) -> PowerBIExport:
    """Read the Power BI export.

    The sheet carries trailing junk - a row holding the 'Applied filters:' note
    and blank spacer rows - which we drop rather than let it pollute the index.
    """
    workbook = openpyxl.load_workbook(source, data_only=True)
    sheet = columns = notes = None
    for candidate in workbook.worksheets:
        header = [c.value for c in candidate[1]]
        found, found_notes = resolve_powerbi_columns(header)
        if all(r in found for r in POWERBI_REQUIRED):
            sheet, columns, notes = candidate, found, found_notes
            break
    if sheet is None:
        raise ValueError(
            "No sheet in this workbook has the expected Power BI columns. It "
            "needs a firm, a practice area, a country or region and a firm "
            "reference, under any of the names the dashboard has used "
            f"({', '.join(POWERBI_REQUIRED)}, or Firm Ref / Country / "
            "Practice Area). Is this the right file?"
        )

    header = [str(c.value).strip() if c.value is not None else "" for c in sheet[1]]
    export = PowerBIExport()
    export.notes.extend(notes)
    missing_metrics = [c for c in EXPORT_COLUMNS.values() if c not in columns]
    if missing_metrics:
        export.notes.append(
            "Export is missing metric column(s): " + ", ".join(missing_metrics)
        )

    for position, raw in enumerate(sheet.iter_rows(min_row=2)):
        row = {
            canonical: (raw[index].value if index < len(raw) else None)
            for canonical, index in columns.items()
        }
        row["_row_order"] = position
        if not row.get("FIRM_NAME") or not row.get("CONCATENATED_PRACTICE_AREA_NAME"):
            if any(v not in (None, "") for v in row.values()):
                export.dropped += 1
            continue
        export.rows.append(row)
        if is_elite_area(row.get("CONCATENATED_PRACTICE_AREA_NAME")):
            # Never compared, and indexing it would let a firm's own "Tax"
            # submission also find "Rio de Janeiro Elite > Tax > Tax".
            continue
        firm = normalise(row.get("FIRM_NAME"))
        region = normalise(row.get("COUNTRY_NAME"))
        for form in practice_area_variants(row.get("CONCATENATED_PRACTICE_AREA_NAME")):
            export.index.setdefault((firm, region, form), []).append(row)

    if export.dropped:
        export.notes.append(
            f"{export.dropped} row(s) without a firm or practice area were ignored "
            "(filter notes and blank spacer rows)."
        )
    duplicates = {k: v for k, v in export.index.items() if len(v) > 1}
    if duplicates:
        export.notes.append(
            f"{len(duplicates)} firm/region/practice-area key(s) appear more than "
            "once in the export; documents matching them are reported as ambiguous."
        )
    return export


def _pick_sheet(workbook, required: list[str]):
    for sheet in workbook.worksheets:
        header = {str(c.value).strip() for c in sheet[1] if c.value is not None}
        if all(col in header for col in required):
            return sheet
    return None


# --- Staff Portal export ----------------------------------------------------


RE_DOWNLOAD_SUFFIX = re.compile(r"^(?P<stem>.*?)\s*\((?P<copy>\d+)\)(?P<ext>\.docx)$", re.I)


def strip_download_suffix(filename: str) -> tuple[str, int | None]:
    """Undo a browser's duplicate-download suffix.

    Downloading the same document twice gives 'name (1).docx'. Returns the
    original name and which copy it was, so the caller can both match the row
    and warn that a second copy means some other document was probably missed.
    """
    match = RE_DOWNLOAD_SUFFIX.match(str(filename or "").strip())
    if not match:
        return str(filename or "").strip(), None
    return match.group("stem") + match.group("ext"), int(match.group("copy"))


@dataclass
class PortalExport:
    by_filename: dict[str, dict] = field(default_factory=dict)
    records: list[dict] = field(default_factory=list)
    by_area: dict[str, list[dict]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)
    hyperlink_count: int = 0
    text_count: int = 0
    without_document: int = 0

    def lookup(self, filename: str) -> dict | None:
        """Find a portal row by filename, tolerating a '(1)' download suffix."""
        name = str(filename or "").strip().lower()
        record = self.by_filename.get(name)
        if record is not None:
            return record
        stripped, copy = strip_download_suffix(name)
        if copy is not None:
            return self.by_filename.get(stripped.lower())
        return None

    def lookup_by_table_code(self, filename: str) -> dict | None:
        """Find a portal row by firm reference and table code, for a file the
        export does not list by name.

        Lawyer submissions are named after the person, and the export can list
        one lawyer's file where the page holds two: "..._Bob_Ray_..._77001_
        submission_2027 Edition_EC900.docx" is absent while Ann Lee's file
        for the same table is listed. The firm reference and the EC code still
        identify the table. Returns None unless exactly one listed document
        shares them, so two known files for one table are never guessed between.
        """
        wanted = _table_code(strip_download_suffix(filename)[0])
        if wanted is None:
            return None
        found = [r for r in self.by_filename.values() if _table_code(r["filename"]) == wanted]
        return found[0] if len(found) == 1 else None

    def records_for_area(self, value: Any, firm_keys: set[str] | None = None) -> list[dict]:
        """Portal rows for a practice area, optionally limited to some firms.

        Tried from the most specific spelling to the least, so a paste naming
        all three levels does not fall back to a bare area name while an exact
        match exists.
        """
        found: list[dict] = []
        for form in sorted(practice_area_variants(value), key=len, reverse=True):
            for record in self.by_area.get(form, []):
                if not any(record is seen for seen in found):
                    found.append(record)
            if found:
                break
        if firm_keys is None:
            return found
        return [r for r in found if _firm_keys_of(r) & firm_keys]


RE_ELITE = re.compile(r"\belite\b", re.I)


def is_elite(practice_group: Any) -> bool:
    """True for an Elite table ("Rio de Janeiro Elite"), which is not recounted.

    Elite tables are lawyer nominations, and the portal and Power BI disagree
    on which lawyers they hold, so they are left out rather than checked. Only
    the practice group decides: a table merely named "Elite Boutique" under
    Mining is an ordinary firm submission.
    """
    return bool(RE_ELITE.search(str(practice_group or "")))


def is_elite_area(practice_area: Any) -> bool:
    """Elite judged from a 'Group > Area' name, as Power BI prints it."""
    return is_elite(str(practice_area or "").split(">", 1)[0])


RE_TABLE_CODE = re.compile(r"_(?P<ref>\d+)_submission_.*_(?P<code>EC\d+)\.docx$", re.I)


def _table_code(filename: Any) -> tuple[str, str] | None:
    """(firm reference, EC table code) from a portal-style file name.

    The EC code names the table, not the document: every firm's Rio de Janeiro
    Elite dispute resolution file ends in EC5814568. Only together with the
    firm reference does it identify one submission slot.
    """
    match = RE_TABLE_CODE.search(str(filename or ""))
    if not match:
        return None
    return match.group("ref"), match.group("code").upper()


def _firm_keys_of(record: dict) -> set[str]:
    """Identifiers a portal row can be matched on: firm reference and name."""
    keys = set()
    if record.get("firm_ref") not in (None, ""):
        keys.add("ref:" + normalise(record["firm_ref"]))
    if record.get("firm"):
        keys.add("name:" + firm_key(record["firm"]))
    return keys


def uri_to_filename(uri: str) -> str:
    """Filename from a File URI, always ending in .docx.

    Most URIs already carry the extension. Where one ends at the document name
    with no extension at all, .docx is appended - these are Word submissions.
    A URI ending in some other extension is left alone so the caller can see it
    is not a submission document and skip it.
    """
    text = str(uri or "").split("?")[0].replace("\\", "/")
    name = unquote(text.rsplit("/", 1)[-1]).strip()
    if not name:
        return name
    if name.lower().endswith(".docx"):
        return name
    if "." not in name.rsplit("/", 1)[-1]:
        return name + ".docx"
    return name


def _workbook_from_csv(source: Any) -> "openpyxl.Workbook":
    """Load a CSV into a workbook so one reader serves both formats.

    The portal exports .xlsx for some publications and .csv for others.
    """
    import csv
    import io

    if hasattr(source, "read"):
        raw = source.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8-sig")
        handle = io.StringIO(raw)
    else:
        handle = open(source, encoding="utf-8-sig", newline="")

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "portal-export"
    with handle:
        for row in csv.reader(handle):
            sheet.append(row)
    return workbook


def _looks_like_csv(source: Any) -> bool:
    """Decide by content, not by file name.

    The app hands this an in-memory buffer with no name attached, so anything
    resting on the extension works when called with a path and fails when
    called from the interface. Every .xlsx is a zip, so its first two bytes are
    "PK"; anything else is read as text.
    """
    try:
        if hasattr(source, "read") and hasattr(source, "seek"):
            here = source.tell()
            magic = source.read(2)
            source.seek(here)
        else:
            with open(source, "rb") as handle:
                magic = handle.read(2)
    except (OSError, ValueError, TypeError):
        return False
    if isinstance(magic, str):
        magic = magic.encode("utf-8", "ignore")
    return magic[:2] != b"PK"


def load_portal(source: Any) -> PortalExport:
    """Read the Staff Portal export and index it by document filename.

    The File URI column may hold plain text or real Excel hyperlinks - in the
    latter case pandas sees only the display text ('DOCX'), so we read the
    workbook through openpyxl and take the hyperlink target when there is one.
    A .csv export is loaded into a workbook first, so one reader serves both.
    """
    if _looks_like_csv(source):
        workbook = _workbook_from_csv(source)
    else:
        workbook = openpyxl.load_workbook(source, data_only=True)
    portal = PortalExport()

    sheet, header, uri_index = _find_uri_column(workbook)
    if sheet is None:
        raise ValueError(
            "Could not find a 'File URI' column, or any column of document links, "
            "in this workbook. Is this the Staff Portal export?"
        )

    def column(name: str) -> int | None:
        return header.index(name) if name in header else None

    cols = {
        "firm": column("Firm"),
        "firm_ref": column("Firm Ref"),
        "region": column("Region"),
        "practice_group": column("Practice Group Name"),
        "practice_area": column("Practice Area Name"),
        "table_name": column("Table Name"),
        "practice_group_id": column("Practice Group ID"),
        "practice_area_id": column("Practice Area ID"),
        "created_at": column("Created At"),
        "portal_id": column("ID"),
        "person": column("Person Name"),
        "submission_type": column("Submission Type"),
    }

    for position, raw in enumerate(sheet.iter_rows(min_row=2)):
        if uri_index >= len(raw):
            continue
        cell = raw[uri_index]
        uri = None
        if cell.hyperlink is not None and cell.hyperlink.target:
            uri = cell.hyperlink.target
            portal.hyperlink_count += 1
        elif cell.value:
            uri = str(cell.value)
            portal.text_count += 1

        filename = uri_to_filename(uri) if uri else ""
        if uri and not filename.lower().endswith(".docx"):
            continue

        record = {
            "filename": filename,
            "uri": uri,
            "_row_order": position,
            "has_document": bool(uri),
        }
        for key, idx in cols.items():
            record[key] = raw[idx].value if idx is not None and idx < len(raw) else None

        portal.records.append(record)
        if filename:
            key = filename.lower()
            existing = portal.by_filename.get(key)
            # The same document can be listed twice with different dates; the
            # later listing is the one to believe.
            if existing is None or _created_at(record) > _created_at(existing):
                portal.by_filename[key] = record
        else:
            portal.without_document += 1

        # Index by every spelling this practice area takes, so a row can be
        # found whether the paste names two levels or three.
        for form in practice_area_variants(_portal_practice_area(record)):
            portal.by_area.setdefault(form, []).append(record)

    portal.notes.append(
        f"Read {len(portal.by_filename)} document filename(s) from the File URI column "
        f"({portal.hyperlink_count} from Excel hyperlinks, {portal.text_count} from plain text); "
        f"{portal.without_document} portal row(s) have no submission document linked."
    )
    return portal


def _find_uri_column(workbook):
    """Locate the File URI column, by header name or by content."""
    for sheet in workbook.worksheets:
        header = [str(c.value).strip() if c.value is not None else "" for c in sheet[1]]
        if "File URI" in header:
            return sheet, header, header.index("File URI")

    # No header row: fall back to the column that actually holds document links.
    for sheet in workbook.worksheets:
        width = min(sheet.max_column, 64)
        scores = [0] * width
        for raw in sheet.iter_rows(min_row=1, max_row=min(60, sheet.max_row)):
            for i, cell in enumerate(raw[:width]):
                target = cell.hyperlink.target if cell.hyperlink is not None else None
                text = target or (str(cell.value) if cell.value else "")
                if ".docx" in text.lower():
                    scores[i] += 1
        if scores and max(scores) > 0:
            return sheet, [""] * width, scores.index(max(scores))
    return None, [], -1


# --- pasted Staff Portal order ----------------------------------------------
#
# Pasted straight into the app rather than saved to a file first. A paste may be
# a single column of practice areas or a whole table copied off the portal page,
# in which case the columns arrive tab-separated.

ORDER_COLUMN_HINTS = {
    "filename": ("file name", "filename", "file", "document name", "document", "file uri", "uri"),
    "firm": ("firm name", "firm"),
    "region": ("region", "country name", "country"),
    "practice_group": ("practice group name", "practice group", "group"),
    "practice_area": ("practice area name", "practice area", "table name", "area", "table"),
}


@dataclass
class PortalOrder:
    """Row order copied off the Staff Portal screen."""

    index: dict[str, int] = field(default_factory=dict)
    values: list[str] = field(default_factory=list)
    size: int = 0
    column_label: str = ""
    notes: list[str] = field(default_factory=list)

    def position(self, row: "ResultRow") -> int | None:
        for key in _order_keys(
            row.filename, row.firm, row.region, None, row.practice_area
        ):
            if key in self.index:
                return self.index[key]
        return None


@dataclass
class PastedOrder:
    """A paste, split into rows and columns, before a column is chosen."""

    rows: list[list[str]] = field(default_factory=list)
    headers: list[str] | None = None
    column_labels: list[str] = field(default_factory=list)
    named: dict[str, int] = field(default_factory=dict)
    suggested: int = 0
    scores: list[int] = field(default_factory=list)
    checked: bool = False

    @property
    def width(self) -> int:
        return len(self.column_labels)

    @property
    def entry_count(self) -> int:
        return len(self.rows)

    def recognised(self, column: int) -> int:
        """How many values in a column the exports recognise."""
        if not self.checked or not 0 <= column < len(self.scores):
            return 0
        return self.scores[column]

    def unrecognised(self, column: int, limit: int = 5) -> list[str]:
        """Values the exports do not know, for showing back to the checker."""
        if not self.checked or not 0 <= column < self.width:
            return []
        seen: list[str] = []
        for row in self.rows:
            value = row[column]
            if value and value not in seen and not self._vocabulary.recognises(value):
                seen.append(value)
                if len(seen) >= limit:
                    break
        return seen

    _vocabulary: "Vocabulary" = field(default_factory=lambda: Vocabulary())


def _order_keys(
    filename: Any,
    firm: Any,
    region: Any,
    practice_group: Any,
    practice_area: Any,
) -> list[str]:
    """Candidate lookup keys, most specific first.

    A pasted screen may carry a filename, or a firm and region, or nothing but
    the practice area, so several shapes are generated and the most specific
    match wins at lookup time.
    """
    keys: list[str] = []
    if filename:
        keys.append("file|" + normalise(filename))

    areas: list[str] = []
    if practice_area:
        areas = sorted(practice_area_variants(practice_area), key=len, reverse=True)
        if practice_group:
            areas += sorted(
                practice_area_variants(f"{practice_group} > {practice_area}"),
                key=len,
                reverse=True,
            )

    if firm and region:
        keys += [f"fra|{normalise(firm)}|{normalise(region)}|{a}" for a in areas]
    if region:
        keys += [f"ra|{normalise(region)}|{a}" for a in areas]
    keys += [f"a|{a}" for a in areas]
    return keys


def read_pasted_order(text: str, vocabulary: "Vocabulary | None" = None) -> PastedOrder:
    """Split a paste into rows and columns and work out which column to use.

    Copying a web table yields tab-separated columns; copying one column yields
    plain lines. Both arrive here. Nothing is matched yet - this only prepares
    the paste so the column can be confirmed before it is used.
    """
    rows: list[list[str]] = []
    for line in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not line.strip():
            continue
        rows.append([cell.strip() for cell in line.split("\t")])

    if not rows:
        return PastedOrder()

    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]

    headers = None
    named: dict[str, int] = {}
    first = [re.sub(r"\s+", " ", cell).strip().lower() for cell in rows[0]]
    for name, hints in ORDER_COLUMN_HINTS.items():
        for index, label in enumerate(first):
            if label and label in hints and name not in named:
                named[name] = index
                break
    # Only an identifying heading proves this is a header row. Portal data
    # contains bare values like "firm" (a submission type) that would otherwise
    # look like a "Firm" heading and swallow the first row of real data.
    if "practice_area" in named or "filename" in named:
        headers = rows[0]
        rows = rows[1:]
    else:
        named = {}

    vocabulary = vocabulary or Vocabulary()
    checked = bool(vocabulary)
    scores = [
        sum(1 for row in rows if vocabulary.recognises(row[index])) if checked else 0
        for index in range(width)
    ]

    # A heading tells us which column was meant; the exports tell us whether it
    # actually holds practice areas. Prefer the heading, but only if its column
    # is recognised - otherwise trust the content over the label.
    suggested = named.get("practice_area", named.get("filename"))
    if checked and max(scores, default=0) > 0:
        if suggested is None or scores[suggested] == 0:
            suggested = scores.index(max(scores))
    if suggested is None:
        suggested = 0

    labels = []
    for index in range(width):
        if headers and index < len(headers) and headers[index]:
            label = f"{index + 1}. {headers[index]}"
        else:
            sample = next((r[index] for r in rows if r[index]), "")
            label = f"Column {index + 1}" + (f" — {sample[:40]}" if sample else "")
        if checked:
            label += f"  ({scores[index]} of {len(rows)} recognised)"
        labels.append(label)

    return PastedOrder(
        rows=rows,
        headers=headers,
        column_labels=labels,
        named=named,
        suggested=min(suggested, width - 1) if width else 0,
        scores=scores,
        checked=checked,
        _vocabulary=vocabulary,
    )


def build_portal_order(pasted: PastedOrder, column: int | None = None) -> PortalOrder:
    """Turn a prepared paste into a lookup, using the chosen column."""
    order = PortalOrder()
    if not pasted.rows:
        raise ValueError("Nothing was pasted, so there is no order to apply.")

    index = pasted.suggested if column is None else column
    if not 0 <= index < pasted.width:
        raise ValueError(f"Column {index + 1} is not in the pasted text.")
    order.column_label = pasted.column_labels[index]

    is_filename = pasted.named.get("filename") == index

    def cell(row: list[str], name: str) -> str | None:
        position = pasted.named.get(name)
        if position is None or position == index:
            return None
        return row[position] or None

    for row in pasted.rows:
        value = row[index]
        if not value:
            continue
        keys = _order_keys(
            value if is_filename else cell(row, "filename"),
            cell(row, "firm"),
            cell(row, "region"),
            cell(row, "practice_group"),
            None if is_filename else value,
        )
        if not keys:
            continue
        for key in keys:
            order.index.setdefault(key, order.size)
        order.values.append(value)
        order.size += 1

    if not order.size:
        raise ValueError(
            f"No usable values in {order.column_label!r}. Pick the column holding "
            "the practice areas."
        )

    if pasted.checked and pasted.recognised(index) == 0:
        best = max(range(pasted.width), key=pasted.recognised, default=index)
        hint = ""
        if pasted.recognised(best) > 0:
            hint = f" {pasted.column_labels[best]!r} looks like the right one."
        raise ValueError(
            f"None of the {order.size} values in {order.column_label!r} are practice "
            f"areas or file names the exports know about, so this column cannot "
            f"order the table.{hint}"
        )

    order.notes.append(
        f"Using {order.size} pasted row(s) from {order.column_label!r} for the row order."
    )
    return order


@dataclass(frozen=True)
class Vocabulary:
    """Every practice area and filename the uploaded exports know about.

    A pasted column can be checked against this the moment it is pasted, so a
    wrong column is caught there rather than showing up later as a table that
    silently failed to reorder.
    """

    areas: frozenset = frozenset()
    filenames: frozenset = frozenset()

    def recognises(self, value: Any) -> bool:
        text = normalise(value)
        if not text:
            return False
        if text in self.areas or text in self.filenames:
            return True
        # A pasted 'Group > Area' counts if the bare area is known.
        return " > " in text and text.split(" > ", 1)[1] in self.areas

    def __bool__(self) -> bool:
        return bool(self.areas or self.filenames)


def build_vocabulary(
    powerbi: PowerBIExport, portal: "PortalExport | None" = None
) -> Vocabulary:
    """Collect the known practice areas and filenames from both exports.

    The Power BI export alone covers the practice areas being checked; the
    Staff Portal export adds every other practice area and filename the portal
    knows, so a paste covering a firm's whole listing is recognised in full.
    """
    areas: set[str] = set()
    filenames: set[str] = set()

    for row in powerbi.rows:
        full = normalise(row.get("CONCATENATED_PRACTICE_AREA_NAME"))
        if not full:
            continue
        areas.add(full)
        if " > " in full:
            areas.add(full.split(" > ", 1)[1])

    if portal:
        # Every portal row, not only the ones with a document: a page row with
        # no submission still appears in a paste and must be recognised.
        for record in portal.records:
            areas |= practice_area_variants(_portal_practice_area(record))
            name = normalise(record.get("filename"))
            if name:
                filenames.add(name)

    return Vocabulary(frozenset(areas), frozenset(filenames))


# --- joining ----------------------------------------------------------------


@dataclass
class ResultRow:
    document: ParsedDocument
    filename: str  # the portal's canonical name, as shown in the table
    source_filename: str = ""  # the name of the file as uploaded, e.g. "... (1).docx"
    firm_ref: str = ""
    firm: str = ""
    region: str = ""
    practice_area: str = ""
    export_row: dict | None = None
    portal_record: dict | None = None
    reference_source: str = ""  # where the firm/region/practice area came from
    undecided_supersession: bool = False
    matched_by_code: bool = False  # portal row found by firm ref + table code, not name
    status: str = "matched"  # matched | unmatched | ambiguous | parse_error | duplicate | superseded
    issues: list[str] = field(default_factory=list)

    def recount(self, key: str) -> int | None:
        return self.document.metrics.get(key).value if self.document.metrics else None

    def export_value(self, key: str) -> int | None:
        if not self.export_row:
            return None
        return to_int(self.export_row.get(EXPORT_COLUMNS[key]))

    def differs(self, key: str) -> bool:
        """True only when both numbers exist and disagree.

        An unparsed recount is never treated as a difference: it is a request
        for a human to look, not a correction.
        """
        mine, theirs = self.recount(key), self.export_value(key)
        return mine is not None and theirs is not None and mine != theirs


def build_results(
    documents: list[ParsedDocument],
    powerbi: PowerBIExport,
    portal: PortalExport | None,
) -> list[ResultRow]:
    results: list[ResultRow] = []
    claimed: dict[str, ResultRow] = {}  # portal filename -> the row that holds it

    for doc in documents:
        # Two names matter: what the portal calls the document, which belongs in
        # the table, and what the file on disk is called, which is what to say
        # when talking about the file itself - a second download keeps its "(1)".
        row = ResultRow(
            document=doc, filename=doc.filename, source_filename=doc.filename
        )

        # The firm, region and practice area are reference data: they belong to
        # the portal record the document was downloaded from, not to whatever a
        # firm typed into the document. Only the six counts are recomputed.
        portal_record = portal.lookup(doc.filename) if portal else None
        if portal and portal_record is None:
            portal_record = portal.lookup_by_table_code(doc.filename)
            if portal_record is not None:
                row.matched_by_code = True
                row.issues.append(
                    "Not in the Staff Portal export under this name. Matched to "
                    f"{_portal_practice_area(portal_record)!r} by firm reference and "
                    "table code."
                )
        if portal_record:
            stripped, copy = strip_download_suffix(doc.filename)
            if copy is not None:
                row.issues.append(
                    f"This file is named as copy {copy} of {stripped!r}. If it was "
                    "downloaded twice, another document was probably missed - check "
                    "the page it came from."
                )
            row.portal_record = portal_record
            row.reference_source = "Staff Portal export"
            if not row.matched_by_code:
                # Matched by code, it is a different document from the one the
                # portal lists, so it keeps its own name.
                row.filename = portal_record["filename"]
            row.firm = str(portal_record.get("firm") or "")
            row.firm_ref = str(portal_record.get("firm_ref") or "")
            row.region = str(portal_record.get("region") or "")
            row.practice_area = _portal_practice_area(portal_record)

            # Keyed on the file itself, so two files matched to one portal row
            # by code are not mistaken for copies of each other.
            key = strip_download_suffix(doc.filename)[0].lower()
            held = claimed.get(key)
            if held is None:
                claimed[key] = row
            else:
                # Whichever file carries the browser's "(1)" suffix is the second
                # copy, regardless of which was read first.
                # Compare the names of the files on disk, not the portal's
                # canonical name - row.filename has already been replaced by it.
                _, incoming_copy = strip_download_suffix(row.source_filename)
                _, held_copy = strip_download_suffix(held.source_filename)
                if incoming_copy is None and held_copy is not None:
                    _mark_duplicate(held, row)
                    claimed[key] = row
                else:
                    _mark_duplicate(row, held)

            _crosscheck_portal(row, doc, portal_record)
        elif portal:
            row.issues.append(
                "Filename not found in the Staff Portal export; falling back to the "
                "details in the document itself."
            )

        # Fall back to the document only where the portal gave us nothing.
        if not row.firm:
            row.firm = doc.firm or ""
            row.reference_source = row.reference_source or "document"
        if not row.region:
            row.region = doc.region or ""
        if not row.practice_area:
            row.practice_area = doc.concatenated_practice_area or ""

        group = (portal_record or {}).get("practice_group") or doc.practice_group
        if is_elite(group) or is_elite_area(row.practice_area):
            row.status = "elite"
            row.issues.append("Elite submission - not recounted.")
            results.append(row)
            continue

        if not doc.ok:
            row.status = "parse_error"
            row.issues.append(doc.error or "Unknown parse error.")
            results.append(row)
            continue

        missing = [
            name
            for name, value in (
                ("firm", row.firm),
                ("region", row.region),
                ("practice area", row.practice_area),
            )
            if not value
        ]
        if missing:
            row.status = "unmatched"
            row.issues.append(
                f"No {' or '.join(missing)} for this document, so it could not be "
                "matched to an export row. The Staff Portal export did not supply "
                + ("it" if len(missing) == 1 else "them")
                + " (is this filename in the export?), and the document's own "
                "practice-area dropdown was left unselected."
            )
            results.append(row)
            continue

        candidates = powerbi.lookup(row.firm, row.region, row.practice_area)

        # The portal and the dashboard do not always spell a practice area the
        # same way - the portal's "Dispute resolution" is the dashboard's
        # "Dispute resolution: Litigation and arbitration". Where the portal's
        # spelling finds nothing, the document's own is worth a try.
        if not candidates and doc.concatenated_practice_area:
            alternative = doc.concatenated_practice_area
            if normalise(alternative) != normalise(row.practice_area):
                by_document = powerbi.lookup(row.firm, row.region, alternative)
                if len(by_document) == 1:
                    candidates = by_document
                    row.issues.append(
                        f"Matched on the practice area in the document "
                        f"({alternative!r}); the Staff Portal export calls it "
                        f"{row.practice_area!r}."
                    )

        if len(candidates) == 1:
            row.export_row = candidates[0]
            _backfill_from_export(row, candidates[0])
        elif len(candidates) > 1:
            row.status = "ambiguous"
            row.issues.append(
                f"{len(candidates)} export rows share this firm, region and practice "
                "area; not comparing against any of them."
            )
        else:
            row.status = "unmatched"
            near = powerbi.lookup_ignoring_region(row.firm, row.practice_area)
            relaxed = powerbi.lookup_relaxed_firm(row.firm, row.practice_area)
            if near:
                regions = sorted({str(r.get("COUNTRY_NAME") or "?") for r in near})
                row.issues.append(
                    "No export row for this firm, region and practice area. The same "
                    f"firm and practice area exist under region(s): {', '.join(regions)}."
                )
            elif relaxed:
                spellings = sorted({str(r.get("FIRM_NAME") or "?") for r in relaxed})
                row.issues.append(
                    f"No export row for firm {row.firm!r}. The export spells this firm "
                    f"{' / '.join(repr(s) for s in spellings)} for this practice area - "
                    "check the firm name before treating this as a genuine mismatch."
                )
            else:
                row.issues.append(
                    "No export row for this firm, region and practice area."
                )

        row.issues.extend(doc.warnings)
        results.append(row)

    return results


def _mark_duplicate(row: ResultRow, counted: ResultRow) -> None:
    """Set a row aside as a second copy of a submission already read."""
    row.status = "duplicate"
    row.issues.append(
        f"This is a second copy of the same submission, counted from "
        f"{counted.source_filename!r}. A duplicate download usually means a "
        "different document was never downloaded."
    )


def _portal_practice_area(record: dict) -> str:
    """How the portal names this submission.

    Usually 'Group > Area'. Where the portal gives the table a name of its own -
    two Transport > Travel submissions distinguished only as "Travel: personal
    injury" and "Travel: regulatory and commercial" - that third level is kept,
    both because the portal screen shows it and because without it the two rows
    would be indistinguishable in the output table.
    """
    group = clean_text(record.get("practice_group"))
    area = clean_text(record.get("practice_area"))
    table = clean_text(record.get("table_name"))

    parts = [p for p in (group, area) if p]
    if table and table.lower() != area.lower():
        parts.append(table)
    return " > ".join(parts)


def _created_at(record: dict | None) -> datetime:
    """The portal's Created At, as a comparable value.

    A missing or unreadable date sorts oldest, so a row with a known date is
    always preferred over one without.
    """
    if not record:
        return datetime.min
    value = record.get("created_at")
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if value:
        try:
            return datetime.fromisoformat(str(value).strip())
        except ValueError:
            return datetime.min
    return datetime.min


def created_at_text(record: dict | None) -> str:
    moment = _created_at(record)
    return "unknown date" if moment == datetime.min else moment.strftime("%Y-%m-%d")


def clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip() if value is not None else ""


def _backfill_from_export(row: ResultRow, export_row: dict) -> None:
    """Fill reference fields the portal could not supply from the matched row."""
    if not row.firm_ref:
        row.firm_ref = clean_text(export_row.get("FIRM_REF"))
    if not row.firm:
        row.firm = clean_text(export_row.get("FIRM_NAME"))
    if not row.region:
        row.region = clean_text(export_row.get("COUNTRY_NAME"))
    if not row.practice_area:
        row.practice_area = clean_text(export_row.get("CONCATENATED_PRACTICE_AREA_NAME"))


def _crosscheck_portal(row: ResultRow, doc: ParsedDocument, record: dict) -> None:
    """Warn when the document's dropdown disagrees with the portal export."""
    portal_area = record.get("practice_area")
    portal_group = record.get("practice_group")
    portal_region = record.get("region")

    if doc.practice_area and portal_area and normalise(doc.practice_area) != normalise(portal_area):
        row.issues.append(
            f"Practice area in the document ({doc.practice_area!r}) differs from the "
            f"Staff Portal export ({portal_area!r})."
        )
    if doc.practice_group and portal_group and normalise(doc.practice_group) != normalise(portal_group):
        row.issues.append(
            f"Practice group in the document ({doc.practice_group!r}) differs from the "
            f"Staff Portal export ({portal_group!r})."
        )
    if doc.region and portal_region and normalise(doc.region) != normalise(portal_region):
        row.issues.append(
            f"Region in the document ({doc.region!r}) differs from the Staff Portal "
            f"export ({portal_region!r})."
        )

    # A typed-in practice area may be a bare area or the full 'Group > Area',
    # so accept either shape before calling it a disagreement.
    typed = doc.typed_practice_area
    if typed and not doc.dropdown_raw and portal_area:
        accepted = {normalise(portal_area), normalise(_portal_practice_area(record))}
        if normalise(typed) not in accepted:
            row.issues.append(
                f"The practice-area dropdown was not used. The typed box reads "
                f"{typed!r}, which does not match the Staff Portal export "
                f"({portal_area!r}) - the export has been used."
            )


# --- corrections made by hand ------------------------------------------------


def override_key(row: "ResultRow", metric: str) -> tuple[str, str]:
    """How a hand-entered value is filed, stable across reruns."""
    return (row.filename.lower(), metric)


def final_value(
    row: "ResultRow", metric: str, overrides: dict | None = None
) -> int | None:
    """What the cell should say: a hand-entered value if there is one.

    Everything downstream - the table, the highlighting, the export - reads the
    value through here, so a figure corrected by hand is highlighted on the same
    terms as one the parser produced.
    """
    if overrides:
        entered = overrides.get(override_key(row, metric), _MISSING)
        if entered is not _MISSING:
            return entered
    return row.recount(metric)


_MISSING = object()


def final_differs(
    row: "ResultRow", metric: str, overrides: dict | None = None
) -> bool:
    """True when the cell disagrees with the export, whoever decided it.

    A hand correction back to the export's own figure is no longer a
    difference, and must stop being highlighted.
    """
    mine = final_value(row, metric, overrides)
    theirs = row.export_value(metric)
    return mine is not None and theirs is not None and mine != theirs


def is_checked(row: "ResultRow", checked: set | None) -> bool:
    return bool(checked and row.filename.lower() in checked)


def display_value(
    row: "ResultRow",
    metric: str,
    overrides: dict | None = None,
    checked: set | None = None,
) -> tuple[int | None, bool]:
    """The figure to show, and whether it is still unverified.

    Where the parser could not establish a count, the best figure available is
    the export's own, so that is what the cell carries. It is amber until
    somebody has looked: a cell reading a phrase alone gave no baseline, so
    counting 4 by hand told you nothing about whether 4 was already there.

    Ticking the row is that look. An amber cell on a ticked row has been seen
    and left as it stands, so it stops being amber.
    """
    established = final_value(row, metric, overrides)
    if established is not None:
        return established, False
    return row.export_value(metric), not is_checked(row, checked)


def needs_attention(
    row: "ResultRow", metric: str, checked: set | None = None
) -> bool:
    """True when the parser counted this but something warrants a look.

    Ticking the row answers that, so the flag clears with it.
    """
    if is_checked(row, checked):
        return False
    found = row.document.metrics.get(metric)
    return bool(found and (found.numbering_gap or found.needs_check))


def review_rows(
    results: list["ResultRow"], overrides: dict | None = None, everything: bool = False
) -> list[dict]:
    """One entry per cell worth a human decision, for the review editor."""
    from parsing import METRIC_COLUMNS, METRIC_KEYS, METRIC_LABELS

    out = []
    for row in results:
        for metric in METRIC_KEYS:
            recount = row.recount(metric)
            value = final_value(row, metric, overrides)
            flags = []
            if recount is None:
                flags.append("count by hand")
            if needs_attention(row, metric):
                flags.append("check the numbering")
            if final_differs(row, metric, overrides):
                flags.append("differs from export")
            if not flags and not everything:
                continue
            out.append(
                {
                    "FILE NAME": row.filename,
                    "PRACTICE AREA": row.practice_area,
                    "Metric": METRIC_COLUMNS[metric],
                    "Power BI": row.export_value(metric),
                    "App counted": recount,
                    "Your value": value,
                    "Why": ", ".join(flags) or "-",
                    "_key": override_key(row, metric),
                    "_metric": metric,
                    "_label": METRIC_LABELS[metric],
                }
            )
    return out


# --- presentation -----------------------------------------------------------

BASE_COLUMNS = ["FIRMREF", "FIRM", "REGION", "PRACTICE AREA", "FILE NAME"]
METRIC_TABLE_COLUMNS = [METRIC_COLUMNS[k] for k in METRIC_KEYS]
TABLE_COLUMNS = BASE_COLUMNS + METRIC_TABLE_COLUMNS

SORT_MODES = {
    "Staff Portal order (pasted)": "portal_paste",
    "Power BI export order": "powerbi",
    "Staff Portal export order": "portal",
    "Practice group, then practice area (portal hierarchy)": "hierarchy",
    "Region, then practice area (A-Z)": "alpha",
    "Order the documents were added": "upload",
}


def sort_results(
    results: list[ResultRow],
    mode: str = "powerbi",
    portal_order: "PortalOrder | None" = None,
) -> list[ResultRow]:
    """Order the rows. Rows with no key for the chosen mode go last, in upload order."""
    if mode == "upload":
        return list(results)
    if mode == "portal_paste" and portal_order is None:
        raise ValueError("Staff Portal order was requested but no order sheet was given.")

    def key(item):
        index, row = item
        if mode == "portal_paste":
            value = portal_order.position(row)
        elif mode == "powerbi":
            value = row.export_row.get("_row_order") if row.export_row else None
        elif mode == "portal":
            value = row.portal_record.get("_row_order") if row.portal_record else None
        elif mode == "hierarchy":
            record = row.portal_record or {}
            group, area = record.get("practice_group_id"), record.get("practice_area_id")
            value = (str(group), str(area)) if group is not None else None
        elif mode == "alpha":
            value = (normalise(row.region), normalise(row.practice_area))
            if not row.region and not row.practice_area:
                value = None
        else:
            raise ValueError(f"Unknown sort mode: {mode!r}")
        # (missing-key flag, key, original position) keeps the sort stable.
        return (value is None, value if value is not None else 0, index)

    return [row for _, row in sorted(enumerate(results), key=key)]


REVIEW_COLUMN = "NEEDS A LOOK"


def review_note(
    row: "ResultRow", overrides: dict | None = None, checked: set | None = None
) -> str:
    """One line saying what on this row wants checking, and why."""
    from parsing import METRIC_COLUMNS, METRIC_KEYS

    if is_checked(row, checked):
        return ""

    parts = []
    for metric in METRIC_KEYS:
        name = METRIC_COLUMNS[metric]
        if final_value(row, metric, overrides) is None:
            original = row.export_value(metric)
            baseline = f", export says {original}" if original is not None else ""
            parts.append(f"{name}: not counted{baseline}")
        elif needs_attention(row, metric, checked):
            found = row.document.metrics.get(metric)
            why = "numbering does not add up"
            if found and found.needs_check:
                why = "a nomination nobody could classify"
            parts.append(f"{name}: {why}")
    return "; ".join(parts)


def results_to_frame(
    results: list[ResultRow],
    show_original: bool = False,
    overrides: dict | None = None,
    numeric: bool = False,
    checked: set | None = None,
) -> pd.DataFrame:
    """One row per firm/practice area, recounted values in the metric columns.

    The first five columns are reference data copied from the exports; only the
    six count columns can differ from what the export said.
    """
    records = []
    for row in results:
        record = {
            "FIRMREF": row.firm_ref,
            "FIRM": row.firm,
            "REGION": row.region,
            "PRACTICE AREA": row.practice_area,
            "FILE NAME": row.filename,
        }
        for key in METRIC_KEYS:
            value, unverified = display_value(row, key, overrides, checked)
            if numeric:
                record[METRIC_COLUMNS[key]] = value
            else:
                text = UNPARSED if value is None else str(value)
                if show_original and final_differs(row, key, overrides):
                    text = f"{value} (was {row.export_value(key)})"
                record[METRIC_COLUMNS[key]] = text
        if show_original:
            for key in METRIC_KEYS:
                record[f"was {METRIC_COLUMNS[key]}"] = row.export_value(key)
        records.append(record)

    columns = list(TABLE_COLUMNS)
    if show_original:
        columns += [f"was {METRIC_COLUMNS[k]}" for k in METRIC_KEYS]
    frame = pd.DataFrame(records, columns=columns)
    if numeric:
        for key in METRIC_KEYS:
            frame[METRIC_COLUMNS[key]] = frame[METRIC_COLUMNS[key]].astype("Int64")
    return frame


def collapse_superseded(results: list[ResultRow]) -> list[ResultRow]:
    """Keep one row per firm, region and practice area: the newest submission.

    A firm often has more than one document for the same slot - an earlier
    submission and a replacement. The portal shows one row; the batch may
    contain both files. The newest by the portal's Created At is the one to
    count, and the others are set aside rather than left to duplicate the row.

    Where the dates cannot decide it, nothing is guessed: a row is still kept so
    the practice area is not lost, but it is flagged for a human to settle.
    """
    groups: dict[tuple[str, str, str], list[ResultRow]] = {}
    for row in results:
        if row.status != "matched" or not (row.firm and row.region and row.practice_area):
            continue
        key = (
            normalise(row.firm),
            normalise(row.region),
            normalise(row.practice_area),
            _slot_person(row),
        )
        groups.setdefault(key, []).append(row)

    for rows in groups.values():
        if len(rows) < 2:
            continue

        ranked = sorted(rows, key=lambda r: _created_at(r.portal_record), reverse=True)
        winner, rest = ranked[0], ranked[1:]
        winner_date = _created_at(winner.portal_record)

        same_name = len({r.filename.lower() for r in rows}) == 1
        undecided = [r for r in rest if _created_at(r.portal_record) == winner_date]

        for row in rest:
            row.status = "superseded"
            row.issues.append(
                f"Same firm, region and practice area as {winner.filename!r} "
                f"({created_at_text(winner.portal_record)}); this one is dated "
                f"{created_at_text(row.portal_record)}. The newer submission is "
                "counted and this row is left out of the table."
            )

        if undecided and not same_name:
            note = (
                f"{len(rows)} submissions share this firm, region and practice area "
                f"and carry the same date ({created_at_text(winner.portal_record)}), "
                f"so which is newer cannot be told from the portal export. "
                f"{winner.filename!r} was counted - check this one by hand."
            )
            winner.issues.append(note)
            winner.undecided_supersession = True
        elif same_name:
            winner.issues.append(
                f"The same document was uploaded {len(rows)} times under different "
                "names; counted once."
            )

    return [r for r in results if r.status != "superseded"]


def _slot_person(row: ResultRow) -> str:
    """Who a lawyer submission is for, so two lawyers in one table stay apart.

    A firm can nominate several lawyers for the same table, each with their own
    document; those are not older and newer versions of one submission. A file
    matched by code may be a lawyer the export does not name, so its own file
    name stands in for the person.
    """
    if row.matched_by_code:
        return "file:" + normalise(strip_download_suffix(row.source_filename)[0])
    record = row.portal_record or {}
    if normalise(record.get("submission_type")) == "lawyer":
        return "person:" + normalise(record.get("person"))
    return ""


@dataclass
class PageRow:
    """One row of the pasted Staff Portal page, and what it maps to."""

    value: str
    position: int = 0  # 1-based line number in the pasted order
    records: list[dict] = field(default_factory=list)
    uploaded: "ResultRow | None" = None
    reached: bool = False  # within the span this batch got down to
    repeat_of: int | None = None  # line whose export row this one repeats

    @property
    def has_document(self) -> bool:
        return any(r["has_document"] for r in self.records)

    @property
    def candidates(self) -> list[str]:
        """Documents the portal holds for this row, newest spelling first."""
        seen: list[str] = []
        for record in self.records:
            name = record.get("filename")
            if name and name not in seen:
                seen.append(name)
        return seen

    @property
    def region(self) -> str:
        regions = {clean_text(r.get("region")) for r in self.records}
        regions.discard("")
        return ", ".join(sorted(regions))

    @property
    def is_elite(self) -> bool:
        return any(is_elite(r.get("practice_group")) for r in self.records)

    @property
    def status(self) -> str:
        if not self.records:
            return "unrecognised"
        if self.is_elite:
            return "elite"
        if not self.has_document:
            return "no document"
        if self.uploaded:
            # A file that could not be read has not been checked, so it must not
            # read as a tick. Nor must an older submission that is not in the
            # table.
            if self.uploaded.status == "parse_error":
                return "could not read"
            if self.uploaded.status == "superseded":
                return "older version"
            return "uploaded"
        if self.repeat_of is not None and self.reached:
            # The export gives this line the same row as an earlier one, so
            # which document it means can only be read off the page itself.
            return "check page"
        return "missing" if self.reached else "not reached"


@dataclass
class Reconciliation:
    """What the pasted portal page said should be here, against what arrived.

    A page is often worked in parts - 50 documents out of 114 - so the whole
    page is not expected at once. `attempted` is how far down the downloadable
    rows this batch reached; only gaps above that line are missing, and the rest
    are simply not started yet.
    """

    rows: list[PageRow] = field(default_factory=list)
    duplicates: list[ResultRow] = field(default_factory=list)
    not_on_page: list[ResultRow] = field(default_factory=list)
    attempted: int = 0

    @property
    def page_size(self) -> int:
        return len(self.rows)

    @property
    def expected(self) -> list[PageRow]:
        """Page rows that do have a submission document to download."""
        return [r for r in self.rows if r.has_document and not r.is_elite]

    @property
    def elite(self) -> list[PageRow]:
        """Elite rows, which are not recounted and so not expected."""
        return [r for r in self.rows if r.status == "elite"]

    @property
    def in_batch(self) -> list[PageRow]:
        """The span of downloadable rows this batch got down to."""
        return self.expected[: self.attempted]

    @property
    def without_document(self) -> list[PageRow]:
        return [r for r in self.rows if r.status == "no document"]

    @property
    def missing(self) -> list[PageRow]:
        return [r for r in self.rows if r.status == "missing"]

    @property
    def not_reached(self) -> list[PageRow]:
        return [r for r in self.rows if r.status == "not reached"]

    @property
    def accounted(self) -> list[PageRow]:
        """Rows whose document is in the batch, counted or set aside as older."""
        return [r for r in self.rows if r.status in ("uploaded", "older version")]

    @property
    def distinct_documents(self) -> int:
        """How many separate files the rows with a document actually point at.

        The portal lists some rows twice against the same document, so there are
        fewer documents to download than there are rows to check. A line the
        export merely repeats is not that: it is most likely a document of its
        own that the export failed to name, so it counts as one.
        """
        names = {
            (r.candidates[0] or "").lower()
            for r in self.expected
            if r.candidates and r.repeat_of is None
        }
        names.discard("")
        return len(names) + len(self.repeats)

    @property
    def repeats(self) -> list[PageRow]:
        """Lines whose export row repeats an earlier line's."""
        return [r for r in self.expected if r.repeat_of is not None]

    @property
    def unclear(self) -> list[PageRow]:
        """Repeated lines no uploaded file could be placed against."""
        return [r for r in self.rows if r.status == "check page"]

    @property
    def not_downloaded(self) -> list[str]:
        """Documents the page names that no uploaded file is, by file name."""
        names: list[str] = []
        for row in self.rows:
            if row.status in ("missing", "not reached"):
                for name in row.candidates[:1]:
                    if name not in names:
                        names.append(name)
        return names

    @property
    def shared_rows(self) -> int:
        """Rows whose document is already listed against another row."""
        return len(self.expected) - self.distinct_documents

    @property
    def unreadable(self) -> list[PageRow]:
        """Rows whose file was uploaded but could not be opened."""
        return [r for r in self.rows if r.status == "could not read"]

    @property
    def unrecognised(self) -> list[PageRow]:
        return [r for r in self.rows if not r.records]

    @property
    def accounted_for(self) -> bool:
        return not (
            self.missing
            or self.not_reached
            or self.duplicates
            or self.unreadable
            or self.unclear
        )


def _choose_record(
    pool: list[dict],
    uploaded_by_name: dict[str, list["ResultRow"]],
    placed: set[int],
) -> tuple[dict | None, "ResultRow | None"]:
    """Pick which portal row a pasted line refers to, and the file covering it.

    Preference: a row whose document was uploaded, then any row that has a
    document, then a row with none. Where a practice area has both a row with a
    document and a row without, assuming the one with a document risks a false
    "missing" warning; assuming the other risks a document never being checked.
    The false alarm is the safer error.

    Each uploaded file covers one line. Several files can share a portal row -
    a lawyer the export lists by name and another matched to it by code - so
    the next file not yet placed is taken.
    """
    for record in pool:
        for match in uploaded_by_name.get((record.get("filename") or "").lower(), []):
            if id(match) not in placed:
                return record, match
    for record in pool:
        if record["has_document"]:
            return record, None
    return (pool[0], None) if pool else (None, None)


def describe_record(record: dict) -> str:
    area = _portal_practice_area(record)
    region = clean_text(record.get("region"))
    return f"{area} ({region})" if region else area


def reconcile(
    results: list[ResultRow], portal: PortalExport, portal_order: PortalOrder
) -> Reconciliation:
    """Check the uploaded batch against the portal page that was pasted in.

    Worked per pasted row rather than per portal record, because one firm and
    practice area can hold several documents in the portal - an older `EC...`
    submission and a current one, sometimes in the same region. What matters is
    whether *a* document for that row was uploaded, not which of them.

    A row with no submission document linked has nothing to check, and must
    never be reported as a missing file.
    """
    report = Reconciliation()

    firm_keys: set[str] = set()
    region_keys: set[str] = set()
    for row in results:
        if row.firm_ref:
            firm_keys.add("ref:" + normalise(row.firm_ref))
        if row.firm:
            firm_keys.add("name:" + firm_key(row.firm))
        if row.region:
            region_keys.add(normalise(row.region))

    # Files matched by name come first, so the file the export names covers
    # the first line for its row and a file matched by code the repeat.
    uploaded_by_name: dict[str, list[ResultRow]] = {}
    for row in sorted(results, key=lambda r: r.matched_by_code):
        if row.portal_record and row.status not in ("duplicate", "elite"):
            uploaded_by_name.setdefault(row.portal_record["filename"].lower(), []).append(row)
    placed: set[int] = set()

    # One pasted line is one row on the portal page, so each portal record is
    # used once. The portal lists some practice areas twice for a firm - once
    # with a submission document and once without - and pasting the area twice
    # must account for both rows.
    consumed: set[int] = set()
    # Where each portal row was first used. The export sometimes repeats a row
    # verbatim, ID and all, where the page has two lines - two lawyers in one
    # table, say - so a second line landing on the same row is not a shared
    # document but one the export failed to name.
    first_line: dict[Any, int] = {}

    for value in portal_order.values:
        records = portal.records_for_area(value, firm_keys or None)
        # Narrow to the regions actually being checked, but only if that leaves
        # something - a page may legitimately cover a region with no uploads.
        if region_keys:
            narrowed = [r for r in records if normalise(r.get("region")) in region_keys]
            if narrowed:
                records = narrowed

        pool = [r for r in records if id(r) not in consumed] or records
        # A line naming an area the firm has both as an ordinary and an Elite
        # table means the ordinary one first; Elite rows are not counted.
        pool.sort(key=lambda r: is_elite(r.get("practice_group")))
        chosen, uploaded = _choose_record(pool, uploaded_by_name, placed)
        position = len(report.rows) + 1
        repeat_of = None
        if chosen is not None:
            consumed.add(id(chosen))
            identity = chosen.get("portal_id") or id(chosen)
            repeat_of = first_line.get(identity)
            first_line.setdefault(identity, position)
        if uploaded is not None:
            placed.add(id(uploaded))

        report.rows.append(
            PageRow(
                value=value,
                position=position,
                records=[chosen] if chosen else [],
                uploaded=uploaded,
                repeat_of=repeat_of,
            )
        )

    report.duplicates = [r for r in results if r.status == "duplicate"]

    accounted = {id(r.uploaded) for r in report.rows if r.uploaded}
    report.not_on_page = [
        row
        for row in results
        if row.status not in ("duplicate", "elite") and id(row) not in accounted
    ]

    # Documents are worked top to bottom, so the number of files given is how
    # many downloadable rows this batch covers - no more. An earlier version
    # stretched this span down to the furthest row that happened to match, which
    # turned every row in between into a false "missing".
    expected = report.expected
    counted = [r for r in results if r.status != "elite"]
    report.attempted = min(len(counted), len(expected))
    for page_row in expected[: report.attempted]:
        page_row.reached = True

    return report


def unordered_rows(results: list[ResultRow], portal_order: PortalOrder) -> list[ResultRow]:
    """Rows the pasted Staff Portal order does not mention."""
    return [r for r in results if portal_order.position(r) is None]


def difference_records(
    results: list[ResultRow], overrides: dict | None = None
) -> list[dict]:
    """Flat list of every disagreement, for the summary and the Changes sheet."""
    from parsing import METRIC_LABELS

    out = []
    for row in results:
        for key in METRIC_KEYS:
            if not final_differs(row, key, overrides):
                continue
            value = final_value(row, key, overrides)
            out.append(
                {
                    "Firm": row.firm,
                    "Region": row.region,
                    "Practice area": row.practice_area,
                    "Filename": row.filename,
                    "Metric": METRIC_LABELS[key],
                    "Export value": row.export_value(key),
                    "Recounted value": value,
                    "Decided by": (
                        "you" if value != row.recount(key) else "the app"
                    ),
                }
            )
    return out
