"""Parse Legal 500 submission documents (.docx) and recount their metrics.

Pure logic: no Streamlit, no file dialogs, no network. Everything here can be
driven from a test with a path or a file-like object.

The anchoring strategy is documented in README.md. In short: the first cell of
a table's first row carries a self-describing label, and that label is the only
reliable way to identify a section. Headings lie (a document in the sample set
files associate nominations under a "leading partners" heading) and the trailing
numbers in labels repeat, so we count tables rather than trusting indices.
"""

from __future__ import annotations

import colorsys
import re
from dataclasses import dataclass, field
from typing import Any, Iterator

from io import BytesIO

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from docx_repair import repair as repair_package

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# --- metric keys, in the order they are shown to the user -------------------

METRIC_KEYS = [
    "matters",
    "active_clients",
    "new_clients",
    "lead_partners",
    "next_gen",
    "associates",
]

EXPORT_COLUMNS = {
    "matters": "NUM_MATTERS",
    "active_clients": "NUM_ACTIVE_CLIENTS",
    "new_clients": "NUM_NEW_CLIENTS",
    "lead_partners": "NUM_LEAD_PARTNERS",
    "next_gen": "NUM_NEXT_GEN",
    "associates": "NUM_ASSOCIATES",
}

# Short headers for the results table, as used in the master sheet.
METRIC_COLUMNS = {
    "matters": "MATTERS",
    "active_clients": "CLIENTS",
    "new_clients": "NEW",
    "lead_partners": "LPs",
    "next_gen": "NGs",
    "associates": "LAs",
}

# Readable names for the parse log, the Changes sheet and error messages.
METRIC_LABELS = {
    "matters": "Work matters",
    "active_clients": "Active clients",
    "new_clients": "New active clients",
    "lead_partners": "Leading partners",
    "next_gen": "Next generation",
    "associates": "Leading associates",
}

# --- table label patterns ---------------------------------------------------

RE_CLIENT_TABLE = re.compile(r"^active key clients", re.I)
# Most firms number the matter boxes ("Publishable matter 7"), but some
# templates leave them unnumbered. The number is therefore optional, and an
# unnumbered label has to end there: the work highlights section is introduced
# by a heading reading "Publishable matter summary", which is not a matter.
# Firms punctuate the label freely: "Publishable - Matter 1", "Publishable
# matter - 4", "Non-publishable matter7", "Non Publishable matter 8". A dash or
# colon may sit on either side of "matter", and the space before the number
# may be missing.
#
# Some firms word the non-publishable box differently: "Confidential matter 5"
# (in the red, non-publishable colour) and "Non-published matter 9". Both are
# non-publishable. A typed slip such as "Son-publishable matter 3" is left to
# the colour check in count_matters, which names it, rather than guessed here.
_MATTER_LABEL = (
    r"^(?:(?P<non>non\s*-?\s*)?publish(?:able|ed)|(?P<confidential>confidential))"
    r"\s*[-–—:]?\s*#?\s*matter"
)
RE_MATTER_TABLE = re.compile(
    _MATTER_LABEL + r"(\s*[-–—:#]?\s*\d+|\s*$)", re.I
)
# A nomination label is "<role>: <descriptor> <number>". Firms invent their own
# descriptors - "leading counsel", "to become next generation partner" - so the
# role before the colon is the reliable part, and the descriptor only has to
# separate the two partner categories. Whitespace before the colon happens too.
RE_NOMINATION_LABEL = re.compile(
    # The separator is usually a colon, but firms also use a dash.
    # "Counsel" is a role of its own, and the form puts it in the leading
    # associates section: "Note that this section can include counsel."
    # The box number sits after the descriptor ("Partner: leading individual 1")
    # or between the role and the colon ("Partner 1: leading individual").
    # The role is written plural by some firms ("Associates: leading associate 1"),
    # which is the section heading's wording carried onto every box.
    r"^(?P<role>partners?|associates?|counsels?)\s*\d*\s*"
    r"[:\u2013\u2014-]\s*(?P<descriptor>\S.*)$",
    re.I,
)

# Some firms drop the role prefix and label the box with the category itself:
# "Next Generation Partner 2", "Rising Star 1". Matched only as a whole label,
# so prose mentioning these words is not mistaken for one.
RE_BARE_NEXT_GEN = re.compile(r"^next\s*gen(eration)?(\s+partner)?s?\s*\d*$", re.I)
RE_BARE_LEAD_PARTNER = re.compile(
    r"^(leading|senior)\s+(partner|individual)s?\s*\d*$", re.I
)
RE_BARE_ASSOCIATE = re.compile(
    r"^(rising\s+stars?|leading\s+(associates?|counsels?))\s*\d*$", re.I
)
# What a nomination description looks like, used only to tell a plural-role
# nomination box from a summary row that happens to share its shape.
RE_NOMINATION_DESCRIPTOR = re.compile(
    r"leading|senior|individual|next\s*gen|rising\s+star|counsel|associate|partner",
    re.I,
)
RE_NEXT_GEN_WORDS = re.compile(r"next\s*gen(eration)?", re.I)
RE_LEADING_WORDS = re.compile(r"leading|rising\s+star|individual|senior", re.I)

UNCLASSIFIED = "unclassified"

# Kept for the older, stricter matching used by reporting tools.
RE_LEAD_PARTNER = re.compile(r"^partner\s*:\s*leading\s+(individual|partner)\b", re.I)
RE_NEXT_GEN = re.compile(r"^partner\s*:\s*next\s+gen(eration)?\b", re.I)
RE_ASSOCIATE = re.compile(
    r"^associate\s*:\s*(rising\s+star|leading\s+associate)\b", re.I
)

NOMINATION_PATTERNS = {
    "lead_partners": RE_LEAD_PARTNER,
    "next_gen": RE_NEXT_GEN,
    "associates": RE_ASSOCIATE,
}


def classify_nomination(label: str) -> str | None:
    """Which category a nomination label belongs to, or None if it is not one.

    An "Associate:" label is an associate whatever the firm calls the role.
    A "Partner:" label has to be sorted into leading or next generation, and
    where the descriptor says neither the label is returned as unclassified
    rather than guessed at - the two categories are not interchangeable.
    """
    match = RE_NOMINATION_LABEL.match(label)
    if not match:
        # A label naming the category on its own, with no role in front.
        if RE_BARE_NEXT_GEN.match(label):
            return "next_gen"
        if RE_BARE_ASSOCIATE.match(label):
            return "associates"
        if RE_BARE_LEAD_PARTNER.match(label):
            return "lead_partners"
        return None
    descriptor = match.group("descriptor")
    role_text = match.group("role").lower()
    # Singular or plural, it is the same role. The plural is the unusual form
    # and also how a summary row reads ("Partners: two"), so it is accepted
    # only where the description names a nomination outright. The singular is
    # the established shape of a nomination box and is still taken on trust,
    # so an invented description is reported rather than dropped.
    if role_text.endswith("s") and not RE_NOMINATION_DESCRIPTOR.search(descriptor):
        return None
    role = role_text.rstrip("s")
    if role in {"associate", "counsel"}:
        return "associates"
    if RE_NEXT_GEN_WORDS.search(descriptor):
        return "next_gen"
    if RE_LEADING_WORDS.search(descriptor):
        return "lead_partners"
    return UNCLASSIFIED

# Section headings, used only as a safety net: if a heading for a nomination
# category is present but no table matched it, we refuse to report 0.
NOMINATION_HEADINGS = {
    "lead_partners": re.compile(r"your team.*:\s*leading\s+partners", re.I),
    "next_gen": re.compile(r"your team.*:\s*next\s+gen(eration)?", re.I),
    "associates": re.compile(r"your team.*:\s*leading\s+associates", re.I),
}

# Rows inside client tables that are not clients.
RE_CLIENT_BOILERPLATE = re.compile(r"to add more clients|insert row below", re.I)
RE_CLIENT_HEADER_ROW = re.compile(
    r"^(client name|name of client|active key clients\b.*)$", re.I
)

# New-client answers. Firms qualify them freely - "Yes (for this work type)",
# "No (existing client)" - so the answer is read by its leading word. The bare
# forms are kept separately so a qualified answer can be reported as such.
#
# Latin American firms answer in their own language: "Sí" or "Si" (Spanish)
# and "Sim" (Portuguese) for yes, "Não" (Portuguese) for no; Spanish "No" is
# already no. "Ye" is a typed slip of "Yes"; not being a bare answer, it is
# named in the note like any qualified one.
RE_AFFIRMATIVE = re.compile(r"^(y(es?)?|s[ií]m?)\b", re.I)
RE_NEGATIVE = re.compile(r"^(no?|n[aã]o)\b", re.I)
RE_NEW_WORD = re.compile(r"^new\b", re.I)
RE_EXISTING_WORD = re.compile(r"^existing\b", re.I)

# "N/A" starts with an N but does not mean "no" - it means nobody answered.
RE_NOT_APPLICABLE = re.compile(r"^n\s*/\s*a\b|^n\.?a\.?$|^not applicable\b", re.I)

RE_BARE_AFFIRMATIVE = re.compile(r"^(y(es)?|s[ií]m?)$", re.I)
RE_BARE_NEGATIVE = re.compile(r"^(no?|n[aã]o)$", re.I)
RE_BARE_NEW = re.compile(r"^new$", re.I)
RE_BARE_EXISTING = re.compile(r"^existing$", re.I)

PLACEHOLDER_DROPDOWN = re.compile(r"^(choose an item|click here to enter text)\.?$", re.I)

# The form invites firms to skip the dropdown: "OR If you have an earlier
# version of Word, type in this box". Plenty do, which is why the dropdown alone
# is not enough to know a document's practice area.
RE_TYPED_BOX = re.compile(r"type in this box", re.I)
RE_TYPED_BOILERPLATE = re.compile(
    r"^(choose one from|contact details|what is the team|head\(s\) of team|"
    r"either select|practice area|country|firm name)\b",
    re.I,
)
TYPED_MAX_LENGTH = 200

# Shown in place of a number the parser could not establish. Worded as an
# instruction rather than a state: the checker needs to count this one by hand.
UNPARSED = "double-check manually"


# --- small helpers ----------------------------------------------------------


def clean(text: str | None) -> str:
    """Collapse whitespace and trim. Word is full of stray runs and newlines."""
    return re.sub(r"\s+", " ", text or "").strip()


def row_cells(row) -> list:
    """Row cells with horizontally-merged duplicates removed.

    python-docx repeats a merged cell once per grid column it spans, so a
    3-column table with a full-width label row returns the same text 3 times.
    De-duplicate on the underlying <w:tc> element rather than on text, so two
    genuinely distinct blank cells are not collapsed into one.
    """
    seen: set[int] = set()
    out = []
    for cell in row.cells:
        key = id(cell._tc)
        if key in seen:
            continue
        seen.add(key)
        out.append(cell)
    return out


def row_text(row) -> list[str]:
    return [clean(c.text) for c in row_cells(row)]


def iter_blocks(document) -> Iterator[tuple[str, Any]]:
    """Yield ('p', Paragraph) / ('tbl', Table) in true document order.

    Walks more than the top level of the body, because a table can be hidden
    from a plain walk in two ways that both occur in real submissions:

    * wrapped in a block-level content control (``w:sdt``), which Word inserts
      around protected or templated regions;
    * nested inside another table's cell.

    Either would make a matter or nomination table silently invisible - counted
    as absent rather than as an empty template, with nothing on screen to say a
    table had been passed over.
    """
    yield from _iter_blocks(document.element.body, document)


def _iter_blocks(parent, document) -> Iterator[tuple[str, Any]]:
    """Walk one level, descending into wrappers and nested tables.

    No "already seen" bookkeeping: the document is a tree, so every table has
    exactly one parent and is reached once. An earlier version tracked visited
    elements by ``id()``, which is unsound here - lxml builds its Python proxies
    on demand, so a collected proxy's ``id()`` can be handed to a different
    element, and that element would then be skipped as a duplicate. It only
    showed up when the walk was consumed lazily rather than into a list.
    """
    for child in parent.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield "p", Paragraph(child, document)
        elif tag == "tbl":
            table = Table(child, document)
            yield "tbl", table
            # Tables nested in this one's cells, in document order after it.
            for row in table.rows:
                for cell in row_cells(row):
                    yield from _iter_blocks(cell._tc, document)
        elif tag == "sdt":
            # A content control is a wrapper; its contents are ordinary blocks.
            for content in child.iterchildren(W + "sdtContent"):
                yield from _iter_blocks(content, document)


def table_label(table: Table) -> str:
    if not table.rows:
        return ""
    cells = row_cells(table.rows[0])
    return clean(cells[0].text) if cells else ""


def table_has_content_below_label(table: Table) -> bool:
    """True if anything below the label row has text."""
    return rows_have_content(table, 1, len(table.rows))


def rows_have_content(table: Table, start: int, end: int) -> bool:
    """True if any row in [start, end) has text."""
    for row in table.rows[start:end]:
        if any(row_text(row)):
            return True
    return False


def label_rows(table: Table, pattern: re.Pattern) -> list[tuple[int, str]]:
    """Every row whose first cell is a label of this kind, with its index.

    Word merges two adjacent tables into one when the empty paragraph between
    them is deleted, which happens readily while editing. The second table's
    label then survives as an ordinary row part-way down the merged table. A
    parser that only reads the first row counts one where there are two, and
    says nothing - so labels are looked for on every row.
    """
    found: list[tuple[int, str]] = []
    for index, row in enumerate(table.rows):
        cells = row_text(row)
        if cells and pattern.match(cells[0]):
            found.append((index, cells[0]))
    return found


def label_segments(
    table: Table, pattern: re.Pattern
) -> list[tuple[str, int, int]]:
    """(label, first row after it, row after the segment) for each label row."""
    labels = label_rows(table, pattern)
    if not labels:
        return []
    boundaries = [index for index, _ in labels] + [len(table.rows)]
    return [
        (label, index + 1, boundaries[position + 1])
        for position, (index, label) in enumerate(labels)
    ]


# --- results ----------------------------------------------------------------


# The same "#Matter 1" spelling has to be read here, or a document using it
# looks like fourteen missing matters rather than fourteen counted ones.
RE_MATTER_NUMBER = re.compile(_MATTER_LABEL + r"\s*[-–—:#]?\s*(?P<number>\d+)", re.I)


@dataclass
class Metric:
    """One recounted metric.

    `value` is None when the metric could not be located with confidence; the
    UI renders that as "unparsed" and never as 0.
    """

    value: int | None = None
    evidence: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    numbering_gap: bool = False  # the labels do not run 1..N
    needs_check: bool = False  # something here could not be settled automatically

    @property
    def is_unparsed(self) -> bool:
        return self.value is None

    @property
    def display(self) -> str:
        return UNPARSED if self.value is None else str(self.value)


@dataclass
class ParsedDocument:
    filename: str
    firm: str | None = None
    region: str | None = None
    practice_group: str | None = None
    practice_area: str | None = None
    dropdown_raw: str | None = None
    typed_practice_area: str | None = None
    metrics: dict[str, Metric] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None

    @property
    def concatenated_practice_area(self) -> str | None:
        """'Group > Area', matching the Power BI export's own format.

        A typed-in practice area often has no group, in which case the bare area
        is returned - it will not match the export, and that is reported rather
        than papered over.
        """
        if self.practice_group and self.practice_area:
            return f"{self.practice_group} > {self.practice_area}"
        return self.practice_area or None

    @property
    def unparsed_metrics(self) -> list[str]:
        return [k for k, m in self.metrics.items() if m.is_unparsed]


# --- header fields ----------------------------------------------------------


def extract_firm(blocks: list[tuple[str, Any]]) -> tuple[str | None, list[str]]:
    """Firm name: the 1x1 table immediately following the 'Firm Name' label."""
    warnings: list[str] = []
    for i, (kind, item) in enumerate(blocks):
        if kind != "p" or clean(item.text).lower() != "firm name":
            continue
        for _, nxt in blocks[i + 1 : i + 4]:
            if isinstance(nxt, Table):
                value = clean(row_cells(nxt.rows[0])[0].text)
                if value:
                    return value, warnings
        warnings.append("Found the 'Firm Name' label but no filled table under it.")
        return None, warnings
    warnings.append("No 'Firm Name' label found in the document.")
    return None, warnings


def extract_practice_dropdown(document) -> tuple[str | None, list[str]]:
    """Read the practice-area dropdown (a Word content control).

    Holds 'Region - Practice group - Practice area'. This is not a paragraph or
    a table, so it is invisible to a plain text walk of the document.
    """
    warnings: list[str] = []
    candidates: list[str] = []
    for sdt in document.element.body.findall(".//" + W + "sdt"):
        alias_el = sdt.find(".//" + W + "alias")
        alias = alias_el.get(W + "val") if alias_el is not None else ""
        text = clean("".join(node.text or "" for node in sdt.findall(".//" + W + "t")))
        if not text or PLACEHOLDER_DROPDOWN.match(text):
            continue
        if alias and "practice" in alias.lower():
            return text, warnings
        candidates.append(text)

    # Fall back to any content control shaped like 'a - b - c'.
    for text in candidates:
        if len(split_dropdown(text)) == 3:
            warnings.append(
                "Practice area read from an unlabelled content control; "
                "verify the region and practice area on this row."
            )
            return text, warnings

    warnings.append("No practice-area dropdown found or it was left unselected.")
    return None, warnings


def extract_typed_practice_area(blocks: list[tuple[str, Any]]) -> str | None:
    """Read the practice area a firm typed instead of using the dropdown.

    The text may sit after the marker on the prompt's own line, or in one of the
    paragraphs just below it. Anything recognisable as another part of the form
    is rejected rather than guessed at.
    """
    for i, (kind, item) in enumerate(blocks):
        if kind != "p" or not RE_TYPED_BOX.search(clean(item.text)):
            continue

        candidates: list[str] = []
        prompt = clean(item.text)
        if "\u25ba" in prompt:  # the black right-pointing marker
            candidates.append(prompt.rsplit("\u25ba", 1)[-1])
        elif ":" in prompt:
            candidates.append(prompt.rsplit(":", 1)[-1])

        # Only the paragraphs immediately below the prompt, and never past the
        # next table - the blocks walk descends into table cells, whose text is
        # not an answer to this question.
        for kind, following in blocks[i + 1 : i + 6]:
            if kind != "p":
                break
            candidates.append(clean(following.text))

        for candidate in candidates:
            text = clean(candidate)
            if not text or len(text) > TYPED_MAX_LENGTH:
                continue
            if RE_TYPED_BOILERPLATE.match(text):
                continue
            return text
        return None
    return None


def split_dropdown(text: str) -> list[str]:
    """Split 'Region - Group - Area: qualifier' into its three parts.

    Only hyphens surrounded by spaces separate the parts; hyphens inside a name
    ('Non-contentious') and the colon-qualified tail must survive intact. The
    split is capped at two so an area containing ' - ' stays whole.
    """
    parts = [clean(p) for p in re.split(r"\s+-\s+", text, maxsplit=2)]
    return [p for p in parts if p]


# --- metric counting --------------------------------------------------------


def count_clients(tables: list[Table]) -> tuple[Metric, Metric]:
    """Count active clients and new active clients across all client tables.

    Publishable and non-publishable tables are both included, per the brief.
    """
    active = Metric(value=0)
    new = Metric(value=0)
    unrecognised: list[str] = []
    qualified: list[str] = []
    blanks = 0
    substitutions = 0
    not_applicable = 0

    client_tables = [t for t in tables if RE_CLIENT_TABLE.match(table_label(t))]
    if not client_tables:
        # The template always ships this table, even when a firm leaves it
        # empty. Its absence means our anchor failed, not that there are no
        # clients, so refuse to report 0.
        msg = "No 'Active key clients' table found - could not locate the client section."
        return Metric(value=None, notes=[msg]), Metric(value=None, notes=[msg])

    for table in client_tables:
        rows_counted = 0
        for row in table.rows[1:]:
            cells = row_text(row)
            name = cells[0] if cells else ""
            answer = unwrap(cells[1] if len(cells) > 1 else "")

            if not name:
                continue
            if RE_CLIENT_BOILERPLATE.search(name):
                continue
            if RE_NOT_APPLICABLE.match(name):
                # "N/A" in the client column is a firm saying there are none,
                # not a client called N/A.
                not_applicable += 1
                continue
            if RE_CLIENT_HEADER_ROW.match(name):
                # Some firms paste an extra header row inside the table. The
                # upstream export counts these as clients; we do not.
                new.notes.append(f"Ignored a repeated header row: {name!r}")
                continue

            rows_counted += 1
            active.value += 1

            if RE_NOT_APPLICABLE.match(answer):
                unrecognised.append(answer)
            elif RE_AFFIRMATIVE.match(answer):
                new.value += 1
                if not RE_BARE_AFFIRMATIVE.match(answer):
                    qualified.append(answer)
            elif RE_NEW_WORD.match(answer):
                new.value += 1
                substitutions += 1
                if not RE_BARE_NEW.match(answer):
                    qualified.append(answer)
            elif RE_NEGATIVE.match(answer):
                if not RE_BARE_NEGATIVE.match(answer):
                    qualified.append(answer)
            elif RE_EXISTING_WORD.match(answer):
                substitutions += 1
                if not RE_BARE_EXISTING.match(answer):
                    qualified.append(answer)
            elif answer == "":
                blanks += 1
            else:
                unrecognised.append(answer)

        active.evidence.append(f"{table_label(table)} table - {rows_counted} client rows")

    new.evidence = list(active.evidence)

    if substitutions:
        new.notes.append(
            f"{substitutions} answer(s) used 'New'/'Existing' instead of Yes/No; "
            "'New' was counted as a new client."
        )
    if blanks:
        new.notes.append(f"{blanks} blank answer(s) counted as not-new.")
    if not_applicable:
        active.notes.append(
            f"{not_applicable} row(s) read 'N/A' in the client column and were not "
            "counted as clients."
        )
    if qualified:
        listed = ", ".join(repr(q) for q in sorted(set(qualified)))
        new.notes.append(
            f"{len(qualified)} qualified answer(s) read by their leading word: "
            f"{listed}. A qualified yes counts as a new client."
        )
    if unrecognised:
        # Neither yes nor no. Counting these as "not new" would be a guess, and
        # a guess that hides itself - so the whole count is handed back.
        listed = ", ".join(repr(u) for u in sorted(set(unrecognised)))
        new.value = None
        new.notes.append(
            f"{len(unrecognised)} answer(s) could not be read as yes or no: "
            f"{listed}. Count this column by hand."
        )
    return active, new


WRAPPERS = {"[": "]", "(": ")", "{": "}", '"': '"', "'": "'", "\u201c": "\u201d"}


def unwrap(answer: str) -> str:
    """An answer without brackets or quotes wrapped around the whole of it.

    "[No]" is "No". "Yes (for this work type)" keeps its closing bracket,
    because the brackets do not enclose the whole answer.
    """
    while len(answer) >= 2 and WRAPPERS.get(answer[0]) == answer[-1]:
        answer = answer[1:-1].strip()
    return answer

def carries_label(table: Table) -> bool:
    """True if a table is introduced by a label the parser recognises."""
    if label_rows(table, RE_MATTER_TABLE) or nomination_segments(table):
        return True
    for row in table.rows:
        cells = row_text(row)
        if cells and RE_CLIENT_TABLE.match(cells[0]):
            return True
    return False


def continuation_of(tables: list[Table], index: int) -> Table | None:
    """The body of a label whose own table ends at the label row.

    The counterpart of a merge: Word also splits one table into two, leaving
    the label alone in a table of its own and its body following as a table
    carrying no label at all. Read literally the label is an empty template and
    the body belongs to nothing, so a filled matter or nomination vanishes with
    nothing on screen to say so.

    A template that is genuinely empty keeps its skeleton rows - 'Name of
    client', 'Matter description' - under its label, so a label row that is the
    last row of its table is the reliable sign of a split rather than of an
    empty box.
    """
    following = tables[index + 1] if index + 1 < len(tables) else None
    if following is None or carries_label(following):
        return None
    if not rows_have_content(following, 0, len(following.rows)):
        return None
    return following


def count_matters(tables: list[Table]) -> Metric:
    """Count work matters: one label per matter, publishable and not.

    Counted per label rather than per table, because two matters can share one
    table when Word has merged them.
    """
    metric = Metric(value=0)
    empty = 0
    merged = 0
    numbered: dict[str, list[int]] = {"Publishable": [], "Non-publishable": []}
    split = 0
    adopted: set[int] = set()
    for index, table in enumerate(tables):
        segments = label_segments(table, RE_MATTER_TABLE)
        if len(segments) > 1:
            merged += len(segments) - 1
        for label, start, end in segments:
            if not rows_have_content(table, start, end):
                # The label may have been split from its body, which then
                # follows as the next table with no label of its own.
                body = None
                if start >= len(table.rows) and index + 1 not in adopted:
                    body = continuation_of(tables, index)
                if body is None:
                    empty += 1
                    continue
                adopted.add(index + 1)
                split += 1
            metric.value += 1
            metric.evidence.append(label)
            seen = RE_MATTER_NUMBER.match(label)
            if seen:
                non_publishable = seen.group("non") or seen.group("confidential")
                kind = "Non-publishable" if non_publishable else "Publishable"
                numbered[kind].append(int(seen.group("number")))

    # A failsafe for labels worded in a way nobody anticipated ("Son-publishable
    # matter 3"): the template colours every matter label, red for
    # non-publishable and green for publishable. Colour alone is not enough -
    # the client tables use the same two colours - so the table must also be
    # laid out as a matter, with a client name and a matter summary.
    by_colour: list[str] = []
    for index, table in enumerate(tables):
        if index in adopted or carries_label(table):
            continue
        colour = matter_colour(table)
        if colour is None or not laid_out_as_matter(table):
            continue
        label = table_label(table)
        metric.value += 1
        metric.evidence.append(label)
        by_colour.append(label)
        number = re.search(r"(\d+)\s*$", label)
        if number:
            numbered[colour].append(int(number.group(1)))
    if by_colour:
        listed = ", ".join(repr(label) for label in by_colour)
        metric.notes.append(
            f"{len(by_colour)} matter(s) counted from the colour of the label box, "
            f"because the label is worded unusually: {listed}. Check it is a matter."
        )

    if metric.value == 0 and empty == 0:
        metric.value = None
        metric.notes.append(
            "No 'Publishable matter' or 'Non-publishable matter' table found - "
            "could not locate the work highlights section."
        )
        return metric

    if empty:
        metric.notes.append(f"{empty} empty matter template(s) skipped.")
    if split:
        metric.notes.append(
            f"{split} matter(s) have their label in a table of their own, with "
            "the details following in the next table - Word splits a table in "
            "two this way. Counted."
        )
    if merged:
        metric.notes.append(
            f"{merged} matter(s) share a table with the matter above them - Word "
            "merges adjacent tables when the paragraph between them is deleted. "
            "Counted separately."
        )

    # A last sanity check: the matters a firm numbers 1..N should come to N.
    # A gap is usually the firm deleting a matter without renumbering, but it is
    # also what an undercount looks like, so it is always reported.
    #
    # Most firms number each kind from 1, but some number straight through
    # both ("Publishable matter 6", then "Non-publishable matter 7"), and some
    # interleave them. Either way is fine as long as one of the two readings
    # runs 1..N; where neither does, the one leaving fewer gaps is reported.
    per_kind = {kind: seen for kind, seen in numbered.items() if seen}
    combined = [n for seen in per_kind.values() for n in seen]
    if not combined or all(_runs_from_one(seen) for seen in per_kind.values()):
        return metric
    if _runs_from_one(combined):
        return metric

    metric.numbering_gap = True
    readings = {"per kind": per_kind, "combined": {"Matters": combined}}
    chosen = min(
        readings.values(),
        key=lambda groups: sum(len(_gaps(seen)[0]) for seen in groups.values()),
    )
    for kind, seen in chosen.items():
        if _runs_from_one(seen):
            continue
        absent, repeated = _gaps(seen)
        detail = []
        if absent:
            detail.append("no " + _ranges(absent))
        if repeated:
            detail.append("repeated " + _ranges(repeated))
        across = " across publishable and non-publishable" if kind == "Matters" else ""
        label = kind if kind == "Matters" else f"{kind} matters"
        metric.notes.append(
            f"{label} are numbered up to {max(seen)}{across} but {len(seen)} were "
            f"counted ({'; '.join(detail)}). Usually the firm deleted one without "
            "renumbering, or mistyped a number, but it is also what a missed "
            "matter looks like, so check this document's matters by hand."
        )
    return metric


def _runs_from_one(numbers: list[int]) -> bool:
    """True if the numbers are exactly 1..N, each once."""
    return sorted(numbers) == list(range(1, len(numbers) + 1))


def _gaps(numbers: list[int]) -> tuple[list[int], list[int]]:
    """(numbers missing below the highest, numbers used more than once)."""
    absent = sorted(set(range(1, max(numbers) + 1)) - set(numbers))
    repeated = sorted({n for n in numbers if numbers.count(n) > 1})
    return absent, repeated


def _ranges(numbers: list[int]) -> str:
    """'5, 7 to 9, 30 to 229': runs written as ranges, so a typo stays short."""
    runs: list[list[int]] = []
    for n in numbers:
        if runs and n == runs[-1][-1] + 1:
            runs[-1].append(n)
        else:
            runs.append([n])
    return ", ".join(
        str(r[0]) if len(r) == 1 else f"{r[0]}, {r[1]}" if len(r) == 2 else f"{r[0]} to {r[-1]}"
        for r in runs
    )


def cell_fill(cell) -> str:
    """A cell's background colour as 'RRGGBB', or '' if it has none."""
    shading = cell._tc.find(f"{W}tcPr/{W}shd")
    fill = (shading.get(f"{W}fill") or "") if shading is not None else ""
    return fill.upper() if re.fullmatch(r"[0-9A-Fa-f]{6}", fill) else ""


def matter_colour(table: Table) -> str | None:
    """'Non-publishable' for a red label box, 'Publishable' for a green one.

    Judged by hue rather than an exact colour, because firms' copies of the
    template drift: the reds seen include F1A3A3, EE0000 and D99594, the greens
    C5E0B3, B3E5A1 and A8D08D. Grey, white, yellow and blue are neither.
    """
    if not table.rows:
        return None
    cells = row_cells(table.rows[0])
    fill = cell_fill(cells[0]) if cells else ""
    if not fill:
        return None
    red, green, blue = (int(fill[i : i + 2], 16) / 255 for i in (0, 2, 4))
    hue, _, saturation = colorsys.rgb_to_hls(red, green, blue)
    if saturation < 0.25:
        return None
    degrees = hue * 360
    if degrees < 15 or degrees > 345:
        return "Non-publishable"
    if 75 < degrees < 160:
        return "Publishable"
    return None


RE_CLIENT_NAME_ROW = re.compile(r"^name of (the )?client", re.I)
RE_MATTER_SUMMARY_ROW = re.compile(r"^matter (summary|description)", re.I)


def laid_out_as_matter(table: Table) -> bool:
    """True if the rows under the label are a matter's: client, then summary."""
    firsts = [(row_text(row) or [""])[0] for row in table.rows[1:9]]
    return any(RE_CLIENT_NAME_ROW.match(text) for text in firsts[:3]) and any(
        RE_MATTER_SUMMARY_ROW.match(text) for text in firsts
    )


def nomination_segments(table: Table) -> list[tuple[str, str, int, int]]:
    """(category, label, first row after it, row after the segment) per label.

    Like matters, two nominations can end up sharing one table after Word has
    merged them, and a merged table can hold labels of different categories.
    """
    labels: list[tuple[int, str, str]] = []
    for index, row in enumerate(table.rows):
        cells = row_text(row)
        if not cells:
            continue
        category = classify_nomination(cells[0])
        if category is not None:
            labels.append((index, category, cells[0]))
    if not labels:
        return []
    boundaries = [index for index, _, _ in labels] + [len(table.rows)]
    return [
        (key, label, index + 1, boundaries[position + 1])
        for position, (index, key, label) in enumerate(labels)
    ]


def count_nominations(
    tables: list[Table], paragraphs: list[str], key: str
) -> Metric:
    """Count nominations for one category.

    A nomination counts only when its Name cell is filled: the sample set
    contains template tables left in place with nothing in them.
    """
    metric = Metric(value=0)
    empty = 0
    merged = 0
    unclassified: list[str] = []

    split = 0
    adopted: set[int] = set()
    for index, table in enumerate(tables):
        segments = nomination_segments(table)
        merged += max(len(segments) - 1, 0)
        for category, label, start, end in segments:
            name = nomination_name(table, start, end)
            if not name and start >= len(table.rows) and index + 1 not in adopted:
                # The label was split from its body, which follows unlabelled.
                body = continuation_of(tables, index)
                if body is not None:
                    name = nomination_name(body, 0, len(body.rows))
                    if name:
                        adopted.add(index + 1)
                        if category == key:
                            split += 1
            if category == UNCLASSIFIED:
                if name:
                    unclassified.append(label)
                continue
            if category != key:
                continue
            if not name:
                empty += 1
                continue
            metric.value += 1
            metric.evidence.append(f"{label} - {name}")

    # Reported before any early return: a label nobody could classify is worth
    # saying even when this category ended up with nothing in it.
    if unclassified and key in ("lead_partners", "next_gen"):
        listed = ", ".join(repr(u) for u in sorted(set(unclassified)))
        metric.needs_check = True
        metric.notes.append(
            f"{len(unclassified)} partner nomination(s) use a description that "
            f"says neither leading nor next generation: {listed}. They are not "
            "counted in either category - decide by hand which this is."
        )

    if metric.value == 0 and empty == 0:
        # Absence is a real zero only when the section is absent too. If the
        # heading is there but no table matched, our label patterns have gone
        # stale and we must not report 0.
        heading = NOMINATION_HEADINGS[key]
        if any(heading.search(p) for p in paragraphs):
            metric.value = None
            metric.notes.append(
                "A heading for this category is present but no matching "
                "nomination table was found."
            )
            return metric
        metric.notes.append("No nominations of this type in the document.")
        return metric

    if empty:
        metric.notes.append(f"{empty} empty nomination template(s) skipped.")
    if split:
        metric.notes.append(
            f"{split} nomination(s) have their label in a table of their own, "
            "with the details following in the next table - Word splits a table "
            "in two this way. Counted."
        )
    if merged:
        metric.notes.append(
            f"{merged} nomination(s) share a table with the one above them; "
            "counted separately."
        )
    return metric


def nomination_name(table: Table, start: int = 1, end: int | None = None) -> str:
    """The nominee's name: the cell under the 'Name' column header.

    `start`/`end` bound the segment belonging to one nomination, so a merged
    table does not hand back the name of the nominee above.
    """
    end = len(table.rows) if end is None else end
    for i in range(start, end):
        cells = row_text(table.rows[i])
        if cells and cells[0].lower() == "name":
            if i + 1 < end:
                following = row_text(table.rows[i + 1])
                return following[0] if following else ""
            return ""
    # Fall back to the conventional layout: label, header, value.
    if start + 1 < end:
        cells = row_text(table.rows[start + 1])
        return cells[0] if cells else ""
    return ""


# --- top level --------------------------------------------------------------


def parse_document(source: Any, filename: str | None = None) -> ParsedDocument:
    """Parse one .docx. `source` is a path or a file-like object."""
    name = filename or getattr(source, "name", None) or str(source)
    name = name.replace("\\", "/").rsplit("/", 1)[-1]
    result = ParsedDocument(filename=name)

    try:
        document = Document(source)
    except Exception as first:  # noqa: BLE001 - surfaced to the user verbatim
        # Some packages are internally inconsistent about the capitalisation of
        # their own part names. Word opens them; Python does not. Rather than
        # lose the document, take a repaired copy and say what was changed.
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            repaired, repairs = repair_package(source)
            document = Document(BytesIO(repaired))
        except Exception:  # noqa: BLE001 - the original failure is the useful one
            result.error = f"Could not open the document: {first}"
            result.metrics = {
                k: Metric(value=None, notes=[result.error]) for k in METRIC_KEYS
            }
            return result
        result.warnings.append(
            "This file had to be repaired before it could be read - "
            + " ".join(repairs)
            + " Nothing in the content was changed, but it is worth opening the "
            "document once to confirm it looks right."
        )

    blocks = list(iter_blocks(document))
    tables = [item for kind, item in blocks if kind == "tbl"]
    paragraphs = [clean(item.text) for kind, item in blocks if kind == "p"]
    paragraphs = [p for p in paragraphs if p]

    result.firm, firm_warnings = extract_firm(blocks)
    result.warnings.extend(firm_warnings)

    dropdown, dropdown_warnings = extract_practice_dropdown(document)
    result.dropdown_raw = dropdown
    result.typed_practice_area = extract_typed_practice_area(blocks)

    if dropdown:
        result.warnings.extend(dropdown_warnings)
        parts = split_dropdown(dropdown)
        if len(parts) == 3:
            result.region, result.practice_group, result.practice_area = parts
        elif len(parts) == 2:
            # Publications outside the UK have no practice group, so the
            # dropdown reads "Region - Practice area".
            result.region, result.practice_area = parts
        else:
            result.warnings.append(
                f"Practice-area dropdown did not split into two or three parts: "
                f"{dropdown!r}"
            )
    elif result.typed_practice_area:
        # The firm typed it instead of selecting it. Take whatever shape it is
        # in: three parts means they copied the full dropdown text, otherwise it
        # is a bare practice area.
        parts = split_dropdown(result.typed_practice_area)
        if len(parts) == 3:
            result.region, result.practice_group, result.practice_area = parts
        else:
            result.practice_area = result.typed_practice_area
        result.warnings.append(
            "The practice-area dropdown was not used; read "
            f"{result.typed_practice_area!r} from the typed box instead."
        )
    else:
        result.warnings.extend(dropdown_warnings)

    if not looks_like_submission(tables):
        result.error = (
            "This does not look like a submission document - none of the expected "
            "tables (clients, matters, nominations) were found."
        )
        result.metrics = {k: Metric(value=None, notes=[result.error]) for k in METRIC_KEYS}
        return result

    active, new = count_clients(tables)
    result.metrics = {
        "matters": count_matters(tables),
        "active_clients": active,
        "new_clients": new,
        "lead_partners": count_nominations(tables, paragraphs, "lead_partners"),
        "next_gen": count_nominations(tables, paragraphs, "next_gen"),
        "associates": count_nominations(tables, paragraphs, "associates"),
    }
    return result


def looks_like_submission(tables: list[Table]) -> bool:
    """Cheap structural sanity check before we trust any count."""
    labels = [table_label(t) for t in tables]
    markers = 0
    if any(RE_CLIENT_TABLE.match(l) for l in labels):
        markers += 1
    if any(RE_MATTER_TABLE.match(l) for l in labels):
        markers += 1
    if any(
        p.match(l) for l in labels for p in NOMINATION_PATTERNS.values()
    ):
        markers += 1
    return markers >= 1


def parse_documents(sources: list[Any]) -> list[ParsedDocument]:
    return [parse_document(s) for s in sources]
