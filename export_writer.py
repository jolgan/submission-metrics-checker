"""Write the results to .xlsx with the highlighting baked into the file.

Streamlit renders colour in the browser, but copying out of a browser loses it.
These fills are real cell formatting, so they survive being pasted into the
shared master sheet.
"""

from __future__ import annotations

from io import BytesIO

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from matching import (
    BASE_COLUMNS,
    ResultRow,
    difference_records,
    display_value,
    final_differs,
    needs_attention,
)
from parsing import METRIC_COLUMNS, METRIC_KEYS, METRIC_LABELS, UNPARSED

GREEN = PatternFill("solid", start_color="92D050", end_color="92D050")
AMBER = PatternFill("solid", start_color="FFE9A8", end_color="FFE9A8")
HEADER = PatternFill("solid", start_color="EDEDED", end_color="EDEDED")


CHECKED_COLUMN = "DOUBLE-CHECKED"
CHECKED_TEXT = "double-checked"


def build_workbook(
    results: list[ResultRow],
    overrides: dict | None = None,
    checked: set | None = None,
) -> bytes:
    """The result table, with the highlighting decided by the final values.

    `overrides` holds figures entered by hand, `checked` the rows marked as
    double-checked. Both are applied here rather than layered on afterwards, so
    a cell corrected back to the export's own figure stops being highlighted.
    """
    workbook = Workbook()
    _write_recount_sheet(workbook.active, results, overrides, checked)
    _write_changes_sheet(workbook.create_sheet("Changes"), results, overrides)
    _write_issues_sheet(workbook.create_sheet("Issues"), results)

    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _write_recount_sheet(sheet, results, overrides=None, checked=None) -> None:
    sheet.title = "Recount"
    metric_headers = [METRIC_COLUMNS[k] for k in METRIC_KEYS]
    headers = BASE_COLUMNS + metric_headers + [CHECKED_COLUMN]

    for col, name in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=col, value=name)
        cell.font = Font(bold=True)
        cell.fill = HEADER
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    for r, row in enumerate(results, start=2):
        values = [row.firm_ref, row.firm, row.region, row.practice_area, row.filename]
        for col, value in enumerate(values, start=1):
            sheet.cell(row=r, column=col, value=value)

        for offset, key in enumerate(METRIC_KEYS):
            col = len(BASE_COLUMNS) + 1 + offset
            value, unverified = display_value(row, key, overrides, checked)
            cell = sheet.cell(row=r, column=col)

            # Amber means "your eyes are needed here", whether because nothing
            # could be counted or because what was counted does not add up.
            # Bold was tried and removed: it travels with a copy into the master
            # sheet and has to be cleared by hand.
            if unverified or needs_attention(row, key, checked):
                # Nothing could be counted, so the cell carries the export's own
                # figure as a starting point and is shaded to say so.
                cell.value = UNPARSED if value is None else value
                cell.fill = AMBER
                notes = row.document.metrics[key].notes if row.document.metrics else []
                heading = (
                    "Not counted from the document. This is the Power BI "
                    "export's own figure, unverified."
                    if unverified
                    else "Counted, but worth checking."
                )
                cell.comment = _comment(
                    heading + ("\n" + "\n".join(notes) if notes else "")
                )
                continue

            cell.value = value
            if final_differs(row, key, overrides):
                cell.fill = GREEN
                decided = "you" if value != row.recount(key) else "the recount"
                cell.comment = _comment(
                    f"Differs from the Power BI export.\n"
                    f"Power BI export: {row.export_value(key)}\n"
                    f"This file: {value}  (decided by {decided})"
                )

        if checked and row.filename.lower() in checked:
            sheet.cell(row=r, column=len(headers), value=CHECKED_TEXT)

    _autosize(sheet, headers)
    sheet.freeze_panes = "A2"


def _write_changes_sheet(sheet, results, overrides=None) -> None:
    records = difference_records(results, overrides)
    headers = [
        "Firm",
        "Region",
        "Practice area",
        "Filename",
        "Metric",
        "Power BI value",
        "Recounted value",
        "Decided by",
    ]
    for col, name in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=col, value=name)
        cell.font = Font(bold=True)
        cell.fill = HEADER

    for r, record in enumerate(records, start=2):
        for col, name in enumerate(headers, start=1):
            cell = sheet.cell(row=r, column=col, value=record[name])
            if name == "Recounted value":
                cell.fill = GREEN

    if not records:
        sheet.cell(row=2, column=1, value="No differences found.")
    _autosize(sheet, headers)
    sheet.freeze_panes = "A2"


def _write_issues_sheet(sheet, results: list[ResultRow]) -> None:
    headers = ["Filename", "Firm", "Practice area", "Status", "Issue"]
    for col, name in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=col, value=name)
        cell.font = Font(bold=True)
        cell.fill = HEADER

    r = 2
    for row in results:
        issues = list(row.issues)
        for key in METRIC_KEYS:
            metric = row.document.metrics.get(key)
            if metric is None:
                continue
            for note in metric.notes:
                issues.append(f"{METRIC_LABELS[key]}: {note}")
        for issue in issues:
            sheet.cell(row=r, column=1, value=row.filename)
            sheet.cell(row=r, column=2, value=row.firm)
            sheet.cell(row=r, column=3, value=row.practice_area)
            sheet.cell(row=r, column=4, value=row.status)
            sheet.cell(row=r, column=5, value=issue)
            r += 1

    if r == 2:
        sheet.cell(row=2, column=1, value="No issues reported.")
    _autosize(sheet, headers)
    sheet.freeze_panes = "A2"


def _comment(text: str) -> Comment:
    comment = Comment(text, "Submission recount")
    comment.width = 320
    comment.height = 110
    return comment


def _autosize(sheet, headers: list[str], limit: int = 48) -> None:
    for col in range(1, len(headers) + 1):
        widest = len(str(headers[col - 1]))
        for row in range(2, min(sheet.max_row, 400) + 1):
            value = sheet.cell(row=row, column=col).value
            if value is not None:
                widest = max(widest, len(str(value)))
        sheet.column_dimensions[get_column_letter(col)].width = min(widest + 2, limit)
