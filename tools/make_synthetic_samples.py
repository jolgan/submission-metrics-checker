"""Generate synthetic submissions and exports for the test suite.

The real submissions are confidential, so the repository ships invented ones.
They copy the structure of the real documents, including the awkward shapes the
parser exists to cope with, and every firm, person and client here is made up.

    python tools/make_synthetic_samples.py

Writes to samples/synthetic/ along with manifest.json, which records what each
document should count. The tests read their expectations from that manifest, so
the two cannot drift apart.

The three documents differ on purpose:

    alpha   plain shape; two nominations carry the same trailing number, and
            one client has no answer in the new-client column
    beta    two matters merged into one table, as Word leaves them when the
            paragraph between two tables is deleted; a nomination sits inside a
            content control; associates use a description of the firm's own
    gamma   New/Existing instead of Yes/No, a header row repeated inside the
            client table, an N/A client row, a qualified "Yes (...)" answer,
            an empty nomination template, and a gap in the matter numbering
"""

from __future__ import annotations

import json
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "samples" / "synthetic"

EDITION = "2027 Edition"


# --- document building blocks ----------------------------------------------


def add_label_table(document, rows, cols=2):
    table = document.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    return table


def wrap_in_content_control(table):
    """Put a table inside a block-level w:sdt, as Word does for templates."""
    element = table._tbl
    sdt = element.makeelement(qn("w:sdt"), {})
    content = element.makeelement(qn("w:sdtContent"), {})
    element.addprevious(sdt)
    sdt.append(content)
    content.append(element)


def add_header(document, firm, region, group, area, team):
    document.add_paragraph("Firm Name")
    add_label_table(document, 1, 1).rows[0].cells[0].text = firm

    document.add_paragraph("Country")
    document.add_paragraph("Practice Area")
    prompt = document.add_paragraph(
        "EITHER select Practice Area from this drop-own list ►"
    )
    # The practice area lives in a dropdown, as it does in the real form.
    run = prompt.add_run()
    sdt = run._r.makeelement(qn("w:sdt"), {})
    props = run._r.makeelement(qn("w:sdtPr"), {})
    alias = run._r.makeelement(qn("w:alias"), {qn("w:val"): "Select Practice "})
    props.append(alias)
    content = run._r.makeelement(qn("w:sdtContent"), {})
    text_run = run._r.makeelement(qn("w:r"), {})
    text = run._r.makeelement(qn("w:t"), {})
    text.text = f"{region} - {group} - {area}"
    text_run.append(text)
    content.append(text_run)
    sdt.append(props)
    sdt.append(content)
    prompt._p.append(sdt)

    document.add_paragraph(
        "OR If you have an earlier version of Word, type in this box: ►"
    )
    document.add_paragraph("What is the Team or Department Name (as used by your firm)")
    add_label_table(document, 1, 1).rows[0].cells[0].text = team


def add_clients(document, heading, rows, extra_header=False):
    document.add_paragraph(heading)
    count = len(rows) + 1 + (1 if extra_header else 0) + 1
    table = add_label_table(document, count, 2)
    table.rows[0].cells[0].text = "Active key clients (over the last 12 months)"
    table.rows[0].cells[1].text = "New client (yes/no)"
    offset = 1
    if extra_header:
        table.rows[1].cells[0].text = "Client name"
        table.rows[1].cells[1].text = "New / Existing"
        offset = 2
    for i, (name, answer) in enumerate(rows):
        table.rows[offset + i].cells[0].text = name
        table.rows[offset + i].cells[1].text = answer
    last = table.rows[-1]
    last.cells[0].text = (
        "To add more clients, right-click in any field and select ‘Insert row below’"
    )
    return table


def add_nomination(document, label, name, location="London", inside_control=False):
    table = add_label_table(document, 5, 3)
    table.rows[0].cells[0].text = label
    table.rows[1].cells[0].text = "Name"
    table.rows[1].cells[1].text = "Location"
    table.rows[1].cells[2].text = "Ranked in previous edition? (Yes/no)"
    table.rows[2].cells[0].text = name
    table.rows[2].cells[1].text = location if name else ""
    table.rows[2].cells[2].text = "N" if name else ""
    table.rows[3].cells[0].text = "Supporting information"
    table.rows[4].cells[0].text = (
        f"{name} advises on matters across the practice." if name else ""
    )
    if inside_control:
        wrap_in_content_control(table)
    return table


def add_matter(document, label, client, sector="Financial Services"):
    table = add_label_table(document, 4, 2)
    table.rows[0].cells[0].text = label
    table.rows[1].cells[0].text = "Name of client"
    table.rows[1].cells[1].text = "Industry sector"
    table.rows[2].cells[0].text = client
    table.rows[2].cells[1].text = sector
    table.rows[3].cells[0].text = "Matter description"
    return table


def add_merged_matters(document, entries):
    """Several matters in one table, as Word leaves them after a merge."""
    table = add_label_table(document, 4 * len(entries), 2)
    for position, (label, client) in enumerate(entries):
        top = position * 4
        table.rows[top].cells[0].text = label
        table.rows[top + 1].cells[0].text = "Name of client"
        table.rows[top + 1].cells[1].text = "Industry sector"
        table.rows[top + 2].cells[0].text = client
        table.rows[top + 2].cells[1].text = "Real Estate"
        table.rows[top + 3].cells[0].text = "Matter description"
    return table


def add_summary(document):
    document.add_paragraph("*NEW* Publishable matter summary")
    table = add_label_table(document, 2, 1)
    table.rows[0].cells[0].text = "Summary (including client name)"
    table.rows[1].cells[0].text = "Advised a client on a transaction."


# --- the three documents ----------------------------------------------------


def build_alpha(path):
    """Plain shape, with a repeated nomination number and a blank answer."""
    document = Document()
    add_header(
        document,
        "Aldgate Fenwick LLP",
        "South East",
        "Real estate",
        "Commercial property: Thames Valley",
        "Real Estate",
    )
    add_clients(
        document,
        "Clients: publishable clients (This will be published in full)",
        [("Harbour Wells Group", "No"), ("Trenton Mills plc", "Yes"),
         ("Calder Vale Estates", "No")],
    )
    add_clients(
        document,
        "Clients: non-publishable clients (This list will not be published)",
        [("Ferrisburgh Holdings", "No"), ("Oakwold Partners", "")],
    )
    document.add_paragraph("Your team - Partners: leading partners")
    add_nomination(document, "Partner: leading individual 1", "Marian Ashgrove")
    # Both carry the trailing number 2, as a real submission did.
    add_nomination(document, "Partner: next generation partner 2", "Desmond Fairlie")
    add_nomination(document, "Partner: next generation partner 2", "Priya Ellwood")
    add_summary(document)
    document.add_paragraph("Your practice - detailed work highlights")
    for n in range(1, 4):
        add_matter(document, f"Publishable matter {n}", f"Client {n} Limited")
    for n in range(1, 3):
        add_matter(document, f"Non-publishable matter {n}", f"Confidential client {n}")
    document.save(path)
    return {
        "matters": 5, "active_clients": 5, "new_clients": 1,
        "lead_partners": 1, "next_gen": 2, "associates": 0,
    }


def build_beta(path):
    """Merged matter tables, a nomination in a content control, odd labels."""
    document = Document()
    add_header(
        document,
        "Aldgate Fenwick LLP",
        "South West",
        "Corporate and commercial",
        "EU and competition",
        "Competition",
    )
    add_clients(
        document,
        "Clients: publishable clients (This will be published in full)",
        [("Vellacott Industries", "Yes"), ("Pemberton Foods", "No"),
         ("Strathmore Energy", "Y"), ("Ilminster Rail", "No")],
    )
    add_clients(
        document,
        "Clients: non-publishable clients (This list will not be published)",
        [("Northgate Chemicals", "No"), ("Ravensworth Bank", "Yes")],
    )
    document.add_paragraph("Your team - Partners: leading partners")
    add_nomination(document, "Partner: leading partner 1", "Colm Radcliffe")
    add_nomination(document, "Partner: leading partner 2", "Ines Beaumont")
    document.add_paragraph("Your team - Partners: next generation partners")
    add_nomination(
        document, "Partner: to become next generation partner 1", "Tobias Wren",
        inside_control=True,
    )
    document.add_paragraph("Your team - Associates: leading associates")
    add_nomination(document, "Associate: leading associate 1", "Nadia Thornbury")
    add_nomination(document, "Associate : leading counsel 2", "Emeka Braithwaite")
    add_summary(document)
    document.add_paragraph("Your practice - detailed work highlights")
    add_matter(document, "Publishable matter 1", "Vellacott Industries")
    # Matters 2 and 3 share one table.
    add_merged_matters(
        document,
        [("Publishable matter 2", "Pemberton Foods"),
         ("Publishable matter 3", "Strathmore Energy")],
    )
    add_matter(document, "Publishable matter 4", "Ilminster Rail")
    for n in range(1, 4):
        add_matter(document, f"Non-publishable matter {n}", f"Confidential client {n}")
    document.save(path)
    return {
        "matters": 7, "active_clients": 6, "new_clients": 3,
        "lead_partners": 2, "next_gen": 1, "associates": 2,
    }


def build_gamma(path):
    """New/Existing answers, a repeated header, N/A, and a numbering gap."""
    document = Document()
    add_header(
        document,
        "Aldgate Fenwick LLP",
        "London",
        "Finance",
        "Banking and finance",
        "Banking",
    )
    add_clients(
        document,
        "Clients: publishable clients (This will be published in full)",
        [("Ashcombe Trust", "Existing"), ("Dunmore Capital", "New"),
         ("Halkirk Leasing", "Yes (for this work type)")],
        extra_header=True,
    )
    add_clients(
        document,
        "Clients: non-publishable clients (This list will not be published)",
        [("N/A", "")],
    )
    document.add_paragraph("Your team - Partners: leading partners")
    add_nomination(document, "Partner: leading partner 1", "Rowan Kesteven")
    document.add_paragraph("Your team - Partners: next generation partners")
    add_nomination(document, "Partner: next generation partner 1", "")
    add_summary(document)
    document.add_paragraph("Your practice - detailed work highlights")
    # Publishable numbering skips 3, as firms do when they delete one.
    for n in (1, 2, 4):
        add_matter(document, f"Publishable matter {n}", f"Client {n} Limited")
    add_matter(document, "Non-publishable matter 1", "Confidential client 1")
    document.save(path)
    return {
        "matters": 4, "active_clients": 3, "new_clients": 2,
        "lead_partners": 1, "next_gen": 0, "associates": 0,
    }


# --- the exports ------------------------------------------------------------


POWERBI_COLUMNS = [
    "MATCH_STATUS", "SUBMISSION_FILE_NAME", "FIRM_NAME", "FIRM_REF",
    "CONCATENATED_PRACTICE_AREA_NAME", "SUBMISSION_PRACTICE_AREA_ID",
    "COUNTRY_NAME", "FIRM_RANKING_TYPE", "FIRM_RANKING_TIER",
    "RANKING_DECISION_STATUS", "RANKING_STATUS", "RANKING_PUBLICATION_STATUS",
    "NUM_REFEREES", "NUM_REFEREES_RESPONDED", "NUM_MATTERS",
    "NUM_ACTIVE_CLIENTS", "NUM_NEW_CLIENTS", "NUM_LEAD_PARTNERS",
    "NUM_NEXT_GEN", "NUM_ASSOCIATES",
]

PORTAL_COLUMNS = [
    "ID", "Firm", "Firm Ref", "Publication", "Edition", "Region",
    "Practice Group Name", "Practice Area Name", "Table Name",
    "Region (Translated)", "Practice Group (Translated)",
    "Practice Area (Translated)", "Table (Translated)", "Language Code",
    "Region ID", "Publication ID", "Editorial Category ID", "Practice Area ID",
    "Practice Group ID", "Table ID", "Created At", "Created By", "Close Date",
    "Submission Type", "Referee Type", "Person Name", "Document", "File URI",
    "Total Referees", "Total People", "Responses", "Referees",
]

BASE = "https://example.invalid/sites/submissions/Shared Documents"


def write_powerbi(path, entries):
    """The export the app checks against. Some figures are deliberately wrong."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Export"
    sheet.append(POWERBI_COLUMNS)
    for entry in entries:
        row = {c: None for c in POWERBI_COLUMNS}
        row.update({
            "MATCH_STATUS": "match",
            "SUBMISSION_FILE_NAME": f"final - {entry['area'].lower()}",
            "FIRM_NAME": entry["firm"],
            "FIRM_REF": entry["ref"],
            "CONCATENATED_PRACTICE_AREA_NAME": f"{entry['group']} > {entry['area']}",
            "COUNTRY_NAME": entry["region"],
            "FIRM_RANKING_TYPE": "firm recommended",
            "FIRM_RANKING_TIER": "2",
            "NUM_MATTERS": entry["export"]["matters"],
            "NUM_ACTIVE_CLIENTS": entry["export"]["active_clients"],
            "NUM_NEW_CLIENTS": entry["export"]["new_clients"],
            "NUM_LEAD_PARTNERS": entry["export"]["lead_partners"],
            "NUM_NEXT_GEN": entry["export"]["next_gen"],
            "NUM_ASSOCIATES": entry["export"]["associates"],
        })
        sheet.append([row[c] for c in POWERBI_COLUMNS])
    # Junk rows, as the real export carries.
    sheet.append([None] * len(POWERBI_COLUMNS))
    sheet.append(["Applied filters:\nEDITION_YEAR is 2027"] + [None] * (len(POWERBI_COLUMNS) - 1))
    workbook.save(path)


def write_portal(path, entries, extra_rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "portal-export"
    sheet.append(PORTAL_COLUMNS)
    for i, entry in enumerate(entries + extra_rows):
        row = {c: None for c in PORTAL_COLUMNS}
        has_document = entry.get("filename") is not None
        row.update({
            "ID": f"synthetic-{i:04d}",
            "Firm": entry["firm"],
            "Firm Ref": entry["ref"],
            "Publication": "United Kingdom - Solicitors",
            "Edition": EDITION,
            "Region": entry["region"],
            "Practice Group Name": entry["group"],
            "Practice Area Name": entry["area"],
            "Table Name": entry.get("table", entry["area"]),
            "Practice Group ID": 100000 + i,
            "Practice Area ID": 200000 + i,
            "Created At": entry.get("created", "2026-03-01 00:00:00"),
            "Submission Type": "firm",
            "Referee Type": "firm",
            "Document": "Y" if has_document else "N",
            "File URI": (
                f"{BASE}/{entry['region']}/{entry['group']}/{entry['filename']}"
                if has_document else None
            ),
        })
        sheet.append([row[c] for c in PORTAL_COLUMNS])
    workbook.save(path)


# --- entry point ------------------------------------------------------------


def main():
    OUT.mkdir(parents=True, exist_ok=True)

    plan = [
        {
            "key": "alpha", "build": build_alpha,
            "firm": "Aldgate Fenwick LLP", "ref": "90001",
            "region": "South East", "group": "Real estate",
            "area": "Commercial property: Thames Valley",
            "filename": "aldgate-fenwick_submission_2027_alpha.docx",
            # the export undercounts the nominations, as the real one does
            "export": {"matters": 5, "active_clients": 5, "new_clients": 1,
                       "lead_partners": 1, "next_gen": 0, "associates": 0},
        },
        {
            "key": "beta", "build": build_beta,
            "firm": "Aldgate Fenwick LLP", "ref": "90001",
            "region": "South West", "group": "Corporate and commercial",
            "area": "EU and competition",
            "filename": "aldgate-fenwick_submission_2027_beta.docx",
            "export": {"matters": 6, "active_clients": 6, "new_clients": 3,
                       "lead_partners": 2, "next_gen": 0, "associates": 0},
        },
        {
            "key": "gamma", "build": build_gamma,
            "firm": "Aldgate Fenwick LLP", "ref": "90001",
            "region": "London", "group": "Finance",
            "area": "Banking and finance",
            "filename": "aldgate-fenwick_submission_2027_gamma.docx",
            "export": {"matters": 4, "active_clients": 4, "new_clients": 0,
                       "lead_partners": 1, "next_gen": 0, "associates": 0},
        },
    ]

    manifest = {
        "documents": {},
        "portal_rows_without_document": [],
        "firms": {},
        "superseded": {},
        "shared_document": {},
    }
    for entry in plan:
        path = OUT / entry["filename"]
        expected = entry["build"](path)
        manifest["documents"][entry["filename"]] = {
            "firm": entry["firm"],
            "firm_ref": entry["ref"],
            "region": entry["region"],
            "practice_area": f"{entry['group']} > {entry['area']}",
            "expected": expected,
            "export": entry["export"],
        }
        manifest["firms"][entry["key"]] = entry["firm"]

    # An earlier submission for the same slot as alpha, so the newest-wins rule
    # has something to work on.
    older_name = "aldgate-fenwick_submission_2027_alpha_v1.docx"
    build_alpha(OUT / older_name)
    manifest["superseded"] = {
        "older": older_name,
        "newer": plan[0]["filename"],
        "practice_area": f"{plan[0]['group']} > {plan[0]['area']}",
    }

    extra = [
        # The same slot, pointing at the earlier document.
        {
            "firm": "Aldgate Fenwick LLP", "ref": "90001", "region": "South East",
            "group": plan[0]["group"], "area": plan[0]["area"],
            "filename": older_name, "created": "2026-01-15 00:00:00",
        },
        # A practice area with no submission document at all.
        {
            "firm": "Aldgate Fenwick LLP", "ref": "90001", "region": "South East",
            "group": "Employment", "area": "Pensions", "filename": None,
        },
        # The portal lists some practice areas twice against the same document.
        {
            "firm": "Aldgate Fenwick LLP", "ref": "90001", "region": "South West",
            "group": plan[1]["group"], "area": plan[1]["area"],
            "filename": plan[1]["filename"], "created": "2026-03-01 00:00:00",
        },
    ]
    # A second firm, present in both exports but with no documents in the
    # batch, so the firm scoping has something to exclude.
    other = {
        "firm": "Brackenmoor Legal", "ref": "90002", "region": "London",
        "group": "Finance", "area": "Banking and finance",
        "filename": "brackenmoor-legal_submission_2027_other.docx",
        "export": {"matters": 1, "active_clients": 1, "new_clients": 0,
                   "lead_partners": 0, "next_gen": 0, "associates": 0},
    }
    extra.append(other)
    manifest["other_firm"] = other["firm"]
    # A submission the portal names with a third level, as it does for some
    # practice areas: group, area, and a table of its own.
    three_level = {
        "firm": "Aldgate Fenwick LLP", "ref": "90001", "region": "South East",
        "group": "Transport", "area": "Travel", "table": "Travel: personal injury",
        "filename": "aldgate-fenwick_submission_2027_alpha.docx",
        "created": "2026-04-20 00:00:00",
    }
    extra.append(three_level)
    manifest["three_level_row"] = "Transport > Travel > Travel: personal injury"
    # A table name punctuated differently from the area it belongs to. Read
    # literally this is a third level, but it names the same thing twice, and
    # the Staff Portal page prints only the bracketed spelling.
    punctuated = {
        "firm": "Aldgate Fenwick LLP", "ref": "90002", "region": "South East",
        "group": "Energy and natural resources",
        "area": "Electricity (and renewable energy)",
        "table": "Electricity and renewable energy",
        "filename": "aldgate-fenwick_submission_2027_alpha.docx",
        "created": "2026-04-20 00:00:00",
    }
    extra.append(punctuated)
    manifest["punctuated_table_row"] = {
        "practice_area": punctuated["area"],
        "table_name": punctuated["table"],
    }
    # A table name built from the group and the area joined together. Nothing
    # repeats, so the name stays three levels deep and the practice area the
    # Staff Portal page prints is the middle one.
    derived = {
        "firm": "Aldgate Fenwick LLP", "ref": "90003", "region": "South East",
        "group": "City focus - Brasilia",
        "area": "Government relations",
        "table": "City focus - Brasilia - Government relations",
        "filename": "aldgate-fenwick_submission_2027_alpha.docx",
        "created": "2026-04-20 00:00:00",
    }
    extra.append(derived)
    manifest["derived_table_row"] = {
        "practice_group": derived["group"],
        "practice_area": derived["area"],
        "table_name": derived["table"],
    }
    manifest["portal_rows_without_document"] = ["Employment > Pensions"]
    manifest["shared_document"] = {
        "filename": plan[1]["filename"],
        "practice_area": f"{plan[1]['group']} > {plan[1]['area']}",
        "portal_rows": 2,
    }

    # The newer alpha is dated after the older one.
    plan[0]["created"] = "2026-04-20 00:00:00"

    write_powerbi(OUT / "powerbi-export.xlsx", plan + [other])
    write_portal(OUT / "portal-export.xlsx", plan, extra)

    (OUT / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
    print(f"written to {OUT}")
    for name in sorted(p.name for p in OUT.iterdir()):
        print("   ", name)


if __name__ == "__main__":
    main()
