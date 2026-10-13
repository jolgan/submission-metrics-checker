"""Role labels inside a matter box are not nominations.

A matter box lists the team that worked on it under "Lead partner(s)", in a
Name / Office / Practice area table. Firms write the role into the name cell:
"Associate: Lucia Lagana", "Counsel: Estefania Sierra". Those read exactly like
nomination labels, and one submission reported nine leading associates where
there was one - shown as a confident correction against the export's figure,
with nothing in "needs a look" to question it.

They are told apart by where they sit. A nomination box has never shared a
table with a matter, in any sample document; a matter's team always does.
"""

import pytest
from docx import Document

from parsing import parse_document


def submission(tmp_path, team_rows, nominations=("Associate: leading associate 1",)):
    document = Document()
    document.add_paragraph("Firm Name")
    document.add_table(rows=1, cols=1).rows[0].cells[0].text = "Example Firm LLP"

    document.add_paragraph("Clients: publishable clients")
    clients = document.add_table(rows=2, cols=2)
    clients.rows[0].cells[0].text = "Active key clients (over the last 12 months)"
    clients.rows[0].cells[1].text = "New client (yes/no)"
    clients.rows[1].cells[0].text = "A Client Ltd"
    clients.rows[1].cells[1].text = "Yes"

    document.add_paragraph("Your team - Associates: leading associates")
    for label in nominations:
        box = document.add_table(rows=3, cols=2)
        box.rows[0].cells[0].text = label
        box.rows[1].cells[0].text = "Name"
        box.rows[1].cells[1].text = "Location"
        box.rows[2].cells[0].text = "Lucia Lagana"
        box.rows[2].cells[1].text = "Mexico City"

    # A matter box, with its team listed inside it.
    matter = document.add_table(rows=5 + len(team_rows), cols=3)
    matter.rows[0].cells[0].text = "Publishable matter 1"
    matter.rows[1].cells[0].text = "Name of client"
    matter.rows[1].cells[1].text = "Industry sector"
    matter.rows[2].cells[0].text = "A Client Ltd"
    matter.rows[3].cells[0].text = "Lead partner(s)"
    matter.rows[4].cells[0].text = "Name"
    matter.rows[4].cells[1].text = "Office"
    matter.rows[4].cells[2].text = "Practice area"
    for index, name in enumerate(team_rows, start=5):
        matter.rows[index].cells[0].text = name
        matter.rows[index].cells[1].text = "Mexico City"
        matter.rows[index].cells[2].text = "Real Estate"

    path = tmp_path / "doc.docx"
    document.save(path)
    return parse_document(path).metrics


TEAM = [
    "Partner: Adrian Villagomez",
    "Associate: Lucia Lagana",
    "Associate: Victoria Vera",
    "Counsel: Estefania Sierra",
]


def test_a_matter_team_does_not_add_associates(tmp_path):
    """The reported shape: one leading associate counted as nine."""
    metrics = submission(tmp_path, TEAM)
    assert metrics["associates"].value == 1
    assert metrics["associates"].evidence == [
        "Associate: leading associate 1 - Lucia Lagana"
    ]


def test_a_matter_team_does_not_add_partners(tmp_path):
    metrics = submission(tmp_path, TEAM)
    assert metrics["lead_partners"].value == 0


def test_the_skipped_labels_are_reported(tmp_path):
    """Counted wrongly it was silent, so being ignored must not be."""
    metrics = submission(tmp_path, TEAM)
    note = " ".join(metrics["associates"].notes)
    assert "inside matter tables" in note
    assert "Associate: Lucia Lagana" in note


def test_the_matter_itself_is_still_counted(tmp_path):
    metrics = submission(tmp_path, TEAM)
    assert metrics["matters"].value == 1


def test_a_nomination_box_on_its_own_is_untouched(tmp_path):
    """The rule keys on sharing a table with a matter, nothing else."""
    metrics = submission(tmp_path, [], nominations=(
        "Associate: leading associate 1",
        "Associate: rising star 2",
    ))
    assert metrics["associates"].value == 2
    assert not any("inside matter tables" in n for n in metrics["associates"].notes)


def test_a_barrister_credited_on_a_matter_is_not_an_associate(tmp_path):
    """Counsel from another chambers, under "Other firms advising"."""
    metrics = submission(tmp_path, [
        "Counsel: Brian Nicholson KC and Kyra Nezami, both of 11 South Square",
    ])
    assert metrics["associates"].value == 1
