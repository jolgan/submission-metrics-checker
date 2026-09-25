# Submission recount

A local Streamlit app that recounts six metrics from law firm submission
documents and compares them against an automated extraction, highlighting only
the figures that disagree.

![The results table](docs/screenshots/results-table.png)

*Screenshots live in `docs/screenshots/`.*

## The problem

Law firms submit Word documents describing their practice: client lists, work
highlights, and nominations for leading partners, next generation partners and
leading associates. An automated system reads those documents and writes six
counts per submission into a Power BI export.

That extraction is unreliable, so a human re-checks it. Done by hand, that
means opening every document, scrolling through tables that run to dozens of
pages, counting rows, and comparing each figure against a spreadsheet, for a
hundred or more documents per firm. It is slow, and it is exactly the kind of
work attention drifts on.

The manual process is documented, and several of its steps are what this app
automates:

| The documented step | What the app does |
|---|---|
| Sort by region, practice group, practice area and created date, then keep the newest of any duplicate | Collapses a firm, region and practice area to one row on the portal's `Created At` |
| Check the order of practice areas matches the corrections spreadsheet | Orders the table from the portal page pasted into the app |
| Split the File URL column on `/` to recover the file names | Reads the File URI directly, hyperlink or plain text |
| Open each submission and check its figures against the spreadsheet | Recounts from the document and shows only what disagrees |

So the tool implements an established procedure rather than a private one, which
matters for anyone picking the checking up.

## The approach

Recount from the source documents, reconcile against the export, and show only
what disagrees.

The app takes the submission documents and both exports, recounts each metric
from the document itself, matches each document to its export row, and produces
one table with the corrected figures. Cells the recount changed are highlighted.
Everything else is left alone, so there is nothing to read where nothing has
changed.

Everything runs locally. The documents are confidential, so nothing is uploaded
and no external service is called.

## How it works

### Parsing

Each metric maps to one Power BI column:

| Column | Export column | Counted from |
|---|---|---|
| `MATTERS` | `NUM_MATTERS` | tables labelled `Publishable matter N` / `Non-publishable matter N` |
| `CLIENTS` | `NUM_ACTIVE_CLIENTS` | rows of the `Active key clients` tables, publishable and not |
| `NEW` | `NUM_NEW_CLIENTS` | the yes/no column of those same tables |
| `LPs` | `NUM_LEAD_PARTNERS` | nominations labelled `Partner: leading ...` |
| `NGs` | `NUM_NEXT_GEN` | nominations labelled `Partner: next generation ...` |
| `LAs` | `NUM_ASSOCIATES` | nominations labelled `Associate: ...` |

Everything anchors on the label in a table's first cell. Section headings are
not usable: one submission files its associate nominations under a heading
reading "Your team - Partners: leading partners". Neither are the numbers in
those labels, which repeat and skip.

The practice area is not in a paragraph or a table either. It is a Word dropdown
holding `Region - Practice group - Practice area`, and firms often leave it
unset and type into the box below instead.

### Matching

Documents are matched to export rows on **firm, region and practice area
together**, all normalised. Reference data (firm, region, practice area, file
name) comes from the Staff Portal export, matched on file name, with the Power
BI export filling any gaps. The document's own fields are a fallback and a
cross-check, never the primary source.

### Output

One row per firm, region and practice area:

```
FIRMREF  FIRM  REGION  PRACTICE AREA  FILE NAME  MATTERS  CLIENTS  NEW  LPs  NGs  LAs
```

Corrected cells are green; cells that could not be counted are amber and read
`double-check manually`. The download writes those fills as real cell
formatting, so the highlighting survives being pasted into a shared sheet.

Paste the portal's own page listing into the app and it also reconciles the
batch: which rows have no document to download, which documents are missing,
and which files are duplicates of one already read.

## Install and run

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python tools/make_synthetic_samples.py
streamlit run app.py
```

Opens at <http://127.0.0.1:8501>. Drag in submission documents, the Power BI
export, and optionally the Staff Portal export.

The app must be restarted, not just refreshed, after changes to `parsing.py`,
`matching.py` or `export_writer.py`. Streamlit keeps imported modules in memory
for the life of the process.

## Project structure

```
app.py                        Streamlit interface, and nothing else
parsing.py                    reads .docx and counts; no Streamlit import
matching.py                   loads both exports, normalises, joins, reconciles
export_writer.py              builds the highlighted .xlsx
docx_repair.py                opens packages with inconsistent internal names
diagnose.py                   command line: what the parser saw in one document
compare_exports.py            command line: what changed between two exports
tools/make_synthetic_samples.py   generates the sample data
samples/synthetic/            invented documents and exports, used by the tests
tests/                        pytest suite
```

Parsing and matching hold no Streamlit import, so both are testable without
launching the app.

## Testing

```bash
python -m pytest tests/ -q
```

295 tests, running against `samples/synthetic/` alone, so a clone can run them
without any confidential data. Expected figures come from
`samples/synthetic/manifest.json`, written by the same script that builds the
documents, so the two cannot drift apart.

Many of the tests pin a specific fault found in a real submission: two matters
sharing one table, an answer reading `Yes (for this work type)`, a client row
reading `N/A`, a nomination labelled `Associate : leading counsel 1`. Each one
was a miscount that looked correct on screen.

## Design decisions

**Match on firm and practice area together, never practice area alone.** In one
firm's export, nine practice areas appeared under more than one region with
different figures. Practice area alone silently matches the wrong row. Region is
part of the key for the same reason.

**Normalise before comparing.** Lowercase, trim, collapse repeated whitespace,
standardise spacing around `>`. The same practice area is written
`Real estate > Planning` in one file and `Real estate>Planning` in another.

**Flag what cannot be counted rather than defaulting to zero.** A metric the
parser cannot establish reads `double-check manually`, not `0`. A zero is
indistinguishable from a real count and quietly becomes a wrong answer; the
marker asks for a human. An unparsed metric is never reported as a difference
from the export, because it is not a correction, it is a question.

**Count labels, not tables.** Word merges two adjacent tables into one when the
paragraph between them is deleted, which firms do while editing. The second
entry's label then survives only as a row part-way down the merged table.
Reading the first row alone counts one where there are two, and says nothing.

**Report ambiguity instead of resolving it.** Where a firm has two submissions
for the same slot, the newer is counted and the older is set aside. Where the
dates are equal, nothing is guessed: one row is kept and flagged for a human.
A nomination whose description says neither leading nor next generation is left
uncounted and named, because the two categories are not interchangeable.

## Limitations

Handles the Legal 500 UK submission template as firms return it, including:
merged tables, nominations inside Word content controls, repeated header rows
inside client tables, `New`/`Existing` in place of `Yes`/`No`, qualified answers
such as `Yes (for this work type)`, `N/A` in a client column, numbering that
skips or repeats, and packages whose internal part names are inconsistently
capitalised.

It would fail on:

- a different submission template, or a substantial redesign of this one. The
  parser anchors on label text such as `Publishable matter 1` and
  `Active key clients`; rename those and it finds nothing.
- a practice area whose row cannot be identified. If a firm leaves the dropdown
  unset and the file name is absent from the Staff Portal export, there is
  nothing left to identify the row by, and it is reported unmatched.
- a nomination described in words that say neither leading nor next generation.
  It is reported, not counted.
- matching a pasted portal page by practice area alone where a firm has the same
  practice area on several rows. The app prefers a row whose document was
  uploaded, but it cannot always tell which row a pasted line meant.

No accuracy figure is claimed here. The app reports where it disagrees with the
export; deciding which is right is still the checker's job.

## Sample data

The repository contains **synthetic sample data only**. Every firm, person,
client and figure in `samples/synthetic/` is invented. Real submissions are
confidential and never leave the machine they are checked on; `.gitignore`
excludes `samples/` along with every spreadsheet and document extension.

## Licence

MIT. See [LICENSE](LICENSE).
