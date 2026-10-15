# Submission recount

A local Streamlit app that recounts six metrics from law firm submission
documents and checks them against the figures an automated system produced,
highlighting only those that disagree.

Used on submissions from the United Kingdom and eight Latin American
countries.

**[Open it in your browser](https://www.jolenegan.com/submission-metrics-checker/)**
- the same code compiled to WebAssembly, with nothing to install. Documents are
read in the page and never leave the machine.

## The problem

Law firms submit Word documents describing their practice: client lists, work
highlights, and nominations for leading partners, next generation partners and
leading associates. An automated system reads those documents and writes six
counts per submission into a Power BI export. Those counts are what this app
checks.

Those counts are unreliable, so a human re-checks them. Done by hand, that
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
| `MATTERS` | `NUM_MATTERS` | matter boxes, however the firm has labelled them: numbered or not, `#Matter`, `Confidential matter`, `Non-published matter` |
| `CLIENTS` | `NUM_ACTIVE_CLIENTS` | rows of the `Active key clients` tables, publishable and not |
| `NEW` | `NUM_NEW_CLIENTS` | the yes/no column of those same tables, in English, Spanish or Portuguese |
| `LPs` | `NUM_LEAD_PARTNERS` | nominations whose role is partner and whose description says leading |
| `NGs` | `NUM_NEXT_GEN` | nominations whose role is partner and whose description says next generation |
| `LAs` | `NUM_ASSOCIATES` | nominations whose role is associate or counsel |

Everything anchors on the label in a table's first cell. Section headings are
not usable: one submission files its associate nominations under a heading
reading "Your team - Partners: leading partners". Neither are the numbers in
those labels, which repeat and skip.

A label is read for its role and its description separately, because firms
write both freely: `Associates: leading associate 1` carries the heading's
plural, `Partner 1: leading individual` numbers the box before the colon, and
`Counsel: leading counsel 1` uses a role the form's own instructions put in
the associates section.

Where the label sits matters as much as what it says. A matter box lists the
team that worked on it, and firms write the role into the name cell -
`Associate: Lucia Lagana` - which reads exactly like a nomination. Those are
left out, identified by the matter label or the `Other key team members`
heading above them, either of which Word may have split into a table of its
own.

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

You need **Python 3.12 or newer** and **git** installed first.

To check, open a terminal and type `python --version`. If it says "Python was
not found" or opens the Microsoft Store, install Python from
<https://www.python.org/downloads/> and tick **"Add Python to PATH"** on the
first screen of the installer.

### Where the commands go

- **Windows:** PowerShell (press Start, type `powershell`).
- **macOS or Linux:** the Terminal app.
- **Windows with WSL:** the Ubuntu app, and follow the macOS/Linux commands.

Pick the block for your system below and **paste the whole block in one go**.
Your terminal will run each line as it arrives. The two blocks are not
interchangeable: the commands genuinely differ between Windows and the rest.

### First time: download and set up

**Windows (PowerShell):**

```powershell
git clone https://github.com/jolgan/submission-metrics-checker.git
cd submission-metrics-checker
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python tools/make_synthetic_samples.py
```

**macOS, Linux or WSL:**

```bash
git clone https://github.com/jolgan/submission-metrics-checker.git
cd submission-metrics-checker
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python tools/make_synthetic_samples.py
```

You only do this once. If the Windows activate line is refused because
scripts are disabled, run this and then try the block again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### Every time: start the app

**Windows (PowerShell):**

```powershell
cd submission-metrics-checker
.venv\Scripts\Activate.ps1
streamlit run app.py
```

**macOS, Linux or WSL:**

```bash
cd submission-metrics-checker
source .venv/bin/activate
streamlit run app.py
```

Then open <http://127.0.0.1:8501> in your browser. Drag in submission
documents, the Power BI export, and optionally the Staff Portal export.

The app runs for as long as that terminal stays open. To **stop** it, click
into the terminal and press **Ctrl+C**. If you close the terminal window the
app stops with it, which is the usual reason it disappears.

Restart it, not just refresh the browser, after changes to `parsing.py`,
`matching.py` or `export_writer.py`. Streamlit keeps imported modules in memory
for as long as it runs.

### If something goes wrong

**"Python was not found"** - Python is not installed, or was installed without
"Add Python to PATH". Reinstall from python.org and tick that box.

**"The term 'source' is not recognized"** - that is the macOS/Linux line. Use
the Windows block.

**"running scripts is disabled on this system"** - run the
`Set-ExecutionPolicy` line above, then activate again.

**Packages install but the app will not start** - the virtual environment was
probably not active. The terminal prompt shows `(.venv)` at the start of the
line when it is. Run the activate line for your system and try again.

**"Port 8501 is already in use"** - the app is already running in another
terminal. Open <http://127.0.0.1:8501>, or start this one on a different port
with `streamlit run app.py --server.port 8502`.

### No installing at all

There is a browser version that needs none of the above:

<https://www.jolenegan.com/submission-metrics-checker/>

It runs the same code inside the browser, so documents stay on your machine.
It is roughly five times slower than running it locally, so it suits small
batches and demonstrations rather than a full firm.

## Project structure

```
app.py                        Streamlit interface, and nothing else
parsing.py                    reads .docx and counts; no Streamlit import
matching.py                   loads both exports, normalises, joins, reconciles
export_writer.py              builds the highlighted .xlsx
docx_repair.py                opens packages with inconsistent internal names
index.html                    runs the same app in a browser, via stlite
check_batch.py                command line: which documents a batch is missing
diagnose.py                   command line: what the parser saw in one document
compare_exports.py            command line: what changed between two exports
tools/make_synthetic_samples.py   generates the sample data
samples/synthetic/            invented documents and exports, used by the tests
tests/                        pytest suite
```

Parsing and matching hold no Streamlit import, so both are testable without
launching the app.

`check_batch.py` answers the question that comes up on every batch: a browser
reports every document downloaded, the folder holds fewer, and nothing says
which. It names them, and sets aside the ones a newer submission has already
replaced so they are not chased for nothing.

```bash
python check_batch.py "Example Firm"
```

## Testing

```bash
python -m pytest tests/ -q
```

559 tests, running against `samples/synthetic/` alone, so a clone can run them
without any confidential data. Expected figures come from
`samples/synthetic/manifest.json`, written by the same script that builds the
documents, so the two cannot drift apart.

Many of the tests pin a specific fault found in a real submission: two matters
sharing one table, a matter whose label Word split away from its body, an
answer reading `Yes (for this work type)`, a client row reading `N/A`, a
nomination labelled `Associate : leading counsel 1`, and a matter's own team
read as nine leading associates. Each one was a miscount that looked correct
on screen.

## Faults it was taught to catch

Every entry below is a real miscount found in a real submission, each one
looking perfectly correct on screen. Each became a test, so a later change
cannot quietly undo the fix - and several did try: adding counsel nominations
for one firm started counting barristers at another, and allowing plural roles
briefly filed associates as partners. Both were caught by tests written for
earlier bugs, before they reached anyone.

### How documents are built

| What the document did | What it cost | What happens now |
|---|---|---|
| Word merged two matter boxes into one table when the paragraph between them was deleted | 20 matters counted as 19, on 68 of 162 rows for one firm | Labels are counted wherever they appear, not only in the first row |
| Word split a matter, leaving its label alone in one table and its body in the next | 22 matters counted as 13 | A label with nothing under it adopts the unlabelled table that follows |
| A nomination sat inside a Word content control, invisible to a plain walk of the document | A whole category read as zero | The walk descends into content controls and nested tables |
| Seven files stored a part as `customXML` where the index said `customXml` | Seven documents would not open at all | The package is rewritten to refer to itself consistently |

### How firms fill them in

| What the document did | What it cost | What happens now |
|---|---|---|
| `Yes (for this work type)` in the new-client column | 9 new clients counted as 6 | Answers are read by their leading word; anything unrecognised is flagged, not guessed |
| `Sí`, `Si`, `Sim`, `Não` | The whole column read as uncountable | Spanish and Portuguese are read alongside English |
| `N/A` written where a client name goes | Counted as a client | Recognised as "none", not as a company |
| A memo to the researcher in place of the client list | Counted as a client | Flagged: the row answers nothing where its neighbours answer |
| `ISSUERS` and `Initial Purchasers` as headings inside a client list | 63 clients counted as 66 | Same flag, judged per table so firms that leave the column blank throughout are not flagged wholesale |
| Matter boxes left unnumbered, or written `#Matter 1`, `Confidential matter`, `Non-published matter` | The section was not found at all | All are recognised, while `Publishable matter summary` is not |
| A matter box left exactly as the form ships it, carrying its field headings and the instruction "Please describe your team's role" | Every document in a firm's batch counted one matter too many | A matter counts when a client is named in it, not when the box holds any text at all |

### How nominations are worded

| What the document did | What it cost | What happens now |
|---|---|---|
| `Counsel: leading counsel 1` | A nomination dropped in silence | Counsel counts as a leading associate, as the form's own instructions say |
| `Associates: leading associate 1`, with the heading's plural | Five nominations dropped | Singular and plural are the same role |
| `Partner 1: leading individual`, numbered before the colon | A firm's leading partners read as zero | The number may sit either side of the colon |
| `Associate: Lucía Laganá` in a matter's team list | One leading associate counted as nine, then as eleven, then thirteen | A role label counts only inside a nomination box, never where it credits someone on a matter |
| A description saying neither leading nor next generation | Would have had to be guessed | Reported by name and counted in neither category |

### Matching and batches

| What happened | What it cost | What happens now |
|---|---|---|
| A practice area named identically under two regions | The wrong export row matched, silently | Firm, region and practice area must all agree |
| The portal's table name differing from its area only by brackets | 11 firms' submissions matched nothing | Levels are compared on their words, and every level below the practice group is indexed |
| A firm resubmitting, leaving both rows on the portal page | Seven files chased that would be set aside on arrival | A row a newer submission replaces is not counted as still to download |
| The Power BI dashboard renamed every column but one | The export would not load | Columns are resolved by alias, falling back to the last six, and a partial rename is refused rather than assumed |
| A firm answering on Chambers' form instead | Three nomination counts read as zero | The form is read on its own terms, and every figure on it is flagged |

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

### A competitor's form

A few firms answer on Chambers' submission form instead. It asks for the same
things in a different shape: one row per lawyer with a `Partner Y/N` column
in place of a box per nomination, and client lists headed `D0 - PUBLISHABLE
CLIENTS` and `E0 - CONFIDENTIAL CLIENTS`. Y reads as a leading partner and N
as a leading associate.

It is recognised by its own section codes and upload address, never by our
template's tables being absent - otherwise a document of ours that failed to
parse would be silently counted the wrong way. Every figure on such a
submission is flagged whatever it counts, and the row reads
`Chambers template used - check every figure`.

### What it handles, and what it does not

Handles the submission template as firms return it, including:
merged tables, nominations inside Word content controls, repeated header rows
inside client tables, `New`/`Existing` in place of `Yes`/`No`, qualified answers
such as `Yes (for this work type)`, `N/A` in a client column, numbering that
skips or repeats, and packages whose internal part names are inconsistently
capitalised.

It would fail on:

- a submission template other than the two it knows. It reads the Legal 500
  form and Chambers' form; anything else anchors on label text such as
  `Publishable matter 1` and `Active key clients`, and renaming those leaves
  it with nothing to find.
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

## Try it with the sample data

Everything needed to see the app work is in `samples/synthetic/`, committed
so it can be downloaded without running anything first:

| File | What it is |
|---|---|
| `aldgate-fenwick_submission_2027_alpha.docx` and three more | submissions to recount |
| `powerbi-export.xlsx` | the automated counts to check against |
| `portal-export.xlsx` | the Staff Portal export, for file names and reference data |

Open the [browser version](https://www.jolenegan.com/submission-metrics-checker/),
drop in the four documents and both spreadsheets, and the table fills in. The
figures are wrong in the export on purpose, so corrections show up green.

For the pasted-page panel, paste these practice areas into the box:

```
Commercial property: Thames Valley
EU and competition
Banking and finance
```

The documents are invented, and so is every firm, person, client and figure
in them. `alpha` and `alpha_v1` are deliberately the same submission twice,
to show the newer one being kept.

## Sample data

The repository contains **synthetic sample data only**. Every firm, person,
client and figure in `samples/synthetic/` is invented. Real submissions are
confidential and never leave the machine they are checked on; `.gitignore`
excludes `samples/` along with every spreadsheet and document extension.

## Licence

MIT. See [LICENSE](LICENSE).
