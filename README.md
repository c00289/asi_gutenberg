# asi_gutenberg

One of the [ASI FuSa tools](https://github.com/c00289/c00289/blob/main/README.md) by Roy Mendoza; the master README there lists every sibling.

Pure-Python document generator for engineering work products. Build a
specification, report, plan, procedure, minutes, memo, safety case, ICD,
analysis, checklist or release notes from a small factory API, and
asi_gutenberg lays it out (section, table and figure numbering, front matter,
table of contents, requirement trace checks) and writes **Markdown**, **PDF**
and **Word** together as a triplet. A document can carry its own figures:
a `diagram` block holds an [asi_escher](../asi_escher) diagram and is
rendered to a PNG beside the Markdown at build time.

Markdown is the source of truth: the PDF (reportlab) and the Word file
(python-docx) are rendered from the Markdown that was just written, never
from the in-memory model, so the three files can never disagree.

| Block kinds | Document kinds |
|---|---|
| heading, paragraph, bullet / numbered / task list, table, field table, code, callout (note, tip, important, warning, caution), requirement, figure, diagram (asi_escher), quote, rule, page break, section break (portrait / landscape), table of contents, raw Markdown | spec, report, plan, procedure, minutes, memo, safety_case, icd, analysis, checklist, release_notes |

## Install

Requires Python 3.13 or newer. The runtime dependencies are reportlab,
python-docx and the sibling asi_escher (which brings matplotlib and Pillow
for the diagrams). asi_escher is declared as an editable path source,
`../asi_escher`, so the two checkouts must sit side by side; see
[ASI_FUSA_TOOLS.md](../asi_fusa_tools/ASI_FUSA_TOOLS.md) for the layout of
the tools and the uv trap that every consumer must declare both sources
itself.

```bash
# for development inside this folder
uv sync

# as a dependency of another project, from a checkout
uv add --editable ../asi_gutenberg      # or: pip install ../asi_gutenberg
```

This repository is meant to sit next to its siblings:

```
OneDrive - Capgemini/
├── asi_fusa_tools/   <- charter: ASI_FUSA_TOOLS.md, no code
├── asi_dewey/        <- primary tool: working knowledge
├── asi_kletz/        <- primary tool: safety plan
├── asi_escher/       <- generator: diagrams (asi_gutenberg depends on it)
├── asi_gutenberg/    <- generator: documents (this repo)
├── asi_hopper/       <- gateway: read-only Jira door
└── Roy's Brain/      <- notes, corpus, datasets (not a repo)
```

Each consumer declares the dependency once with a uv path source, for
example `uv add --editable ../asi_gutenberg` from any sibling folder. OneDrive
rejects hardlinks from the uv cache, so `pyproject.toml` pins
`tool.uv.link-mode = "copy"`; pass `--link-mode=copy` to ad-hoc uv commands
run elsewhere under OneDrive.

## Quick start

```python
from asi_gutenberg import Document, build

d = Document("E-stop Reset Procedure", "procedure",
             doc_id="ASI-FS-PROC-004", revision="0.1", status="Draft",
             date="2026-10-27", author="Roy", project="WA485 Autonomous Wheel Loader",
             approvers=[("Functional Safety Manager", "Max Stone")])
d.revision("0.1", "2026-10-27", "Roy", "Initial draft")

d.h1("Purpose")
d.p("This procedure restores the loader to **STANDBY** after an emergency stop.")
d.h1("Steps")
d.numbered_list(["Confirm the exclusion zone is clear.",
                 "Release every latched E-stop button.",
                 "Press RESET on the safety HMI and wait for the green lamp."])
d.warning("Do not reset while anyone is inside the exclusion zone.")
d.requirement("PROC-04-01", "Two-person reset",
              "A reset shall be witnessed by a second trained operator.", plr="c", status="Approved")

paths = build(d, "out")        # out/e_stop_reset_procedure.{md,pdf,docx}
print(paths["pages"], "pages,", paths["issues"] or "no issues")
```

`build()` returns the three output paths, the PDF page count, the layout
statistics, any issues found (for example a requirement that traces to an
id that does not exist, or a diagram whose layout had to degrade) and one
entry per rendered diagram under `diagrams`.

### Step by step

The pipeline is plain functions you can call yourself:

```python
from asi_gutenberg import layout, to_markdown, from_markdown, render_pdf, render_docx, LayoutOptions

layout(d, LayoutOptions(number_depth=2, approvals=False))
md = to_markdown(d)
parsed = from_markdown(md)
render_pdf(parsed, "out/doc.pdf")
render_docx(parsed, "out/doc.docx")
```

`Document.to_dict()` / `Document.from_dict()` give a JSON interchange form,
and `asi_gutenberg.theme.THEME` is a plain dictionary of colours, fonts and
sizes that either renderer accepts overrides for (`render_pdf(..., theme={...})`).
The palette is the asi_escher palette, so diagrams match the text around them.

## Diagrams

A `diagram` block carries an asi_escher diagram; `build()` renders it to a
PNG beside the Markdown and the three outputs show it as a numbered figure,
sharing the counter with `figure` blocks:

```python
from asi_escher import Diagram
from asi_gutenberg import Document, build

d = Document("E-stop Reset Procedure", "procedure", doc_id="ASI-FS-PROC-004")
flow = Diagram("E-stop handling", "activity")
flow.lane("op", "Operator"); flow.lane("sc", "Safety controller")
flow.start(lane="op"); flow.action("press", "Press E-stop", lane="op")
flow.action("inhibit", "Open drive inhibit relay", lane="sc"); flow.end(lane="sc")
flow.chain("start", "press", "inhibit", "end")

d.h1("Overview")
d.diagram(flow, "E-stop handling across operator and safety controller")
result = build(d, "out")   # out/e_stop_reset_procedure.{md,pdf,docx} + out/figures/e_stop_handling.png
print(result["diagrams"])  # [{"name": "e_stop_handling", "png": "...", "size": (w, h), "nodes": 4, "edges": 3, "issues": []}]
```

`Document.diagram(spec, caption="", width=0.0, name="", formats=("png",), dpi=160)`
takes a `Diagram` or its `to_dict()` form. The block stores the plain-data
spec, never the object, so `Document.to_dict()` / `from_dict()` carry it and
an agent can send a document spec with diagrams inside it as one JSON file;
a dict spec is validated through `asi_escher.Diagram.from_dict` when the
block is added, so an unknown node kind fails at once. `name` is the output
stem (default: a slug of the diagram title; two diagrams may not share a
stem), `formats` adds `svg`, `dot` or `mmd` files beside the PNG, and `dpi`
sets the PNG resolution (160 matches asi_escher; raise it for print).

Conventions:

- Diagrams are written to `<out_dir>/<figures_dir>/<stem>.png`, with
  `figures_dir="figures"` by default (`build(d, out, figures_dir="img")`
  changes it; `""` writes beside the Markdown).
- The Markdown line is `![Figure n: caption](figures/stem.png)`: a relative
  path with forward slashes, so GitHub renders it and the whole output folder
  can be moved as one unit. The PDF and Word renderers embed the file through
  the same relative path (`base_dir`); if it is missing they draw a
  placeholder, as for any `figure`.
- `build(..., diagram_theme={...})` overrides entries of `asi_escher.THEME`
  for every diagram, so a document with its own palette can restyle its
  figures in the same call.
- `LayoutOptions(mermaid=True)` adds a ```` ```mermaid ```` fence after each
  image so the `.md` shows the diagram on GitHub even when the PNG is not
  committed; the parser skips that fence, so the PDF and Word files are
  unchanged.
- A degraded diagram layout (asi_escher's `label_overlap` or
  `edge_through_node`) is never hidden and never raised: asi_escher's
  `LayoutWarning` is captured, and each issue appears as
  `"diagram <stem>: <message>"` in `result["issues"]` and, in full, under
  `result["diagrams"][i]["issues"]`.
- asi_escher (and matplotlib with it) is imported only when a dict spec is
  validated or a diagram is rendered; `import asi_gutenberg` stays light.

The step is also a plain function: `render_diagrams(d, out_dir,
figures_dir="figures", theme=None, mermaid=False)` after `layout(d)` and
before `to_markdown(d)`.

## Importing a Word document, and parameters

Since 0.3.0 a controlled document that exists only as a Word file can be brought in and
parametrised. `from_docx` reads the `.docx` into the model, faithfully: the title, every
`Heading n` as a section (numbered when Word numbered them), paragraphs, consecutive list
paragraphs as one list, every table (a two-column table of labels becomes a `Field` / `Value`
table), inline pictures as figures written beside the source, the `Document ID` and
`Document owner` rows of a front table as `doc_id` and `author`, and a footer date stamp when
it is unambiguous. What it cannot keep it says in `issues` (page orientation per section,
header and footer text, fonts). A section break that changes the page orientation becomes a
`section` block (`<!-- section: landscape -->`), so a plan table on landscape pages inside a portrait
document comes out the same way in the PDF and the Word file. The result is a Markdown source that
rebuilds to the triplet.

A parameter is `{{ name }}` anywhere text goes: metadata, headings, paragraphs, list items,
table cells and captions, callouts. A filled value is written as `==value==` and both renderers
highlight it (the theme's `mark_bg`, a green), so a replacement is easy to tell from the text around
it; `--no-mark` switches that off. `parameters(doc)` lists them; `fill(doc, values)` returns a
filled copy and refuses, naming them, the ones it has no value for; `build(doc, out,
params=values)` fills first and reports anything left as `unfilled parameter: <name>`.

```bash
python -m asi_gutenberg import "QF-17 Safety Plan.docx" ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md
python -m asi_gutenberg params ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md
python -m asi_gutenberg build ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out --param program=WA485 --landscape
python -m asi_gutenberg build ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out --params-file values.json --landscape
python -m asi_gutenberg fill ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out/qf-17-safety-plan.docx --params-file values.json --pdf
python -m asi_gutenberg fill ../asi_standards/qf-17-safety-plan/qf-17-safety-plan.md out/qf-17-safety-plan.docx --params-file values.json --pdf --abridge-headers headings.json
```

### The template exactly: filling the Word original in place

The build above renders the source with Gutenberg's own theme. When the document must look exactly like
the controlled template (the owner, 2026-10-29: "match the template exactly"), `fill_docx(template,
source, values, out)` and the `fill` command write the Word original itself with the source's text put
into its own runs: every style, header, footer, table, column width and page orientation stays the
original's, because nothing else is touched. The body is walked as the importer walks it, so each
paragraph and cell is found again in order; the source must still line up unit for unit, else the fill
refuses and names the first unit that does not. Only the span that changed inside a paragraph or cell
is rewritten, in new runs after the run where it starts, so a bold prefix or an italic tail keep their
formatting; a cell written line by line gets one paragraph per line, the cell's own. The result is
`Filled`: the Word file, the PDF when Word printed one, how many units changed, the placeholders left
unfilled, the `issues` in words, and (when Word ran) the tables split for abridged headers and Word's
page count; the `fill` command prints the same as one JSON object.

| `fill` flag | What it does |
|---|---|
| `--template DOCX` | the Word original (default: the one `.docx` under `source/` beside the Markdown, the asi_standards layout) |
| `--param NAME=VALUE`, `--params-file JSON` | the values, as for `build`; a later file or `--param` wins |
| `--allow-unfilled` | fill even when a placeholder has no value (it stays, and is listed) |
| `--mark` | shade every filled value green on its own run, as the rendered documents do; off, the file is the template exactly |
| `--no-repeat-header` | do not repeat each table's first row at the top of every page it spans (on by default: `repeat_header`) |
| `--pdf` | also print the PDF beside the Word file through Word itself (Windows with Word installed): Word's own print of the filled file, not a rendering |
| `--abridge-headers JSON` | verbatim column heading to short form; the table shows the verbatim header on its first page and the short forms on every later page (needs `--pdf`) |

**A link in a cell.** The owner, 2026-10-29: "could we turn the name of the safety deliverable into a
clickable URL and highlight it blue instead of green so it pops?" The source writes an inline link the
way every rendered output already reads it, `[Safety Plan]({{ row_1_url }})`, the url a parameter like
any other (filled, it is never marked). In the filled Word original the link segment becomes a real
`w:hyperlink` (an external relationship to the url) whose run carries the paragraph's own run properties
plus Word's `Hyperlink` character style (added to the document when the template has none), Word's blue
`#0563C1`, a single underline and the theme's pale blue shading `link_bg` (`#CFE2F3`); a mark never
applies inside a link, so it is blue, not green, and the PDF Word prints keeps it clickable. A link whose
url is empty or still holds its `{{ placeholder }}` collapses to its plain text, no link and no mark, so
the cell is the template's exactly, and the url is listed as unfilled. The Word rendering and the PDF
rendering paint a link the same way, blue over the pale blue shading. For the QF-17 every plan row's
Safety Deliverables cell is `[<name>]({{ row_N_url }})`, the row's first work product URL as asi_kletz
supplies it (`python -m asi_gutenberg fill ... --mark`).

**Verbatim, then abridged.** A wide plan table carries wordy column headings, and repeating them on every
page costs a third of each (the owner, 2026-10-29: "it would be amazing if the safety plan table headers
were verbatim in the first page and then abridged in all subsequent, because they are wordy").
`fill_docx(..., abridged_headers={verbatim: short})` and `fill --abridge-headers headings.json` (a JSON
object, verbatim heading to short form) do that: a table whose first row is wholly in the mapping (every
cell with text has a short form; headings are compared with their whitespace normalised, so a heading
that wraps over lines is written as one line) shows its verbatim header on its first page, once, and the
short forms at the top of every later page. A table only partly in the mapping is refused naming the
headings without a short form (a typo, not a table to skip); so is a mapping no table's first row is in.
Word cannot repeat a different header on the later pages of one table, and the row where page one ends is
known only once Word has laid the document out, so this is a Word pass in the same session that prints
the PDF: Word reports the page of every row, the table is split after the last row of its first page, the
first part keeps the verbatim header with its repeat turned off, the second gets a new first row with the
short forms in the verbatim header's own formatting and the repeat turned on, both parts keep the same
column widths with nothing but a 1pt empty paragraph between them, and the Word file is saved with the
split before the PDF is printed from it. A table that ends on its first page is left alone. The result's
`abridged` lists the tables split (by their first heading) and `pages` is Word's page count; without
Word, or without `--pdf`, the file keeps the verbatim header repeating and `issues` says "abridged
headers need Word". For the QF-17 the mapping is its ten headings (`PDP Phases` to `PDP`, `Safety
Lifecycle Phase` to `Lifecycle`, and so on); the plan table's first page keeps the verbatim header and
the document goes from 19 pages to 15.

**The Word session.** Everything Word itself does, the abridge pass and the print, is one session per
call (`word_session` in `word.py`, exported with `WordSession` and the plain `pdf_with_word`): a
PowerShell script drives Word over COM, opens the filled file, splits the planned tables, saves, reports
one JSON line and exports the PDF to a working name beside the target, which takes the target's place
only once the print is there, so a failed print leaves whatever PDF was there (asi_kletz writes
Gutenberg's rendering first and keeps it when Word cannot print). The Word process the session starts is
ended, and gone, before the session returns; one the person has open is never touched. Since 0.9.1 the
session's activation and open are retried on the RPC errors a Word still shutting down from the last
session causes, so two prints in a row both succeed. Every failure is an issue in words on the result,
never an exception: a machine without Word (or not Windows) gets its Word file and the reason.

The parametrised standards themselves live in **asi_standards**, a private sibling repo, never
here: one folder per document, the original Word file preserved byte for byte under `source/`
with its SHA-256 in `SHA256SUMS`, and the imported Markdown source beside it. The first is the
owner's QF-17 Safety Plan. The test suite checks that checkout when it is beside this repo and
skips it otherwise, so this public repo holds the machinery and none of the documents.

## Demos

`demos/catalog.py` holds fifteen builders, listed in `DEMOS` in this order
(the name is the output stem; the last column is the number of `diagram`
blocks the document carries):

| Name | Document | Diagrams |
|---|---|---|
| `spec_safety_requirements` | functional safety requirements specification | 2 |
| `icd_safety_plc` | safety PLC interface control document | 1 |
| `report_hara` | hazard analysis and risk assessment report | 1 |
| `report_fmea_brake` | brake system FMEA report | 1 |
| `report_post_test` | post-test report | 1 |
| `report_stopping_distance` | stopping-distance analysis report | 0 |
| `plan_functional_safety` | functional safety plan | 1 |
| `plan_verification_validation` | verification and validation plan | 0 |
| `procedure_estop_reset` | E-stop reset procedure | 1 |
| `checklist_release` | release checklist | 0 |
| `minutes_safety_review` | safety review minutes | 0 |
| `memo_change_notice` | change notice memo | 0 |
| `release_notes_fw` | firmware release notes | 0 |
| `safety_case_outline` | safety-case outline | 1 |
| `stress_long_document` | configurable stress-test document | 0 |

Render them all with:

```bash
uv run python demos/build_all.py          # writes out/<name>.{md,pdf,docx}
```

Eight of the documents carry asi_escher's bundled example diagrams as
`diagram` blocks, nine blocks in all (captions end in `(asi_escher)`); they
are rendered into `out/figures/` by the same call.

## Development

```bash
uv sync                      # creates .venv with the dev group
uv run pytest                # pipeline round trip, diagram block, every demo
uv run ruff check src tests
```

The suite is `tests/test_smoke.py` (Markdown round trip, the triplet is
written, a dangling trace is reported, every demo builds with its diagram
files in place), `tests/test_diagrams.py` (the `diagram` block: model
round trip, numbering, rendering, issue capture, portability, import
isolation), `tests/test_import_params.py` (the Word importer, the
parameters and their marks, the `section` block, the command line, and the
QF-17 source when asi_standards is checked out beside this repo),
`tests/test_fill_docx.py` (the in-place fill, the marks, the links, the
abridged headers, the command line; the Word prints and the abridge pass
when Word is installed, the QF-17 when asi_standards is beside this repo)
and `tests/test_version.py`. It runs with asi_escher's `LayoutWarning` as
an error, so a warning that escaped `render_diagrams` would fail it.

Release: bump `version` in `pyproject.toml`, add a `CHANGELOG.md` entry with
its compare link, tag `vX.Y.Z`. `tests/test_version.py` fails until the three
agree (`asi_gutenberg.__version__` reads the installed metadata, so `uv sync`
after the bump).
