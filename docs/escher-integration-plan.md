# asi_gutenberg — Plan: integrate asi_escher (diagrams as document blocks)

**Status:** Implemented in asi_gutenberg 0.2.0 (2026-10-27); kept as the design record
**Written:** 2026-10-27
**Owner:** FuSa engineer
**Target version:** asi_gutenberg 0.2.0
**Family context:** [ASI_FUSA_TOOLS.md](../../asi_fusa_tools/ASI_FUSA_TOOLS.md)
**Consumers waiting on this:** [asi_dewey](../../asi_dewey/docs/escher-gutenberg-integration-plan.md)
and [asi_kletz](../../asi_kletz/docs/escher-gutenberg-integration-plan.md), phase 4 of each.

---

## 1. Goal

A Gutenberg document can **carry its own diagrams**. An author (a person, a primary tool, or an
agent sending a JSON spec) adds an asi_escher diagram as a block; `build()` renders it to a PNG
beside the Markdown and the three outputs show it as a numbered figure:

```python
from asi_escher import Diagram
from asi_gutenberg import Document, build

d = Document("E-stop Reset Procedure", "procedure", doc_id="ASI-FS-PROC-004")
flow = Diagram("E-stop handling", "activity")
flow.lane("op", "Operator"); flow.lane("sc", "Safety controller")
...
d.h1("Overview")
d.diagram(flow, "E-stop handling across operator and safety controller")   # <- new
build(d, "out")       # out/e_stop_reset_procedure.{md,pdf,docx} + out/figures/e_stop_handling.png
```

Today the only link between the two generators is the palette (Gutenberg's `THEME` reuses the
Escher colours) and the demo catalog's `figure(...)` calls that name Escher example files which
render as placeholders unless someone copies the PNGs by hand.

Out of scope: a Gutenberg CLI, Mermaid rendering in the PDF, editing diagrams through Gutenberg,
and any dependency in the other direction (Escher stays at the bottom of the stack).

## 2. What exists today (verified 2026-10-27)

| In Gutenberg | Relevant fact |
|---|---|
| `model.py` | `BLOCK_KINDS` is a closed set; `Figure(path, caption, width)` gets `number` at layout; `_block_to_dict` / `_block_from_dict` carry every block through JSON; the module imports neither reportlab nor python-docx by design. |
| `layout.py` | Numbers figures in document order (`f += 1`); `_stats` counts `figures`. |
| `markdown.py` | A figure is written as `![Figure n: caption](path)` and parsed back into a `Figure` from the same line. |
| `pipeline.py` | `build()` = `layout` → write `.md` → `from_markdown` → `render_pdf` / `render_docx` with `base_dir=out_dir`. The renderers never see the in-memory model. |
| `render_pdf.py`, `render_docx.py` | `_b_figure` embeds the file if it exists under `base_dir`, else draws a placeholder. No change needed if the file is there. |
| `theme.py` | Palette already equals Escher's; docstrings still say `asi_pydraw`. |
| `demos/catalog.py` | Eight `figure("<escher example>.png", "... (asi_pydraw)")` calls; the names match Escher's `EXAMPLES` keys exactly (`flowchart_fusa_lifecycle` is the key of the function `flowchart_lifecycle`). |
| `tests/test_smoke.py` | Pipeline round trip, triplet written, dangling trace reported, every demo builds. |

| In Escher | Relevant fact |
|---|---|
| `build(d, out_dir, name, opts, dpi, formats)` | Returns paths per format plus `size`, `nodes`, `edges`, `issues` (`LayoutIssue` objects). |
| `Diagram.to_dict()` / `Diagram.from_dict()` | Complete JSON interchange of the model, no layout state. |
| `layout()` | Emits one `LayoutWarning` per diagram with issues; the Escher test suite runs with that warning as an error. |
| `render()` | Writes to paths or file objects; no global matplotlib state. |
| `examples.EXAMPLES` | Sixteen `(name, builder)` pairs; `asi_escher.examples` is importable. |

## 3. Design

### 3.1 Dependency

```powershell
cd asi_gutenberg
uv add --editable ../asi_escher
```

Escher becomes a **hard dependency** (matplotlib and Pillow come with it). Alternative considered:
an optional extra `asi-gutenberg[diagrams]`. Rejected for now: every consumer in the family wants
figures, and an optional import adds a branch to every path that touches a `diagram` block. The
import stays lazy (inside `pipeline.py` and the `diagram()` helper), so `import asi_gutenberg` does
not load matplotlib.

### 3.2 The `diagram` block

A new block kind, `"diagram"`, added to `BLOCK_KINDS`:

```python
@dataclass
class DiagramBlock:
    spec: dict                 # asi_escher.Diagram.to_dict(), plain data
    caption: str = ""
    width: float = 0.0         # fraction of text width, 0 = auto, same as Figure
    name: str = ""             # output stem, defaults to a slug of spec["title"]
    formats: tuple = ("png",)  # extra Escher formats to write beside the PNG, e.g. ("png", "svg")
    dpi: int = 160
    # layout results
    number: int = 0            # shares the figure counter
    path: str = ""             # relative path of the rendered PNG, set by the pipeline
    kind: str = field(default="diagram", init=False)
```

The block stores the **plain-data spec**, not the `Diagram` object, so `model.py` keeps its promise
of depending on nothing, `to_dict` / `from_dict` stay complete, and an agent can send a document
spec with diagrams inside it as one JSON file.

Factory method on `Document`:

```python
def diagram(self, spec, caption="", width=0.0, name="", formats=("png",), dpi=160) -> DiagramBlock:
    """``spec`` is an ``asi_escher.Diagram`` or its ``to_dict()`` form."""
```

When given a `Diagram`, it calls `spec.to_dict()` immediately; when given a dict it validates by
round-tripping through `asi_escher.Diagram.from_dict` (lazy import) so an unknown node kind fails at
`add` time, as every other block does.

### 3.3 Layout

`layout()` treats a `diagram` block like a figure: it takes the next figure number, and `_stats`
counts it under `figures` and under a new `diagrams` key. `validate()` checks that `name`, when
given, is a file-system-safe stem and unique among the document's diagram blocks.

### 3.4 Rendering in the pipeline

`build()` gains two keyword arguments and one step:

```python
def build(d, out_dir, name=None, opts=None, theme=None,
          figures_dir="figures", diagram_theme=None) -> dict:
    layout(d, opts)
    rendered = render_diagrams(d, out_dir, figures_dir, diagram_theme)   # new, before to_markdown
    md = to_markdown(d)
    ...
```

`render_diagrams` (new module `diagrams.py`) does, per block:

1. `Diagram.from_dict(block.spec)`;
2. `asi_escher.build(diagram, os.path.join(out_dir, figures_dir), block.name or slug, formats=("png", *extra), dpi=block.dpi)`
   under `warnings.catch_warnings()` with `LayoutWarning` recorded, never propagated;
3. `block.path = f"{figures_dir}/{stem}.png"` (forward slashes, relative to the Markdown file, so the
   `.md` is portable and GitHub renders it);
4. every `LayoutIssue` becomes a string in `doc.issues`: `"diagram <stem>: <issue.message>"`.

Rendering happens **before** the Markdown is written because the Markdown needs the path, and before
the PDF and Word renderers run because they need the file. The renderers themselves do not change:
the parsed Markdown presents the diagram as an ordinary `Figure` whose file now exists under
`base_dir`.

The `diagram_theme` dict is passed through to `asi_escher.render` so a consumer that overrides the
document palette can override the diagram palette in the same call.

`build()`'s result gains a `"diagrams"` list, one entry per block in document order:
`{"name", "png", "svg"?, "size", "nodes", "edges", "issues"}`, taken from the Escher result. Dewey
and Kletz put it in their `Result.data` so a degraded layout is visible without opening the PDF.

### 3.5 Markdown

`to_markdown` writes a `diagram` block exactly as a figure, `![Figure n: caption](figures/stem.png)`.
Nothing new to parse: `from_markdown` yields a `Figure`, and the document that comes back from the
Markdown is still a complete description of what was rendered.

Optional, off by default: `LayoutOptions(mermaid=True)` appends a fenced ```` ```mermaid ```` block
with `to_mermaid(diagram)` after the image so the `.md` shows the diagram on GitHub even when the
PNG is not committed. Left off because the parser would turn it into a `code` block in the PDF; if
wanted, the parser learns to skip a mermaid fence that directly follows a figure.

### 3.6 Serialisation

```json
{"kind": "diagram", "spec": {"title": "E-stop handling", "kind": "activity", "nodes": [...], "edges": [...]},
 "caption": "E-stop handling", "width": 0.8, "name": "estop", "formats": ["png", "svg"]}
```

`_block_to_dict` writes it; `_block_from_dict` reads it; `Document.from_dict` therefore accepts a
document spec with diagrams from any process that never imported Escher.

### 3.7 Demos

Each `d.figure("<name>.png", "... (asi_pydraw)")` in `demos/catalog.py` becomes
`d.diagram(examples_by_name["<name>"](), "... (asi_escher)")` where `examples_by_name` is
`dict(asi_escher.examples.EXAMPLES)`. The eight names already match. `demos/build_all.py` then
produces documents with real figures and no placeholders, and prints diagram issues beside
document issues.

### 3.8 Docs and metadata

- README: a "Diagrams" section with the example above, the `figures_dir` convention, and the note
  that the PNG path in the Markdown is relative so the triplet can be moved as a folder.
- `theme.py` and `layout.py` docstrings: `asi_pydraw` → `asi_escher`.
- `pyproject.toml`: version `0.2.0`, `asi-escher` dependency and source, keyword `diagrams`.
- CHANGELOG `0.2.0` entry.

## 4. Delivery phases

| Phase | Content | Done when |
|---|---|---|
| 1. Dependency and block | `uv add --editable ../asi_escher`; `DiagramBlock`, `Document.diagram`, `BLOCK_KINDS`, `to_dict` / `from_dict`; layout numbering and stats; `validate`. | A document with a diagram block lays out, serialises, and comes back equal. |
| 2. Pipeline | `diagrams.py` with `render_diagrams`; `build()` arguments and ordering; issues merged; warning capture. | `build()` writes `figures/<stem>.png` and the PDF embeds it (no placeholder text in the PDF). |
| 3. Demos | Catalog switched to `d.diagram(...)`; captions updated; `build_all` reports diagram issues. | `uv run python demos/build_all.py` shows no placeholders and no issues. |
| 4. Release | README, docstrings, CHANGELOG, version, tag `v0.2.0`. | `uv run pytest` and `uv run ruff check src tests` clean; Dewey and Kletz can start their phase 4. |

## 5. Files touched

```
pyproject.toml, uv.lock            asi-escher dependency and path source, version 0.2.0
src/asi_gutenberg/model.py         DiagramBlock, BLOCK_KINDS, Document.diagram, (de)serialisation, validate
src/asi_gutenberg/layout.py        numbering, stats, mermaid option on LayoutOptions
src/asi_gutenberg/diagrams.py      new: render_diagrams
src/asi_gutenberg/pipeline.py      build() arguments and the render step
src/asi_gutenberg/markdown.py      diagram writes as a figure line (one branch)
src/asi_gutenberg/theme.py         docstring
src/asi_gutenberg/__init__.py      export DiagramBlock
demos/catalog.py, demos/build_all.py
tests/test_smoke.py, tests/test_diagrams.py (new)
README.md, CHANGELOG.md
```

`render_pdf.py` and `render_docx.py` are untouched.

## 6. Tests

- Round trip: a document with two diagram blocks survives `to_dict` → `from_dict` → `to_dict`
  unchanged; an invalid spec (unknown node kind) raises `ValueError` at `diagram()`.
- Layout: diagram blocks and figure blocks share one numbering sequence in document order;
  `stats["figures"]` counts both, `stats["diagrams"]` counts diagrams only; duplicate `name` fails
  `validate`.
- Pipeline: `build()` creates `figures/<stem>.png` (and `.svg` when asked); the Markdown line is
  `![Figure 1: caption](figures/stem.png)`; `result["diagrams"]` has one entry per block with
  `nodes` and `edges` matching the spec; the parsed document has a `Figure` at that position; the
  PDF has no `figure placeholder` text (extract with `pypdf` if available, otherwise assert the PDF
  size grows by more than the PNG's compressed size would allow for a placeholder).
- Issues: a diagram built to produce a `label_overlap` (two long labels on parallel edges) yields
  `issues == ["diagram <stem>: ..."]`, no Python warning escapes (`pytest.warns` is not triggered,
  `-W error::asi_escher.layout.LayoutWarning` set in `pyproject.toml` for the Gutenberg suite too).
- Portability: moving the output folder and rendering the Markdown elsewhere still resolves the
  figure (path is relative, forward slashes).
- Demos: every demo builds, and for every `diagram` block the rendered file exists; no demo has
  issues.
- Isolation: `import asi_gutenberg` does not import `matplotlib` (checked through `sys.modules`).

## 7. Risks and decisions

| Risk | Decision |
|---|---|
| Pulling matplotlib into every Gutenberg consumer | Accepted: every family consumer wants figures. Lazy import keeps start-up cost out of paths that do not render. |
| Markdown no longer "complete" without the PNG | The PNG is produced in the same `build()` call and the path is relative; a `mermaid` option can add a text form later (3.5). |
| Figure numbering changes when a diagram is inserted | That is how figures already behave; numbers live beside the text and are recomputed at layout. |
| Escher `LayoutWarning` escaping into consumers' test suites | Captured in `render_diagrams`; surfaced as `doc.issues` strings only. |
| Demo captions and docstrings still say `asi_pydraw` | Fixed in phase 3 and 4. |
| Very large diagrams exceeding the page | The existing `_b_figure` scaling (72 % of page height) already handles it; Escher's `size` is reported in `build()`'s result under `diagrams` for consumers that want to warn earlier. |

## 8. Open questions

None. Decided 2026-10-27:

- `build()` **does** return per-diagram results under `diagrams` (section 3.4).
- **PNG, not SVG, in the Word file.** python-docx has no native SVG; the PNG default stays 160 dpi
  to match Escher, and `dpi` is per block for anything going to print.
- **No Gutenberg CLI in 0.2.0.** The family's command surface is the primary tools; a Gutenberg CLI
  is a separate plan if an agent ever needs a triplet without Python.

## 9. Change log

- 2026-10-28 Status set to implemented (shipped in 0.2.0 on 2026-10-27); the family-context
  link follows `ASI_FUSA_TOOLS.md` to `asi_fusa_tools/`.
- 2026-10-27 Draft 0.2: questions decided: `diagrams` in the build result, PNG in Word, no CLI in
  0.2.0. No open questions remain.
- 2026-10-27 Draft 0.1: goal, verified starting point, dependency decision, `diagram` block, layout,
  pipeline ordering, Markdown form, serialisation, demos, phases, files, tests, risks, open questions.
