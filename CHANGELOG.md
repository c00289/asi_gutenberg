# Changelog

All notable changes to asi_gutenberg are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

## [0.3.1] - 2026-10-29

The first parameter of the QF-17 Safety Plan is its first sentence, and the import had not seen it.

### Fixed
- **A paragraph's text is every run** (`importer.paragraph_text`): content controls (`w:sdt`, how Word
  binds the Title paragraph to the file's title property), fields and text boxes included, line and
  tab breaks kept, the `mc:Fallback` copy of a text box skipped; `Paragraph.text` of python-docx skips
  a content control, which is how "Safety plan (SP) Template" went missing. The title also falls back
  to the file's title property. Table cells read the same way.

### Tests
- `test_a_title_in_a_content_control_is_read`.

## [0.3.0] - 2026-10-29

The owner, 2026-10-29: "bring that document into Gutenberg; we will use Gutenberg to parametrise
the document; move the doc into Gutenberg and preserve it along with its checksum."

### Added
- **`from_docx`** (`importer.py`): a Word file into the model, in body order: the title, headings
  as sections (numbered when Word numbered them), paragraphs, consecutive list paragraphs as one
  list, every table (a two-column table of labels as a `Field` / `Value` table; line breaks inside
  a cell kept), inline pictures as figures written to `figures/` beside the source, `Document ID`
  and `Document owner` rows as `doc_id` and `author`, an unambiguous footer date stamp as `date`;
  NBSP and the Symbol-font bullet cleaned. `Imported` carries the document, the figure files and
  the `issues` for what is not kept (mixed page orientations, header or footer text).
- **Parameters** (`params.py`): `{{ name }}` placeholders anywhere text goes; `parameters`, `fill`
  (a copy, strict by default), `unfilled`; `build(..., params=, strict_params=)` fills first and
  reports leftovers as `unfilled parameter: <name>`. Plain text in the Markdown, so it round-trips.
- **`python -m asi_gutenberg`** with `import`, `params` and `build --param name=value [--landscape]`.
- **Landscape pages**: `THEME["orientation"]` (`portrait` | `landscape`), honoured by `page_size_pt`, both
  renderers and the Word section; a wide table reads better that way.

### Fixed
- A ten-column table on a portrait page gave one column a negative width and the PDF failed
  (`_col_widths`): every column now keeps a floor (`MIN_COL_PT`, 40 pt) and the others shrink in
  proportion; past the point where even that fits, the columns are equal and the cells wrap.
- **asi_standards**, a private sibling repo, holds the parametrised standards (the owner's decision:
  this repo stays public and never holds a controlled document): one folder per document, the Word
  original preserved under `source/` with `SHA256SUMS`, the imported Markdown source beside it. The
  first is the QF-17 Safety Plan; a test checks its checksum and its import when that checkout is
  beside this one and skips otherwise.

### Tests
- `tests/test_import_params.py`: the importer on a Word file built in the test (shape, cleaning,
  the picture, the date), the Markdown round trip and build, `fill` everywhere and strict,
  `build` with parameters, the command line, and the QF-17 source.

## [0.2.1] - 2026-10-28

Broken-windows sweep: no change to the Python API or the generated
documents.

### Fixed

- `tests/test_smoke.py` and `tests/test_diagrams.py` read the generated
  Markdown and PDF through a bare `open(...).read()`, which left seventeen
  `ResourceWarning: unclosed file` lines in the pytest warnings summary;
  both reads now use `with`.
- `theme.py` still called the requirement colours the "pydraw requirement
  node"; it is the asi_escher requirement node.
- README: "nine of the documents carry diagrams" was wrong (eight documents
  carry nine diagram blocks); the demo list is now a table naming every
  `demos/catalog.py` entry with its output stem and diagram count; the
  sibling layout tree shows the six repositories of the family as
  `ASI_FUSA_TOOLS.md` lists them; the Tests paragraph names the two test
  files; the family-document link points at the file's new home,
  `../asi_fusa_tools/ASI_FUSA_TOOLS.md`.
- `docs/escher-integration-plan.md`: the status line still said "ready to
  implement" for a plan that shipped in 0.2.0; it now says implemented. Its
  family-document link points at the new home as well.
- This changelog has compare links at its foot, as the sibling repositories
  have.
- `demos/build_all.py`: a `noqa: BLE001` for a rule the ruff configuration
  never selects is now a plain comment.

### Changed

- `uv.lock` follows the sibling asi_escher to 0.1.1.

## [0.2.0] - 2026-10-27

A document can now carry its own diagrams. asi_escher is a runtime
dependency (declared as an editable path source to the sibling checkout, see
`docs/escher-integration-plan.md`), imported lazily: `import asi_gutenberg`
still does not load matplotlib.

### Added

- `diagram` block (`DiagramBlock`, `Document.diagram(spec, caption, width,
  name, formats, dpi)`, `Document.diagrams()`): an asi_escher diagram held as
  its plain-data `to_dict()` form. A `Diagram` object is converted on the
  spot; a dict is validated through `asi_escher.Diagram.from_dict` so an
  unknown node kind fails when the block is added. The block shares the
  figure counter with `figure`, is counted under `stats["figures"]` and a new
  `stats["diagrams"]`, and travels through `Document.to_dict()` /
  `from_dict()`, so a document spec with figures is one JSON file.
- `render_diagrams(doc, out_dir, figures_dir="figures", theme=None,
  mermaid=False)` in the new `asi_gutenberg.diagrams` module: renders every
  diagram block with asi_escher into `<out_dir>/<figures_dir>/<stem>.png`
  (plus `svg`, `dot`, `mmd` when the block asks), sets the block's relative
  forward-slash path for the Markdown, captures asi_escher's `LayoutWarning`
  and records each layout issue as `"diagram <stem>: <message>"` in
  `doc.issues`.
- `build(..., figures_dir="figures", diagram_theme=None)` renders the
  diagrams between layout and the Markdown, so the `.md` points at files
  that exist and the PDF and Word renderers embed them (no placeholder).
  The result gains `"diagrams"`: one entry per block in document order with
  `name`, the written paths, `size`, `nodes`, `edges` and `issues`.
- `LayoutOptions(mermaid=True)`: after each rendered diagram the Markdown
  also carries a `mermaid` fence with the same diagram as text, so GitHub
  shows it even when the PNG is not committed; the parser skips that fence,
  so the PDF and Word files are unchanged.
- `Document.validate()` checks that diagram names are file-system-safe
  stems and that no two diagram blocks would render to the same file.
- Tests: `tests/test_diagrams.py` (round trip, add-time validation, shared
  numbering, pipeline output, PDF without placeholder, issue capture with
  warnings as errors, moved output folder, import isolation) and demo
  checks that every diagram file exists and no demo has issues. The suite
  runs with `LayoutWarning` as an error.

### Changed

- Demos carry asi_escher's bundled example diagrams as `diagram` blocks
  instead of `figure` blocks naming files that were never there; captions
  say `(asi_escher)`; `build_all.py` prints the diagram count per document.
- Docstrings and comments that still said `asi_pydraw` now say `asi_escher`.

### Fixed

- The demo safety requirements specification reported thirteen dangling
  trace links on every build: its `allocated="Safety PLC, ..."` attribute
  named elements, but `allocated` is a trace attribute whose values are
  requirement ids. The attribute is now `allocation`.

## [0.1.0] - 2026-10-27

First release in its own repository. The package was started inside the
Roy's Brain repository (folder `asi_gutenberg`, never tagged there) and moved
next to it, beside `asi_dewey` and `asi_escher`, with the same project layout
as those two. The Python API is unchanged by the move.

### Added

- `Document` model with eleven document kinds (spec, report, plan, procedure,
  minutes, memo, safety case, ICD, analysis, checklist, release notes) and a
  factory API for headings, paragraphs, lists, task lists, tables, field
  tables, code, callouts, requirements, figures, quotes, rules, page breaks,
  an inline table of contents and raw Markdown.
- Layout step (`layout`, `LayoutOptions`): section, table and figure
  numbering, appendix lettering, GitHub-style anchors, requirement registry
  with trace-link checking, generated front matter (document control,
  approvals, revision history, table of contents) and document statistics.
  The author's blocks are never rewritten; numbers live beside the text.
- GitHub-flavoured Markdown writer and parser (`to_markdown`,
  `from_markdown`) that round-trip every block kind.
- PDF renderer on reportlab (`render_pdf`) and Word renderer on python-docx
  (`render_docx`) sharing one `THEME` dictionary (the asi_escher palette, so
  dropped-in diagrams match the text around them).
- One-call pipeline `build()` that writes `.md`, parses it back and renders
  `.pdf` and `.docx` from the parsed Markdown so the three outputs always
  agree.
- JSON interchange via `Document.to_dict()` / `Document.from_dict()`.
- Fifteen demo documents in `demos/catalog.py` themed around the WA485
  autonomous wheel loader safety programme, including a stress-test
  generator, and `demos/build_all.py` to render them all.
- `py.typed` marker so host type-checkers see the package as typed.
- Smoke tests covering the pipeline round trip and the demo catalog.

[Unreleased]: https://example.invalid/asi-gutenberg/compare/v0.2.1...HEAD
[0.2.1]: https://example.invalid/asi-gutenberg/compare/v0.2.0...v0.2.1
[0.2.0]: https://example.invalid/asi-gutenberg/compare/v0.1.0...v0.2.0
[0.1.0]: https://example.invalid/asi-gutenberg/releases/tag/v0.1.0
