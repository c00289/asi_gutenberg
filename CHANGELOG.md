# Changelog

All notable changes to asi_gutenberg are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

## [0.7.0] - 2026-10-29

The owner, 2026-10-29, on the exported QF-17: "can we make it exactly like the template? I still want all
the columns to pop up in every page but otherwise we need to match the template exactly." The rendered
triplet carries Gutenberg's theme by design; the controlled document wants the Word original's own look.

### Added
- **`fill_docx(template, source, values, out, *, strict, mark, repeat_header, pdf)`** and the **`fill`**
  command (`fill_docx.py`, `__main__.py`): the Word original written in place with the Markdown source's
  text, its `{{ name }}` placeholders filled. The body is walked as `from_docx` walks it, so every
  paragraph, list item, heading and table cell is found again in order as a unit; the source must line up
  unit for unit (same kinds, same table shapes) or the fill refuses naming the first unit that does not.
  For a unit whose text differs, the longest common prefix and suffix stay in their runs and only the span
  between them is rewritten, in new runs after the run where it starts, which lends them its formatting;
  a span that crosses paragraphs in a cell is written one paragraph per line, the cell's own paragraphs
  reused and the last cloned. A title inside a content control is written and the file's title property
  with it. Every table's first row repeats at the top of each page (`repeat_header`, default on); `mark`
  shades a filled value mint on its own run (default off: the template exactly); `pdf` prints the PDF
  through Word itself (`pdf_with_word`, PowerShell COM, Windows only; the Word process it starts is ended,
  one the person has open is never touched; a failure is an issue, not an error). Returns `Filled` (`docx`,
  `pdf`, `replaced`, `unfilled`, `issues`). Exported from the package.
- `tests/test_fill_docx.py`: a synthetic original (bold prefix, italic tail, a cell of lines, an empty
  cell, a landscape section, a footer) filled and read back; marks on the values only; strict and loose
  unfilled placeholders; nothing to change changes nothing; a source out of line is refused; the command
  line; the PDF through Word when it is installed; the owner's QF-17 filled whole and read back cell for
  cell (when asi_standards is checked out beside).

### Changed
- `__main__`: `--param` / `--params-file` loading shared by `build` and `fill` (`_values`); `fill` finds
  the one `.docx` under `source/` beside the Markdown when `--template` is not given (the asi_standards
  layout).

## [0.6.1] - 2026-10-29

Broken-windows sweep after 0.6.0: no change to the Python API or the generated documents.

### Fixed
- **`--param` without `=`** on the command line left by `SystemExit`, a bare line on stderr; it is refused
  as `{"ok": false, "error": ...}` with exit code 1, as the command line promises for every error.
- `render_docx._para_shading` had no caller since the marks went on runs (0.4.0); gone.
- README: the block-kinds table names the `section` block (0.6.0); the Tests paragraph names
  `tests/test_import_params.py`, the third test file since 0.3.0.
- The command line's docstring pointed its examples at `templates/qf-17/`, a folder that never existed;
  they point at the asi_standards checkout, as the README does.
- The CI comment, copied from a sibling, promised GUI tests on Tk; there are none here. It says instead
  that the QF-17 test skips in CI, where the private asi_standards is not checked out.
- The QF-17 test's docstring said 48 rows; the assertion, and the table, say 49.
- `uv.lock` still recorded this project at 0.2.1; refreshed.
- This changelog's compare links pointed at an `example.invalid` placeholder and stopped at 0.2.1; they
  point at the GitHub repository, one per tagged version. Versions 0.1.0 and 0.2.0 predate this
  repository (its history starts at the `v0.2.1` tag) and have no tag here, so their headings are unlinked.
- `asi_gutenberg.__version__` (the installed metadata) and `tests/test_version.py`, which checks that
  `pyproject.toml`, the top CHANGELOG entry and `__version__` agree, so a half-done bump fails the suite;
  the sibling asi_escher 0.1.2 does the same.

### Changed
- `uv.lock` follows the sibling asi_escher to 0.1.2.

### Tests
- `test_a_bad_param_is_refused_as_json`, `test_version_is_the_same_everywhere`.

## [0.6.0] - 2026-10-29

The owner, 2026-10-29: "the report was portrait; we should strive to make the Word product as close
as possible to the original." The QF-17 original is portrait with its plan table on landscape pages.

### Added
- **The `section` block** (`model.Section`, `Document.section_break(orientation)`): a new page in the
  other orientation from here on, as a Word section break does; `<!-- section: landscape -->` in the
  Markdown. The PDF gets a third page template, the turned page with the same margins, and the header,
  footer and watermark draw to the page in use; a table on a turned page is as wide as that page. The
  Word file gets a real new section with the orientation and the page size swapped and the margins kept.
- **The importer reads section breaks**: a paragraph whose properties close a section hands the next
  section's orientation on as a `section` block when it differs; the "mixes page orientations" issue is
  gone, there is nothing left to warn about. `--landscape` on the command line still turns a whole
  document.

### Tests
- `test_a_section_block_turns_the_page_and_back` (the round trip, the three pages, the PDF's two page
  sizes, the Word sections' orientations), `test_from_docx_reads_the_section_breaks`.

## [0.5.0] - 2026-10-29

### Added
- **`--params-file JSON`** on `python -m asi_gutenberg build` (repeatable): a JSON object of parameter
  values, as a tool writes them (asi_kletz's `plan export` writes the Safety Plan's); a later file or a
  `--param` wins. A file that is not an object is refused with its path.

### Tests
- `test_the_command_line_takes_a_params_file`.

## [0.4.1] - 2026-10-29

### Fixed
- The output file stem came from the marked title (`__wheel_loader___safety_plan__sp_`); it comes from
  the plain title (`wheel_loader_safety_plan__sp_`).

## [0.4.0] - 2026-10-29

The owner, 2026-10-29: "when replacing parameters in the Word document the parametrised value, whether
it is the vehicle type or some name, must be highlighted in some way that easily tells us that it is a
replacement, just not yellow."

### Added
- **`==mark==`**, an inline highlight (`inline.py`: `Span.mark`, the token, `to_pdf_markup`): both
  renderers paint it with the theme's new `mark_bg` (`#D6F0E3`, a mint), the PDF as a background colour
  on the run, Word as run shading; the Markdown carries the marks as plain text.
- **Marked fills**: `fill(..., mark=True)` (the default) writes every filled value as `==value==` (an empty
  value stays unmarked); `build(..., mark_params=True)` and the command line's `--no-mark` to switch it off.

### Tests
- `test_marked_fills_are_highlighted_in_every_output`: the parser, `plain`, the PDF markup, the fill,
  the Word runs' shading colour in a paragraph and in a table cell, and the unmarked build.

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

[Unreleased]: https://github.com/c00289/asi_gutenberg/compare/v0.7.0...HEAD
[0.7.0]: https://github.com/c00289/asi_gutenberg/compare/v0.6.1...v0.7.0
[0.6.1]: https://github.com/c00289/asi_gutenberg/compare/v0.6.0...v0.6.1
[0.6.0]: https://github.com/c00289/asi_gutenberg/compare/v0.5.0...v0.6.0
[0.5.0]: https://github.com/c00289/asi_gutenberg/compare/v0.4.1...v0.5.0
[0.4.1]: https://github.com/c00289/asi_gutenberg/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/c00289/asi_gutenberg/compare/v0.3.1...v0.4.0
[0.3.1]: https://github.com/c00289/asi_gutenberg/compare/v0.3.0...v0.3.1
[0.3.0]: https://github.com/c00289/asi_gutenberg/compare/v0.2.1...v0.3.0
[0.2.1]: https://github.com/c00289/asi_gutenberg/releases/tag/v0.2.1
