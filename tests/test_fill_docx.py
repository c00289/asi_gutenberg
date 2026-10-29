"""The in-place fill (0.7.0): the Word original with the source's text written into its own runs."""

from __future__ import annotations

import json
import os
import pathlib
import subprocess
import sys

import pytest

from asi_gutenberg import fill_docx, from_docx, from_markdown, to_markdown
from asi_gutenberg.theme import THEME

REPO = pathlib.Path(__file__).resolve().parent.parent
QF17 = REPO.parent / "asi_standards" / "qf-17-safety-plan"
PROGRAM_FILES = pathlib.Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
WINWORD = PROGRAM_FILES / "Microsoft Office" / "root" / "Office16" / "WINWORD.EXE"
GREEN = THEME["mark_bg"].lstrip("#")  # the mark colour, whatever the theme says it is


def _template(path: pathlib.Path) -> pathlib.Path:
    """A small Word original shaped like a controlled form: a title, control table, a bold prefix before a
    blank, bullets, a landscape section holding a table whose cells have lines and blanks, a footer."""
    import docx
    from docx.enum.section import WD_ORIENT, WD_SECTION

    d = docx.Document()
    d.add_paragraph("Widget Safety Plan Template", style="Title")
    t = d.add_table(rows=2, cols=2)
    t.style = "Table Grid"
    t.cell(0, 0).text, t.cell(0, 1).text = "Document ID", "QF-99"
    t.cell(1, 0).text, t.cell(1, 1).text = "Document owner", "__________"
    d.add_paragraph("Purpose", style="Heading 1")
    p = d.add_paragraph()
    p.add_run("Note: ").bold = True
    p.add_run("this plan applies to ___________ program(s), ")
    p.add_run("always.").italic = True
    d.add_paragraph("Perform systematic planning.", style="List Bullet")
    d.add_paragraph("Assign the activities.", style="List Bullet")
    d.add_paragraph("Plan", style="Heading 1")
    wide = d.add_section(WD_SECTION.NEW_PAGE)
    wide.orientation = WD_ORIENT.LANDSCAPE
    wide.page_width, wide.page_height = wide.page_height, wide.page_width
    t2 = d.add_table(rows=3, cols=3)
    t2.style = "Table Grid"
    for c, h in enumerate(("Phase", "Deliverable", "Deadline")):
        t2.cell(0, c).text = h
    t2.cell(1, 0).text, t2.cell(1, 1).text, t2.cell(1, 2).text = "1", "Safety Plan", "__________"
    t2.cell(2, 0).text = "2"
    t2.cell(2, 1).paragraphs[0].text = "Role A"
    t2.cell(2, 1).add_paragraph("Role B")
    back = d.add_section(WD_SECTION.NEW_PAGE)
    back.orientation = WD_ORIENT.PORTRAIT
    back.page_width, back.page_height = back.page_height, back.page_width
    d.add_paragraph("Resolution", style="Heading 1")
    d.add_paragraph("Closing text.")
    d.sections[0].footer.paragraphs[0].text = "Proprietary"
    d.save(str(path))
    return path


def _source(template: pathlib.Path, tmp_path: pathlib.Path) -> str:
    """The template's Markdown source with its blanks named, as the owner names them."""
    src = from_docx(str(template), str(tmp_path / "figures")).document
    src.meta.title = "{{ vehicle }} Safety Plan"
    tables = [b for b in src.blocks if b.kind == "table"]
    tables[0].rows[1][1] = "{{ owner }}"
    para = next(b for b in src.blocks if b.kind == "paragraph")
    para.text = para.text.replace("___________", "the {{ program }}")
    tables[1].rows[0][2] = "{{ deadline }}"
    tables[1].rows[1][1] = "Role A: {{ a }}\nRole B: {{ b }}\nRole C: {{ c }}"
    tables[1].rows[1][2] = "{{ status }}"
    return to_markdown(src)


VALUES = {"vehicle": "Wheel Loader", "owner": "Owner One", "program": "Alpha", "deadline": "2027-01-31",
          "a": "Ann", "b": "Ben", "c": "Cy", "status": "Pending"}


def _texts(cell) -> list[str]:
    return [p.text for p in cell.paragraphs]


def test_fill_writes_the_values_into_the_original_and_nothing_else(tmp_path):
    import docx
    from docx.oxml.ns import qn

    template = _template(tmp_path / "widget.docx")
    result = fill_docx(str(template), _source(template, tmp_path), VALUES, str(tmp_path / "out" / "widget.docx"))
    assert result.replaced == 4 and result.unfilled == [] and result.issues == [] and result.pdf is None
    out = docx.Document(result.docx)
    # the title, in its paragraph and in the file's properties
    assert out.paragraphs[0].text == "Wheel Loader Safety Plan" and out.paragraphs[0].style.name == "Title"
    assert out.core_properties.title == "Wheel Loader Safety Plan"
    # the bold prefix and the italic tail keep their runs; only the blank between them changed
    note = next(p for p in out.paragraphs if p.text.startswith("Note:"))
    assert note.text == "Note: this plan applies to the Alpha program(s), always."
    assert note.runs[0].text == "Note: " and note.runs[0].bold
    assert note.runs[-1].text == "always." and note.runs[-1].italic
    assert not any(r.bold or r.italic for r in note.runs[1:-1])
    # the control table, the plan table: text, lines as paragraphs, an empty cell filled, styles kept
    control, plan = out.tables
    assert control.cell(1, 1).text == "Owner One" and control.style.name == "Table Grid"
    assert plan.cell(1, 2).text == "2027-01-31" and plan.cell(2, 2).text == "Pending"
    assert _texts(plan.cell(2, 1)) == ["Role A: Ann", "Role B: Ben", "Role C: Cy"]
    assert plan.cell(2, 1).paragraphs[2].style.name == plan.cell(2, 1).paragraphs[0].style.name
    assert plan.style.name == "Table Grid"
    # every table's first row repeats on each page; the sections and the footer are the original's
    for table in out.tables:
        assert table.rows[0]._tr.find(".//" + qn("w:tblHeader")) is not None
    assert [s.orientation.name for s in out.sections] == ["PORTRAIT", "LANDSCAPE", "PORTRAIT"]
    assert out.sections[0].footer.paragraphs[0].text == "Proprietary"
    assert [p.style.name for p in out.paragraphs if p.style.name.startswith("List")] == ["List Bullet", "List Bullet"]
    # read back through the importer, the file says exactly what the filled source says
    again = from_docx(result.docx, str(tmp_path / "figures2")).document
    assert again.title == "Wheel Loader Safety Plan"
    assert [b for b in again.blocks if b.kind == "table"][1].rows[1][1] == "Role A: Ann\nRole B: Ben\nRole C: Cy"


def test_marks_shade_the_values_only(tmp_path):
    import docx
    from docx.oxml.ns import qn

    template = _template(tmp_path / "widget.docx")
    result = fill_docx(str(template), _source(template, tmp_path), VALUES, str(tmp_path / "marked.docx"), mark=True)
    out = docx.Document(result.docx)
    note = next(p for p in out.paragraphs if p.text.startswith("Note:"))

    def fill_of(run) -> str | None:
        shd = run._r.find(qn("w:rPr") + "/" + qn("w:shd")) if run._r.find(qn("w:rPr")) is not None else None
        return shd.get(qn("w:fill")) if shd is not None else None

    shaded = [r.text for r in note.runs if fill_of(r) == GREEN]
    assert shaded == ["Alpha"] and fill_of(note.runs[0]) is None
    plan = out.tables[1]
    assert [r.text for p in plan.cell(2, 1).paragraphs for r in p.runs if fill_of(r) == GREEN] == ["Ann", "Ben", "Cy"]
    plain = fill_docx(str(template), _source(template, tmp_path), VALUES, str(tmp_path / "plain.docx"))
    assert not any(fill_of(r) == GREEN for p in docx.Document(plain.docx).paragraphs for r in p.runs)


def test_unfilled_placeholders_stay_or_refuse(tmp_path):
    import docx

    template = _template(tmp_path / "widget.docx")
    source = _source(template, tmp_path)
    with pytest.raises(ValueError, match="no value for parameter"):
        fill_docx(str(template), source, {"vehicle": "X"}, str(tmp_path / "strict.docx"), strict=True)
    result = fill_docx(str(template), source, {"vehicle": "Wheel Loader"}, str(tmp_path / "loose.docx"))
    assert result.unfilled == ["a", "b", "c", "deadline", "owner", "program", "status"]
    out = docx.Document(result.docx)
    assert out.tables[0].cell(1, 1).text == "{{ owner }}" and out.paragraphs[0].text == "Wheel Loader Safety Plan"


def test_nothing_to_change_changes_nothing(tmp_path):
    template = _template(tmp_path / "widget.docx")
    untouched = to_markdown(from_docx(str(template), str(tmp_path / "figures")).document)
    result = fill_docx(str(template), untouched, {}, str(tmp_path / "same.docx"), repeat_header=False)
    assert result.replaced == 0
    before = from_docx(str(template), str(tmp_path / "f1")).document.to_dict()
    after = from_docx(result.docx, str(tmp_path / "f2")).document.to_dict()
    assert before == after


def test_a_cloned_line_keeps_the_runs_properties(tmp_path):
    """0.7.1: a cell written line by line clones its last paragraph for the extra lines; the clone's runs are
    emptied, not dropped, so the text written there carries their properties (bold, size, font) as the first
    line's does, and an empty paragraph with no run takes its paragraph mark's. Before, both went into a bare
    run and took the paragraph style's look: the extra role lines of the QF-17 came out in the body font."""
    import docx
    from docx.oxml.ns import qn
    from docx.shared import Pt

    d = docx.Document()
    d.add_paragraph("Roles", style="Title")
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "Role", "Who"
    run = t.cell(1, 0).paragraphs[0].add_run("Role A: ____")
    run.bold, run.font.size, run.font.name = True, Pt(7), "Arial"
    blank = t.cell(1, 1).paragraphs[0]  # no run at all; the paragraph mark says bold
    mark = blank._p.get_or_add_pPr()
    rpr = mark.makeelement(qn("w:rPr"), {})
    rpr.append(rpr.makeelement(qn("w:b"), {}))
    mark.append(rpr)
    template = tmp_path / "roles.docx"
    d.save(str(template))
    src = from_docx(str(template), str(tmp_path / "figures")).document
    [table] = [b for b in src.blocks if b.kind == "table"]
    table.rows[0][0] = "Role A: {{ a }}" + chr(10) + "Role B: {{ b }}" + chr(10) + "Role C: {{ c }}"
    table.rows[0][1] = "{{ who }}"
    result = fill_docx(str(template), to_markdown(src), {"a": "Ann", "b": "Ben", "c": "Cy", "who": "Ann"},
                       str(tmp_path / "filled.docx"))
    out = docx.Document(result.docx).tables[0]
    assert _texts(out.cell(1, 0)) == ["Role A: Ann", "Role B: Ben", "Role C: Cy"]
    for paragraph in out.cell(1, 0).paragraphs:
        for r in paragraph.runs:
            if r.text:
                assert (r.bold, r.font.size, r.font.name) == (True, Pt(7), "Arial"), paragraph.text
    [who] = [r for r in out.cell(1, 1).paragraphs[0].runs if r.text]
    assert who.text == "Ann" and who.bold


def test_a_failed_print_leaves_the_pdf_that_was_there(tmp_path, monkeypatch):
    """0.7.1: Word prints to a working name and the target changes hands only once the print is there, so a
    print that fails (no Word, a COM error) leaves the PDF that was at the target, which asi_kletz counts on:
    its export writes Gutenberg's rendering first and keeps it when Word cannot print. Before, the target was
    removed before Word was asked, and a failed print left no PDF at all."""
    import subprocess

    template = _template(tmp_path / "widget.docx")
    rendered = tmp_path / "printed.pdf"
    rendered.write_bytes(b"%PDF-1.4 the rendering of the source")

    def no_word(*args, **kwargs):
        raise OSError("powershell is not here")

    monkeypatch.setattr(subprocess, "run", no_word)  # what fill_docx.py calls, by module attribute
    result = fill_docx(str(template), _source(template, tmp_path), VALUES, str(tmp_path / "printed.docx"), pdf=True)
    assert result.pdf is None and result.issues == ["pdf through Word failed: powershell is not here"]
    assert rendered.read_bytes() == b"%PDF-1.4 the rendering of the source"
    assert not (tmp_path / "printed.printing.pdf").exists()


def test_a_source_out_of_line_with_the_original_is_refused(tmp_path):
    template = _template(tmp_path / "widget.docx")
    src = from_markdown(_source(template, tmp_path))
    src.blocks = [b for b in src.blocks if b.kind != "list"]  # two bullets gone: the units no longer line up
    with pytest.raises(ValueError, match="unit 5: the Word file has a text, the source a table"):
        fill_docx(str(template), src, VALUES, str(tmp_path / "bad.docx"))
    src = from_markdown(_source(template, tmp_path))
    table = [b for b in src.blocks if b.kind == "table"][1]
    table.rows.append(["3", "Extra", "x"])
    with pytest.raises(ValueError, match="rows"):
        fill_docx(str(template), src, VALUES, str(tmp_path / "bad2.docx"))


def test_the_command_line_fills(tmp_path):
    template = _template(tmp_path / "source" / "widget.docx") if (tmp_path / "source").mkdir() is None else None
    md = tmp_path / "widget.md"
    md.write_text(_source(template, tmp_path), encoding="utf-8")
    (tmp_path / "values.json").write_text(json.dumps(VALUES), encoding="utf-8")
    run = subprocess.run([sys.executable, "-m", "asi_gutenberg", "fill", str(md), str(tmp_path / "out" / "w.docx"),
                          "--params-file", str(tmp_path / "values.json"), "--param", "program=Beta"],
                         capture_output=True, text=True, encoding="utf-8")
    assert run.returncode == 0, run.stderr
    payload = json.loads(run.stdout)
    assert payload["ok"] and payload["verb"] == "fill" and payload["replaced"] == 4 and payload["pdf"] is None
    assert payload["template"] == str(template.resolve()) and os.path.isfile(payload["docx"])
    import docx
    note = next(p.text for p in docx.Document(payload["docx"]).paragraphs if p.text.startswith("Note"))
    assert "the Beta program" in note
    run = subprocess.run([sys.executable, "-m", "asi_gutenberg", "fill", str(md), str(tmp_path / "x.docx")],
                         capture_output=True, text=True, encoding="utf-8")
    assert run.returncode == 1 and "no value for parameter" in json.loads(run.stdout)["error"]


@pytest.mark.skipif(not WINWORD.is_file(), reason="Word is not installed here")
def test_word_prints_the_pdf(tmp_path):
    template = _template(tmp_path / "widget.docx")
    result = fill_docx(str(template), _source(template, tmp_path), VALUES, str(tmp_path / "printed.docx"), pdf=True)
    assert result.issues == [] and result.pdf == str(tmp_path / "printed.pdf")
    assert pathlib.Path(result.pdf).stat().st_size > 1000


@pytest.mark.skipif(not (QF17 / "qf-17-safety-plan.md").is_file(),
                    reason="asi_standards (private) is not checked out beside this repo")
def test_qf_17_fills_whole(tmp_path):
    """The owner's controlled document: every unit lines up, every value lands, nothing else moves."""
    import docx
    from docx.oxml.ns import qn

    from asi_gutenberg import fill, parameters

    source = (QF17 / "qf-17-safety-plan.md").read_text(encoding="utf-8")
    template = next((QF17 / "source").glob("*.docx"))
    doc = from_markdown(source)
    values = {n: f"<{n}>" for n in parameters(doc)}
    values["vehicle_type"] = "Wheel Loader"
    result = fill_docx(str(template), source, values, str(tmp_path / "qf-17.docx"))
    assert result.unfilled == [] and result.replaced >= 3
    out = docx.Document(result.docx)
    assert out.core_properties.title == "Wheel Loader Safety Plan (SP)"
    assert [s.orientation.name for s in out.sections] == ["PORTRAIT", "LANDSCAPE", "PORTRAIT"]
    plan = out.tables[2]
    assert len(plan.rows) == 50 and len(plan.columns) == 10 and plan.style.name == "Grid Table 6 Colorful"
    assert plan.rows[0]._tr.find(".//" + qn("w:tblHeader")) is not None
    durations = [plan.cell(r, 8).text for r in range(1, len(plan.rows))]
    statuses = [plan.cell(r, 9).text for r in range(1, len(plan.rows))]
    assert durations[0] == "<row_1_points> points\n<row_1_sprint_end>"
    assert any("<row_48_status>" in s for s in statuses)  # the Status column stacks the Jira key over the status
    # read back, the file says exactly what the filled source says, cell for cell
    from asi_gutenberg.fill_docx import _source_units

    want = _source_units(fill(doc, values, mark=False))
    got = _source_units(from_docx(result.docx, str(tmp_path / "figures")).document)
    assert len(want) == len(got)
    for w, g in zip(want, got, strict=True):
        assert w == g
