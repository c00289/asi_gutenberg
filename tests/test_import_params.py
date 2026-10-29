"""The Word importer and the parameters (0.3.0): a Word file becomes a source, its blanks become names."""

from __future__ import annotations

import json
import os
import pathlib
import subprocess
import sys

import pytest

from asi_gutenberg import Document, build, fill, from_docx, from_markdown, parameters, to_markdown, unfilled

REPO = pathlib.Path(__file__).resolve().parent.parent
# the owner's parametrised standards live in the private sibling repo asi_standards, never in this public one;
# the test runs when that checkout is beside this one and skips otherwise
QF17 = REPO.parent / "asi_standards" / "qf-17-safety-plan" / "source" / "QF-17 Safety Plan.docx"


def _word_file(path: pathlib.Path, *, picture: bool = False) -> pathlib.Path:
    """A small Word document shaped like a controlled QMS form: control tables, sections, lists, a wide table."""
    import docx
    from docx.shared import Inches

    d = docx.Document()
    d.add_paragraph("Widget Safety Plan", style="Title")
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "Document ID", "QF-99"
    t.cell(1, 0).text, t.cell(1, 1).text = "Document owner", "{{ owner }}"
    d.add_paragraph("Purpose", style="Heading 1")
    d.add_paragraph("The purpose of this plan is to plan {{ program }}.")
    d.add_paragraph("Perform systematic planning.", style="List Bullet")
    d.add_paragraph("Assign the activities.", style="List Bullet")
    d.add_paragraph("Scope", style="Heading 1")
    d.add_paragraph("Details", style="Heading 2")
    d.add_paragraph("This plan applies to ___________ program(s).")
    t2 = d.add_table(rows=3, cols=3)
    for c, h in enumerate(("Phase", "Deliverable", "Deadline")):
        t2.cell(0, c).text = h
    t2.cell(1, 0).text, t2.cell(1, 1).text, t2.cell(1, 2).text = "1", "Safety Plan\nsecond line", "{{ deadline_1 }}"
    t2.cell(2, 0).text, t2.cell(2, 1).text, t2.cell(2, 2).text = "2", " HARA report", "__________"
    if picture:
        png = path.parent / "pic.png"
        png.write_bytes(bytes.fromhex(
            "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da6364f8cfc00000"
            "0201010022a8ef520000000049454e44ae426082"))
        d.add_picture(str(png), width=Inches(1))
    section = d.sections[0]
    section.footer.paragraphs[0].text = "28 / 10 / 2026  -  9:00AM  |"
    d.save(str(path))
    return path


def test_from_docx_reads_the_shape_faithfully(tmp_path):
    result = from_docx(str(_word_file(tmp_path / "widget.docx", picture=True)))
    doc = result.document
    assert doc.title == "Widget Safety Plan" and doc.meta.doc_id == "QF-99" and doc.meta.author == "{{ owner }}"
    assert doc.meta.date == "2026-10-28"  # the footer stamp, day above 12 so the order is certain
    kinds = [b.kind for b in doc.blocks]
    assert kinds == ["table", "heading", "paragraph", "list", "heading", "heading", "paragraph", "table", "figure"]
    control = doc.blocks[0]
    assert control.columns == ["Field", "Value"]
    assert control.rows == [["Document ID", "QF-99"], ["Document owner", "{{ owner }}"]]
    assert doc.blocks[2].text == "The purpose of this plan is to plan {{ program }}."  # NBSP cleaned
    assert [i.text for i in doc.blocks[3].items] == ["Perform systematic planning.", "Assign the activities."]
    assert (doc.blocks[4].level, doc.blocks[5].level) == (1, 2) and doc.blocks[4].text == "Scope"
    wide = doc.blocks[7]
    assert wide.columns == ["Phase", "Deliverable", "Deadline"]
    assert wide.rows == [["1", "Safety Plan\nsecond line", "{{ deadline_1 }}"], ["2", "• HARA report", "__________"]]
    assert doc.blocks[8].path == "figures/widget_1.png" and os.path.isfile(result.figures[0])
    assert result.issues == []  # one orientation, no header text beyond the stamp
    assert parameters(doc) == ["deadline_1", "owner", "program"]


def test_imported_source_round_trips_and_builds(tmp_path):
    result = from_docx(str(_word_file(tmp_path / "widget.docx")))
    md = to_markdown(result.document)
    again = from_markdown(md)
    assert [b.kind for b in again.blocks] == [b.kind for b in result.document.blocks]
    assert again.blocks[7].rows[0][1] == "Safety Plan\nsecond line"  # <br> in the cell survives the round trip
    assert "{{ program }}" in md and "{{ deadline_1 }}" in md
    out = build(result.document, str(tmp_path / "out"), "widget")
    for ext in ("md", "pdf", "docx"):
        assert os.path.getsize(out[ext]) > 0
    assert sorted(i for i in out["issues"] if i.startswith("unfilled")) == [
        "unfilled parameter: deadline_1", "unfilled parameter: owner", "unfilled parameter: program",
    ]


def test_fill_substitutes_everywhere_and_strict_refuses_a_missing_value():
    d = Document("{{ program }} Safety Plan", "plan", doc_id="QF-{{ number }}", author="{{ owner }}")
    d.h1("Scope of {{ program }}")
    d.p("Owner: {{owner}}.")
    d.bullets(["Item {{ program }}"])
    d.table(["Deliverable", "{{ column }}"], [["Plan", "{{ deadline }}"]], caption="Plan {{ program }}")
    d.note("Mind {{ program }}", title="{{ owner }}")
    assert parameters(d) == ["column", "deadline", "number", "owner", "program"]
    with pytest.raises(ValueError, match="no value for parameter\\(s\\): column, deadline"):
        fill(d, {"program": "WA485", "number": 17, "owner": "the FSE"})
    done = fill(d, {"program": "WA485", "number": 17, "owner": "the FSE", "column": "When", "deadline": None})
    assert done.title == "WA485 Safety Plan" and done.meta.doc_id == "QF-17" and done.meta.author == "the FSE"
    assert done.blocks[0].text == "Scope of WA485" and done.blocks[1].text == "Owner: the FSE."
    assert done.blocks[2].items[0].text == "Item WA485"
    assert done.blocks[3].columns == ["Deliverable", "When"] and done.blocks[3].rows == [["Plan", ""]]
    assert done.blocks[3].caption == "Plan WA485" and done.blocks[4].title == "the FSE"
    assert parameters(d) == ["column", "deadline", "number", "owner", "program"]  # the original is untouched
    assert unfilled(done) == []
    partial = fill(d, {"program": "WA485"}, strict=False)
    assert unfilled(partial) == ["column", "deadline", "number", "owner"]


def test_build_with_params_fills_and_reports_leftovers(tmp_path):
    d = Document("{{ program }} Plan", "plan")
    d.h1("Scope")
    d.p("For {{ program }}, owned by {{ owner }}.")
    with pytest.raises(ValueError, match="owner"):
        build(d, str(tmp_path), "strict", params={"program": "WA485"})
    out = build(d, str(tmp_path), "loose", params={"program": "WA485"}, strict_params=False)
    text = pathlib.Path(out["md"]).read_text(encoding="utf-8")
    assert "For WA485, owned by {{ owner }}." in text and "unfilled parameter: owner" in out["issues"]
    full = build(d, str(tmp_path), "full", params={"program": "WA485", "owner": "the FSE"})
    assert "{{" not in pathlib.Path(full["md"]).read_text(encoding="utf-8") and full["issues"] == []


def test_the_command_line_imports_lists_and_builds(tmp_path):
    word = _word_file(tmp_path / "widget.docx")
    source = tmp_path / "src" / "widget.md"

    def run(*args):
        return subprocess.run([sys.executable, "-m", "asi_gutenberg", *args], capture_output=True, text=True)

    imported = run("import", str(word), str(source))
    assert imported.returncode == 0, imported.stdout + imported.stderr
    payload = json.loads(imported.stdout)
    assert payload["ok"] and payload["parameters"] == ["deadline_1", "owner", "program"] and source.is_file()
    listed = json.loads(run("params", str(source)).stdout)
    assert listed["parameters"] == ["deadline_1", "owner", "program"]
    refused = run("build", str(source), str(tmp_path / "out"), "--param", "program=WA485")
    assert refused.returncode == 1 and "no value for parameter(s): deadline_1, owner" in refused.stdout
    built = run("build", str(source), str(tmp_path / "out"), "--param", "program=WA485", "--param", "owner=FSE",
                "--param", "deadline_1=2026-11-10")
    assert built.returncode == 0, built.stdout + built.stderr
    result = json.loads(built.stdout)
    assert result["issues"] == [] and os.path.getsize(result["pdf"]) > 0


@pytest.mark.skipif(not QF17.is_file(), reason="asi_standards (private) is not checked out beside this repo")
def test_qf_17_safety_plan_imports_whole(tmp_path):
    """The owner's controlled document (asi_standards/qf-17-safety-plan): the sections, the tables, the 48 rows."""
    import hashlib

    manifest = (QF17.parent / "SHA256SUMS").read_text(encoding="utf-8").split()[0]
    assert hashlib.sha256(QF17.read_bytes()).hexdigest() == manifest  # preserved byte for byte
    result = from_docx(str(QF17), str(tmp_path / "figures"))
    doc = result.document
    assert doc.meta.doc_id == "QF-17"
    headings = [b.text for b in doc.blocks if b.kind == "heading" and b.level == 1]
    assert headings[:3] == ["Purpose", "Scope", "Safety Plan"] and headings[-1] == "Resolution Summary"
    tables = [b for b in doc.blocks if b.kind == "table"]
    assert len(tables) == 3 and tables[0].rows[0][0] == "Document ID"
    plan = tables[2]
    assert plan.columns[:3] == ["PDP Phases", "Safety Lifecycle Phase", "Safety Deliverables"] and len(plan.rows) == 49
    assert any("mixes page orientations" in i for i in result.issues)


def test_a_wide_table_builds_portrait_and_landscape(tmp_path):
    """A ten-column plan table on a portrait page starved a column (negative width, 0.2.1); now every column keeps
    a floor and the page can be landscape."""
    import docx
    from docx.enum.section import WD_ORIENT

    from asi_gutenberg.theme import page_size_pt

    d = Document("Wide", "plan")
    d.h1("Plan")
    columns = [f"Column {i} with a long heading" for i in range(10)]
    d.table(columns, [[f"cell {i} " * 6 for i in range(10)] for _ in range(5)])
    portrait = build(d, str(tmp_path), "portrait")
    assert os.path.getsize(portrait["pdf"]) > 0 and portrait["pages"] >= 1
    landscape = build(d, str(tmp_path), "landscape", theme={"orientation": "landscape"})
    assert page_size_pt({"orientation": "landscape"}) == (792.0, 612.0) and page_size_pt({}) == (612.0, 792.0)
    section = docx.Document(landscape["docx"]).sections[0]
    assert section.orientation == WD_ORIENT.LANDSCAPE and section.page_width > section.page_height
    assert docx.Document(portrait["docx"]).sections[0].orientation == WD_ORIENT.PORTRAIT
