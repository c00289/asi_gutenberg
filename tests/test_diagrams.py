"""The ``diagram`` block: model round trip, layout numbering, rendering through
asi_escher inside ``build()``, issue capture, portability and import isolation."""
from __future__ import annotations

import base64
import os
import re
import shutil
import subprocess
import sys
import warnings
import zlib

import pytest

from asi_gutenberg import (
    DiagramBlock,
    Document,
    LayoutOptions,
    build,
    from_markdown,
    layout,
    render_diagrams,
    render_pdf,
    to_markdown,
)


# ---------------------------------------------------------------- helpers --
def _flow(title: str = "Login flow"):
    from asi_escher import Diagram

    d = Diagram(title, "flowchart")
    d.start()
    d.action("check", "Check credentials")
    d.decision("ok", "Valid?")
    d.end()
    d.chain("start", "check", "ok")
    d.edge("ok", "end", "yes")
    d.edge("ok", "check", "no")
    return d


def _modes():
    from asi_escher import Diagram

    d = Diagram("Modes", "state")
    d.start()
    d.state("manual", "Manual")
    d.state("auto", "Autonomous")
    d.chain("start", "manual")
    d.edge("manual", "auto", "arm")
    d.edge("auto", "manual", "disarm")
    return d


def _crowded():
    """Copied from the asi_escher suite: a tall back-edge label in a narrow lane
    that cannot be placed without overlapping, so layout reports an issue."""
    from asi_escher import Diagram

    d = Diagram("crowded", "activity")
    d.lane("l1", "One")
    d.lane("l2", "Two")
    d.action("a", "A", lane="l1")
    d.action("b", "B", lane="l1")
    d.action("c", "C", lane="l1")
    d.chain("a", "b", "c")
    d.edge("c", "a", "line one\nline two\nline three\nline four")
    d.action("x", "X", lane="l2")
    d.action("y", "Y", lane="l2")
    d.action("z", "Z", lane="l2")
    d.chain("x", "y", "z")
    return d


def _doc(title: str = "Diagram Test") -> Document:
    d = Document(title, "report", doc_id="T-D", date="2026-10-27", author="tests")
    d.h1("Intro")
    d.p("Some text.")
    return d


def _pdf_text(path: str) -> str:
    """Text content of a reportlab PDF without pypdf: inflate every Flate stream
    and keep the string operands, so ``(figure placeholder) Tj`` is searchable."""
    try:
        from pypdf import PdfReader  # type: ignore
        return "\n".join(p.extract_text() or "" for p in PdfReader(path).pages)
    except ImportError:
        pass
    with open(path, "rb") as f:
        raw = f.read()
    out = []
    for m in re.finditer(rb"stream\r?\n(.*?)endstream", raw, re.S):
        data = m.group(1).strip()
        for decode in (lambda b: zlib.decompress(base64.a85decode(b, adobe=True)),
                       zlib.decompress, lambda b: b):
            try:
                out.append(decode(data).decode("latin-1"))
                break
            except Exception:
                continue
    return "\n".join(out)


# ------------------------------------------------------------ round trip --
def test_dict_round_trip_with_two_diagrams():
    d = _doc()
    d.diagram(_flow(), "Login", width=0.8, name="login", formats=("png", "svg"))
    d.diagram(_modes().to_dict(), "Modes", dpi=200)
    spec = d.to_dict()
    kinds = [b["kind"] for b in spec["blocks"]]
    assert kinds == ["heading", "paragraph", "diagram", "diagram"]
    assert spec["blocks"][2]["formats"] == ["png", "svg"]
    assert spec["blocks"][2]["name"] == "login" and spec["blocks"][2]["width"] == 0.8
    assert "formats" not in spec["blocks"][3] and spec["blocks"][3]["dpi"] == 200
    assert spec["blocks"][2]["spec"]["nodes"][1]["label"] == "Check credentials"
    back = Document.from_dict(spec)
    assert back.to_dict() == spec
    blocks = back.diagrams()
    assert len(blocks) == 2 and all(isinstance(b, DiagramBlock) for b in blocks)
    assert blocks[0].stem == "login" and blocks[1].stem == "modes"
    assert blocks[0].formats == ("png", "svg") and blocks[1].formats == ("png",)


def test_spec_is_copied_not_shared():
    d = _doc()
    src = _flow().to_dict()
    b = d.diagram(src, "x")
    src["nodes"].clear()
    assert b.spec["nodes"]
    out = d.to_dict()
    out["blocks"][2]["spec"]["nodes"].clear()
    assert b.spec["nodes"]


def test_invalid_spec_fails_at_add_time():
    d = _doc()
    bad = _flow().to_dict()
    bad["nodes"][1]["kind"] = "teapot"
    with pytest.raises(ValueError, match="unknown node kind"):
        d.diagram(bad, "bad")
    assert d.diagrams() == []
    with pytest.raises(TypeError):
        d.diagram("not a diagram")
    with pytest.raises(ValueError, match="unknown diagram format"):
        d.diagram(_flow(), formats=("png", "jpeg"))
    with pytest.raises(ValueError, match="file-system-safe"):
        d.diagram(_flow(), name="../escape")
    with pytest.raises(ValueError, match="dpi"):
        d.diagram(_flow(), dpi=0)
    with pytest.raises(ValueError, match="title"):
        Document.from_dict({"meta": {"title": "x"}, "blocks": [{"kind": "diagram", "spec": {}}]})


# ---------------------------------------------------------------- layout --
def test_figures_and_diagrams_share_one_numbering():
    d = _doc()
    d.figure("a.png", "first")
    d.diagram(_flow(), "second")
    d.figure("b.png", "third")
    d.diagram(_modes(), "fourth")
    layout(d)
    numbered = [(b.kind, b.number) for b in d.blocks if b.kind in ("figure", "diagram")]
    assert numbered == [("figure", 1), ("diagram", 2), ("figure", 3), ("diagram", 4)]
    assert d.stats["figures"] == 4
    assert d.stats["diagrams"] == 2
    assert d.blocks[-1].display_caption == "Figure 4: fourth"


def test_duplicate_stem_fails_validate():
    d = _doc()
    d.diagram(_flow(), "a", name="same")
    d.diagram(_modes(), "b", name="same")
    with pytest.raises(ValueError, match="same"):
        d.validate()
    d2 = _doc()
    d2.diagram(_flow("Twice"), "a")
    d2.diagram(_modes().to_dict() | {"title": "Twice"}, "b")
    with pytest.raises(ValueError, match="twice"):
        layout(d2)
    d2.blocks[-1].name = "other"
    layout(d2)


def test_unrendered_diagram_cannot_be_written():
    d = _doc()
    d.diagram(_flow(), "x")
    layout(d)
    with pytest.raises(ValueError, match="not been rendered"):
        to_markdown(d)


# -------------------------------------------------------------- pipeline --
def test_build_renders_diagram_into_figures_dir(tmp_path):
    d = _doc()
    d.diagram(_flow(), "Login flow (asi_escher)", name="login", formats=("png", "svg"))
    d.figure("missing.png", "a file that is not there")
    result = build(d, str(tmp_path), "doc")
    png = tmp_path / "figures" / "login.png"
    assert png.is_file() and png.stat().st_size > 1000
    assert (tmp_path / "figures" / "login.svg").is_file()
    md = (tmp_path / "doc.md").read_text(encoding="utf-8")
    assert "![Figure 1: Login flow (asi_escher)](figures/login.png)" in md
    assert "\\" not in md.split("](figures/")[1].split(")")[0]
    assert result["figures"] == 2 and result["diagrams"] and len(result["diagrams"]) == 1
    entry = result["diagrams"][0]
    assert entry["name"] == "login"
    assert os.path.samefile(entry["png"], png) and entry["svg"].endswith("login.svg")
    spec = d.diagrams()[0].spec
    assert entry["nodes"] == len(spec["nodes"]) and entry["edges"] == len(spec["edges"])
    assert entry["issues"] == [] and len(entry["size"]) == 2
    assert result["issues"] == []
    parsed = from_markdown(md)
    figs = [b for b in parsed.blocks if b.kind == "figure"]
    assert figs[0].path == "figures/login.png" and figs[0].number == 1
    assert figs[0].caption == "Login flow (asi_escher)"
    text = _pdf_text(str(tmp_path / "doc.pdf"))
    assert text.count("figure placeholder") == 1          # the missing file only
    assert "missing.png" in text and "login.png" not in text


def test_build_without_diagrams_is_unchanged(tmp_path):
    result = build(_doc(), str(tmp_path), "plain")
    assert result["diagrams"] == [] and result["figures"] == 0
    assert not (tmp_path / "figures").exists()


def test_build_passes_theme_and_figures_dir(tmp_path):
    d = _doc()
    d.diagram(_modes(), "Modes")
    result = build(d, str(tmp_path), "doc", figures_dir="img", diagram_theme={"bg": "#000000"})
    assert (tmp_path / "img" / "modes.png").is_file()
    assert "![Figure 1: Modes](img/modes.png)" in (tmp_path / "doc.md").read_text(encoding="utf-8")
    assert result["diagrams"][0]["png"].endswith(os.path.join("img", "modes.png"))
    d = _doc()
    d.diagram(_modes(), "Modes")
    build(d, str(tmp_path / "flat"), "doc", figures_dir="")
    assert (tmp_path / "flat" / "modes.png").is_file()
    assert "](modes.png)" in (tmp_path / "flat" / "doc.md").read_text(encoding="utf-8")


def test_mermaid_option_adds_text_form_the_renderers_skip(tmp_path):
    d = _doc()
    d.diagram(_modes(), "Modes")
    build(d, str(tmp_path), "doc", opts=LayoutOptions(mermaid=True))
    md = (tmp_path / "doc.md").read_text(encoding="utf-8")
    assert "![Figure 1: Modes](figures/modes.png)\n\n```mermaid\n" in md
    fence = md.split("```mermaid\n", 1)[1].split("```", 1)[0]
    assert "stateDiagram" in fence and "manual --> auto : arm" in fence
    parsed = from_markdown(md)
    assert [b.kind for b in parsed.blocks if b.kind in ("figure", "code")] == ["figure"]


def test_render_diagrams_step_alone(tmp_path):
    d = _doc()
    d.diagram(_flow(), "Login", formats=("png", "dot", "mmd"))
    layout(d)
    entries = render_diagrams(d, str(tmp_path))
    assert [e["name"] for e in entries] == ["login_flow"]
    for fmt in ("png", "dot", "mmd"):
        assert os.path.isfile(entries[0][fmt])
    assert d.diagrams()[0].path == "figures/login_flow.png"
    assert "![Figure 1: Login](figures/login_flow.png)" in to_markdown(d)


# ---------------------------------------------------------------- issues --
def test_layout_issue_is_reported_and_no_warning_escapes(tmp_path):
    d = _doc()
    d.diagram(_crowded(), "Crowded")
    with warnings.catch_warnings():
        warnings.simplefilter("error")                   # any warning fails the test
        result = build(d, str(tmp_path), "doc")
    assert result["issues"] and all(i.startswith("diagram crowded: ") for i in result["issues"])
    assert result["issues"] == d.issues
    entry = result["diagrams"][0]
    assert entry["issues"] and entry["issues"][0]["kind"] == "label_overlap"
    assert entry["issues"][0]["edge"] == ("c", "a")
    assert entry["issues"][0]["message"] in result["issues"][0]
    assert (tmp_path / "figures" / "crowded.png").is_file()


def test_trace_issues_and_diagram_issues_are_merged(tmp_path):
    d = _doc()
    d.requirement("R-1", "Child", "Shall derive.", traces="R-404")
    d.diagram(_crowded(), "Crowded")
    result = build(d, str(tmp_path), "doc")
    assert any("R-404" in i for i in result["issues"])
    assert any(i.startswith("diagram crowded") for i in result["issues"])


# ----------------------------------------------------------- portability --
def test_moved_output_folder_still_resolves_the_figure(tmp_path):
    d = _doc()
    d.diagram(_flow(), "Login", name="login")
    build(d, str(tmp_path / "a"), "doc")
    shutil.move(str(tmp_path / "a"), str(tmp_path / "b"))
    md = (tmp_path / "b" / "doc.md").read_text(encoding="utf-8")
    parsed = from_markdown(md)
    render_pdf(parsed, str(tmp_path / "b" / "again.pdf"), base_dir=str(tmp_path / "b"))
    assert "figure placeholder" not in _pdf_text(str(tmp_path / "b" / "again.pdf"))
    assert (tmp_path / "b" / "again.pdf").stat().st_size > 20_000


# -------------------------------------------------------------- isolation --
def test_import_does_not_load_matplotlib():
    code = ("import sys, asi_gutenberg, asi_gutenberg.diagrams; "
            "bad = [m for m in ('matplotlib', 'asi_escher') if m in sys.modules]; "
            "print(bad); sys.exit(1 if bad else 0)")
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stdout + proc.stderr
