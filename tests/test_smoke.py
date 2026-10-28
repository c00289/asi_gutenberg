"""Smoke tests: the pipeline round-trips a small document and every demo builds."""
from __future__ import annotations

import os
import sys

import pytest

from asi_gutenberg import Document, build, from_markdown, layout, to_markdown

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "demos"))
from catalog import DEMOS  # noqa: E402


def _small() -> Document:
    d = Document("Smoke Test", "report", doc_id="T-1", date="2026-10-27", author="tests")
    d.revision("0.1", "2026-10-27", "tests", "Initial")
    d.h1("Intro")
    d.p("Hello **world**.")
    d.table(["a", "b"], [["1", "2"]], caption="Numbers")
    d.requirement("R-1", "Something", "The thing shall work.", plr="c", status="Draft")
    return d


def test_markdown_round_trip():
    d = layout(_small())
    md = to_markdown(d)
    parsed = from_markdown(md)
    assert parsed.title == d.title
    kinds = [b.kind for b in parsed.blocks]
    for kind in ("heading", "paragraph", "table", "requirement"):
        assert kind in kinds
    assert "Hello **world**." in md
    assert "Hello **world**." in to_markdown(layout(parsed))


def test_build_writes_triplet(tmp_path):
    result = build(_small(), str(tmp_path), "smoke")
    for ext in ("md", "pdf", "docx"):
        assert os.path.getsize(result[ext]) > 0
    assert result["pages"] >= 1
    assert result["issues"] == []


def test_dangling_trace_is_reported():
    d = _small()
    d.requirement("R-2", "Child", "Shall derive.", traces="R-404")
    layout(d)
    assert any("R-404" in issue for issue in d.issues)


@pytest.mark.parametrize("name,builder", DEMOS, ids=[n for n, _ in DEMOS])
def test_demo_builds(name, builder, tmp_path):
    d = builder()
    result = build(d, str(tmp_path), name)
    assert result["pages"] >= 1
    for ext in ("md", "pdf", "docx"):
        assert os.path.exists(result[ext])
    # every diagram block was rendered to the file its Markdown line points at
    with open(result["md"], encoding="utf-8") as f:
        md = f.read()
    assert len(result["diagrams"]) == len(d.diagrams())
    for block, entry in zip(d.diagrams(), result["diagrams"], strict=True):
        assert block.path.startswith("figures/") and "\\" not in block.path
        assert os.path.isfile(os.path.join(str(tmp_path), block.path))
        assert os.path.samefile(entry["png"], os.path.join(str(tmp_path), block.path))
        assert f"]({block.path})" in md
        assert entry["issues"] == []
    assert result["issues"] == []
