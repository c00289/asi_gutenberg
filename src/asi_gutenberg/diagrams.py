"""Render the ``diagram`` blocks of a document with asi_escher.

``render_diagrams(doc, out_dir)`` turns every :class:`~asi_gutenberg.model.DiagramBlock`
back into an ``asi_escher.Diagram``, lays it out, writes
``<out_dir>/<figures_dir>/<stem>.png`` (and any extra formats the block asks
for), stores the relative forward-slash path in ``block.path`` so the
Markdown stays portable, and records every degraded layout as a string in
``doc.issues``.  asi_escher's ``LayoutWarning`` is captured here and never
propagates: a host that runs with warnings as errors sees the issue text in
``doc.issues`` and in the returned list instead.

asi_escher (and matplotlib with it) is imported inside the functions, so
importing this module costs nothing.
"""
from __future__ import annotations

import os
import warnings
from dataclasses import asdict

from .model import DiagramBlock, Document


def render_diagrams(doc: Document, out_dir: str, figures_dir: str = "figures",
                    theme: dict | None = None, mermaid: bool = False) -> list[dict]:
    """Render every diagram block of ``doc`` into ``out_dir/figures_dir``.

    ``theme`` overrides entries of ``asi_escher.THEME`` for every diagram.
    With ``mermaid`` set, each block also receives its Mermaid source in
    ``block.mermaid`` so :func:`asi_gutenberg.to_markdown` can append it as
    a fenced block.

    Returns one dict per diagram block in document order: ``name`` (the file
    stem), one path per written format (``png``, ``svg``, ``dot``, ``mmd``),
    ``size`` (canvas width, height in px), ``nodes``, ``edges`` and
    ``issues`` (plain-data copies of asi_escher's ``LayoutIssue`` records:
    ``kind``, ``message``, ``edge``, ``node``).  The same issues are appended
    to ``doc.issues`` as ``"diagram <stem>: <message>"``.
    """
    blocks = [b for b in doc.blocks if b.kind == "diagram"]
    if not blocks:
        return []
    target = os.path.join(out_dir, figures_dir) if figures_dir else out_dir
    os.makedirs(target, exist_ok=True)
    return [_render_one(block, doc, target, figures_dir, theme, mermaid) for block in blocks]


def _render_one(block: DiagramBlock, doc: Document, target: str, figures_dir: str,
                theme: dict | None, mermaid: bool) -> dict:
    # lazy: this is the only place the document engine pulls in matplotlib
    from asi_escher import Diagram, LayoutWarning, layout, render, to_dot, to_mermaid

    stem = block.stem
    diagram = Diagram.from_dict(block.spec)
    paths = {fmt: os.path.join(target, f"{stem}.{fmt}") for fmt in block.formats}
    with warnings.catch_warnings():
        # the issues are read from the diagram below; the warning must not
        # reach a host that treats warnings as errors
        warnings.simplefilter("ignore", LayoutWarning)
        layout(diagram)
    render(diagram, png=paths.get("png"), svg=paths.get("svg"), dpi=block.dpi, theme=theme)
    if "dot" in paths:
        with open(paths["dot"], "w", encoding="utf-8") as f:
            f.write(to_dot(diagram))
    if "mmd" in paths:
        with open(paths["mmd"], "w", encoding="utf-8") as f:
            f.write(to_mermaid(diagram))
    block.path = (f"{figures_dir}/{stem}.png" if figures_dir else f"{stem}.png").replace("\\", "/")
    block.mermaid = to_mermaid(diagram) if mermaid else ""
    doc.issues.extend(f"diagram {stem}: {i.message}" for i in diagram.issues)
    result: dict = {"name": stem}
    result.update(paths)
    result["size"] = (round(diagram.width), round(diagram.height))
    result["nodes"] = len(diagram.real_nodes())
    result["edges"] = len(diagram.edges)
    result["issues"] = [asdict(i) for i in diagram.issues]
    return result
