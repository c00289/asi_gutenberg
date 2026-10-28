"""One-call pipeline: layout -> render diagrams -> write Markdown -> parse it
back -> render PDF + Word."""
from __future__ import annotations

import os

from .diagrams import render_diagrams
from .layout import LayoutOptions, layout
from .markdown import from_markdown, to_markdown
from .model import Document
from .render_docx import render_docx
from .render_pdf import render_pdf


def build(d: Document, out_dir: str, name: str | None = None,
          opts: LayoutOptions | None = None, theme: dict | None = None,
          figures_dir: str = "figures", diagram_theme: dict | None = None) -> dict:
    """Lay out `d` and write <name>.md/.pdf/.docx into out_dir.

    ``diagram`` blocks are rendered first, by asi_escher, into
    ``<out_dir>/<figures_dir>/<stem>.png`` (plus any extra formats the block
    asks for); ``diagram_theme`` overrides entries of ``asi_escher.THEME`` for
    every diagram.  The Markdown file is then written and parsed back; the
    PDF and the Word file are rendered from the parsed Markdown, never from
    the in-memory model, so the three outputs can never disagree.

    Returns the three output paths, ``pages``, the layout statistics,
    ``issues`` (trace-link problems and degraded diagram layouts, as strings)
    and ``diagrams``: one entry per diagram block in document order with
    ``name``, the written file paths, ``size``, ``nodes``, ``edges`` and its
    own ``issues``.
    """
    name = name or "".join(c if c.isalnum() else "_" for c in d.title.lower())
    os.makedirs(out_dir, exist_ok=True)
    opts = opts or LayoutOptions()
    layout(d, opts)
    diagrams = render_diagrams(d, out_dir, figures_dir, diagram_theme, mermaid=opts.mermaid)
    paths = {
        "md": os.path.join(out_dir, name + ".md"),
        "pdf": os.path.join(out_dir, name + ".pdf"),
        "docx": os.path.join(out_dir, name + ".docx"),
    }
    md = to_markdown(d)
    with open(paths["md"], "w", encoding="utf-8", newline="\n") as f:
        f.write(md)
    parsed = from_markdown(md)
    pdf = render_pdf(parsed, paths["pdf"], theme=theme, base_dir=out_dir)
    render_docx(parsed, paths["docx"], theme=theme, base_dir=out_dir)
    paths["pages"] = pdf["pages"]
    paths.update(d.stats)
    paths["issues"] = list(d.issues)
    paths["diagrams"] = diagrams
    return paths
