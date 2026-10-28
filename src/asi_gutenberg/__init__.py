"""asi_gutenberg - pure-Python document generator.

Builds specifications, reports, plans, procedures, minutes and safety cases
from a small factory API, lays them out (section / table / figure numbering,
front matter, table of contents), writes GitHub-flavoured Markdown as the
source of truth, then renders that Markdown to PDF (reportlab) and Word
(python-docx).  Markdown, PDF and Word are produced together as a triplet
and always agree because both binary formats are derived from the Markdown.

A document can carry its own figures: a ``diagram`` block holds an
asi_escher diagram as plain data and ``build()`` renders it to a PNG beside
the Markdown.  asi_escher (and matplotlib with it) is imported only when a
diagram is validated or rendered, never by ``import asi_gutenberg``.
"""
from .diagrams import render_diagrams
from .layout import LayoutOptions, layout
from .markdown import from_markdown, to_markdown
from .model import (
    Callout,
    Code,
    DiagramBlock,
    DocMeta,
    Document,
    Figure,
    Heading,
    ListBlock,
    ListItem,
    PageBreak,
    Paragraph,
    Quote,
    Raw,
    Requirement,
    Revision,
    Rule,
    Table,
    Toc,
)
from .pipeline import build
from .render_docx import render_docx
from .render_pdf import render_pdf

__all__ = ["Document", "Heading", "Paragraph", "ListBlock", "ListItem", "Table", "Code",
           "Callout", "Requirement", "Figure", "DiagramBlock", "Quote", "Rule", "PageBreak", "Toc",
           "Raw", "Revision", "DocMeta", "layout", "LayoutOptions", "to_markdown", "from_markdown",
           "render_diagrams", "render_pdf", "render_docx", "build"]
