"""Import a Word document (.docx) into the :class:`Document` model, faithfully.

The owner, 2026-10-29: "bring that document into Gutenberg; we will use Gutenberg to parametrise it."
A controlled document that exists only as a Word file becomes a Gutenberg source (the Markdown that
``to_markdown`` writes) that rebuilds to the triplet, and whose blanks can then become parameters
(see :mod:`asi_gutenberg.params`).

What is read, in body order:

- ``Title`` paragraph (or the file stem) as the title; ``Heading n`` paragraphs as sections, numbered
  when Word numbered them; every other non-empty paragraph as a paragraph, consecutive list
  paragraphs (``List Bullet``, ``List Number`` or Word numbering) as one list; empty paragraphs are
  dropped.
- Every table as a ``table`` block (its first row the header); a two-column table whose first column
  reads as labels becomes a ``fields`` table (``Field`` / ``Value``). Line breaks inside a cell are
  kept (the Markdown writer turns them into ``<br>``). Non-breaking spaces become spaces, the
  Symbol-font bullet becomes ``•``, trailing whitespace goes.
- Inline pictures as ``figure`` blocks, their bytes written to ``figures_dir`` (``figures/`` beside the
  Markdown when not given) and the path kept relative, as the pipeline wants.
- Document control: a ``Document ID`` row anywhere in a front table sets ``doc_id``; a ``Document
  owner`` row sets ``author``; a footer or body line that is only a date sets ``date`` when it is
  unambiguous (ISO, or a dd / mm / yyyy stamp with a day above 12).

Not read: page orientation per section (Gutenberg lays out one page size), headers and footers
beyond the date stamp, fonts and colours (the theme has them), comments and tracked changes. Each is
reported in ``issues`` so nothing is lost silently. python-docx is imported lazily, like reportlab.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Any

from .model import Document

_NBSP = " "
_SYMBOL_BULLET = ""  # the Symbol-font bullet Word writes as a private-use character
_HEADING = re.compile(r"^Heading (\d)$")
_LIST_STYLE = re.compile(r"^List (Bullet|Number|Paragraph)", re.I)
_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_STAMP = re.compile(r"^(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})\b")
_DOC_ID = re.compile(r"^document\s*id\b", re.I)
_OWNER = re.compile(r"^document\s*owner\b", re.I)
_LABELISH = re.compile(r"^[A-Za-z][A-Za-z0-9 /&()._-]{0,40}:?$")


@dataclass
class Imported:
    """What :func:`from_docx` found: the document, the figure files it wrote and what it could not keep."""

    document: Document
    figures: list[str] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)


def clean(text: str) -> str:
    """One cell or paragraph: NBSP to space, the Symbol bullet to ``•``, every line stripped, blank lines dropped."""
    text = text.replace(_NBSP, " ").replace(_SYMBOL_BULLET, "•").replace("\r", "")
    lines = [ln.strip() for ln in text.split("\n")]
    return "\n".join(ln for ln in lines if ln)


def _date_of(text: str) -> str | None:
    """An ISO date, or a ``dd / mm / yyyy`` stamp whose day is above 12 (so the order is certain); else None."""
    text = clean(text).split("\n")[0].strip() if text else ""
    if _ISO.match(text):
        return text
    m = _STAMP.match(text)
    if m:
        a, b, year = int(m.group(1)), int(m.group(2)), m.group(3)
        if a > 12 and b <= 12:
            return f"{year}-{b:02d}-{a:02d}"
        if b > 12 and a <= 12:
            return f"{year}-{a:02d}-{b:02d}"
    return None


def paragraph_text(p, qn) -> str:
    """Every run of a paragraph, content controls (``w:sdt``), fields and text boxes included; ``Paragraph.text``
    skips a content control, which is how a Word title bound to the file's title property went missing
    (0.3.1). A ``mc:Fallback`` subtree repeats a text box for old Word versions, so it is skipped."""
    fallback = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
    parts: list[str] = []
    for node in p._p.iter():
        tag = node.tag
        if tag == qn("w:t"):
            skip = False
            parent = node.getparent()
            while parent is not None:
                if parent.tag == fallback:
                    skip = True
                    break
                parent = parent.getparent()
            if not skip:
                parts.append(node.text or "")
        elif tag == qn("w:br") or tag == qn("w:cr"):
            parts.append("\n")
        elif tag == qn("w:tab"):
            parts.append("\t")
    return "".join(parts)


def _is_fields_table(rows: list[list[str]]) -> bool:
    return bool(rows) and all(len(r) == 2 for r in rows) and all(_LABELISH.match(r[0].split("\n")[0]) for r in rows)


def from_docx(path: str, figures_dir: str | None = None, *, kind: str = "plan",
              title: str | None = None) -> Imported:
    """Read ``path`` into a :class:`Document`; see the module docstring for what is kept.

    ``figures_dir`` is where inline pictures are written (default: ``figures`` beside ``path``); the
    figure blocks keep the path relative to ``path``'s folder, which is where the Markdown source
    is expected to live. ``kind`` is Gutenberg's document kind (``plan`` for a safety plan).
    """
    import docx  # lazy: ``import asi_gutenberg`` never loads python-docx
    from docx.oxml.ns import qn
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph

    src = docx.Document(path)
    base = os.path.dirname(os.path.abspath(path))
    stem = os.path.splitext(os.path.basename(path))[0]
    figures_dir = figures_dir or os.path.join(base, "figures")
    issues: list[str] = []
    figures: list[str] = []
    meta: dict[str, Any] = {"doc_id": "", "author": "", "date": ""}

    def record_control(rows: list[list[str]]) -> None:
        for r in rows:
            if len(r) >= 2 and _DOC_ID.match(r[0]) and not meta["doc_id"]:
                meta["doc_id"] = r[1].split("\n")[0]
            if len(r) >= 2 and _OWNER.match(r[0]) and not meta["author"]:
                meta["author"] = r[1].split("\n")[0]

    # -- first pass: the title and the date stamp -------------------------------------------
    body_title = clean(src.core_properties.title or "")
    for p in src.paragraphs:
        style = p.style.name if p.style is not None else ""
        if style == "Title" and clean(paragraph_text(p, qn)):
            body_title = clean(paragraph_text(p, qn))
            break
    for section in src.sections:
        for p in list(section.footer.paragraphs) + list(section.header.paragraphs):
            found = _date_of(paragraph_text(p, qn))
            if found and not meta["date"]:
                meta["date"] = found
    doc = Document(title or body_title or stem, kind, toc=False, front_matter=False, numbered=False)

    # -- second pass: the body in order ----------------------------------------------------
    pending_list: list[str] = []
    pending_ordered = False
    numbered_headings = 0
    heading_count = 0

    def flush_list() -> None:
        nonlocal pending_list
        if pending_list:
            doc.bullets(pending_list, ordered=pending_ordered)
            pending_list = []

    def add_pictures(p: DocxParagraph) -> bool:
        wrote = False
        for blip in p._p.iter(qn("a:blip")):
            rid = blip.get(qn("r:embed"))
            if not rid or rid not in src.part.related_parts:
                continue
            part = src.part.related_parts[rid]
            ext = os.path.splitext(part.partname)[1] or ".png"
            os.makedirs(figures_dir, exist_ok=True)
            name = f"{_slug(stem)}_{len(figures) + 1}{ext}"
            target = os.path.join(figures_dir, name)
            with open(target, "wb") as f:
                f.write(part.blob)
            figures.append(target)
            doc.figure(os.path.relpath(target, base).replace(os.sep, "/"))
            wrote = True
        return wrote

    for child in src.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            p = DocxParagraph(child, src)
            style = p.style.name if p.style is not None else ""
            text = clean(paragraph_text(p, qn))
            if add_pictures(p):
                flush_list()
                continue
            if not text:
                continue
            m = _HEADING.match(style)
            if style == "Title":
                continue  # already the title
            if m:
                flush_list()
                level = min(int(m.group(1)), 4)
                heading_count += 1
                if child.find(".//" + qn("w:numPr")) is not None:
                    numbered_headings += 1
                doc.section(text, level)
                continue
            is_list = bool(_LIST_STYLE.match(style)) or child.find(".//" + qn("w:numPr")) is not None
            if is_list:
                ordered = "Number" in style
                if pending_list and ordered != pending_ordered:
                    flush_list()
                pending_ordered = ordered
                pending_list.append(text)
                continue
            flush_list()
            found = _date_of(text) if style.startswith("Footer") else None
            if found:
                meta["date"] = meta["date"] or found
                continue
            doc.p(text)
        elif tag == "tbl":
            flush_list()
            t = DocxTable(child, src)
            rows = [[clean("\n".join(paragraph_text(cp, qn) for cp in c.paragraphs)) for c in r.cells] for r in t.rows]
            rows = [r for r in rows if any(r)]
            if not rows:
                continue
            record_control(rows)
            if _is_fields_table(rows):
                doc.fields(rows)
            else:
                width = max(len(r) for r in rows)
                header = rows[0] + [""] * (width - len(rows[0]))
                body = [r + [""] * (width - len(r)) for r in rows[1:]]
                doc.table(header, body)
        else:
            continue  # sectPr and the like
    flush_list()

    doc.meta.doc_id = meta["doc_id"]
    doc.meta.author = meta["author"]
    doc.meta.date = meta["date"]
    doc.numbered = numbered_headings > 0 and numbered_headings >= heading_count / 2
    orientations = {str(s.orientation) for s in src.sections}
    if len(orientations) > 1:
        issues.append("the Word file mixes page orientations; Gutenberg lays out one page size, "
                      "so wide tables are portrait (build with --landscape)")
    for section in src.sections:
        extra = [clean(p.text) for p in list(section.header.paragraphs) + list(section.footer.paragraphs)]
        extra = [t for t in extra if t and not _date_of(t)]
        if extra:
            issues.append("header or footer text not imported: " + " | ".join(extra)[:120])
            break
    return Imported(doc, figures, issues)


def _slug(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]+", "_", text).strip("_").lower() or "figure"
