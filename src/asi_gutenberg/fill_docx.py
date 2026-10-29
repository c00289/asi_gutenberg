"""Fill a Word original in place: the template exactly, the parameters' values written into its runs.

The owner, 2026-10-29: "can we make it exactly like the template? I still want all the columns to pop up
in every page but otherwise we need to match the template exactly."

:func:`fill_docx` takes the Word original (the controlled document, never edited), the Markdown source
:func:`~asi_gutenberg.importer.from_docx` read from it (its blanks since named ``{{ name }}``) and the
values, and writes a copy of the Word file in which only the text the source changed is different: every
style, header, footer, table, column width and page orientation is the original's, because nothing else is
touched. The first row of every table repeats at the top of each page the table spans, so the columns show
on every page. A PDF can follow through Word itself (``pdf=True``, Windows with Word installed): Word's own
print of the filled file, not a rendering.

How the text is found again: the Word body is walked exactly as the importer walks it, so the paragraphs
and cells that became the source's blocks come up in the same order, as *units* (a paragraph, a list item,
a heading, a table cell); the source must still line up unit for unit (a paragraph dropped from the source,
a table of another shape: the fill refuses, naming the first unit that does not). For a unit whose text
differs from the filled source, the longest common prefix and suffix stay and only the span between them is
rewritten, in new runs after the run where it starts, which lends them its formatting, so a bold ``Note:``
before a change or the run after it keep their own. A cell written line by line (as the role columns were)
gets one paragraph per line, the cell's own paragraphs reused and the last cloned for extra lines, so the
cell keeps its paragraph formatting too. With ``mark=True`` a filled value gets the theme's mint shading on
its own run, as the rendered documents paint it.
"""

from __future__ import annotations

import copy
import os
import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from typing import Any

from .importer import _date_of, clean, from_docx, paragraph_text
from .markdown import from_markdown
from .model import Document
from .params import fill, unfilled
from .theme import THEME

_MARK = re.compile(r"==(.+?)==")
_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_NBSP = " "
_SYMBOL_BULLET = ""
WD_EXPORT_PDF = 17  # Word's WdExportFormat.wdExportFormatPDF


@dataclass
class Filled:
    """What :func:`fill_docx` wrote: the Word file, the PDF when Word printed one, how many units changed."""

    docx: str
    pdf: str | None = None
    replaced: int = 0
    unfilled: list[str] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)


@dataclass
class _Unit:
    """One text unit of the Word body: the ``w:p`` elements that hold it (a paragraph's one, a cell's few);
    a table unit holds its cells instead, in row order, None for a cell the importer padded or a merge."""

    paragraphs: list
    cells: list | None = None
    tc: Any = None  # a cell's element, held so its identity stays its own while the walk dedupes merges


def fill_docx(template: str, source: Document | str, values: dict[str, Any], out_path: str, *,
              strict: bool = False, mark: bool = False, repeat_header: bool = True, pdf: bool = False) -> Filled:
    """Write ``out_path``: ``template`` with the text of ``source`` (a Markdown string or the parsed
    :class:`Document`), its ``{{ name }}`` placeholders filled from ``values``.

    ``strict`` refuses a placeholder without a value (else it stays, listed in ``unfilled``); ``mark``
    shades every filled value; ``repeat_header`` makes each table's first row repeat on every page;
    ``pdf`` prints ``<out_path stem>.pdf`` through Word (an issue, not an error, when Word cannot).
    """
    import docx
    from docx.oxml.ns import qn

    doc = from_markdown(source) if isinstance(source, str) else source
    filled = fill(doc, values, strict=strict, mark=True)
    left = unfilled(filled)
    want = _source_units(filled)
    with tempfile.TemporaryDirectory() as tmp:
        imported = from_docx(template, tmp)
    _align(_source_units(imported.document), want)

    word = docx.Document(template)
    title_unit, units = _walk(word, qn)
    if len(units) != len(want):  # the walk and the importer disagree: a bug here, never a silent misfill
        raise ValueError(f"the Word body walk found {len(units)} units, the importer {len(want)}")
    replaced = 0
    for unit, (_kind, new) in zip(units, want, strict=True):
        if _rewrite(unit, new, mark, qn):
            replaced += 1
    if title_unit is not None and filled.title:
        if _rewrite(title_unit, filled.title, mark, qn):
            replaced += 1
        word.core_properties.title = _plain(filled.title)
    if repeat_header:
        for table in word.tables:
            if table.rows:
                _repeat_header(table.rows[0]._tr, qn)
    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    word.save(out_path)
    result = Filled(out_path, replaced=replaced, unfilled=left)
    if pdf:
        target = os.path.splitext(out_path)[0] + ".pdf"
        issue = pdf_with_word(out_path, target)
        if issue:
            result.issues.append(issue)
        else:
            result.pdf = target
    return result


# -- the units, on both sides ------------------------------------------------------------------------
def _source_units(doc: Document) -> list[tuple[str, Any]]:
    """The document's text units in body order: ``("text", str)`` per paragraph, heading or list item,
    ``("table", rows)`` per table with the header row first (a fields table has none: its ``Field`` /
    ``Value`` header is the importer's, not the Word file's)."""
    units: list[tuple[str, Any]] = []
    for b in doc.blocks:
        if b.kind in ("heading", "paragraph"):
            units.append(("text", b.text))
        elif b.kind == "list":
            for item in _items(b.items):
                units.append(("text", item))
        elif b.kind == "table":
            rows = [list(r) for r in b.rows]
            if list(b.columns) != ["Field", "Value"]:
                rows.insert(0, list(b.columns))
            units.append(("table", rows))
    return units


def _items(items) -> list[str]:
    out: list[str] = []
    for it in items:
        out.append(it.text)
        out.extend(_items(it.children))
    return out


def _align(have: list[tuple[str, Any]], want: list[tuple[str, Any]]) -> None:
    for i, (h, w) in enumerate(zip(have, want, strict=False)):
        if h[0] != w[0]:
            raise ValueError(f"unit {i + 1}: the Word file has a {h[0]}, the source a {w[0]}")
        shapes = zip(h[1], w[1], strict=False)
        if h[0] == "table" and (len(h[1]) != len(w[1]) or any(len(a) != len(b) for a, b in shapes)):
            raise ValueError(f"unit {i + 1}: the Word table is {len(h[1])} rows, the source's {len(w[1])}")
    if len(have) != len(want):
        raise ValueError(f"the Word file has {len(have)} units, the source {len(want)}")


def _walk(word, qn) -> tuple[_Unit | None, list[_Unit]]:
    """The Word body as units, in the importer's order and by its rules: a paragraph with a picture is a
    figure (no text), an empty one is nothing, a dated footer line is the date, the first ``Title`` is the
    title; a table's rows with any text, cell by cell, a merged cell once."""
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph as DocxParagraph

    title: _Unit | None = None
    units: list[_Unit] = []
    for child in word.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            p = DocxParagraph(child, word)
            style = p.style.name if p.style is not None else ""
            text = clean(paragraph_text(p, qn))
            rids = [blip.get(qn("r:embed")) for blip in child.iter(qn("a:blip"))]
            if any(rid and rid in word.part.related_parts for rid in rids):
                continue
            if not text:
                continue
            if style == "Title":
                if title is None:
                    title = _Unit([child])
                continue
            if style.startswith("Footer") and _date_of(text):
                continue
            units.append(_Unit([child]))
        elif tag == "tbl":
            t = DocxTable(child, word)
            texts = [[clean("\n".join(paragraph_text(cp, qn) for cp in c.paragraphs)) for c in r.cells]
                     for r in t.rows]
            kept = [i for i, r in enumerate(texts) if any(r)]
            if not kept:
                continue
            width = max(len(texts[i]) for i in kept)
            seen: set[int] = set()  # ids of elements the units below hold, so lxml cannot reuse them
            cells: list[_Unit | None] = []
            for i in kept:
                row = t.rows[i].cells
                for j in range(width):
                    if j >= len(row) or id(row[j]._tc) in seen:
                        cells.append(None)
                        continue
                    seen.add(id(row[j]._tc))
                    cells.append(_Unit([cp._p for cp in row[j].paragraphs], tc=row[j]._tc))
            units.append(_Unit([], cells))
    return title, units


# -- the rewrite -----------------------------------------------------------------------------------
def _plain(marked: str) -> str:
    return _MARK.sub(lambda m: m.group(1), marked)


def _segments(marked: str) -> list[tuple[str, bool]]:
    """A filled text as ``(text, marked)`` pieces in order, a ``==value==`` the marked ones."""
    out: list[tuple[str, bool]] = []
    pos = 0
    for m in _MARK.finditer(marked):
        if m.start() > pos:
            out.append((marked[pos:m.start()], False))
        out.append((m.group(1), True))
        pos = m.end()
    if pos < len(marked):
        out.append((marked[pos:], False))
    return out


def _cut(segments: list[tuple[str, bool]], start: int, end: int) -> list[tuple[str, bool]]:
    """The segments between plain-text offsets ``start`` and ``end``."""
    out: list[tuple[str, bool]] = []
    pos = 0
    for text, marked in segments:
        lo, hi = max(start, pos), min(end, pos + len(text))
        if hi > lo:
            out.append((text[lo - pos:hi - pos], marked))
        pos += len(text)
    return out


def _pieces(p_el, qn) -> list[tuple[Any, str]]:
    """A paragraph's text-bearing nodes in order, each with the text it stands for as the importer reads it."""
    out: list[tuple[Any, str]] = []
    for node in p_el.iter():
        tag = node.tag
        if tag == qn("w:t"):
            parent = node.getparent()
            skip = False
            while parent is not None:
                if parent.tag == _FALLBACK:
                    skip = True
                    break
                parent = parent.getparent()
            if not skip:
                text = (node.text or "").replace(_NBSP, " ").replace(_SYMBOL_BULLET, "•").replace("\r", "")
                out.append((node, text))
        elif tag in (qn("w:br"), qn("w:cr")):
            out.append((node, "\n"))
        elif tag == qn("w:tab"):
            out.append((node, "\t"))
    return out


def _rewrite(unit: _Unit, new: Any, mark: bool, qn) -> bool:
    """Bring the unit's text to ``new`` (a marked string; for a table, its rows of cell strings) with the
    least change; True when anything was written."""
    if unit.cells is not None:
        changed = False
        for cell, text in zip(unit.cells, (c for row in new for c in row), strict=True):
            if cell is not None and _rewrite(cell, text, mark, qn):
                changed = True
        return changed
    old = clean("\n".join("".join(t for _, t in _pieces(p, qn)) for p in unit.paragraphs))
    plain = _plain(new)
    if old == plain:
        return False
    segments = _segments(new)
    if len(unit.paragraphs) == 1 and "\n" not in plain:
        _edit_paragraph(unit.paragraphs[0], segments, mark, qn)
        return True
    # one paragraph per line, the cell's own paragraphs first, the last cloned for the extra lines
    lines = plain.split("\n")
    line_segments: list[list[tuple[str, bool]]] = []
    pos = 0
    for line in lines:
        line_segments.append(_cut(segments, pos, pos + len(line)))
        pos += len(line) + 1
    paragraphs = list(unit.paragraphs)
    last = paragraphs[-1]
    while len(paragraphs) < len(lines):
        clone = copy.deepcopy(last)
        for node, _ in _pieces(clone, qn):
            node.getparent().remove(node)
        last.addnext(clone)
        paragraphs.append(clone)
        last = clone
    for p_el in paragraphs[len(lines):]:
        p_el.getparent().remove(p_el)
    for p_el, segs in zip(paragraphs[:len(lines)], line_segments, strict=True):
        _edit_paragraph(p_el, segs, mark, qn)
    return True


def _edit_paragraph(p_el, segments: list[tuple[str, bool]], mark: bool, qn) -> None:
    """Rewrite one paragraph's text to ``segments``: the common prefix and suffix stay in their runs, the
    span between goes into new runs after the run where it starts, which lends them its formatting."""
    from docx.oxml import OxmlElement

    pieces = _pieces(p_el, qn)
    raw = "".join(t for _, t in pieces)
    new = "".join(t for t, _ in segments)
    if raw == new:
        return
    a = 0
    while a < min(len(raw), len(new)) and raw[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(raw), len(new)) - a and raw[-1 - b] == new[-1 - b]:
        b += 1
    end = len(raw) - b
    middle = _cut(segments, a, len(new) - b)

    # every node the span [a, end) touches, from the original offsets: a text node keeps what is before
    # ``a`` (then it anchors the insertion) and what is after ``end`` (which must follow the inserted
    # runs, so it goes to a tail run when the same node held the prefix); a break or tab inside goes
    anchor = None  # the run holding the last prefix text
    prefix_node = None
    tail_children: list = []
    offset = 0
    first_touched = None  # the run of the first node at or after ``a``, when nothing anchors before it
    for node, text in pieces:
        lo, hi = offset, offset + len(text)
        offset = hi
        if hi <= a or lo >= end:
            if first_touched is None and lo >= a:
                first_touched = node.getparent()
            continue
        run = node.getparent()
        if first_touched is None and lo >= a:
            first_touched = run
        if node.tag == qn("w:t"):
            prefix = text[:a - lo] if lo < a else ""
            suffix = text[end - lo:] if hi > end else ""
            if prefix:
                node.text = prefix
                node.set(_XML_SPACE, "preserve")
                anchor, prefix_node = run, node
                if suffix:
                    t = OxmlElement("w:t")
                    t.text = suffix
                    t.set(_XML_SPACE, "preserve")
                    tail_children.append(t)
            elif suffix:
                node.text = suffix
                node.set(_XML_SPACE, "preserve")
            else:
                run.remove(node)
        else:
            run.remove(node)
    if anchor is not None:  # what the anchor run still holds after the prefix must follow the new runs
        siblings = list(anchor)
        for sibling in siblings[siblings.index(prefix_node) + 1:]:
            if sibling.tag != qn("w:rPr"):
                anchor.remove(sibling)
                tail_children.append(sibling)
        source, before = anchor, None
    elif first_touched is not None:  # the span starts the text: the new runs go before the first run
        source, before = first_touched, first_touched
    elif pieces:  # an append at the very end
        source, before = pieces[-1][0].getparent(), None
    else:  # an empty paragraph: its own last run (a cloned line keeps its runs, emptied, so their properties
        # carry to the text written there), else a new run with the paragraph mark's properties, as Word
        # gives text typed into an empty paragraph
        runs = p_el.findall(qn("w:r"))
        if runs:
            source = runs[-1]
        else:
            source = OxmlElement("w:r")
            ppr = p_el.find(qn("w:pPr"))
            mark_rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
            if mark_rpr is not None:
                source.append(copy.deepcopy(mark_rpr))
            p_el.append(source)
        before = None
    # the new text, run by run, then the tail
    new_runs = []
    for text, marked in middle:
        run = OxmlElement("w:r")
        rpr = source.find(qn("w:rPr"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        if marked and mark:
            _shade_run(run, qn)
        for i, line in enumerate(text.split("\n")):
            if i:
                run.append(OxmlElement("w:br"))
            if line:
                t = OxmlElement("w:t")
                t.text = line
                t.set(_XML_SPACE, "preserve")
                run.append(t)
        new_runs.append(run)
    if tail_children:
        tail = OxmlElement("w:r")
        rpr = source.find(qn("w:rPr"))
        if rpr is not None:
            tail.append(copy.deepcopy(rpr))
        for child in tail_children:
            tail.append(child)
        new_runs.append(tail)
    if before is not None:
        for run in new_runs:
            before.addprevious(run)
    else:
        previous = source
        for run in new_runs:
            previous.addnext(run)
            previous = run
    _show_content(source, qn)


def _shade_run(run, qn) -> None:
    from docx.oxml import OxmlElement

    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run.insert(0, rpr)
    for old in rpr.findall(qn("w:shd")) + rpr.findall(qn("w:highlight")):
        rpr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), THEME["mark_bg"].lstrip("#"))
    rpr.append(shd)


def _show_content(run, qn) -> None:
    """A content control around the run shows its text, not its placeholder prompt."""
    parent = run.getparent()
    while parent is not None:
        if parent.tag == qn("w:sdt"):
            pr = parent.find(qn("w:sdtPr"))
            if pr is not None:
                for flag in pr.findall(qn("w:showingPlcHdr")):
                    pr.remove(flag)
        parent = parent.getparent()


def _repeat_header(tr, qn) -> None:
    from docx.oxml import OxmlElement

    trpr = tr.find(qn("w:trPr"))
    if trpr is None:
        trpr = OxmlElement("w:trPr")
        ex = tr.find(qn("w:tblPrEx"))
        if ex is not None:
            ex.addnext(trpr)
        else:
            tr.insert(0, trpr)
    if trpr.find(qn("w:tblHeader")) is None:
        trpr.append(OxmlElement("w:tblHeader"))


# -- the PDF, through Word -------------------------------------------------------------------------
def pdf_with_word(docx_path: str, pdf_path: str, timeout: float = 180) -> str | None:
    """Print ``docx_path`` to ``pdf_path`` with Word (COM automation through PowerShell, Windows only);
    None when it did, else the reason as one line, and ``pdf_path`` as it was (0.7.1)."""
    if os.name != "nt":
        return "pdf through Word needs Windows with Word installed"
    # Word prints to a working name beside the target, which takes the target's place only once the print
    # is there: a failed print leaves whatever ``pdf_path`` already was (a rendering of the source, say)
    printing = os.path.splitext(os.path.abspath(pdf_path))[0] + ".printing.pdf"
    if os.path.isfile(printing):
        os.remove(printing)
    # the export is the work; Word sometimes drops the COM link right after it (RPC_E_DISCONNECTED on
    # Close or Quit), which is not a failed print, so those two are tried, never trusted, and the Word
    # process this call started (never one the person has open) is ended if Quit did not reach it
    script = (
        "$ErrorActionPreference = 'Stop'; "
        "$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id); "
        "$w = New-Object -ComObject Word.Application; $w.Visible = $false; "
        f"try {{ $d = $w.Documents.Open('{_ps(docx_path)}', $false, $true); "
        f"$d.ExportAsFixedFormat('{_ps(printing)}', {WD_EXPORT_PDF}) }} "
        "finally { try { $d.Close(0) } catch {}; try { $w.Quit() } catch {}; Start-Sleep -Milliseconds 500; "
        "Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } "
        "| Stop-Process -Force -ErrorAction SilentlyContinue }"
    )
    try:
        run = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                             capture_output=True, text=True, timeout=timeout)
    except (OSError, subprocess.TimeoutExpired) as exc:
        return f"pdf through Word failed: {exc}"
    if not os.path.isfile(printing):
        detail = (run.stderr or run.stdout or "").strip().splitlines()
        return "pdf through Word failed: " + (detail[0] if detail else f"exit {run.returncode}")
    os.replace(printing, pdf_path)
    return None


def _ps(path: str) -> str:
    return os.path.abspath(path).replace("'", "''")
