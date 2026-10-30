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
cell keeps its paragraph formatting too. With ``mark=True`` a filled value gets the theme's green shading on
its own run, as the rendered documents paint it.

Abridged headers (0.8.0; the owner, 2026-10-29, on the QF-17 plan table, ten wordy columns over eighteen
landscape pages: "it would be amazing if the safety plan table headers were verbatim in the first page and
then abridged in all subsequent, because they are wordy"). Word cannot repeat a different header on the
later pages of one table, and the row where page one ends is known only once Word has laid the document
out, so ``abridged_headers`` (verbatim heading -> short form) is a Word pass in the same session that
prints the PDF: for each table whose first row is wholly in the mapping, Word reports the page of every
row; a table that ends on its first page is left alone; else it is split after the last row of its first
page, the first part keeps the verbatim header (shown once, not repeating), the second part gets a new
first row with the short forms in the verbatim header's run formatting, repeating on every later page, and
the two parts keep the same column widths with nothing but a 1pt empty paragraph between them. The Word
file is saved with the split and the PDF printed from it. Without Word (or without ``pdf=True``) the file
keeps the verbatim repeating header and the result says so in ``issues``.

Links (0.9.0; the owner, 2026-10-29: "could we turn the name of the safety deliverable into a clickable URL
and highlight it blue instead of green so it pops?"). The source writes an inline link as the rendered
outputs already read it, ``[Safety Plan]({{ row_1_url }})``; once the parameters are filled, a link segment
becomes a ``w:hyperlink`` (an external relationship, the run inside it carrying the paragraph's run
properties plus the ``Hyperlink`` character style, added to the document when the template lacks it, Word's
blue, a single underline and the theme's pale blue ``link_bg`` shading); a mark never applies inside a link,
blue, not green. A link whose url is empty or still holds a ``{{ placeholder }}`` collapses to its plain
text: no link, the cell as the template had it, but shaded as a gap (below).

Attention (0.10.0; the owner, 2026-10-29: "when the link is microsoft.com because of default, or blank,
highlight in pinkish red instead of blue; and when no Jira ticket is attached also highlight as pinkish
red"). A third mark, ``!!value!!``, written by :func:`~asi_gutenberg.params.fill` for the parameters
``warn`` names and for the text of a link whose url is in ``warn_urls`` (``[!!text!!](url)``): with
``mark=True`` a warn segment is shaded the theme's pinkish red ``warn_bg`` instead of the green, a link
that is one is still a ``w:hyperlink`` to its url in the Hyperlink style but in the dark red ``warn_fg``
over ``warn_bg`` instead of blue over pale blue, and a link whose url is empty or unfilled is its plain
text shaded ``warn_bg``, so a deliverable that points nowhere yet is seen. With ``mark=False`` nothing is
shaded and a link is blue: the template exactly.
"""

from __future__ import annotations

import copy
import os
import re
import tempfile
from collections.abc import Iterable
from dataclasses import dataclass, field
from typing import Any

from .importer import _date_of, clean, from_docx, paragraph_text
from .markdown import from_markdown
from .model import Document
from .params import fill, unfilled
from .theme import THEME
from .word import (  # noqa: F401  re-exported: asi_kletz imports WordSession and patches word_session here
    ABRIDGE_NEEDS_WORD,
    WD_EXPORT_PDF,
    WordSession,
    pdf_with_word,
    word_session,
)

_MARK = re.compile(r"==(.+?)==")
_WARN = re.compile(r"!!(\S(?:.*?\S)??)!!")  # !!attention!!, as inline.py reads it (0.10.0)
_LINK = re.compile(r"\[([^\]]+)\]\(([^)]*)\)")  # an inline link, as inline.py reads it, the url filled or not
_INLINE = re.compile(r"\[(?P<ltext>[^\]]+)\]\((?P<href>[^)]*)\)|==(?P<mark>.+?)==|!!(?P<warn>\S(?:.*?\S)??)!!")
HYPERLINK_COLOUR = "0563C1"  # Word's own Hyperlink blue
_RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_NBSP = " "
_SYMBOL_BULLET = ""


@dataclass
class Filled:
    """What :func:`fill_docx` wrote: the Word file, the PDF when Word printed one, how many units changed;
    ``abridged`` names (by their first heading) the tables split for abridged headers and ``pages`` is
    Word's page count, both known only when Word ran (0.8.0)."""

    docx: str
    pdf: str | None = None
    replaced: int = 0
    unfilled: list[str] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)
    abridged: list[str] = field(default_factory=list)
    pages: int | None = None


@dataclass
class _Unit:
    """One text unit of the Word body: the ``w:p`` elements that hold it (a paragraph's one, a cell's few);
    a table unit holds its cells instead, in row order, None for a cell the importer padded or a merge."""

    paragraphs: list
    cells: list | None = None
    tc: Any = None  # a cell's element, held so its identity stays its own while the walk dedupes merges


def fill_docx(template: str, source: Document | str, values: dict[str, Any], out_path: str, *,
              strict: bool = False, mark: bool = False, repeat_header: bool = True, pdf: bool = False,
              abridged_headers: dict[str, str] | None = None, warn: Iterable[str] = (),
              warn_urls: Iterable[str] = ()) -> Filled:
    """Write ``out_path``: ``template`` with the text of ``source`` (a Markdown string or the parsed
    :class:`Document`), its ``{{ name }}`` placeholders filled from ``values``.

    ``strict`` refuses a placeholder without a value (else it stays, listed in ``unfilled``); ``mark``
    shades every filled value, the ones ``warn`` names pinkish red instead of green, a link whose url is
    in ``warn_urls`` dark red over pinkish red instead of blue, a link with no url as pinkish red text
    (0.10.0); ``repeat_header`` makes each table's first row repeat on every page;
    ``pdf`` prints ``<out_path stem>.pdf`` through Word (an issue, not an error, when Word cannot);
    ``abridged_headers`` maps a table's verbatim column headings (whitespace-normalised) to short forms:
    a table whose first row is wholly in the mapping shows the verbatim header on its first page and the
    short forms on every later page (a Word pass with the print: it needs ``pdf=True`` and Word, else an
    issue; a bad mapping, or one no table's first row is wholly in, is a ``ValueError``).
    """
    import docx
    from docx.oxml.ns import qn

    mapping = _mapping(abridged_headers)
    doc = from_markdown(source) if isinstance(source, str) else source
    filled = fill(doc, values, strict=strict, mark=True, warn=warn, warn_urls=warn_urls)
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
        if _rewrite(unit, new, mark, qn, word.part):
            replaced += 1
    if title_unit is not None and filled.title:
        if _rewrite(title_unit, filled.title, mark, qn, word.part):
            replaced += 1
        word.core_properties.title = _plain(filled.title)
    if any(True for _ in word.element.body.iter(qn("w:hyperlink"))):
        _ensure_hyperlink_style(word, qn)
    if repeat_header:
        for table in word.tables:
            if table.rows:
                _repeat_header(table.rows[0]._tr, qn)
    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    plan = _abridge_plan(word, mapping, qn) if mapping is not None else []
    recommended = False
    if plan and pdf:  # Word opens a file "read-only recommended" read-only whatever it is asked; the pass
        recommended = _lift_read_only_recommended(word, qn)  # needs to write, and the session puts it back
    word.save(out_path)
    result = Filled(out_path, replaced=replaced, unfilled=left)
    if pdf:
        target = os.path.splitext(out_path)[0] + ".pdf"
        done = word_session(out_path, target, plan, read_only_recommended=recommended)
        result.issues.extend(done.issues)
        result.pdf = target if done.printed else None
        result.abridged, result.pages = done.abridged, done.pages
    elif plan:
        result.issues.append(ABRIDGE_NEEDS_WORD)
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
Segment = tuple[str, bool, str | None, bool]  # (text, marked, href, warn): a run of the rewritten text


def _plain(marked: str) -> str:
    """The text without its marks, attention marks and links: what the Word file reads back."""
    return _WARN.sub(lambda m: m.group(1), _MARK.sub(lambda m: m.group(1), _LINK.sub(lambda m: m.group(1), marked)))


def _href(url: str) -> str | None:
    """A link's target once filled; None when it is empty or still a ``{{ placeholder }}`` (the link
    collapses to its text)."""
    url = _plain(url).strip()
    return url if url and "{{" not in url else None


def _segments(marked: str) -> list[Segment]:
    """A filled text as ``(text, marked, href, warn)`` pieces in order: a ``==value==`` the marked ones, a
    ``!!value!!`` the attention ones (0.10.0), a ``[text](url)`` the linked ones (never marked, blue not
    green; attention inside one stays, a gap over a link; one whose url is unfilled is its text, a gap)."""
    out: list[Segment] = []
    pos = 0
    for m in _INLINE.finditer(marked):
        if m.start() > pos:
            out.append((marked[pos:m.start()], False, None, False))
        if m.group("mark") is not None:
            out.append((m.group("mark"), True, None, False))
        elif m.group("warn") is not None:
            out.append((m.group("warn"), False, None, True))
        else:
            href = _href(m.group("href"))
            for text, _, _, warn in _segments(m.group("ltext")):
                out.append((text, False, None, True) if href is None else (text, False, href, warn))
        pos = m.end()
    if pos < len(marked):
        out.append((marked[pos:], False, None, False))
    return out


def _cut(segments: list[Segment], start: int, end: int) -> list[Segment]:
    """The segments between plain-text offsets ``start`` and ``end``."""
    out: list[Segment] = []
    pos = 0
    for text, marked, href, warn in segments:
        lo, hi = max(start, pos), min(end, pos + len(text))
        if hi > lo:
            out.append((text[lo - pos:hi - pos], marked, href, warn))
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


def _rewrite(unit: _Unit, new: Any, mark: bool, qn, part=None) -> bool:
    """Bring the unit's text to ``new`` (a marked string; for a table, its rows of cell strings) with the
    least change; True when anything was written. ``part`` is the document part a link's relationship
    goes into (0.9.0); a text that reads the same but gains a link, or attention shading (0.10.0, with
    ``mark``), is a change."""
    if unit.cells is not None:
        changed = False
        for cell, text in zip(unit.cells, (c for row in new for c in row), strict=True):
            if cell is not None and _rewrite(cell, text, mark, qn, part):
                changed = True
        return changed
    old = clean("\n".join("".join(t for _, t in _pieces(p, qn)) for p in unit.paragraphs))
    plain = _plain(new)
    segments = _segments(new)
    if old == plain and not any(href or (warn and mark) for _, _, href, warn in segments):
        return False
    if len(unit.paragraphs) == 1 and "\n" not in plain:
        _edit_paragraph(unit.paragraphs[0], segments, mark, qn, part)
        return True
    # one paragraph per line, the cell's own paragraphs first, the last cloned for the extra lines
    lines = plain.split("\n")
    line_segments: list[list[Segment]] = []
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
        _edit_paragraph(p_el, segs, mark, qn, part)
    return True


def _edit_paragraph(p_el, segments: list[Segment], mark: bool, qn, part=None) -> None:
    """Rewrite one paragraph's text to ``segments``: the common prefix and suffix stay in their runs, the
    span between goes into new runs after the run where it starts, which lends them its formatting. A
    linked segment goes into a ``w:hyperlink`` (0.9.0), so the prefix and suffix never reach into one:
    text that reads the same but becomes a link is rewritten; an attention segment (0.10.0), with
    ``mark``, is its own shaded run the same way."""
    from docx.oxml import OxmlElement

    pieces = _pieces(p_el, qn)
    raw = "".join(t for _, t in pieces)
    new = "".join(t for t, _, _, _ in segments)
    bounded = [(text, href, warn) for text, _, href, warn in segments if href or (warn and mark)]
    if raw == new and not bounded:
        return
    a = 0
    while a < min(len(raw), len(new)) and raw[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(raw), len(new)) - a and raw[-1 - b] == new[-1 - b]:
        b += 1
    if bounded:  # the first link's start and the last link's end bound the prefix and the suffix
        offset, first, last = 0, len(new), 0
        for text, _, href, warn in segments:
            if href or (warn and mark):
                first, last = min(first, offset), max(last, offset + len(text))
            offset += len(text)
        a = min(a, first)
        b = min(b, len(new) - last)
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
    for text, marked, href, warn in middle:
        run = OxmlElement("w:r")
        rpr = source.find(qn("w:rPr"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
        if href:
            _link_run(run, qn, warn=warn and mark)
        elif warn and mark:
            _shade_run(run, qn, THEME["warn_bg"])
        elif marked and mark:
            _shade_run(run, qn)
        for i, line in enumerate(text.split("\n")):
            if i:
                run.append(OxmlElement("w:br"))
            if line:
                t = OxmlElement("w:t")
                t.text = line
                t.set(_XML_SPACE, "preserve")
                run.append(t)
        if href:
            if part is None:
                raise ValueError("a link needs the document part its relationship goes into")
            link = OxmlElement("w:hyperlink")
            link.set(qn("r:id"), part.relate_to(href, _RT_HYPERLINK, is_external=True))
            link.set(qn("w:history"), "1")
            link.append(run)
            run = link
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


def _shade_run(run, qn, colour: str | None = None) -> None:
    """Shade the run ``colour`` (default the theme's green ``mark_bg``; the pinkish red ``warn_bg`` for
    attention, 0.10.0), any shading or highlight it carried gone."""
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
    shd.set(qn("w:fill"), (colour or THEME["mark_bg"]).lstrip("#"))
    rpr.append(shd)


def _link_run(run, qn, warn: bool = False) -> None:
    """The run inside a hyperlink: the ``Hyperlink`` character style, Word's blue, a single underline and
    the theme's pale blue shading; any mark shading, highlight, colour or underline the source run carried
    goes, so a link is blue, never green (0.9.0). With ``warn`` (0.10.0: a link to a placeholder) the text
    is the theme's dark red ``warn_fg`` over the pinkish red ``warn_bg`` instead, still a link. The
    properties keep the schema's order."""
    from docx.oxml import OxmlElement

    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run.insert(0, rpr)
    for tag in ("w:rStyle", "w:color", "w:highlight", "w:u", "w:shd"):
        for old in rpr.findall(qn(tag)):
            rpr.remove(old)
    style = OxmlElement("w:rStyle")
    style.set(qn("w:val"), "Hyperlink")
    rpr.insert(0, style)
    colour = rpr.get_or_add_color()
    colour.set(qn("w:val"), THEME["warn_fg"].lstrip("#") if warn else HYPERLINK_COLOUR)
    u = rpr.get_or_add_u()
    u.set(qn("w:val"), "single")
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), THEME["warn_bg" if warn else "link_bg"].lstrip("#"))
    rpr.insert_element_before(shd, "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
                              "w:eastAsianLayout", "w:specVanish", "w:oMath")


def _ensure_hyperlink_style(word, qn) -> bool:
    """Word's ``Hyperlink`` character style (blue, underlined), added when the template has none, so the
    links the fill writes read as Word's own; True when it was added."""
    from docx.enum.style import WD_STYLE_TYPE
    from docx.shared import RGBColor

    styles = word.styles
    for style in styles.element.findall(qn("w:style")):
        name = style.find(qn("w:name"))
        if style.get(qn("w:styleId")) == "Hyperlink" or (name is not None and name.get(qn("w:val")) == "Hyperlink"):
            return False
    style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER)
    style.font.color.rgb = RGBColor.from_string(HYPERLINK_COLOUR)
    style.font.underline = True
    style.element.set(qn("w:styleId"), "Hyperlink")
    return True


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


# -- abridged headers: which tables, and the short forms -----------------------------------------------
def _heading(text: str) -> str:
    """A column heading as the mapping names it: NBSP to space, every run of whitespace (line breaks
    included) to one space, stripped."""
    return " ".join(text.replace(_NBSP, " ").split())


def _mapping(abridged_headers: dict[str, str] | None) -> dict[str, str] | None:
    """The mapping with its headings normalised; a ValueError when it is not ``{verbatim: short}`` of
    non-empty strings."""
    if abridged_headers is None:
        return None
    if not isinstance(abridged_headers, dict) or not abridged_headers:
        raise ValueError("abridged_headers: a non-empty mapping of verbatim heading -> short form")
    out: dict[str, str] = {}
    for key, value in abridged_headers.items():
        if not isinstance(key, str) or not isinstance(value, str) or not _heading(key) or not _heading(value):
            raise ValueError(f"abridged_headers: {key!r}: {value!r} is not a verbatim heading and its short form")
        out[_heading(key)] = _heading(value)
    return out


def _abridge_plan(word, mapping: dict[str, str], qn) -> list[dict[str, Any]]:
    """The tables whose first row is wholly in the mapping (every cell with text has a short form; an empty
    cell stays empty), each as Word sees it: its 1-based index among the top-level tables, its first
    heading, and the short forms cell by cell. A table only partly in the mapping is a ValueError naming
    the headings without a short form (a typo in the mapping, not a table to skip); no table at all, too."""
    from docx.text.paragraph import Paragraph as DocxParagraph

    plan: list[dict[str, Any]] = []
    partial: list[str] = []
    index = 0
    for tbl in word.element.body.iter(qn("w:tbl")):
        if any(a.tag == qn("w:tbl") for a in tbl.iterancestors()):
            continue  # a nested table: not one of Word's Document.Tables
        index += 1
        rows = tbl.findall(qn("w:tr"))
        if len(rows) < 2:
            continue
        headings = []
        for tc in rows[0].findall(qn("w:tc")):
            paragraphs = [DocxParagraph(cp, word) for cp in tc.findall(qn("w:p"))]
            headings.append(_heading(clean("\n".join(paragraph_text(cp, qn) for cp in paragraphs))))
        missing = [h for h in headings if h and h not in mapping]
        if len(missing) == len([h for h in headings if h]):
            continue  # none of its headings is in the mapping: not a table to abridge
        if missing:
            partial.append(f"table {index} ({headings[0] or '?'}): " + ", ".join(repr(h) for h in missing))
            continue
        plan.append({"index": index, "first": headings[0], "short": [mapping[h] if h else "" for h in headings]})
    if partial:
        raise ValueError("abridged_headers: headings without a short form: " + "; ".join(partial))
    if not plan:
        raise ValueError("abridged_headers: no table's first row is in the mapping")
    return plan


def _lift_read_only_recommended(word, qn) -> bool:
    """Drop a ``w:writeProtection w:recommended`` setting; True when there was one."""
    settings = word.settings.element
    wp = settings.find(qn("w:writeProtection"))
    if wp is None or wp.get(qn("w:recommended")) not in ("1", "true", "on"):
        return False
    settings.remove(wp)
    return True
