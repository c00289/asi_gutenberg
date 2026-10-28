"""Document layout: numbering, anchors, front matter and statistics.

The analogue of asi_escher's layered layout step.  ``layout(doc)`` assigns
section numbers (``1``, ``1.2``, ``A.3``), table and figure numbers (a
``diagram`` block takes the next figure number), GitHub style anchors,
builds the generated front matter (document control, approvals, revision
history, table of contents) into ``doc.front``, registers requirements and
checks their trace links, and fills ``doc.stats``.  The author's blocks are
never rewritten: numbers live in ``number`` fields next to the original
text, exactly as asi_escher keeps ``Node.label`` untouched and fills
``Node.lines``.
"""
from __future__ import annotations

import re

from .model import Document, Heading, PageBreak, Table, Toc

# requirement attributes whose values are comma separated requirement ids
TRACE_ATTRS = {"traces", "parent", "derives", "refines", "satisfies", "verifies", "allocated"}


class LayoutOptions:
    """Knobs for :func:`layout` and :func:`asi_gutenberg.build`.

    number_depth:          deepest heading level that receives a number
    toc_depth:             deepest heading level listed in the table of contents
    control_table:         generate the document-control table
    approvals:             generate the approvals table (when approvers are set)
    revision_history:      generate the revision-history table
    page_break_after_toc:  put a page break after the table of contents
    mermaid:               after each rendered diagram, also write its Mermaid
                           source as a fenced block in the Markdown so GitHub
                           shows the diagram even when the PNG is not committed;
                           the parser skips that fence, so the PDF and Word
                           files are unchanged
    """

    def __init__(self, number_depth: int = 3, toc_depth: int = 3,
                 control_table: bool = True, approvals: bool = True,
                 revision_history: bool = True, page_break_after_toc: bool = True,
                 mermaid: bool = False):
        self.number_depth = number_depth
        self.toc_depth = toc_depth
        self.control_table = control_table
        self.approvals = approvals
        self.revision_history = revision_history
        self.page_break_after_toc = page_break_after_toc
        self.mermaid = mermaid


def layout(doc: Document, opts: LayoutOptions | None = None) -> Document:
    opts = opts or LayoutOptions()
    doc.validate()
    _number_headings(doc, opts)
    _number_captions(doc)
    _anchors(doc)
    _register_requirements(doc)
    doc.front = _front_matter(doc, opts)
    doc.outline = [b for b in doc.blocks if b.kind == "heading"]
    doc.stats = _stats(doc)
    return doc


# ------------------------------------------------------------- numbering --
def _number_headings(doc: Document, opts: LayoutOptions):
    counters = [0] * 5
    saved = None               # section counters while inside the appendices
    appendix = 0
    in_appendix = False
    for b in doc.blocks:
        if b.kind != "heading":
            continue
        b.number = ""
        if not b.numbered or not doc.numbered:
            continue
        if b.appendix:
            appendix += 1
            if not in_appendix:
                saved = list(counters)
            in_appendix = True
            counters = [0] * 5
            counters[1] = appendix
            b.number = chr(ord("A") + appendix - 1)
            continue
        if b.level > opts.number_depth:
            continue
        if in_appendix and b.level == 1:
            in_appendix = False
            counters = saved or [0] * 5
        counters[b.level] += 1
        for i in range(b.level + 1, 5):
            counters[i] = 0
        parts = []
        for i in range(1, b.level + 1):
            if i == 1 and in_appendix:
                parts.append(chr(ord("A") + counters[1] - 1))
            else:
                parts.append(str(counters[i]))
        b.number = ".".join(parts)


def _number_captions(doc: Document):
    t = f = 0
    for b in doc.blocks:
        if b.kind == "table":
            t += 1
            b.number = t
        elif b.kind in ("figure", "diagram"):
            f += 1
            b.number = f


def slug(text: str) -> str:
    """GitHub-style heading anchor."""
    s = text.strip().lower()
    s = re.sub(r"[^\w\s-]", "", s)
    s = re.sub(r"\s", "-", s)
    return s


def _anchors(doc: Document):
    seen: dict[str, int] = {}
    for b in doc.blocks:
        if b.kind != "heading":
            continue
        base = slug(b.display)
        n = seen.get(base, 0)
        seen[base] = n + 1
        b.anchor = base if n == 0 else f"{base}-{n}"


# ---------------------------------------------------------- requirements --
def _register_requirements(doc: Document):
    doc.requirements = {}
    doc.issues = []
    for b in doc.blocks:
        if b.kind == "requirement":
            doc.requirements[b.req_id] = b
    for b in doc.blocks:
        if b.kind != "requirement":
            continue
        for k, v in b.attrs.items():
            if k.lower() in TRACE_ATTRS:
                for ref in re.split(r"[,\s]+", v.strip()):
                    if ref and ref not in doc.requirements:
                        doc.issues.append(f"{b.req_id}: {k} references {ref!r}, which is not defined in this document")


# ---------------------------------------------------------- front matter --
def _front_matter(doc: Document, opts: LayoutOptions) -> list:
    front: list = []
    m = doc.meta
    if doc.front_matter:
        if opts.control_table:
            rows = [["Document ID", m.doc_id], ["Title", m.title], ["Type", m.kind_label],
                    ["Revision", m.revision], ["Status", m.status], ["Date", m.date],
                    ["Author", m.author], ["Project", m.project],
                    ["Organisation", m.organisation], ["Classification", m.classification]]
            rows = [r for r in rows if r[1]]
            front.append(Heading("Document Control", 1, numbered=False))
            front.append(Table(["Field", "Value"], rows))
        if opts.approvals and m.approvers:
            front.append(Heading("Approvals", 1, numbered=False))
            front.append(Table(["Role", "Name", "Signature", "Date"],
                               [[role, name, "", ""] for role, name in m.approvers]))
        if opts.revision_history and doc.revisions:
            front.append(Heading("Revision History", 1, numbered=False))
            front.append(Table(["Rev", "Date", "Author", "Description"],
                               [[r.rev, r.date, r.author, r.description] for r in doc.revisions]))
    if doc.toc:
        if front:
            front.append(PageBreak())
        front.append(Toc(opts.toc_depth))
        if opts.page_break_after_toc:
            front.append(PageBreak())
    return front


# ------------------------------------------------------------------ stats --
def _words(s: str) -> int:
    return len(s.split())


def _list_words(items) -> int:
    return sum(_words(i.text) + _list_words(i.children) for i in items)


def _stats(doc: Document) -> dict[str, int]:
    words = sections = tables = figures = diagrams = reqs = code = lists = callouts = 0
    for b in doc.blocks:
        k = b.kind
        if k == "heading":
            sections += 1
            words += _words(b.text)
        elif k == "paragraph" or k == "quote" or k == "raw":
            words += _words(b.text)
        elif k == "list":
            lists += 1
            words += _list_words(b.items)
        elif k == "table":
            tables += 1
            words += sum(_words(c) for r in b.rows for c in r) + sum(_words(c) for c in b.columns)
        elif k == "figure":
            figures += 1
        elif k == "diagram":
            figures += 1
            diagrams += 1
        elif k == "requirement":
            reqs += 1
            words += _words(b.title) + _words(b.text)
        elif k == "code":
            code += 1
        elif k == "callout":
            callouts += 1
            words += _words(b.text)
    return {"words": words, "sections": sections, "tables": tables, "figures": figures,
            "diagrams": diagrams, "requirements": reqs, "code_blocks": code, "lists": lists,
            "callouts": callouts}
