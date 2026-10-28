"""Document data model: metadata, revision history and content blocks.

A :class:`Document` is an ordered list of blocks plus document-control
metadata.  Nothing in this module depends on reportlab or python-docx, so a
document can be built, validated and serialised (``to_dict`` / ``from_dict``)
in any process; numbering, Markdown emission and PDF / Word rendering are
separate steps (see :mod:`asi_gutenberg.layout`, :mod:`asi_gutenberg.markdown`,
:mod:`asi_gutenberg.render_pdf` and :mod:`asi_gutenberg.render_docx`).

The factory surface mirrors asi_escher: one object, short verb-like helpers
(``doc.section``, ``doc.p``, ``doc.bullets``, ``doc.table`` ...) that append
blocks and return them, strict validation on unknown kinds, and a plain-data
form for JSON specs.

A ``diagram`` block carries an asi_escher diagram as its plain-data spec
(``Diagram.to_dict()``), so a document with figures is still one JSON file;
the pipeline renders the spec to a PNG next to the Markdown (see
:mod:`asi_gutenberg.diagrams`).  asi_escher itself is imported lazily and only
when a spec has to be validated or rendered.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

# Document kinds.  The kind is printed on the cover page and in the front
# matter; it does not change the content model.
DOC_KINDS = {
    "spec", "report", "plan", "procedure", "minutes", "memo", "safety_case",
    "icd", "analysis", "checklist", "release_notes",
}

KIND_LABELS = {
    "spec": "Specification",
    "report": "Report",
    "plan": "Plan",
    "procedure": "Procedure",
    "minutes": "Meeting Minutes",
    "memo": "Memorandum",
    "safety_case": "Safety Case",
    "icd": "Interface Control Document",
    "analysis": "Analysis",
    "checklist": "Checklist",
    "release_notes": "Release Notes",
}

# Block kinds understood by the layout engine, the Markdown writer / parser
# and both renderers.
BLOCK_KINDS = {
    "heading", "paragraph", "list", "table", "code", "callout", "requirement",
    "figure", "diagram", "quote", "rule", "pagebreak", "toc", "raw",
}

# Output formats asi_escher can write for a diagram block (mirrors
# ``asi_escher.pipeline.FORMATS``; kept here so this module stays import-free).
DIAGRAM_FORMATS = ("png", "svg", "dot", "mmd")
DIAGRAM_DPI = 160

_STEM = re.compile(r"^[A-Za-z0-9_][A-Za-z0-9_.-]*$")

CALLOUT_KINDS = {"note", "tip", "important", "warning", "caution"}

ALIGNMENTS = {"left", "center", "right"}

DOC_STATUSES = {"Draft", "In Review", "Approved", "Released", "Superseded", "Obsolete"}


# ----------------------------------------------------------------- blocks --
@dataclass
class Heading:
    text: str
    level: int = 1
    numbered: bool = True
    appendix: bool = False
    # layout results
    number: str = ""
    anchor: str = ""
    kind: str = field(default="heading", init=False)

    @property
    def display(self) -> str:
        return f"{self.number} {self.text}" if self.number else self.text


@dataclass
class Paragraph:
    text: str
    kind: str = field(default="paragraph", init=False)


@dataclass
class ListItem:
    text: str
    children: list = field(default_factory=list)
    checked: bool | None = None   # None = plain item, True/False = task item


@dataclass
class ListBlock:
    items: list
    ordered: bool = False
    kind: str = field(default="list", init=False)


@dataclass
class Table:
    columns: list
    rows: list
    caption: str = ""
    align: list = field(default_factory=list)   # per column: left|center|right
    # layout results
    number: int = 0
    kind: str = field(default="table", init=False)

    @property
    def display_caption(self) -> str:
        if not self.caption:
            return ""
        return f"Table {self.number}: {self.caption}" if self.number else self.caption


@dataclass
class Code:
    text: str
    lang: str = ""
    kind: str = field(default="code", init=False)


@dataclass
class Callout:
    style: str
    text: str
    title: str = ""
    kind: str = field(default="callout", init=False)


@dataclass
class Requirement:
    req_id: str
    title: str
    text: str
    attrs: dict = field(default_factory=dict)
    kind: str = field(default="requirement", init=False)


@dataclass
class Figure:
    path: str
    caption: str = ""
    width: float = 0.0        # fraction of text width, 0 = auto
    # layout results
    number: int = 0
    kind: str = field(default="figure", init=False)

    @property
    def display_caption(self) -> str:
        if not self.caption:
            return ""
        return f"Figure {self.number}: {self.caption}" if self.number else self.caption


@dataclass
class DiagramBlock:
    """An asi_escher diagram carried as plain data and rendered at build time.

    ``spec`` is ``asi_escher.Diagram.to_dict()``: title, kind, lanes, nodes
    and edges, nothing about layout.  The block shares the figure counter
    with :class:`Figure`; ``path`` is filled by the pipeline with the relative
    (forward-slash) location of the rendered PNG, and the Markdown then shows
    the block exactly as a figure.
    """
    spec: dict
    caption: str = ""
    width: float = 0.0            # fraction of text width, 0 = auto, same as Figure
    name: str = ""                # output stem, defaults to a slug of spec["title"]
    formats: tuple = ("png",)     # Escher formats to write; "png" is always among them
    dpi: int = DIAGRAM_DPI
    # layout results
    number: int = 0               # shares the figure counter
    path: str = ""                # relative path of the rendered PNG, set by the pipeline
    mermaid: str = ""             # Mermaid source, set by the pipeline when asked for
    kind: str = field(default="diagram", init=False)

    @property
    def title(self) -> str:
        return str(self.spec.get("title", ""))

    @property
    def stem(self) -> str:
        """File stem of the rendered files: ``name`` or a slug of the title."""
        return self.name or diagram_stem(self.title)

    @property
    def display_caption(self) -> str:
        if not self.caption:
            return ""
        return f"Figure {self.number}: {self.caption}" if self.number else self.caption


def diagram_stem(title: str) -> str:
    """File-system-safe stem derived from a diagram title (same rule as asi_escher)."""
    s = re.sub(r"[^a-z0-9]+", "_", title.lower()).strip("_")
    return s or "diagram"


@dataclass
class Quote:
    text: str
    kind: str = field(default="quote", init=False)


@dataclass
class Rule:
    kind: str = field(default="rule", init=False)


@dataclass
class PageBreak:
    kind: str = field(default="pagebreak", init=False)


@dataclass
class Toc:
    depth: int = 3
    kind: str = field(default="toc", init=False)


@dataclass
class Raw:
    """Verbatim Markdown.  The escape hatch for anything the model lacks."""
    text: str
    kind: str = field(default="raw", init=False)


@dataclass
class Revision:
    rev: str
    date: str
    author: str
    description: str


@dataclass
class DocMeta:
    title: str
    kind: str = "report"
    subtitle: str = ""
    doc_id: str = ""
    revision: str = "0.1"
    status: str = "Draft"
    date: str = ""
    author: str = ""
    organisation: str = "ASI Robots"
    project: str = ""
    classification: str = ""
    approvers: list = field(default_factory=list)   # [(role, name), ...]
    keywords: list = field(default_factory=list)

    @property
    def kind_label(self) -> str:
        return KIND_LABELS.get(self.kind, self.kind.replace("_", " ").title())


# --------------------------------------------------------------- document --
class Document:
    """An ordered list of content blocks plus document-control metadata.

    Parameters
    ----------
    title:
        Human-readable title.  Used as the default output file stem.
    kind:
        One of :data:`DOC_KINDS`.
    toc:
        Insert a table of contents after the front matter.
    front_matter:
        Generate the cover page, document-control and revision-history tables.
    numbered:
        Number level 1..3 headings (``1``, ``1.2``, ``1.2.3``) at layout time.
    """

    def __init__(self, title: str, kind: str = "report", *, subtitle: str = "",
                 doc_id: str = "", revision: str = "0.1", status: str = "Draft",
                 date: str = "", author: str = "", organisation: str = "ASI Robots",
                 project: str = "", classification: str = "",
                 approvers: list | None = None, keywords: list | None = None,
                 toc: bool = True, front_matter: bool = True, numbered: bool = True):
        if kind not in DOC_KINDS:
            raise ValueError(f"unknown document kind {kind!r}; expected one of {sorted(DOC_KINDS)}")
        if status not in DOC_STATUSES:
            raise ValueError(f"unknown status {status!r}; expected one of {sorted(DOC_STATUSES)}")
        self.meta = DocMeta(title=title, kind=kind, subtitle=subtitle, doc_id=doc_id,
                            revision=revision, status=status, date=date, author=author,
                            organisation=organisation, project=project,
                            classification=classification,
                            approvers=[tuple(a) for a in (approvers or [])],
                            keywords=list(keywords or []))
        self.toc = toc
        self.front_matter = front_matter
        self.numbered = numbered
        self.blocks: list = []
        self.revisions: list[Revision] = []
        # layout results
        self.front: list = []            # generated front-matter blocks
        self.outline: list[Heading] = []
        self.requirements: dict[str, Requirement] = {}
        self.issues: list[str] = []
        self.stats: dict[str, int] = {}

    # convenience accessors
    @property
    def title(self) -> str:
        return self.meta.title

    # --- construction helpers -------------------------------------------
    def add(self, block) -> Any:
        """Append any block object.  Raises on unknown block kinds."""
        if getattr(block, "kind", None) not in BLOCK_KINDS:
            raise ValueError(f"unknown block kind {getattr(block, 'kind', None)!r}")
        self.blocks.append(block)
        return block

    def section(self, text: str, level: int = 1, numbered: bool | None = None) -> Heading:
        if not 1 <= level <= 4:
            raise ValueError(f"heading level must be 1..4, got {level}")
        return self.add(Heading(text, level, self.numbered if numbered is None else numbered))

    def h1(self, text: str, **kw) -> Heading: return self.section(text, 1, **kw)
    def h2(self, text: str, **kw) -> Heading: return self.section(text, 2, **kw)
    def h3(self, text: str, **kw) -> Heading: return self.section(text, 3, **kw)
    def h4(self, text: str, **kw) -> Heading: return self.section(text, 4, **kw)

    def appendix(self, text: str) -> Heading:
        """A level-1 heading numbered with letters (``A``, ``B`` ...)."""
        return self.add(Heading(text, 1, True, appendix=True))

    def p(self, text: str) -> Paragraph:
        return self.add(Paragraph(text))

    paragraph = p

    def bullets(self, items, ordered: bool = False) -> ListBlock:
        """Bulleted list.  Items are strings or ``(text, [children])`` tuples."""
        return self.add(ListBlock(_items(items), ordered=ordered))

    def numbered_list(self, items) -> ListBlock:
        return self.bullets(items, ordered=True)

    def tasks(self, items) -> ListBlock:
        """Task list.  Items are ``(text, done)`` tuples or plain strings (open)."""
        out = []
        for it in items:
            if isinstance(it, tuple) and len(it) == 2 and isinstance(it[1], bool):
                out.append(ListItem(it[0], checked=it[1]))
            else:
                out.append(ListItem(str(it), checked=False))
        return self.add(ListBlock(out))

    def table(self, columns, rows, caption: str = "", align: list | None = None) -> Table:
        columns = [str(c) for c in columns]
        norm = []
        for r in rows:
            r = ["" if c is None else str(c) for c in r]
            if len(r) != len(columns):
                raise ValueError(f"table row {r!r} has {len(r)} cells, expected {len(columns)}")
            norm.append(r)
        align = list(align or [])
        for a in align:
            if a not in ALIGNMENTS:
                raise ValueError(f"unknown alignment {a!r}; expected one of {sorted(ALIGNMENTS)}")
        return self.add(Table(columns, norm, caption, align))

    def fields(self, pairs, caption: str = "") -> Table:
        """Two-column key / value table (``Field`` / ``Value``)."""
        items = pairs.items() if isinstance(pairs, dict) else pairs
        return self.table(["Field", "Value"], [[k, v] for k, v in items], caption)

    def code(self, text: str, lang: str = "") -> Code:
        return self.add(Code(text.rstrip("\n"), lang))

    def callout(self, style: str, text: str, title: str = "") -> Callout:
        if style not in CALLOUT_KINDS:
            raise ValueError(f"unknown callout style {style!r}; expected one of {sorted(CALLOUT_KINDS)}")
        return self.add(Callout(style, text, title))

    def note(self, text, title=""): return self.callout("note", text, title)
    def tip(self, text, title=""): return self.callout("tip", text, title)
    def important(self, text, title=""): return self.callout("important", text, title)
    def warning(self, text, title=""): return self.callout("warning", text, title)
    def caution(self, text, title=""): return self.callout("caution", text, title)

    def requirement(self, req_id: str, title: str, text: str, **attrs) -> Requirement:
        """A traceable requirement block.  ``attrs`` become the attribute line
        (e.g. ``plr="d"``, ``status="Draft"``, ``traces="SG-01"``)."""
        return self.add(Requirement(req_id, title, text, {k: str(v) for k, v in attrs.items()}))

    def figure(self, path: str, caption: str = "", width: float = 0.0) -> Figure:
        return self.add(Figure(path, caption, width))

    def diagram(self, spec, caption: str = "", width: float = 0.0, name: str = "",
                formats: tuple = ("png",), dpi: int = DIAGRAM_DPI) -> DiagramBlock:
        """An asi_escher diagram rendered into the document as a numbered figure.

        ``spec`` is an ``asi_escher.Diagram`` or its ``to_dict()`` form.  A
        ``Diagram`` is converted to plain data at once; a dict is validated by
        round-tripping it through ``asi_escher.Diagram.from_dict`` (imported
        lazily), so an unknown node kind fails here, as every other block
        does, not at build time.  ``name`` is the output stem (defaults to a
        slug of the title), ``formats`` lists the Escher formats to write
        beside the PNG (``png``, ``svg``, ``dot``, ``mmd``) and ``dpi`` is the
        PNG resolution.
        """
        if hasattr(spec, "to_dict"):
            spec = spec.to_dict()
        elif isinstance(spec, dict):
            from asi_escher import Diagram  # lazy: pulls in matplotlib
            Diagram.from_dict(spec)             # raises ValueError on a bad spec
            spec = _copy_spec(spec)
        else:
            raise TypeError(f"diagram spec must be an asi_escher.Diagram or a dict, got {type(spec).__name__}")
        formats = _diagram_formats(formats)
        if name and not _STEM.match(name):
            raise ValueError(f"diagram name {name!r} is not a file-system-safe stem")
        if int(dpi) <= 0:
            raise ValueError(f"diagram dpi must be positive, got {dpi!r}")
        return self.add(DiagramBlock(spec, caption, width, name, formats, int(dpi)))

    def diagrams(self) -> list[DiagramBlock]:
        return [b for b in self.blocks if b.kind == "diagram"]

    def quote(self, text: str) -> Quote:
        return self.add(Quote(text))

    def rule(self) -> Rule:
        return self.add(Rule())

    def page_break(self) -> PageBreak:
        return self.add(PageBreak())

    def toc_here(self, depth: int = 3) -> Toc:
        """Place the table of contents explicitly instead of after the front matter."""
        self.toc = False
        return self.add(Toc(depth))

    def raw(self, markdown: str) -> Raw:
        return self.add(Raw(markdown))

    def revision(self, rev: str, date: str, author: str, description: str) -> Revision:
        r = Revision(rev, date, author, description)
        self.revisions.append(r)
        return r

    # --- queries -----------------------------------------------------------
    def headings(self) -> list[Heading]:
        return [b for b in self.blocks if b.kind == "heading"]

    def validate(self) -> None:
        """Raise ``ValueError`` if the document cannot be laid out consistently.

        Checked: requirement ids are unique, callout styles and table
        alignments are known, heading levels do not jump by more than one,
        diagram names are file-system-safe stems and no two diagram blocks
        would render to the same file.
        """
        seen = set()
        stems: dict[str, str] = {}
        prev = 0
        for b in self.blocks:
            if b.kind == "requirement":
                if b.req_id in seen:
                    raise ValueError(f"duplicate requirement id {b.req_id!r}")
                seen.add(b.req_id)
            elif b.kind == "callout" and b.style not in CALLOUT_KINDS:
                raise ValueError(f"unknown callout style {b.style!r}")
            elif b.kind == "table":
                for r in b.rows:
                    if len(r) != len(b.columns):
                        raise ValueError(f"table {b.caption!r}: row {r!r} has {len(r)} cells, "
                                         f"expected {len(b.columns)}")
            elif b.kind == "heading":
                if b.level > prev + 1 and prev != 0 and not b.appendix:
                    raise ValueError(f"heading {b.text!r} is level {b.level} after level {prev}")
                prev = b.level
            elif b.kind == "diagram":
                if not isinstance(b.spec, dict) or not b.spec.get("title"):
                    raise ValueError("diagram spec needs a 'title'")
                if b.name and not _STEM.match(b.name):
                    raise ValueError(f"diagram name {b.name!r} is not a file-system-safe stem")
                if b.stem in stems:
                    raise ValueError(f"diagram {b.title!r} would render to {b.stem!r}, "
                                     f"already used by diagram {stems[b.stem]!r}; give one a name")
                stems[b.stem] = b.title

    # --- serialisation -----------------------------------------------------
    def to_dict(self) -> dict[str, Any]:
        """Plain-data form of the document (no layout results)."""
        m = self.meta
        meta: dict[str, Any] = {"title": m.title, "kind": m.kind}
        for k in ("subtitle", "doc_id", "revision", "status", "date", "author",
                  "organisation", "project", "classification"):
            v = getattr(m, k)
            if v:
                meta[k] = v
        if m.approvers:
            meta["approvers"] = [list(a) for a in m.approvers]
        if m.keywords:
            meta["keywords"] = list(m.keywords)
        out: dict[str, Any] = {"meta": meta, "toc": self.toc, "front_matter": self.front_matter,
                               "numbered": self.numbered}
        if self.revisions:
            out["revisions"] = [[r.rev, r.date, r.author, r.description] for r in self.revisions]
        out["blocks"] = [_block_to_dict(b) for b in self.blocks]
        return out

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Document:
        """Build a document from the structure produced by :meth:`to_dict`."""
        meta = dict(data.get("meta", {}))
        if "title" not in meta:
            raise ValueError("document spec needs meta.title")
        d = cls(meta.pop("title"), meta.pop("kind", "report"),
                toc=data.get("toc", True), front_matter=data.get("front_matter", True),
                numbered=data.get("numbered", True), **meta)
        for r in data.get("revisions", []):
            d.revision(*r)
        for b in data.get("blocks", []):
            d.add(_block_from_dict(b))
        return d


# ---------------------------------------------------------------- helpers --
def _items(items) -> list[ListItem]:
    out = []
    for it in items:
        if isinstance(it, ListItem):
            out.append(it)
        elif isinstance(it, tuple) and len(it) == 2 and isinstance(it[1], (list, tuple)):
            out.append(ListItem(str(it[0]), _items(it[1])))
        else:
            out.append(ListItem(str(it)))
    return out


def _item_to_dict(it: ListItem) -> dict:
    d: dict[str, Any] = {"text": it.text}
    if it.children:
        d["children"] = [_item_to_dict(c) for c in it.children]
    if it.checked is not None:
        d["checked"] = it.checked
    return d


def _item_from_dict(d: dict) -> ListItem:
    return ListItem(d["text"], [_item_from_dict(c) for c in d.get("children", [])], d.get("checked"))


def _diagram_formats(formats) -> tuple:
    """Normalise a diagram's format list: a tuple, known names, ``png`` first."""
    if isinstance(formats, str):
        formats = (formats,)
    out = ["png"]
    for f in formats:
        if f not in DIAGRAM_FORMATS:
            raise ValueError(f"unknown diagram format {f!r}; expected a subset of {DIAGRAM_FORMATS}")
        if f not in out:
            out.append(f)
    return tuple(out)


def _copy_spec(spec: dict) -> dict:
    """Deep copy of a diagram spec built from JSON-shaped data only."""
    def cp(v):
        if isinstance(v, dict):
            return {k: cp(x) for k, x in v.items()}
        if isinstance(v, (list, tuple)):
            return [cp(x) for x in v]
        return v
    return cp(spec)


def _block_to_dict(b) -> dict:
    k = b.kind
    if k == "heading":
        d = {"kind": k, "text": b.text, "level": b.level}
        if not b.numbered:
            d["numbered"] = False
        if b.appendix:
            d["appendix"] = True
        return d
    if k == "paragraph":
        return {"kind": k, "text": b.text}
    if k == "list":
        d = {"kind": k, "items": [_item_to_dict(i) for i in b.items]}
        if b.ordered:
            d["ordered"] = True
        return d
    if k == "table":
        d = {"kind": k, "columns": list(b.columns), "rows": [list(r) for r in b.rows]}
        if b.caption:
            d["caption"] = b.caption
        if b.align:
            d["align"] = list(b.align)
        return d
    if k == "code":
        return {"kind": k, "text": b.text, "lang": b.lang}
    if k == "callout":
        d = {"kind": k, "style": b.style, "text": b.text}
        if b.title:
            d["title"] = b.title
        return d
    if k == "requirement":
        return {"kind": k, "id": b.req_id, "title": b.title, "text": b.text, "attrs": dict(b.attrs)}
    if k == "figure":
        d = {"kind": k, "path": b.path, "caption": b.caption}
        if b.width:
            d["width"] = b.width
        return d
    if k == "diagram":
        d = {"kind": k, "spec": _copy_spec(b.spec), "caption": b.caption}
        if b.width:
            d["width"] = b.width
        if b.name:
            d["name"] = b.name
        if tuple(b.formats) != ("png",):
            d["formats"] = list(b.formats)
        if b.dpi != DIAGRAM_DPI:
            d["dpi"] = b.dpi
        return d
    if k == "quote":
        return {"kind": k, "text": b.text}
    if k == "toc":
        return {"kind": k, "depth": b.depth}
    if k == "raw":
        return {"kind": k, "text": b.text}
    return {"kind": k}


def _block_from_dict(d: dict):
    k = d.get("kind")
    if k == "heading":
        return Heading(d["text"], d.get("level", 1), d.get("numbered", True), d.get("appendix", False))
    if k == "paragraph":
        return Paragraph(d["text"])
    if k == "list":
        return ListBlock([_item_from_dict(i) for i in d["items"]], d.get("ordered", False))
    if k == "table":
        return Table([str(c) for c in d["columns"]], [[str(c) for c in r] for r in d["rows"]],
                     d.get("caption", ""), list(d.get("align", [])))
    if k == "code":
        return Code(d["text"], d.get("lang", ""))
    if k == "callout":
        if d["style"] not in CALLOUT_KINDS:
            raise ValueError(f"unknown callout style {d['style']!r}")
        return Callout(d["style"], d["text"], d.get("title", ""))
    if k == "requirement":
        return Requirement(d["id"], d["title"], d["text"], {a: str(v) for a, v in d.get("attrs", {}).items()})
    if k == "figure":
        return Figure(d["path"], d.get("caption", ""), d.get("width", 0.0))
    if k == "diagram":
        spec = d.get("spec")
        if not isinstance(spec, dict) or not spec.get("title"):
            raise ValueError("diagram block needs a 'spec' with a 'title'")
        return DiagramBlock(_copy_spec(spec), d.get("caption", ""), d.get("width", 0.0),
                            d.get("name", ""), _diagram_formats(d.get("formats", ("png",))),
                            int(d.get("dpi", DIAGRAM_DPI)))
    if k == "quote":
        return Quote(d["text"])
    if k == "rule":
        return Rule()
    if k == "pagebreak":
        return PageBreak()
    if k == "toc":
        return Toc(d.get("depth", 3))
    if k == "raw":
        return Raw(d["text"])
    raise ValueError(f"unknown block kind {k!r}; expected one of {sorted(BLOCK_KINDS)}")
