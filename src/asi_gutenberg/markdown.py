"""Markdown is the source of truth: write a Document as GitHub-flavoured
Markdown and read it back.

``to_markdown(doc)`` emits YAML front matter (document control metadata),
the generated front-matter sections, a table of contents between
``<!-- toc -->`` markers and then every block using plain GFM: ATX headings,
pipe tables, fenced code, ``> [!NOTE]`` alerts, task lists and images.  The
file renders correctly on GitHub, in VS Code and in any Markdown viewer.  A
``diagram`` block is written exactly as a figure, ``![Figure n: caption]
(figures/stem.png)``, pointing at the PNG the pipeline rendered next to the
file (optionally followed by a ``mermaid`` fence with the same diagram as
text, which the parser skips).

``from_markdown(text)`` parses that dialect (and ordinary hand-written
Markdown) back into a :class:`Document`.  The PDF and Word renderers consume
the parsed document, so both are *derived from the Markdown* rather than from
the in-memory model: whatever the ``.md`` says is what the ``.pdf`` and
``.docx`` show.
"""
from __future__ import annotations

import re

from .model import (
    CALLOUT_KINDS,
    DOC_KINDS,
    DOC_STATUSES,
    Callout,
    Code,
    Document,
    Figure,
    Heading,
    ListBlock,
    ListItem,
    PageBreak,
    Paragraph,
    Quote,
    Requirement,
    Rule,
    Table,
    Toc,
)

PAGEBREAK = "<!-- pagebreak -->"
TOC_OPEN = "<!-- toc -->"
TOC_CLOSE = "<!-- /toc -->"


# ================================================================= writer ==
def to_markdown(doc: Document) -> str:
    out: list[str] = []
    out.append(_front_matter_yaml(doc))
    out.append(f"# {doc.meta.title}")
    if doc.meta.subtitle:
        out.append(f"*{doc.meta.subtitle}*")
    out.append("")
    toc_depth = 3
    for b in doc.front + doc.blocks:
        if b.kind == "toc":
            toc_depth = b.depth
            out.append(_toc_markdown(doc, toc_depth))
        else:
            out.append(_block_markdown(b))
        out.append("")
    text = "\n".join(out)
    return re.sub(r"\n{3,}", "\n\n", text).rstrip("\n") + "\n"


def _yaml_str(v: str) -> str:
    reserved = v.lower() in ("yes", "no", "true", "false", "null")
    if v == "" or re.search(r'[:#\[\]{}&*!|>\'"%@`,]|^\s|\s$', v) or reserved:
        return '"' + v.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return v


def _front_matter_yaml(doc: Document) -> str:
    m = doc.meta
    lines = ["---", f"title: {_yaml_str(m.title)}", f"kind: {m.kind}"]
    for k in ("subtitle", "doc_id", "revision", "status", "date", "author", "organisation",
              "project", "classification"):
        v = getattr(m, k)
        if v:
            lines.append(f"{k}: {_yaml_str(str(v))}")
    if m.approvers:
        lines.append("approvers:")
        for role, name in m.approvers:
            lines.append(f"  - [{_yaml_str(role)}, {_yaml_str(name)}]")
    if m.keywords:
        lines.append("keywords: [" + ", ".join(_yaml_str(k) for k in m.keywords) + "]")
    lines.append(f"numbered: {'true' if doc.numbered else 'false'}")
    lines.append("generator: asi_gutenberg")
    lines.append("---")
    return "\n".join(lines)


def _toc_markdown(doc: Document, depth: int) -> str:
    lines = [TOC_OPEN, "## Contents", ""]
    start = next((i for i, b in enumerate(doc.blocks) if b.kind == "toc"), -1)
    for h in doc.blocks[start + 1:]:
        if h.kind != "heading" or h.level > depth:
            continue
        lines.append("  " * (h.level - 1) + f"- [{h.display}](#{h.anchor})")
    lines.append(TOC_CLOSE)
    return "\n".join(lines)


def _cell(s: str) -> str:
    return s.replace("|", "\\|").replace("\n", "<br>")


def _block_markdown(b) -> str:
    k = b.kind
    if k == "heading":
        return "#" * (b.level + 1) + " " + b.display
    if k == "paragraph":
        return b.text
    if k == "list":
        return _list_markdown(b.items, b.ordered, 0)
    if k == "table":
        align = list(b.align) + ["left"] * (len(b.columns) - len(b.align))
        sep = []
        for a in align:
            sep.append({"left": ":---", "center": ":---:", "right": "---:"}[a])
        lines = ["| " + " | ".join(_cell(c) for c in b.columns) + " |",
                 "| " + " | ".join(sep) + " |"]
        for r in b.rows:
            lines.append("| " + " | ".join(_cell(c) for c in r) + " |")
        if b.caption:
            lines.append("")
            lines.append(f"*{b.display_caption}*")
        return "\n".join(lines)
    if k == "code":
        fence = "````" if "```" in b.text else "```"
        return f"{fence}{b.lang}\n{b.text}\n{fence}"
    if k == "callout":
        lines = [f"> [!{b.style.upper()}]"]
        if b.title:
            lines.append(f"> **{b.title}**")
        lines.extend(f"> {ln}" if ln else ">" for ln in b.text.split("\n"))
        return "\n".join(lines)
    if k == "requirement":
        lines = [f"> **{b.req_id}** \u2014 {b.title}"]
        lines.extend(f"> {ln}" if ln else ">" for ln in b.text.split("\n"))
        if b.attrs:
            lines.append(">")
            lines.append("> _" + " \u00b7 ".join(f"{a} = {v}" for a, v in b.attrs.items()) + "_")
        return "\n".join(lines)
    if k == "figure":
        alt = b.display_caption or b.path
        return f"![{alt}]({b.path})"
    if k == "diagram":
        if not b.path:
            raise ValueError(f"diagram {b.title!r} has not been rendered; run build() or render_diagrams() first")
        alt = b.display_caption or b.path
        line = f"![{alt}]({b.path})"
        if b.mermaid:
            line += f"\n\n```mermaid\n{b.mermaid.rstrip()}\n```"
        return line
    if k == "quote":
        return "\n".join(f"> {ln}" if ln else ">" for ln in b.text.split("\n"))
    if k == "rule":
        return "---"
    if k == "pagebreak":
        return PAGEBREAK
    if k == "raw":
        return b.text
    if k == "toc":
        return TOC_OPEN + "\n" + TOC_CLOSE
    raise ValueError(f"cannot write block kind {k!r}")


def _list_markdown(items: list[ListItem], ordered: bool, depth: int) -> str:
    lines = []
    for i, it in enumerate(items, 1):
        marker = f"{i}." if ordered else "-"
        box = ""
        if it.checked is not None:
            box = "[x] " if it.checked else "[ ] "
        indent = "   " if ordered else "  "
        lines.append(indent * depth + f"{marker} {box}{it.text}")
        if it.children:
            lines.append(_list_markdown(it.children, ordered, depth + 1))
    return "\n".join(lines)


# ================================================================= parser ==
_H = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
_NUM = re.compile(r"^((?:\d{1,3}(?:\.\d{1,3})+)|(?:\d{1,2})|(?:[A-Z](?:\.\d{1,3})*))\s+(\S.*)$")
_LIST = re.compile(r"^(\s*)([-*+]|\d+[.)])\s+(\[([ xX])\]\s+)?(.*)$")
_IMG = re.compile(r"^!\[(.*?)\]\((\S+?)(?:\s+\"[^\"]*\")?\)\s*$")
_RULE = re.compile(r"^\s*([-*_])(\s*\1){2,}\s*$")
_FENCE = re.compile(r"^(`{3,}|~{3,})\s*(\S*)\s*$")
_ALERT = re.compile(r"^\[!(\w+)\]\s*$")
_REQ = re.compile(r"^\*\*([^*]+)\*\*\s*(?:\u2014|--|-|:)\s*(.*)$")
_ATTRS = re.compile(r"^_(.+)_$")
_TCAP = re.compile(r"^\*Table (\d+): (.*)\*$")
_COMMENT = re.compile(r"^<!--.*?-->\s*$")


def from_markdown(text: str) -> Document:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    meta, i = _parse_front_matter(lines)
    title = meta.pop("title", None)
    kind = meta.pop("kind", "report")
    if kind not in DOC_KINDS:
        kind = "report"
    numbered = _as_bool(meta.pop("numbered", True))
    meta.pop("generator", None)
    approvers = meta.pop("approvers", [])
    keywords = meta.pop("keywords", [])
    status = meta.pop("status", "Draft")
    if status not in DOC_STATUSES:
        status = "Draft"
    allowed = {"subtitle", "doc_id", "revision", "date", "author", "organisation", "project",
               "classification"}
    kw = {k: str(v) for k, v in meta.items() if k in allowed}
    doc = Document(title or "Untitled", kind, status=status, approvers=approvers,
                   keywords=keywords, toc=False, front_matter=False, numbered=numbered, **kw)
    blocks = doc.blocks
    para: list[str] = []
    have_title = title is not None
    seen_title_line = False

    def flush():
        nonlocal para
        if not para:
            return
        txt = " ".join(ln.strip() for ln in para).strip()
        para = []
        if not txt:
            return
        m = _TCAP.match(txt)
        if m and blocks and blocks[-1].kind == "table" and not blocks[-1].caption:
            blocks[-1].caption = m.group(2)
            blocks[-1].number = int(m.group(1))
            return
        if (not blocks and have_title and txt.startswith("*") and txt.endswith("*")
                and txt.count("*") == 2):
            # the italic line under the title is the subtitle (already in the
            # front matter when the file was written by to_markdown)
            if not doc.meta.subtitle:
                doc.meta.subtitle = txt[1:-1]
            return
        blocks.append(Paragraph(txt))

    n = len(lines)
    while i < n:
        ln = lines[i]
        s = ln.strip()
        # -- blank
        if not s:
            flush()
            i += 1
            continue
        # -- comments / markers
        if s == TOC_OPEN:
            flush()
            j = i + 1
            while j < n and lines[j].strip() != TOC_CLOSE:
                j += 1
            blocks.append(Toc())
            i = j + 1
            continue
        if s == PAGEBREAK:
            flush()
            blocks.append(PageBreak())
            i += 1
            continue
        if _COMMENT.match(s):
            i += 1
            continue
        # -- fenced code
        m = _FENCE.match(s)
        if m:
            flush()
            fence, lang = m.group(1), m.group(2)
            j = i + 1
            body = []
            closer = fence[0] * len(fence)
            while j < n and not (lines[j].strip().startswith(closer) and lines[j].strip().rstrip(fence[0]) == ""):
                body.append(lines[j])
                j += 1
            if lang == "mermaid" and not para and blocks and blocks[-1].kind == "figure":
                # the text form of the figure just above (LayoutOptions(mermaid=True)):
                # the image already shows it, so the PDF and Word files skip it
                i = j + 1
                continue
            blocks.append(Code("\n".join(body), lang))
            i = j + 1
            continue
        # -- heading
        m = _H.match(ln)
        if m:
            flush()
            level = len(m.group(1))
            txt = m.group(2)
            if level == 1 and not seen_title_line and not blocks:
                # the single H1 is the document title, not a section
                if not have_title:
                    doc.meta.title = txt
                    have_title = True
                seen_title_line = True
                i += 1
                continue
            lvl = max(1, min(4, level - 1 if have_title else level))
            h = Heading(txt, lvl, numbered=False)
            mm = _NUM.match(txt) if numbered else None
            if mm:
                h.number, h.text, h.numbered = mm.group(1), mm.group(2), True
                h.appendix = h.number[0].isalpha() and "." not in h.number
            from .layout import slug
            h.anchor = slug(h.display)
            blocks.append(h)
            i += 1
            continue
        # -- table
        if s.startswith("|") and i + 1 < n and _is_sep(lines[i + 1]):
            flush()
            header = _split_row(s)
            align = []
            for c in _split_row(lines[i + 1].strip()):
                c = c.strip()
                align.append("center" if c.startswith(":") and c.endswith(":") else
                             "right" if c.endswith(":") else "left")
            rows = []
            j = i + 2
            while j < n and lines[j].strip().startswith("|"):
                r = _split_row(lines[j].strip())
                r = (r + [""] * len(header))[:len(header)]
                rows.append(r)
                j += 1
            blocks.append(Table(header, rows, "", align))
            i = j
            continue
        # -- blockquote: callout, requirement or quote
        if s.startswith(">"):
            flush()
            q = []
            j = i
            while j < n and lines[j].strip().startswith(">"):
                q.append(re.sub(r"^\s*>\s?", "", lines[j]))
                j += 1
            blocks.append(_quote_block(q))
            i = j
            continue
        # -- list
        m = _LIST.match(ln)
        if m and not _RULE.match(s):
            flush()
            j = i
            raw = []
            while j < n:
                l2 = lines[j]
                if _LIST.match(l2) or (l2.strip() and l2.startswith("  ") and raw):
                    raw.append(l2)
                    j += 1
                else:
                    break
            blocks.append(_list_block(raw))
            i = j
            continue
        # -- image
        m = _IMG.match(s)
        if m:
            flush()
            alt, path = m.group(1), m.group(2)
            fig = Figure(path, alt)
            mm = re.match(r"^Figure (\d+): (.*)$", alt)
            if mm:
                fig.number, fig.caption = int(mm.group(1)), mm.group(2)
            elif alt == path:
                fig.caption = ""
            blocks.append(fig)
            i += 1
            continue
        # -- rule
        if _RULE.match(s):
            flush()
            blocks.append(Rule())
            i += 1
            continue
        # -- paragraph text
        para.append(ln)
        i += 1
    flush()
    return doc


def _as_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("1", "true", "yes", "on")


def _is_sep(ln: str) -> bool:
    s = ln.strip()
    if not s.startswith("|"):
        return False
    cells = _split_row(s)
    return bool(cells) and all(re.fullmatch(r"\s*:?-{1,}:?\s*", c) for c in cells)


def _split_row(s: str) -> list[str]:
    s = s.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|") and not s.endswith("\\|"):
        s = s[:-1]
    cells = re.split(r"(?<!\\)\|", s)
    return [c.strip().replace("\\|", "|").replace("<br>", "\n").replace("<br/>", "\n") for c in cells]


def _quote_block(q: list[str]):
    while q and not q[-1].strip():
        q.pop()
    if not q:
        return Quote("")
    m = _ALERT.match(q[0].strip())
    if m and m.group(1).lower() in CALLOUT_KINDS:
        body = q[1:]
        title = ""
        if body and re.fullmatch(r"\*\*[^*]+\*\*", body[0].strip()):
            title = body[0].strip()[2:-2]
            body = body[1:]
        return Callout(m.group(1).lower(), "\n".join(body).strip(), title)
    m = _REQ.match(q[0].strip())
    if m:
        body = q[1:]
        attrs: dict[str, str] = {}
        if body:
            am = _ATTRS.match(body[-1].strip())
            if am and "=" in am.group(1):
                for part in am.group(1).split(" \u00b7 "):
                    if "=" in part:
                        a, v = part.split("=", 1)
                        attrs[a.strip()] = v.strip()
                body = body[:-1]
        return Requirement(m.group(1).strip(), m.group(2).strip(), "\n".join(body).strip(), attrs)
    return Quote("\n".join(q).strip())


def _list_block(raw: list[str]) -> ListBlock:
    root = ListItem("", [])
    stack: list[tuple[int, ListItem]] = [(-1, root)]
    ordered: bool | None = None
    for ln in raw:
        m = _LIST.match(ln)
        if not m:
            # continuation line
            stack[-1][1].text += " " + ln.strip()
            continue
        indent = len(m.group(1).replace("\t", "    "))
        marker = m.group(2)
        if ordered is None:
            ordered = marker[0].isdigit()
        checked = None
        if m.group(3):
            checked = m.group(4).lower() == "x"
        item = ListItem(m.group(5).strip(), [], checked)
        while stack and stack[-1][0] >= indent:
            stack.pop()
        stack[-1][1].children.append(item)
        stack.append((indent, item))
    return ListBlock(root.children, bool(ordered))


def _parse_front_matter(lines: list[str]) -> tuple[dict, int]:
    if not lines or lines[0].strip() != "---":
        return {}, 0
    meta: dict = {}
    i = 1
    key = None
    while i < len(lines):
        ln = lines[i]
        if ln.strip() == "---":
            return meta, i + 1
        if re.match(r"^\s+-\s+", ln) and key:
            v = ln.strip()[1:].strip()
            if v.startswith("[") and v.endswith("]"):
                meta.setdefault(key, []).append([_unq(x) for x in _split_list(v[1:-1])])
            else:
                meta.setdefault(key, []).append(_unq(v))
        elif ":" in ln and not ln.startswith(" "):
            key, v = ln.split(":", 1)
            key, v = key.strip(), v.strip()
            if v == "":
                meta[key] = []
            elif v.startswith("[") and v.endswith("]"):
                meta[key] = [_unq(x) for x in _split_list(v[1:-1])]
            else:
                meta[key] = _unq(v)
        i += 1
    return {}, 0


def _split_list(s: str) -> list[str]:
    return [x for x in re.split(r",\s*(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)", s) if x.strip()]


def _unq(s: str) -> str:
    s = s.strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
        return s[1:-1].replace('\\"', '"').replace("\\\\", "\\")
    return s
