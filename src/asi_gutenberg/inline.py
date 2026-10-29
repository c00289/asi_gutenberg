"""Inline Markdown: ``**bold**``, ``*italic*``, ``code``, ``==mark==``, links and breaks.

``parse(text)`` turns a run of inline Markdown into a flat list of
:class:`Span` objects that both renderers consume: reportlab gets them as
Paragraph markup (:func:`to_pdf_markup`), python-docx as styled runs.
"""
from __future__ import annotations

import re
from dataclasses import dataclass

# characters outside the WinAnsi set of the standard PDF fonts
PDF_FALLBACK = {
    "≥": ">=", "≤": "<=", "≠": "!=", "→": "->", "←": "<-",
    "⇒": "=>", "✔": "v", "✓": "v", "✗": "x", "☐": "[ ]",
    "☑": "[x]", "☒": "[x]", "²": "2", "³": "3", "μ": "u",
    "Ω": "ohm", "Δ": "delta", "≈": "~", "−": "-", "•": "•",
}


@dataclass
class Span:
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    strike: bool = False
    href: str | None = None
    br: bool = False
    mark: bool = False  # ==highlighted==: a filled parameter, or anything a person wants to stand out (0.4.0)


_TOKEN = re.compile(
    r"(?P<esc>\\[\\`*_\[\]()~<>#|-])"
    r"|(?P<code>`+)(?P<codetext>.+?)(?P=code)"
    r"|(?P<br><br\s*/?>|  \n|\n)"
    r"|(?P<link>\[(?P<ltext>[^\]]+)\]\((?P<href>[^)\s]+)(?:\s+\"[^\"]*\")?\))"
    r"|(?P<bi>\*\*\*(?P<bitext>.+?)\*\*\*)"
    r"|(?P<b>\*\*(?P<btext>.+?)\*\*|__(?P<btext2>.+?)__)"
    r"|(?P<i>\*(?P<itext>[^*\n]+?)\*|(?<![\w])_(?P<itext2>[^_\n]+?)_(?![\w]))"
    r"|(?P<s>~~(?P<stext>.+?)~~)"
    r"|(?P<mk>==(?P<mktext>.+?)==)",
    re.DOTALL,
)


def parse(text: str, bold=False, italic=False, strike=False, href=None, mark=False) -> list[Span]:
    out: list[Span] = []
    pos = 0
    for m in _TOKEN.finditer(text):
        if m.start() > pos:
            out.append(Span(text[pos:m.start()], bold, italic, False, strike, href, mark=mark))
        g = m.groupdict()
        if g["esc"]:
            out.append(Span(g["esc"][1], bold, italic, False, strike, href, mark=mark))
        elif g["code"]:
            out.append(Span(g["codetext"], bold, italic, True, strike, href, mark=mark))
        elif g["br"]:
            out.append(Span("", br=True))
        elif g["link"]:
            out.extend(parse(g["ltext"], bold, italic, strike, g["href"], mark))
        elif g["bi"]:
            out.extend(parse(g["bitext"], True, True, strike, href, mark))
        elif g["b"]:
            out.extend(parse(g["btext"] if g["btext"] is not None else g["btext2"], True, italic, strike, href, mark))
        elif g["i"]:
            out.extend(parse(g["itext"] if g["itext"] is not None else g["itext2"], bold, True, strike, href, mark))
        elif g["s"]:
            out.extend(parse(g["stext"], bold, italic, True, href, mark))
        elif g["mk"]:
            out.extend(parse(g["mktext"], bold, italic, strike, href, True))
        pos = m.end()
    if pos < len(text):
        out.append(Span(text[pos:], bold, italic, False, strike, href, mark=mark))
    return [s for s in out if s.text or s.br]


def plain(text: str) -> str:
    """Inline Markdown with all markup removed."""
    return "".join("\n" if s.br else s.text for s in parse(text))


def _xml(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def pdf_safe(s: str) -> str:
    return "".join(PDF_FALLBACK.get(c, c) for c in s)


def to_pdf_markup(text: str, mono: str = "Courier", link_colour: str = "#2F6DB5",
                  code_bg: str = "#F1F4F8", mark_bg: str = "#A0D8A0") -> str:
    """Inline Markdown -> reportlab Paragraph markup."""
    parts = []
    for s in parse(text):
        if s.br:
            parts.append("<br/>")
            continue
        t = _xml(pdf_safe(s.text))
        if s.code:
            t = f'<font face="{mono}" backColor="{code_bg}">{t}</font>'
        elif s.mark:
            t = f'<font backColor="{mark_bg}">{t}</font>'
        if s.bold:
            t = f"<b>{t}</b>"
        if s.italic:
            t = f"<i>{t}</i>"
        if s.strike:
            t = f"<strike>{t}</strike>"
        if s.href:
            t = f'<a href="{_xml(s.href)}" color="{link_colour}"><u>{t}</u></a>'
        parts.append(t)
    return "".join(parts)
