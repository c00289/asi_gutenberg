"""Render a Document to PDF with reportlab (platypus).

Cover page with a colour band and status chip, running header / footer with
"Page x of y", a clickable table of contents and PDF outline (bookmarks),
numbered headings, zebra tables with repeating header rows, tinted callouts,
requirement cards, shaded code blocks, task lists, figures (embedded when the
file exists, a placeholder otherwise) and a diagonal watermark while the
document is a draft.  Only the standard Type 1 fonts are used so the PDF is
reproducible on any machine without font installation.
"""
from __future__ import annotations

import os

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    Image,
    KeepTogether,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Preformatted,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.doctemplate import IndexingFlowable
from reportlab.platypus.flowables import HRFlowable
from reportlab.platypus.tableofcontents import TableOfContents

from .inline import pdf_safe, plain, to_pdf_markup
from .model import Document
from .theme import CALLOUTS, FIGURE_PLACEHOLDER, REQUIREMENT, STATUS_COLOURS, THEME, page_size_pt

MIN_COL_PT = 40.0  # the narrowest a table column gets before the whole table goes equal-width


def render_pdf(doc: Document, path: str, theme: dict | None = None, base_dir: str = ".") -> dict:
    """Write ``doc`` to ``path``.  Returns ``{"pages": n}``."""
    th = dict(THEME)
    if theme:
        th.update(theme)
    r = _PdfRenderer(doc, th, base_dir)
    return r.build(path)


def wants_cover(doc: Document) -> bool:
    """A cover page is produced when the document carries control metadata."""
    m = doc.meta
    return bool(doc.front_matter or doc.front or m.doc_id or m.author or m.date or m.approvers)


# ================================================================ flowables ==
class _PageCounter(IndexingFlowable):
    """Zero-size flowable that forces another multiBuild pass until the page
    count drawn in the footers equals the real page count."""

    def __init__(self):
        IndexingFlowable.__init__(self)
        self.total = 0
        self.template = None

    def wrap(self, aW, aH):
        return 0, 0

    def draw(self):
        pass

    def isSatisfied(self):
        return self.template is not None and self.template.page == self.total

    def afterBuild(self):
        if self.template is not None:
            self.total = self.template.page


class _TaskItem(Flowable):
    """A task-list line: vector check box + wrapped paragraph."""

    def __init__(self, para: Paragraph, checked: bool, indent: float, ink: str, accent: str):
        Flowable.__init__(self)
        self.para, self.checked, self.indent = para, checked, indent
        self.ink, self.accent = ink, accent
        self.box = 8.0

    def wrap(self, aW, aH):
        w, h = self.para.wrap(aW - self.indent - 16, aH)
        self.width, self.height = aW, max(h, self.box + 2) + 2
        return self.width, self.height

    def split(self, aW, aH):
        return []

    def draw(self):
        c = self.canv
        x = self.indent
        top = self.height - 2
        y = top - self.para.style.fontSize * 0.95
        c.saveState()
        c.setLineWidth(0.9)
        if self.checked:
            c.setFillColor(colors.HexColor(self.accent))
            c.setStrokeColor(colors.HexColor(self.accent))
            c.roundRect(x, y, self.box, self.box, 1.5, stroke=1, fill=1)
            c.setStrokeColor(colors.white)
            c.setLineWidth(1.4)
            c.line(x + 1.8, y + 4.0, x + 3.4, y + 2.2)
            c.line(x + 3.4, y + 2.2, x + 6.4, y + 6.2)
        else:
            c.setStrokeColor(colors.HexColor(self.ink))
            c.roundRect(x, y, self.box, self.box, 1.5, stroke=1, fill=0)
        c.restoreState()
        self.para.drawOn(c, x + 16, self.height - 2 - self.para.height)


# ================================================================= renderer ==
class _PdfRenderer:
    def __init__(self, doc: Document, th: dict, base_dir: str):
        self.doc = doc
        self.th = th
        self.base_dir = base_dir
        self.W, self.H = page_size_pt(th)
        self.margin = th["margin_in"] * inch
        self.text_w = self.W - 2 * self.margin
        self.styles = self._styles()
        self._toc_seen = False
        self._toc_depth = 3
        self._key = 0
        self.counter = _PageCounter()

    # ------------------------------------------------------------- styles --
    def _styles(self) -> dict[str, ParagraphStyle]:
        th = self.th
        body = th["body_pt"]
        lead = body * th["leading"]
        ink = colors.HexColor(th["ink"])
        accent = colors.HexColor(th["accent"])
        muted = colors.HexColor(th["muted"])
        s = {}
        s["body"] = ParagraphStyle("body", fontName=th["pdf_font"], fontSize=body, leading=lead,
                                   textColor=ink, spaceAfter=body * 0.6)
        s["cell"] = ParagraphStyle("cell", parent=s["body"], fontSize=body - 1, leading=(body - 1) * 1.3,
                                   spaceAfter=0)
        s["cell_head"] = ParagraphStyle("cell_head", parent=s["cell"], fontName=th["pdf_font_bold"],
                                        textColor=colors.white)
        s["small"] = ParagraphStyle("small", parent=s["body"], fontSize=th["small_pt"],
                                    leading=th["small_pt"] * 1.3, textColor=muted, spaceAfter=0)
        s["caption"] = ParagraphStyle("caption", parent=s["small"], fontName=th["pdf_font_italic"],
                                      alignment=TA_CENTER, spaceBefore=3, spaceAfter=body)
        s["code"] = ParagraphStyle("code", fontName=th["pdf_mono"], fontSize=th["code_pt"],
                                   leading=th["code_pt"] * 1.25, textColor=colors.HexColor(th["code_ink"]))
        s["quote"] = ParagraphStyle("quote", parent=s["body"], fontName=th["pdf_font_italic"],
                                    textColor=muted, spaceAfter=0)
        for lvl, key in ((1, "h1_pt"), (2, "h2_pt"), (3, "h3_pt"), (4, "h4_pt")):
            size = th[key]
            s[f"h{lvl}"] = ParagraphStyle(
                f"h{lvl}", fontName=th["pdf_font_bold"], fontSize=size, leading=size * 1.25,
                textColor=accent if lvl <= 2 else ink, spaceBefore=size * (1.1 if lvl == 1 else 0.9),
                spaceAfter=size * 0.45, keepWithNext=1)
        s["title"] = ParagraphStyle("title", fontName=th["pdf_font_bold"], fontSize=th["title_pt"],
                                    leading=th["title_pt"] * 1.15, textColor=ink, spaceAfter=10)
        s["subtitle"] = ParagraphStyle("subtitle", fontName=th["pdf_font"], fontSize=14, leading=18,
                                       textColor=muted, spaceAfter=18)
        s["kind"] = ParagraphStyle("kind", fontName=th["pdf_font_bold"], fontSize=9.5, leading=12,
                                   textColor=accent, spaceAfter=14)
        s["toc_title"] = ParagraphStyle("toc_title", parent=s["h1"], spaceBefore=0)
        s["toc1"] = ParagraphStyle("toc1", fontName=th["pdf_font_bold"], fontSize=body, leading=lead,
                                   textColor=ink, spaceBefore=4)
        s["toc2"] = ParagraphStyle("toc2", fontName=th["pdf_font"], fontSize=body, leading=lead,
                                   textColor=ink, leftIndent=16)
        s["toc3"] = ParagraphStyle("toc3", fontName=th["pdf_font"], fontSize=body - 0.5, leading=lead,
                                   textColor=muted, leftIndent=32)
        s["meta_k"] = ParagraphStyle("meta_k", parent=s["small"], fontName=th["pdf_font_bold"])
        s["meta_v"] = ParagraphStyle("meta_v", parent=s["body"], spaceAfter=0)
        return s

    def _m(self, text: str) -> str:
        th = self.th
        return to_pdf_markup(text, th["pdf_mono"], th["accent"], th["code_bg"], mark_bg=th["mark_bg"],
                             link_bg=th.get("link_bg", "#CFE2F3"))

    # ------------------------------------------------------------- build --
    def build(self, path: str) -> dict:
        doc = self.doc
        frame = Frame(self.margin, self.margin, self.text_w, self.H - 2 * self.margin,
                      id="body", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        cover_frame = Frame(self.margin + 0.45 * inch, self.margin, self.text_w - 0.45 * inch,
                            self.H - 2 * self.margin, id="cover", leftPadding=0, rightPadding=0,
                            topPadding=0, bottomPadding=0)
        # the other orientation, for a ``section`` block: the same margins on a page turned sideways
        turned = Frame(self.margin, self.margin, self.H - 2 * self.margin, self.W - 2 * self.margin,
                       id="turned", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        templates = [PageTemplate(id="Cover", frames=[cover_frame], onPage=self._on_cover),
                     PageTemplate(id="Body", frames=[frame], onPage=self._on_body),
                     PageTemplate(id="Turned", frames=[turned], onPage=self._on_body, pagesize=(self.H, self.W))]
        template = _DocTemplate(path, pagesize=(self.W, self.H), pageTemplates=templates,
                                leftMargin=self.margin, rightMargin=self.margin,
                                topMargin=self.margin, bottomMargin=self.margin,
                                title=doc.meta.title, author=doc.meta.author,
                                subject=doc.meta.kind_label,
                                keywords=", ".join(doc.meta.keywords), creator="asi_gutenberg")
        self.counter.template = template
        story: list = [self.counter]
        if wants_cover(doc):
            story.extend(self._cover_story())
            story.append(NextPageTemplate("Body"))
            story.append(PageBreak())
        else:
            story.append(NextPageTemplate("Body"))
            story.append(PageBreak())
            story.extend(self._title_story())
        for b in doc.front + doc.blocks:
            story.extend(self._block(b))
        template.multiBuild(story)
        return {"pages": template.page}

    # -------------------------------------------------------------- pages --
    def _on_cover(self, canv, template_doc):
        th, m = self.th, self.doc.meta
        canv.saveState()
        band_w = 0.5 * inch
        canv.setFillColor(colors.HexColor(th["band"]))
        canv.rect(0, 0, band_w, self.H, stroke=0, fill=1)
        canv.setFillColor(colors.HexColor(th["accent"]))
        canv.rect(band_w, self.H - 2.2 * inch, 0.12 * inch, 1.1 * inch, stroke=0, fill=1)
        # organisation and classification at the bottom
        canv.setFont(th["pdf_font_bold"], 10)
        canv.setFillColor(colors.HexColor(th["ink"]))
        canv.drawString(self.margin + 0.45 * inch, self.margin * 0.9, pdf_safe(m.organisation))
        if m.classification:
            canv.setFont(th["pdf_font"], 8.5)
            canv.setFillColor(colors.HexColor(th["muted"]))
            canv.drawString(self.margin + 0.45 * inch, self.margin * 0.9 - 13, pdf_safe(m.classification))
        # status chip
        chip = colors.HexColor(STATUS_COLOURS.get(m.status, th["muted"]))
        label = m.status.upper()
        canv.setFont(th["pdf_font_bold"], 8.5)
        tw = canv.stringWidth(label, th["pdf_font_bold"], 8.5) + 16
        x = self.W - self.margin - tw
        y = self.margin * 0.9 - 4
        canv.setFillColor(chip)
        canv.roundRect(x, y, tw, 16, 4, stroke=0, fill=1)
        canv.setFillColor(colors.white)
        canv.drawCentredString(x + tw / 2, y + 4.5, label)
        canv.restoreState()

    def _on_body(self, canv, template_doc):
        th, m = self.th, self.doc.meta
        W, H = canv._pagesize  # the page being drawn: a ``section`` block may have turned it
        text_w = W - 2 * self.margin
        canv.saveState()
        # header
        y = H - self.margin * 0.6
        canv.setFont(th["pdf_font"], 7.5)
        canv.setFillColor(colors.HexColor(th["muted"]))
        canv.drawString(self.margin, y, pdf_safe(m.title)[:110])
        right = " · ".join(x for x in (m.doc_id, m.kind_label) if x)
        canv.drawRightString(W - self.margin, y, pdf_safe(right))
        canv.setStrokeColor(colors.HexColor(th["rule"]))
        canv.setLineWidth(0.6)
        canv.line(self.margin, y - 5, W - self.margin, y - 5)
        # footer ("of N" comes from the previous multiBuild pass)
        y = self.margin * 0.55
        canv.line(self.margin, y + 11, W - self.margin, y + 11)
        left = " · ".join(x for x in (m.organisation, m.classification,
                                           f"Rev {m.revision}" if m.revision else "") if x)
        canv.drawString(self.margin, y, pdf_safe(left))
        total = max(self.counter.total, template_doc.page)
        canv.drawRightString(W - self.margin, y, f"Page {template_doc.page} of {total}")
        del text_w
        # watermark
        if m.status == "Draft":
            canv.setFont(th["pdf_font_bold"], 110)
            canv.setFillColor(colors.HexColor(th["watermark"]))
            canv.translate(W / 2, H / 2)
            canv.rotate(45)
            canv.drawCentredString(0, -30, "DRAFT")
        canv.restoreState()

    # -------------------------------------------------------------- cover --
    def _cover_story(self) -> list:
        s, m = self.styles, self.doc.meta
        story = [Spacer(1, 1.9 * inch)]
        kind = m.kind_label.upper()
        if m.project:
            kind = f"{pdf_safe(m.project).upper()} · {kind}"
        story.append(Paragraph(_xml(kind), s["kind"]))
        story.append(Paragraph(_xml(pdf_safe(m.title)), s["title"]))
        if m.subtitle:
            story.append(Paragraph(_xml(pdf_safe(m.subtitle)), s["subtitle"]))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor(self.th["accent"]),
                                spaceBefore=4, spaceAfter=16))
        rows = [(k, v) for k, v in (("Document ID", m.doc_id), ("Revision", m.revision),
                                    ("Status", m.status), ("Date", m.date), ("Author", m.author),
                                    ("Project", m.project)) if v]
        if rows:
            story.append(self._meta_table(rows))
        if m.approvers:
            story.append(Spacer(1, 18))
            story.append(Paragraph("APPROVALS", s["kind"]))
            story.append(self._meta_table(list(m.approvers)))
        return story

    def _meta_table(self, rows) -> Table:
        s = self.styles
        data = [[Paragraph(_xml(k), s["meta_k"]), Paragraph(_xml(pdf_safe(v)), s["meta_v"])] for k, v in rows]
        t = Table(data, colWidths=[1.3 * inch, self.text_w - 0.45 * inch - 1.3 * inch])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                               ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
        return t

    def _title_story(self) -> list:
        s, m = self.styles, self.doc.meta
        story = [Paragraph(_xml(m.kind_label.upper()), s["kind"]),
                 Paragraph(_xml(pdf_safe(m.title)), s["title"])]
        if m.subtitle:
            story.append(Paragraph(_xml(pdf_safe(m.subtitle)), s["subtitle"]))
        return story

    # ------------------------------------------------------------- blocks --
    def _block(self, b) -> list:
        k = b.kind
        fn = getattr(self, f"_b_{k}", None)
        if fn is None:
            return []
        out = fn(b)
        return out if isinstance(out, list) else [out]

    def _b_heading(self, b):
        self._key += 1
        key = f"h{self._key}"
        p = Paragraph(self._m(b.display), self.styles[f"h{min(b.level, 4)}"])
        p._gutenberg_heading = (b.level - 1, pdf_safe(plain(b.display)), key,
                                self._toc_seen and b.level <= self._toc_depth)
        out = [p]
        if b.level == 1:
            out.append(HRFlowable(width="100%", thickness=0.8, color=colors.HexColor(self.th["accent"]),
                                  spaceBefore=0, spaceAfter=6))
        return out

    def _b_paragraph(self, b):
        return Paragraph(self._m(b.text), self.styles["body"])

    def _b_quote(self, b):
        p = Paragraph(self._m(b.text), self.styles["quote"])
        t = Table([["", p]], colWidths=[4, self.text_w - 4])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), colors.HexColor(self.th["quote_bar"])),
                               ("LEFTPADDING", (1, 0), (1, 0), 10), ("TOPPADDING", (0, 0), (-1, -1), 2),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 2), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        return [t, Spacer(1, 6)]

    def _b_raw(self, b):
        return Paragraph(self._m(b.text), self.styles["body"])

    def _b_rule(self, b):
        return HRFlowable(width="100%", thickness=0.6, color=colors.HexColor(self.th["rule"]),
                          spaceBefore=6, spaceAfter=8)

    def _b_pagebreak(self, b):
        return PageBreak()

    def _b_section(self, b):
        """A new page in the block's orientation: the template whose page size matches it."""
        wants_landscape = b.orientation == "landscape"
        is_landscape = self.W > self.H
        template = "Body" if wants_landscape == is_landscape else "Turned"
        self._page_turned = template == "Turned"
        return [NextPageTemplate(template), PageBreak()]

    def _b_toc(self, b):
        self._toc_seen = True
        self._toc_depth = b.depth
        toc = TableOfContents()
        toc.levelStyles = [self.styles["toc1"], self.styles["toc2"], self.styles["toc3"]][:max(1, b.depth)]
        toc.dotsMinLevel = 0
        return [Paragraph("Contents", self.styles["toc_title"]),
                HRFlowable(width="100%", thickness=0.8, color=colors.HexColor(self.th["accent"]),
                           spaceBefore=0, spaceAfter=8), toc]

    def _b_code(self, b):
        pre = Preformatted(pdf_safe(b.text), self.styles["code"], maxLineLength=92)
        t = Table([[pre]], colWidths=[self.text_w])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(self.th["code_bg"])),
                               ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(self.th["rule"])),
                               ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                               ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
        return [t, Spacer(1, 8)]

    def _b_list(self, b):
        out = []
        self._list_items(b.items, b.ordered, 0, out)
        out.append(Spacer(1, 4))
        return out

    def _list_items(self, items, ordered, depth, out):
        th, s = self.th, self.styles
        body = th["body_pt"]
        for i, it in enumerate(items, 1):
            if it.checked is not None:
                st = ParagraphStyle(f"task{depth}", parent=s["body"], spaceAfter=0)
                para = Paragraph(self._m(it.text), st)
                out.append(_TaskItem(para, it.checked, 18 * depth + 4, th["ink"], th["accent"]))
                out.append(Spacer(1, body * 0.25))
            else:
                st = ParagraphStyle(f"li{depth}", parent=s["body"], leftIndent=18 * (depth + 1),
                                    bulletIndent=18 * depth + 4, spaceAfter=body * 0.25)
                bullet = f"{i}." if ordered else ["•", "–", "·"][depth % 3]
                out.append(Paragraph(self._m(it.text), st, bulletText=bullet))
            if it.children:
                self._list_items(it.children, ordered, depth + 1, out)

    def _b_table(self, b):
        th, s = self.th, self.styles
        ncol = len(b.columns)
        widths = self._col_widths(b)
        data = [[Paragraph(self._m(c), s["cell_head"]) for c in b.columns]]
        for r in b.rows:
            data.append([Paragraph(self._m(c), s["cell"]) for c in r])
        for ci, a in enumerate(list(b.align)[:ncol]):
            if a != "left":
                for row in data:
                    row[ci].style = ParagraphStyle("al", parent=row[ci].style,
                                                   alignment={"center": 1, "right": 2}[a])
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(th["accent"])),
            ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor(th["accent_dark"])),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [None, colors.HexColor(th["zebra_b"])]),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(th["rule"])),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5)]))
        out = [t]
        if b.caption:
            out.append(Paragraph(self._m(b.display_caption), s["caption"]))
        else:
            out.append(Spacer(1, 10))
        return out

    @property
    def text_w(self) -> float:
        """The text width of the page in use: the turned page's when a ``section`` block turned it."""
        if getattr(self, "_page_turned", False):
            return self.H - 2 * self.margin
        return self._text_w

    @text_w.setter
    def text_w(self, value: float) -> None:
        self._text_w = value

    def _col_widths(self, b) -> list[float]:
        ncol = len(b.columns)
        weights = []
        for ci in range(ncol):
            lens = [len(plain(b.columns[ci]))] + [len(plain(r[ci])) for r in b.rows]
            longest = max(lens)
            avg = sum(lens) / len(lens)
            w = min(longest, 48) * 0.6 + avg * 0.4
            weights.append(max(w, 5.0))
        total = sum(weights)
        widths = [self.text_w * w / total for w in weights]
        # a wide table (a 10-column safety plan on a portrait page) must not starve a column: every
        # column keeps at least MIN_COL_PT, the rest shrink in proportion; past the point where even
        # that does not fit, the columns are equal and the renderer wraps every cell
        floor = min(MIN_COL_PT, self.text_w / ncol)
        for _ in range(ncol):
            short = [i for i, w in enumerate(widths) if w < floor]
            if not short:
                break
            fixed = len(short) * floor
            rest = [i for i in range(ncol) if i not in short]
            rest_total = sum(widths[i] for i in rest) or 1.0
            for i in short:
                widths[i] = floor
            for i in rest:
                widths[i] = max(floor, (self.text_w - fixed) * widths[i] / rest_total)
        scale = self.text_w / sum(widths)
        return [w * scale for w in widths]

    def _b_callout(self, b):
        th, s = self.th, self.styles
        label, fill, stroke = CALLOUTS[b.style]
        title = b.title or label
        head = ParagraphStyle("co_head", parent=s["body"], fontName=th["pdf_font_bold"],
                              textColor=colors.HexColor(stroke), spaceAfter=2)
        body = ParagraphStyle("co_body", parent=s["body"], spaceAfter=0)
        content = [Paragraph(self._m(title), head)]
        for para in b.text.split("\n\n"):
            content.append(Paragraph(self._m(para.replace("\n", " ")), body))
        t = Table([["", content]], colWidths=[5, self.text_w - 5])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), colors.HexColor(stroke)),
                               ("BACKGROUND", (1, 0), (1, 0), colors.HexColor(fill)),
                               ("LEFTPADDING", (1, 0), (1, 0), 10), ("RIGHTPADDING", (1, 0), (1, 0), 10),
                               ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                               ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        return [KeepTogether([t]), Spacer(1, 10)]

    def _b_requirement(self, b):
        th, s = self.th, self.styles
        fill, stroke = REQUIREMENT
        head = ParagraphStyle("rq_head", parent=s["body"], fontName=th["pdf_font_bold"],
                              textColor=colors.HexColor(stroke), spaceAfter=0)
        stereo = ParagraphStyle("rq_stereo", parent=s["small"], textColor=colors.HexColor(stroke))
        body = ParagraphStyle("rq_body", parent=s["body"], spaceAfter=0)
        top = [Paragraph("«requirement»", stereo),
               Paragraph(f"{_xml(b.req_id)} — {self._m(b.title)}", head)]
        rows = [["", top], ["", Paragraph(self._m(b.text.replace("\n", "<br/>")), body)]]
        if b.attrs:
            attrs = " · ".join(f"<b>{_xml(k)}</b> {_xml(pdf_safe(v))}" for k, v in b.attrs.items())
            rows.append(["", Paragraph(attrs, s["small"])])
        t = Table(rows, colWidths=[5, self.text_w - 5])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), colors.HexColor(stroke)),
                               ("BACKGROUND", (1, 0), (1, 0), colors.HexColor(fill)),
                               ("BOX", (1, 0), (-1, -1), 0.5, colors.HexColor(stroke)),
                               ("LINEBELOW", (1, 0), (1, 0), 0.5, colors.HexColor(stroke)),
                               ("LEFTPADDING", (1, 0), (1, -1), 10), ("RIGHTPADDING", (1, 0), (1, -1), 10),
                               ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                               ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        return [KeepTogether([t]), Spacer(1, 10)]

    def _b_figure(self, b):
        s = self.styles
        path = b.path if os.path.isabs(b.path) else os.path.join(self.base_dir, b.path)
        frac = b.width or 0.9
        out = []
        if os.path.isfile(path):
            try:
                iw, ih = ImageReader(path).getSize()
                w = self.text_w * frac
                h = w * ih / iw
                max_h = (self.H - 2 * self.margin) * 0.72
                if h > max_h:
                    h, w = max_h, max_h * iw / ih
                out.append(Image(path, width=w, height=h))
            except Exception:
                out.append(self._placeholder(b))
        else:
            out.append(self._placeholder(b))
        if b.caption:
            out.append(Paragraph(self._m(b.display_caption), s["caption"]))
        else:
            out.append(Spacer(1, 10))
        return [KeepTogether(out)]

    def _placeholder(self, b):
        fill, stroke = FIGURE_PLACEHOLDER
        st = ParagraphStyle("ph", parent=self.styles["small"], alignment=TA_CENTER)
        p = Paragraph(f"figure placeholder<br/><font face='{self.th['pdf_mono']}'>{_xml(b.path)}</font>", st)
        t = Table([[p]], colWidths=[self.text_w * 0.7], rowHeights=[1.1 * inch])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(fill)),
                               ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(stroke)),
                               ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        return t


class _DocTemplate(BaseDocTemplate):
    def afterFlowable(self, flowable):
        info = getattr(flowable, "_gutenberg_heading", None)
        if info is None:
            return
        level, text, key, in_toc = info
        self.canv.bookmarkPage(key)
        self.canv.addOutlineEntry(text, key, level=min(level, 3), closed=level >= 1)
        if in_toc:
            self.notify("TOCEntry", (level, _xml(text), self.page, key))

    def handle_documentBegin(self):
        super().handle_documentBegin()
        self.canv.showOutline()


def _xml(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
