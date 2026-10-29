"""Shared typography and colour theme for the PDF and Word renderers.

The palette is the asi_escher palette, so diagrams rendered into a document
(see the ``diagram`` block) or dropped in as ``figure`` files share the ink,
accent and tint colours of the text around them.  Everything here is plain
data; renderers copy the dict and apply overrides
(``render_pdf(..., theme={...})``).
"""
from __future__ import annotations

THEME = {
    # paper
    "page": "letter",            # letter | a4
    "orientation": "portrait",   # portrait | landscape (0.3.0): a wide table wants landscape
    "margin_in": 1.0,            # inches, all sides
    # colour
    "ink": "#1F2933",
    "muted": "#52606D",
    "accent": "#2F6DB5",         # headings, table header, links
    "accent_dark": "#1F4E85",
    "rule": "#D9DEE5",
    "band": "#1F2933",           # cover band
    "zebra_a": "#FFFFFF",
    "zebra_b": "#F1F4F8",
    "code_bg": "#F1F4F8",
    "code_ink": "#1F2933",
    "quote_bar": "#B8C2CC",
    "watermark": "#E3E8EE",
    # fonts (PDF uses the standard Type 1 faces so output is identical on
    # every machine; Word uses the metric-compatible system faces)
    "pdf_font": "Helvetica",
    "pdf_font_bold": "Helvetica-Bold",
    "pdf_font_italic": "Helvetica-Oblique",
    "pdf_font_bolditalic": "Helvetica-BoldOblique",
    "pdf_mono": "Courier",
    "docx_font": "Arial",
    "docx_mono": "Consolas",
    # sizes in points
    "body_pt": 10.0,
    "small_pt": 8.5,
    "code_pt": 8.5,
    "h1_pt": 16.0,
    "h2_pt": 13.0,
    "h3_pt": 11.5,
    "h4_pt": 10.5,
    "title_pt": 28.0,
    "leading": 1.4,
}

# Callout styles -> (label, fill, stroke).  Same tints as the asi_escher
# node kinds: note=action blue, tip=state green, important=block purple,
# warning=decision amber, caution=requirement red.
CALLOUTS = {
    "note": ("Note", "#E6F0FA", "#2F6DB5"),
    "tip": ("Tip", "#E8F5EC", "#2E8B57"),
    "important": ("Important", "#F1EDFA", "#6B46C1"),
    "warning": ("Warning", "#FFF3DC", "#C98A1B"),
    "caution": ("Caution", "#FDF0F0", "#B83232"),
}

REQUIREMENT = ("#FDF0F0", "#B83232")      # fill, stroke (asi_escher requirement node)
FIGURE_PLACEHOLDER = ("#F4F4F4", "#5F6B7A")

# Status -> chip colour used on the cover page and in the document control table.
STATUS_COLOURS = {
    "Draft": "#C98A1B",
    "In Review": "#6B46C1",
    "Approved": "#2E8B57",
    "Released": "#2F6DB5",
    "Superseded": "#5F6B7A",
    "Obsolete": "#B83232",
}


def page_size_pt(theme: dict) -> tuple[float, float]:
    """(width, height) in points for the theme's paper size and orientation."""
    w, h = (595.2756, 841.8898) if theme.get("page", "letter").lower() == "a4" else (612.0, 792.0)
    if str(theme.get("orientation", "portrait")).lower() == "landscape":
        return h, w
    return w, h
