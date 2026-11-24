# -*- coding: utf-8 -*-
"""Design system for the Persian deck: palette, RTL text helpers, native shapes."""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml
from lxml import etree

CREAM = RGBColor(0xFB, 0xF7, 0xEF)
INK = RGBColor(0x17, 0x21, 0x1E)
EMERALD = RGBColor(0x0E, 0x4D, 0x45)
EMERALD_DK = RGBColor(0x08, 0x2F, 0x2B)
GOLD = RGBColor(0xC4, 0x9A, 0x2B)
GOLD_L = RGBColor(0xE4, 0xCC, 0x8A)
MUTED = RGBColor(0x5F, 0x6F, 0x69)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xE6, 0xDF, 0xCF)
SAND = RGBColor(0xF3, 0xEC, 0xDC)
FONT = "IRANSans"

SW, SH = 13.333, 7.5
MX = 0.78                     # side margin
CW = SW - 2 * MX              # content width
NUM_TOTAL = 15

FA_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def fa(n):
    return str(n).translate(FA_DIGITS)


def rtl(p):
    """Force right-to-left paragraph direction."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("rtl", "1")
    pPr.set("algn", "r")
    return p


def _cs(run, name=FONT):
    rPr = run._r.get_or_add_rPr()
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = etree.SubElement(rPr, qn("a:cs"))
    cs.set("typeface", name)
    latin = rPr.find(qn("a:latin"))
    if latin is not None:
        latin.set("typeface", name)


def add_text(slide, x, y, w, h, text, size=15, color=INK, bold=False,
             align=PP_ALIGN.RIGHT, line=1.45, space_after=0, anchor=MSO_ANCHOR.TOP,
             italic=False, font=FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line
        p.space_after = Pt(space_after)
        rtl(p)
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
        r.font.name = font
        _cs(r, font)
    return box


def rect(slide, x, y, w, h, fill=CARD, line_color=LINE, line_w=1.0, radius=None,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, shadow=False):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line_color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line_color
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = shadow
    if sp.has_text_frame:
        sp.text_frame.word_wrap = True
    return sp


def bg(slide, color=CREAM):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def alpha_fill(sp, pct):
    """Set transparency on a solid-filled shape (pct = opacity percent)."""
    clr = sp.fill.fore_color._xFill
    srgb = clr.find(qn("a:srgbClr"))
    if srgb is None:
        return
    for old in srgb.findall(qn("a:alpha")):
        srgb.remove(old)
    a = etree.SubElement(srgb, qn("a:alpha"))
    a.set("val", str(int(pct * 1000)))


def star(slide, cx, cy, d, color=GOLD):
    sp = slide.shapes.add_shape(MSO_SHAPE.STAR_8_POINT, Inches(cx - d / 2), Inches(cy - d / 2),
                                Inches(d), Inches(d))
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def header(slide, kicker, title, num, dark=False):
    """Kicker + title + gold rule, right aligned (RTL)."""
    kcol = GOLD_L if dark else GOLD
    tcol = CREAM if dark else INK
    add_text(slide, MX, 0.46, CW, 0.32, kicker, size=11.5, color=kcol, bold=True)
    add_text(slide, MX, 0.80, CW, 0.78, title, size=28, color=tcol, bold=True, line=1.15)
    rect(slide, SW - MX - 1.55, 1.66, 1.55, 0.055, fill=GOLD, line_color=None,
         shape=MSO_SHAPE.RECTANGLE)


def footer(slide, num, dark=False):
    col = GOLD_L if dark else MUTED
    rect(slide, MX, 6.93, CW, 0.02, fill=LINE if not dark else RGBColor(0x1C, 0x4A, 0x44),
         line_color=None, shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, MX, 7.03, 4.0, 0.28, "اصول و ارتباط دولت اسلامی با دولت غیر اسلامی",
             size=9.5, color=col)
    add_text(slide, SW - MX - 1.6, 7.03, 1.6, 0.28, fa(num) + " / " + fa(NUM_TOTAL),
             size=9.5, color=col, bold=True, align=PP_ALIGN.LEFT)


def badge(slide, x, y, d, label, fill=EMERALD, color=CREAM, size=13):
    sp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    sp.shadow.inherit = False
    tf = sp.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = color
    r.font.name = FONT
    _cs(r)
    return sp


def card(slide, x, y, w, h, title, body, num=None, title_size=15.5, body_size=13,
         accent=GOLD, fill=CARD, title_color=EMERALD_DK, body_color=MUTED, line=1.5):
    rect(slide, x, y, w, h, fill=fill, radius=0.06)
    rect(slide, x + w - 0.075, y + 0.22, 0.055, h - 0.44, fill=accent, line_color=None,
         shape=MSO_SHAPE.RECTANGLE)
    tx = x + 0.34
    tw = w - 0.34 - 0.42
    top = y + 0.26
    if num is not None:
        badge(slide, x + w - 0.62, y + 0.24, 0.42, num, fill=EMERALD)
        tw = w - 0.34 - 0.86
    add_text(slide, tx, top, tw, 0.38, title, size=title_size, color=title_color,
             bold=True, line=1.15)
    add_text(slide, tx, top + 0.44, w - 0.34 - 0.40, max(0.24, h - 0.92), body,
             size=body_size, color=body_color, line=line)


def card_list(slide, x, y, w, h, title, items, num=None, size=12.5, gap=0.44,
              accent=GOLD, fill=CARD, title_color=EMERALD_DK, body_color=MUTED):
    """Card with a title and a list of diamond-bulleted items (bidi-safe)."""
    rect(slide, x, y, w, h, fill=fill, radius=0.06)
    rect(slide, x + w - 0.075, y + 0.22, 0.055, h - 0.44, fill=accent, line_color=None,
         shape=MSO_SHAPE.RECTANGLE)
    tw = w - 0.34 - 0.42
    if num is not None:
        badge(slide, x + w - 0.62, y + 0.24, 0.42, num, fill=EMERALD)
        tw = w - 0.34 - 0.86
    add_text(slide, x + 0.34, y + 0.26, tw, 0.38, title, size=15.5, color=title_color,
             bold=True, line=1.15)
    cy = y + 0.78
    for it in items:
        d = 0.095
        sp = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(x + w - 0.31),
                                    Inches(cy + 0.09), Inches(d), Inches(d))
        sp.fill.solid()
        sp.fill.fore_color.rgb = accent
        sp.line.fill.background()
        sp.shadow.inherit = False
        add_text(slide, x + 0.34, cy, w - 0.34 - 0.46, gap, it, size=size,
                 color=body_color, line=1.4)
        cy += gap
    return cy


def verse(slide, x, y, w, h, text, ref, size=15.5):
    rect(slide, x, y, w, h, fill=EMERALD, line_color=None, radius=0.05)
    star(slide, x + w - 0.45, y + 0.42, 0.32, color=GOLD)
    add_text(slide, x + 0.42, y + 0.36, w - 1.0, max(0.4, h - 0.86), text, size=size,
             color=CREAM, line=1.7, align=PP_ALIGN.RIGHT)
    add_text(slide, x + 0.42, y + h - 0.46, w - 0.84, 0.3, ref, size=11, color=GOLD_L,
             bold=True)


def bullets(slide, x, y, w, items, size=15, gap=0.52, color=INK, marker=GOLD):
    """Gold-diamond bulleted list, RTL: diamond hugs the right edge of each line."""
    cy = y
    for it in items:
        d = 0.105
        sp = slide.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(x + w - d), Inches(cy + 0.1),
                                    Inches(d), Inches(d))
        sp.fill.solid()
        sp.fill.fore_color.rgb = marker
        sp.line.fill.background()
        sp.shadow.inherit = False
        add_text(slide, x, cy, w - 0.26, gap, it, size=size, color=color, line=1.4)
        cy += gap
    return cy


def note(slide, x, y, w, h, text, size=12.5):
    rect(slide, x, y, w, h, fill=SAND, line_color=None, radius=0.1)
    add_text(slide, x + 0.3, y + 0.16, w - 0.6, h - 0.3, text, size=size, color=EMERALD_DK,
             line=1.5, anchor=MSO_ANCHOR.MIDDLE)
