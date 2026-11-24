# -*- coding: utf-8 -*-
"""Approximate PNG renderer for visual inspection of the generated deck."""
import os
import sys
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

S = 96  # px per inch
REG = "assets/Vazirmatn-Regular.ttf"
BOLD = "assets/Vazirmatn-Bold.ttf"
_cache = {}


def font(px, bold):
    key = (px, bold)
    if key not in _cache:
        _cache[key] = ImageFont.truetype(BOLD if bold else REG, px)
    return _cache[key]


def shape_text(el):
    return "".join(t.text or "" for t in el.iter() if t.tag.endswith("}t"))


def color_of(fill_el):
    from pptx.dml.color import RGBColor
    if fill_el is None:
        return None, 1.0
    alpha = 1.0
    a = fill_el.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
    if a is not None:
        alpha = int(a.get("val")) / 100000.0
    clr = fill_el.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
    if clr is None:
        return None, alpha
    v = clr.get("val")
    return (int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16)), alpha


def walk(shapes, img, draw, ox=0, oy=0):
    for sh in shapes:
        x = int(Emu(sh.left).inches * S) + ox
        y = int(Emu(sh.top).inches * S) + oy
        w = int(Emu(sh.width).inches * S)
        h = int(Emu(sh.height).inches * S)
        if sh.shape_type == 13 or sh.__class__.__name__ == "Picture":
            try:
                pic = Image.open(sh.image.blob and __import__("io").BytesIO(sh.image.blob))
                pic = pic.convert("RGB").resize((max(1, w), max(1, h)))
                img.paste(pic, (x, y))
            except Exception:
                pass
            continue
        fill_rgb, alpha = None, 1.0
        try:
            if sh.fill.type is not None and sh.fill.type == 1:
                fill_rgb = tuple(sh.fill.fore_color.rgb)
            if sh.fill.type == 5:
                pass
            if sh.fill.type is not None and sh.fill.type == 1:
                pass
        except Exception:
            fill_rgb = None
        try:
            spPr = sh._element.spPr
            solid = spPr.find(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
            if solid is not None:
                fill_rgb, alpha = color_of(solid)
        except Exception:
            pass
        line_rgb, lw = None, 1
        try:
            if sh.line.color and sh.line.color.rgb is not None:
                line_rgb = tuple(sh.line.color.rgb)
                lw = max(1, int((sh.line.width.pt if sh.line.width else 1) * S / 72))
        except Exception:
            line_rgb = None
        if fill_rgb or line_rgb:
            layer = Image.new("RGBA", (max(1, w), max(1, h)), (0, 0, 0, 0))
            ld = ImageDraw.Draw(layer)
            nc = sh.name
            col = (fill_rgb or (255, 255, 255)) + (int(255 * alpha),)
            ol = (line_rgb + (255,)) if line_rgb else None
            if "Oval" in nc or "oval" in nc:
                ld.ellipse([0, 0, w - 1, h - 1], fill=col, outline=ol, width=lw)
            elif "Star" in nc:
                import math
                cx, cy, R = w / 2, h / 2, min(w, h) / 2
                pts = []
                for i in range(16):
                    r = R if i % 2 == 0 else R * 0.42
                    a = math.pi * i / 8 - math.pi / 2
                    pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
                ld.polygon(pts, fill=col)
            elif "Rounded" in nc:
                ld.rounded_rectangle([0, 0, w - 1, h - 1], radius=min(w, h) * 0.12,
                                     fill=col, outline=ol, width=lw)
            else:
                ld.rectangle([0, 0, w - 1, h - 1], fill=col, outline=ol, width=lw)
            img.paste(layer, (x, y), layer)
        if sh.has_text_frame and sh.text_frame.text.strip():
            draw_text(sh, draw, x, y, w, h)
        if sh.shape_type == 6:  # group
            walk(sh.shapes, img, draw)


def draw_text(sh, draw, x, y, w, h):
    tf = sh.text_frame
    blocks = []
    total = 0
    for p in tf.paragraphs:
        if not p.runs:
            continue
        r0 = p.runs[0]
        size = r0.font.size.pt if r0.font.size else 15
        bold = bool(r0.font.bold)
        rgb = tuple(r0.font.color.rgb) if r0.font.color and r0.font.color.rgb else (0, 0, 0)
        txt = "".join(r.text for r in p.runs)
        spacing = float(p.line_spacing) if p.line_spacing else 1.4
        align = p.alignment
        lines = wrap(txt, size, w)
        lh = size * 1.2 * spacing * S / 72
        blocks.append((lines, size, bold, rgb, align, lh))
        total += lh * len(lines) + ((p.space_after.pt * S / 72) if p.space_after else 0)
    cy = y
    if tf.vertical_anchor == MSO_ANCHOR.MIDDLE:
        cy = y + max(0, (h - total) / 2)
    for lines, size, bold, rgb, align, lh in blocks:
        f = font(max(8, int(round(size * S / 72))), bold)
        for ln in lines:
            disp = ln
            tw = draw.textlength(disp, font=f, direction="rtl", language="fa")
            if align == PP_ALIGN.CENTER:
                tx = x + (w - tw) / 2
            elif align == PP_ALIGN.LEFT:
                tx = x
            else:
                tx = x + w - tw
            draw.text((tx, cy + (lh - size * S / 72) / 2), disp, font=f, fill=rgb,
                      direction="rtl", language="fa")
            cy += lh


def wrap(txt, size, width_px):
    f = font(max(8, int(round(size * S / 72))), False)
    out = []
    for para in txt.split("\n"):
        if not para.strip():
            out.append("")
            continue
        cur = ""
        for word in para.split(" "):
            trial = (cur + " " + word).strip()
            if f.getlength(trial, direction="rtl", language="fa") <= width_px * 0.99:
                cur = trial
            else:
                if cur:
                    out.append(cur)
                cur = word
        out.append(cur)
    return out


def main():
    prs = Presentation(sys.argv[1])
    outdir = sys.argv[2] if len(sys.argv) > 2 else "preview"
    os.makedirs(outdir, exist_ok=True)
    W = int(prs.slide_width.inches * S)
    H = int(prs.slide_height.inches * S)
    for i, slide in enumerate(prs.slides, 1):
        img = Image.new("RGB", (W, H), (250, 247, 240))
        try:
            bgc, _ = color_of(slide.background.fill._xPr.find(
                ".//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr").getparent())
            if bgc:
                img = Image.new("RGB", (W, H), bgc)
        except Exception:
            pass
        draw = ImageDraw.Draw(img)
        walk(slide.shapes, img, draw)
        img.save(os.path.join(outdir, "slide-%02d.png" % i))
    print("rendered", len(prs.slides.__iter__.__self__._sldIdLst), "slides ->", outdir)


if __name__ == "__main__":
    main()
