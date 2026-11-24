# -*- coding: utf-8 -*-
"""Report text blocks whose estimated rendered height exceeds their box."""
import sys
from pptx import Presentation
from pptx.util import Emu
from PIL import ImageFont

FONT_PATH = sys.argv[2] if len(sys.argv) > 2 else "assets/Vazirmatn-Regular.ttf"
PRS = Presentation(sys.argv[1])


def wrap_lines(text, size_pt, width_in):
    font = ImageFont.truetype(FONT_PATH, max(8, int(size_pt * 1.333)))
    maxw = width_in * 96 * 0.96
    lines = 0
    for para in text.split("\n"):
        if not para.strip():
            lines += 1
            continue
        cur = ""
        n = 1
        for word in para.split(" "):
            trial = (cur + " " + word).strip()
            if font.getlength(trial) <= maxw:
                cur = trial
            else:
                n += 1
                cur = word
        lines += n
    return lines


bad = 0
for idx, slide in enumerate(PRS.slides, 1):
    for sh in slide.shapes:
        if not sh.has_text_frame or not sh.text_frame.text.strip():
            continue
        text = sh.text_frame.text
        size = 15
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size:
                    size = r.font.size.pt
                    break
            break
        width_in = Emu(sh.width).inches
        height_in = Emu(sh.height).inches
        lines = wrap_lines(text, size, width_in)
        spacing = 1.5
        for p in sh.text_frame.paragraphs:
            if p.line_spacing:
                spacing = float(p.line_spacing)
                break
        need = lines * size * spacing / 72.0
        if need > height_in + 0.02:
            bad += 1
            print("OVERFLOW slide %d: need %.2fin have %.2fin (%d lines) :: %s"
                  % (idx, need, height_in, lines, text[:58].replace("\n", " ")))
print("blocks over capacity:", bad)
