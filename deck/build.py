# -*- coding: utf-8 -*-
import os
import sys
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import slides_a as A
import slides_b as B
import slides_c as C

OUT = "دولت-اسلامی-و-دولت-غیر-اسلامی.pptx"


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    for fn in (A.s01_cover, A.s02_roadmap, A.s03_concepts, A.s04_sources, A.s05_engagement,
               B.s06_dawah, B.s07_peace, B.s08_treaty, B.s09_dignity, B.s10_justice,
               C.s11_politics, C.s12_economy, C.s13_minorities, C.s14_war, C.s15_conclusion):
        fn(prs)
    prs.save(OUT)
    print("slides:", len(prs.slides.__iter__.__self__._sldIdLst), "->", OUT)


if __name__ == "__main__":
    main()
