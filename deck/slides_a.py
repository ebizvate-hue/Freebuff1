# -*- coding: utf-8 -*-
"""Slides 1-5: cover, roadmap, concepts, sources, principle of engagement."""
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from theme import (add_text, rect, bg, star, header, footer, card, verse, bullets, note,
                   badge, alpha_fill, CREAM, INK, EMERALD, EMERALD_DK, GOLD, GOLD_L, MUTED,
                   CARD, LINE, SAND, SW, SH, MX, CW, fa)

COVER = "assets/cover-pattern.jpg"


def s01_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, EMERALD_DK)
    s.shapes.add_picture(COVER, 0, 0, width=Inches(SW), height=Inches(SH))
    veil = rect(s, 0, 0, SW, SH, fill=EMERALD_DK, line_color=None,
                shape=MSO_SHAPE.RECTANGLE)
    alpha_fill(veil, 86)
    add_text(s, 4.9, 1.42, 7.15, 0.34, "فقه سیاسی • روابط بین‌الملل اسلامی", size=12,
             color=GOLD_L, bold=True, align=PP_ALIGN.RIGHT)
    add_text(s, 4.9, 1.92, 7.15, 2.5,
             "اصول و ارتباط\nدولت اسلامی\nبا دولت غیر اسلامی", size=40, color=CREAM,
             bold=True, line=1.32, align=PP_ALIGN.RIGHT)
    rect(s, 12.26, 1.5, 0.055, 3.95, fill=GOLD, line_color=None, shape=MSO_SHAPE.RECTANGLE)
    add_text(s, 5.4, 4.66, 6.65, 0.9,
             "مبانی، اصول حاکم، عرصه‌های عملی و چالش‌های معاصر", size=15, color=GOLD_L,
             align=PP_ALIGN.RIGHT, line=1.6)
    star(s, 5.1, 6.35, 0.42)
    add_text(s, 7.5, 6.45, 4.55, 0.32, "ارائه تحلیلی | ۱۵ اسلاید", size=10.5, color=MUTED,
             align=PP_ALIGN.RIGHT)


def _alpha(pct):
    from pptx.oxml import parse_xml
    return parse_xml(
        '<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="%d000"/>'
        % pct)


def s02_roadmap(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "نقشه راه", "مسیر بحث در پنج گام", 2)
    items = [
        ("۱", "مفاهیم و مبانی", "دولت اسلامی، دولت غیراسلامی و تقسیم‌بندی سه‌گانه سرزمین‌ها"),
        ("۲", "اصول حاکم بر ارتباط", "دعوت، صلح، وفای به عهد، عزت و نفی سبیل، کرامت و مصلحت"),
        ("۳", "عرصه‌های عملی", "دیپلماسی، پیمان‌ها، اقتصاد، فرهنگ و تبادل علمی"),
        ("۴", "حقوق و اخلاق", "اقلیت‌های دینی، آزادی عقیده و اخلاق جنگ در فقه"),
        ("۵", "چالش‌های معاصر", "دولت ـ ملت، حقوق بین‌الملل سکولار و جمع‌بندی اصول"),
    ]
    y = 2.05
    for n, t, d in items:
        rect(s, MX, y, CW, 0.78, fill=CARD, radius=0.12)
        rect(s, SW - MX - 0.06, y + 0.12, 0.05, 0.54, fill=GOLD, line_color=None,
             shape=MSO_SHAPE.RECTANGLE)
        badge(s, MX + CW - 0.64, y + 0.18, 0.42, n, fill=EMERALD)
        add_text(s, MX + CW - 3.85, y + 0.19, 3.0, 0.4, t, size=15.5, color=EMERALD_DK,
                 bold=True)
        add_text(s, MX + 0.35, y + 0.22, 7.4, 0.4, d, size=12.5, color=MUTED)
        y += 0.94
    footer(s, 2)


def s03_concepts(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "مفاهیم پایه", "دولت اسلامی و دولت غیر اسلامی", 3)
    w = (CW - 0.7) / 3
    card(s, MX + 2 * (w + 0.35), 2.0, w, 2.55, "دولت اسلامی",
         "جامعه سیاسی که در آن حاکمیت، قانون‌گذاری و جهت‌گیری سیاست داخلی و خارجی "
         "بر پایه شریعت اسلامی تنظیم می‌شود و مسلمانان در آن از امنیت و آزادی "
         "عبادت برخوردارند.", num="۱")
    card(s, MX + (w + 0.35), 2.0, w, 2.55, "دولت غیر اسلامی",
         "کشوری که حاکمیت اسلامی در آن نیست. فقه، رابطه با آن را بر اساس نوع "
         "رفتارش تعیین می‌کند: دولت معاهد، دولت حربی و دولت امان‌یافته.", num="۲")
    card(s, MX, 2.0, w, 2.55, "سه‌گانه فقهی",
         "دارالاسلام (قلمرو حاکمیت اسلامی)، دارالعهد یا دارالصلح (قلمرو پیمان) و "
         "دارالحرب (قلمروی بدون پیمان و درگیری). این تقسیم، توصیفی ـ تحلیلی است "
         "نه نژادی یا دینی.", num="۳")
    note(s, MX, 4.85, CW, 1.1,
         "نکته: ملاک رابطه در فقه، «رفتار سیاسی دولت‌ها» است، نه صرفِ تفاوت دین. "
         "از این‌رو یک دولت غیر اسلامی می‌تواند طرف صلح، پیمان و همکاری باشد و "
         "یک دولت مدعیِ اسلام نیز اگر پیمان‌شکنی کند، مشروعیت رابطه را از دست می‌دهد.")
    add_text(s, MX, 6.2, CW, 0.45,
             "دولت اسلامی در این چارچوب، پیش از آن‌که «قلمرو» باشد، «مسئولیت» است: "
             "مسئولیت اقامه عدل، حفظ کرامت انسان و پایبندی به تعهدات.",
             size=13, color=EMERALD, bold=True)
    footer(s, 3)


def s04_sources(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "مبانی", "چهار سرچشمه برای تنظیم رابطه با دولت غیر اسلامی", 4)
    w = (CW - 1.05) / 4
    data = [
        ("۱", "قرآن کریم", "آیات صلح، دعوت، عدالت و وفای به عهد شالوده اصول را می‌سازند؛ "
                            "مانند انفال ۶۱ و ممتحنه ۸."),
        ("۲", "سنت و سیره", "صلح حدیبیه، نامه‌های پیامبر (ص) به پادشاهان، پذیرش هیئت‌ها "
                            "و پناه‌دادن به مشرکان."),
        ("۳", "عقل و مصلحت", "تشخیص مصلحت نظام اسلامی در چارچوب احکام؛ ابزار تطبیق اصول "
                             "با شرایط زمان."),
        ("۴", "عهد و عرف", "پیمان‌ها، کنوانسیون‌ها و عرف‌های پذیرفته‌شده بین‌المللی، تا "
                          "حدی که مخالف شرع نباشد."),
    ]
    for i, (n, t, d) in enumerate(data):
        card(s, MX + (3 - i) * (w + 0.35), 2.0, w, 2.75, t, d, num=n, body_size=12.5)
    rect(s, MX, 5.05, CW, 1.5, fill=EMERALD, line_color=None, radius=0.05)
    add_text(s, MX + 0.55, 5.35, CW - 1.1, 0.42, "قاعده راهبردی", size=12, color=GOLD_L, bold=True)
    add_text(s, MX + 0.55, 5.78, CW - 1.1, 0.55,
             "اصول ثابت‌اند، روش‌ها متغیر؛ تغییر ابزار دیپلماسی هرگز مجوز عبور از "
             "«عزت»، «عدالت» و «وفای به عهد» نیست.", size=15, color=CREAM, bold=True)
    star(s, MX + 0.42, 5.8, 0.3, color=GOLD)
    footer(s, 4)


def s05_engagement(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "اصل یکم", "اصلِ ارتباط: تعامل، نه انزوا", 5)
    bullets(s, MX + 2 * (CW / 3) + 0.2, 2.05, CW / 3, [
        "انسان موجودی اجتماعی و ارتباط‌طلب است؛ اسلام آیین فطرت است.",
        "رسالت اسلامی جهانی است و پیام آن بدون ارتباط با دیگر ملت‌ها منتقل نمی‌شود.",
        "نیاز متقابل اقتصادی، علمی و امنیتی، همزیستی را ضروری می‌کند.",
        "انزوا، زمینه سوءتفاهم، تهدید و منازعه را افزایش می‌دهد.",
    ], size=14.5, gap=0.86)
    bullets(s, MX + (CW / 3) + 0.1, 2.05, CW / 3, [
        "ارتباط، ابزار «شناخت» است؛ «تعارف» در قرآن مقدمه همکاری است.",
        "پیامبر (ص) با حاکمان غیرمسلمان مکاتبه و پیمان بست.",
        "مسلمانان در پناه دولت‌های غیراسلامی (حبشه) نیز زیستند.",
        "اصل، ارتباط است؛ قطع رابطه، استثنای موقت و مصلحتی است.",
    ], size=14.5, gap=0.86)
    verse(s, MX, 2.05, CW / 3 - 0.1, 3.55,
          "یا أَیُّهَا النّاسُ إِنّا خَلَقْناکُمْ مِنْ ذَکَرٍ وَ أُنْثی وَ جَعَلْناکُمْ "
          "شُعُوباً وَ قَبائِلَ لِتَعارَفُوا",
          "حجرات، ۱۳", size=13.5)
    note(s, MX, 5.95, CW, 0.75,
         "جمع‌بندی: ارتباط با دولت غیر اسلامی «اصل» است و انزوا «استثنا»؛ "
         "استثنا تنها با دلیل شرعی و مصلحت روشن قابل دفاع است.")
    footer(s, 5)