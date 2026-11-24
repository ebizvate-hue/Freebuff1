# -*- coding: utf-8 -*-
"""Slides 11-15: political practice, economy/culture, minorities, war ethics, conclusion."""
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from theme import (add_text, rect, bg, star, header, footer, card, card_list, verse,
                   bullets, note, badge, alpha_fill, CREAM, INK, EMERALD, EMERALD_DK, GOLD,
                   GOLD_L, MUTED, CARD, LINE, SAND, SW, SH, MX, CW, fa)


def s11_politics(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "عرصه نخست", "ابزارهای سیاسی و دیپلماتیک", 11)
    tools = [
        ("سفارت و نمایندگی", "اعزام و پذیرش فرستاده با تضمین مصونیت"),
        ("امان و پناهندگی", "پذیرش پناه‌جو و تعهد به عدم استرداد"),
        ("معاهده و کنوانسیون", "پیمان صلح، همکاری و حل اختلاف"),
        ("میانجی‌گری و داوری", "حکمیت ثالث برای پایان منازعه"),
        ("سازمان‌های بین‌المللی", "عضویت فعال و همکاری چندجانبه"),
    ]
    y = 2.0
    for i, (t, d) in enumerate(tools):
        rect(s, MX + 4.35, y, CW - 4.35, 0.72, fill=CARD, radius=0.14)
        badge(s, MX + CW - 0.64, y + 0.15, 0.42, fa(i + 1), fill=EMERALD)
        add_text(s, MX + CW - 3.25, y + 0.18, 2.4, 0.4, t, size=13.5, color=EMERALD_DK,
                 bold=True)
        add_text(s, MX + 4.55, y + 0.21, 3.9, 0.4, d, size=11.5, color=MUTED)
        y += 0.86
    rect(s, MX, 2.0, 4.05, 4.05, fill=EMERALD, line_color=None, radius=0.05)
    add_text(s, MX + 0.4, 2.25, 3.25, 0.35, "الگوی تاریخی", size=11.5, color=GOLD_L, bold=True)
    add_text(s, MX + 0.4, 2.7, 3.25, 2.2,
             "نامه‌های پیامبر (ص) به هرقل، خسرو و نجاشی\n\n"
             "پذیرش هیئت‌های نمایندگی قبایل (وفود)\n\n"
             "پناه‌دادن به مشرکان و تعهد به بازگرداندن امانت",
             size=12.5, color=CREAM, line=1.65)
    star(s, MX + 0.55, 5.6, 0.3, color=GOLD)
    add_text(s, MX, 6.25, CW, 0.4,
             "قاعده: ابزار دیپلماتیک، وسیله است نه هدف؛ مشروعیت آن به انطباق با اصول بسته است.",
             size=12.5, color=EMERALD, bold=True)
    footer(s, 11)


def s12_economy(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "عرصه دوم", "اقتصاد و فرهنگ: میدان‌های پیوند روزمره", 12)
    w = (CW - 0.4) / 2
    card_list(s, MX + w + 0.4, 2.0, w, 3.1, "روابط اقتصادی", [
        "تجارت و امنیت راه‌های بازرگانی (سابقه ایلاف قریش)",
        "جزیه و خراج به‌عنوان نظم مالی، نه ابزار تحقیر",
        "پایبندی به قرارداد، وزن و پیمانه در معامله",
        "ممنوعیت ربا و پرهیز از وابستگی سلطه‌آور",
    ], size=13, gap=0.66)
    card_list(s, MX, 2.0, w, 3.1, "روابط فرهنگی و علمی", [
        "گفت‌وگوی علمی و ترجمه میراث ملت‌ها (بیت‌الحکمه)",
        "تبادل پزشکی، فنی و آموزشی به‌عنوان قدرت نرم",
        "حفاظت از اماکن مقدس طرف مقابل در جنگ و صلح",
        "پرهیز از توهین به مقدسات دیگران",
    ], size=13, gap=0.66)
    note(s, MX, 5.45, CW, 1.15,
         "منطق حاکم: هر پیوندی که «عزت» را حفظ کند و «استقلال» را تهدید نکند، "
         "مشروع و حتی مطلوب است. اقتصاد و فرهنگ، ابزار «قدرت نرم» دولت اسلامی در "
         "عرصه بین‌الملل به شمار می‌روند؛ همان‌جا که جنگ سخت، آخرین گزینه است.")
    footer(s, 12)


def s13_minorities(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "عرصه سوم", "اقلیت‌های دینی و آزادی عقیده", 13)
    bullets(s, MX + CW / 2 + 0.2, 2.0, CW / 2 - 0.2, [
        "اهل کتاب می‌توانند با عقد ذمه در پناه دولت اسلامی زیست کنند.",
        "امنیت جان، مال و ناموس آنان تضمین می‌شود و عبادتگاه‌هایشان محترم است.",
        "در امور خانوادگی و احوال شخصیه، به آیین خود عمل می‌کنند.",
        "از تکالیف ویژه مسلمانان (جهاد و زکات) معاف و در برابر، جزیه می‌پردازند.",
        "سند مدینه نمونه‌ای از پیمان چندفرهنگی با یهودیان و مشرکان مدینه است.",
    ], size=14, gap=0.8)
    verse(s, MX, 2.0, CW / 2 - 0.35, 1.5,
          "لَکُمْ دینُکُمْ وَ لِیَ دینِ", "کافرون، ۶", size=17)
    note(s, MX, 3.75, CW / 2 - 0.35, 1.6,
         "«لا إِکْراهَ فِی الدِّینِ»: آزادی پذیرش دین، پایه رابطه است. "
         "اجبار در اعتقاد، از نظر قرآن بی‌معناست و با ماهیت دعوت ناسازگار است. "
         "سیاست خارجی نیز نمی‌تواند «تغییر باور ملت‌ها» را هدف نظامی خود کند.",
         size=12)
    add_text(s, MX, 6.15, CW, 0.5,
             "پیام: رفتار با اقلیت‌ها، آینه اعتبار بین‌المللی دولت اسلامی است؛ "
             "نقض حقوق آنان به «وهن دین» می‌انجامد.", size=13, color=EMERALD, bold=True)
    footer(s, 13)


def s14_war(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s)
    header(s, "استثنا", "جنگ دفاعی و اخلاق نبرد", 14)
    rect(s, MX, 1.95, CW, 0.9, fill=SAND, line_color=None, radius=0.08)
    add_text(s, MX + 0.4, 2.12, CW - 0.8, 0.35,
             "«وَ قاتِلُوا فِی سَبیلِ اللّهِ الَّذینَ یُقاتِلُونَکُمْ وَ لا تَعْتَدُوا» (بقره، ۱۹۰)",
             size=14, color=EMERALD_DK, bold=True)
    w = (CW - 0.4) / 2
    card_list(s, MX + w + 0.4, 2.95, w, 2.9, "شرایط جواز", [
        "تجاوز نظامی یا تهدید قطعی به امنیت و موجودیت",
        "فتنه و آشوب علیه مسلمانان و نقض آزادی دعوت",
        "پیمان‌شکنی طرف مقابل پس از هشدار و اعلام",
        "فقدان راه‌حل مسالمت‌آمیز و حکمیت",
        "تناسب در پاسخ و هدفِ رفع تجاوز، نه کشورگشایی",
    ], size=13, gap=0.42, accent=EMERALD)
    card_list(s, MX, 2.95, w, 2.9, "خطوط سرخ", [
        "کشتن زنان، کودکان، پیران و بی‌گناهان",
        "مثله‌کردن، شکنجه و کشتار اسیران",
        "قطع درختان و نابودی مزارع و منابع زیستی",
        "پیمان‌شکنی پنهانی و خیانت در امان",
        "آزار کسانی که در جنگ مشارکت ندارند",
    ], size=13, gap=0.42, accent=GOLD)
    note(s, MX, 6.2, CW, 0.6,
         "نتیجه: در فقه اسلامی، جنگ «ابزار سیاست» نیست؛ «دفاع» است، و در دفاع نیز "
         "اخلاق، قید الزام‌آور است.", size=12.5)
    footer(s, 14)


def s15_conclusion(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, EMERALD_DK)
    s.shapes.add_picture("assets/cover-pattern.jpg", 0, 0, width=Inches(SW), height=Inches(SH))
    veil = rect(s, 0, 0, SW, SH, fill=EMERALD_DK, line_color=None, shape=MSO_SHAPE.RECTANGLE)
    alpha_fill(veil, 88)
    header(s, "جمع‌بندی", "چالش‌های معاصر و چهار ستون ثابت", 15, dark=True)
    chal = [
        ("۱", "دولت ـ ملت یا امت", "تنش میان مرزهای ملی و همبستگی فراملی مسلمانان"),
        ("۲", "حقوق بین‌الملل سکولار", "نیاز به خوانش فقهی از معاهدات و نهادهای جهانی"),
        ("۳", "تحریم و فشار اقتصادی", "توازن میان تعامل و پرهیز از سلطه"),
        ("۴", "رسانه و جنگ نرم", "امنیت فرهنگی و مرزبانی از هویت دینی"),
    ]
    w = (CW - 1.05) / 4
    for i, (n, t, d) in enumerate(chal):
        x = MX + (3 - i) * (w + 0.35)
        rect(s, x, 2.0, w, 1.85, fill=EMERALD, line_color=RGBColor_hex(), radius=0.1)
        badge(s, x + w - 0.55, 2.12, 0.4, n, fill=GOLD, color=EMERALD_DK, size=12)
        add_text(s, x + 0.25, 2.6, w - 0.6, 0.4, t, size=13, color=CREAM, bold=True)
        add_text(s, x + 0.25, 3.0, w - 0.6, 0.75, d, size=11, color=GOLD_L, line=1.5)
    add_text(s, MX, 4.05, CW, 0.4, "چهار ستون ثابت رابطه با دولت غیر اسلامی",
             size=12.5, color=GOLD_L, bold=True)
    pillars = [("دعوت", "تبلیغ با حکمت و اقناع"), ("صلح", "همزیستی به‌عنوان اصل"),
               ("عهد", "پایبندی بی‌قید‌و‌شرط به پیمان"), ("عزت", "استقلال و نفی سلطه")]
    pw = (CW - 0.9) / 4
    for i, (t, d) in enumerate(pillars):
        x = MX + (3 - i) * (pw + 0.3)
        rect(s, x, 4.5, pw, 1.2, fill=CREAM, line_color=None, radius=0.08)
        add_text(s, x + 0.2, 4.62, pw - 0.4, 0.4, t, size=17, color=EMERALD_DK, bold=True)
        add_text(s, x + 0.2, 5.1, pw - 0.4, 0.5, d, size=10.5, color=MUTED)
    add_text(s, MX, 6.0, CW, 0.75,
             "دولت اسلامی با دولت غیر اسلامی، بر پایه «تعامل عزتمندانه» رابطه می‌سازد: "
             "دعوت بی‌اجبار، صلح بی‌سازش، پیمان بی‌خیانت و تعامل بی‌تبعیت.",
             size=14.5, color=CREAM, bold=True, line=1.6, align=PP_ALIGN.CENTER)
    footer(s, 15, dark=True)


def RGBColor_hex():
    from pptx.dml.color import RGBColor
    return RGBColor(0x2A, 0x5D, 0x55)


def _alpha(pct):
    from pptx.oxml import parse_xml
    return parse_xml(
        '<a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="%d000"/>'
        % pct)
