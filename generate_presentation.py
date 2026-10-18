#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Taqdimot generatori: "Fleshkalar (USB) qanday qilib ma'lumot saqlaydi?"

Ushbu skript tashqi kutubxonalarsiz (faqat Python standart kutubxonasi)
to'liq yaroqli .pptx (OpenXML / OOXML) faylini yaratadi.

Sabab: sandbox muhitida PyPI (python-pptx) ga ulanish cheklangan (403 Forbidden).
.pptx fayli aslida XML qismlardan iborat ZIP arxividir, shu bois uni
standart `zipfile` moduli yordamida qo'lda, lekin to'liq spetsifikatsiyaga
mos ravishda yaratamiz.

Talablar:
  * Qat'iy 10 ta slayd.
  * Faqat "Times New Roman" shrifti.
  * Eng kichik shrift o'lchami 12 pt (sarlavha va muhim nuqtalar kattaroq).
  * 4 qismli tuzilma (3 asosiy reja + 1 xulosa).
  * Har bir slaydda Unsplash rasm URL manzillari matn ko'rinishida mavjud.
  * Matn sof, akademik o'zbek (lotin) tilida.
"""

import os
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# EMU (English Metric Units) hisoblari. 1 pt = 12700 EMU. Slayd 16:9.
# ---------------------------------------------------------------------------
EMU_PER_PT = 12700
SLIDE_W = 12192000  # 13.333 inch (16:9)
SLIDE_H = 6858000   # 7.5 inch

# Mavzuga mos ranglar palitrasi (raqamli texnologiya / mikroelektronika)
C_BG_DARK = "0B1E3B"      # chuqur ko'k fon
C_PANEL = "12294A"        # panel ko'ki
C_ACCENT = "27E0C8"       # neon turkuaz urg'u
C_ACCENT2 = "5AA9FF"      # yorqin ko'k
C_TEXT = "EAF2FF"         # oqimtir matn
C_MUTED = "9FB8D8"        # kulrang-ko'k ikkilamchi matn
C_TITLE_BG = "061329"     # titul foni

FONT = "Times New Roman"


# ---------------------------------------------------------------------------
# Matnli "run" va "paragraph" yaratuvchi yordamchi funksiyalar
# ---------------------------------------------------------------------------
def run(text, size_pt, color=C_TEXT, bold=False, italic=False):
    """Bitta matn 'run' obyektini qaytaradi (dict)."""
    return {
        "text": text,
        "size": size_pt,
        "color": color,
        "bold": bold,
        "italic": italic,
    }


def para(runs, align="l", bullet=False, space_before=600, line=110000):
    """Paragraf obyekti. `runs` - run dict'lari ro'yxati."""
    if isinstance(runs, dict):
        runs = [runs]
    return {
        "runs": runs,
        "align": align,
        "bullet": bullet,
        "space_before": space_before,
        "line": line,
    }


def run_xml(r):
    sz = int(round(r["size"])) * 100  # pt -> OOXML yuzdan birlar
    b = ' b="1"' if r["bold"] else ' b="0"'
    i = ' i="1"' if r["italic"] else ' i="0"'
    return (
        f'<a:r><a:rPr lang="uz-UZ" sz="{sz}"{b}{i} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{r["color"]}"/></a:solidFill>'
        f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/>'
        f'</a:rPr><a:t>{escape(r["text"])}</a:t></a:r>'
    )


def para_xml(p):
    buchar = ""
    if p["bullet"]:
        buchar = (
            f'<a:buFont typeface="{FONT}"/>'
            f'<a:buChar char="&#8226;"/>'
        )
    else:
        buchar = '<a:buNone/>'
    marL = ' marL="285750" indent="-285750"' if p["bullet"] else ''
    ppr = (
        f'<a:pPr{marL} algn="{p["align"]}">' 
        f'<a:lnSpc><a:spcPct val="{p["line"]}"/></a:lnSpc>'
        f'<a:spcBef><a:spcPts val="{p["space_before"]}"/></a:spcBef>'
        f'{buchar}</a:pPr>'
    )
    runs = "".join(run_xml(r) for r in p["runs"])
    if not p["runs"]:
        runs = (
            f'<a:endParaRPr lang="uz-UZ"><a:latin typeface="{FONT}"/>'
            f'<a:cs typeface="{FONT}"/></a:endParaRPr>'
        )
    return f'<a:p>{ppr}{runs}</a:p>'


def textbox(sp_id, name, x, y, cx, cy, paragraphs, fill=None, anchor="t"):
    """Bitta matnli shakl (text box) uchun <p:sp> XML qaytaradi."""
    body = "".join(para_xml(p) for p in paragraphs)
    fill_xml = ""
    if fill:
        fill_xml = (
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        )
    else:
        fill_xml = '<a:noFill/>'
    return (
        f'<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'{fill_xml}'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" '
        f'bIns="45720" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
        f'<a:lstStyle/>'
        f'{body}'
        f'</p:txBody></p:sp>'
    )


def rect(sp_id, name, x, y, cx, cy, color):
    """Dekorativ rang bo'lagi (fon panellari / urg'u chiziqlari)."""
    return (
        f'<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )


def slide_xml(shapes, bg=C_BG_DARK):
    sp_tree = "".join(shapes)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld>'
        f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{sp_tree}'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:overrideClrMapping '
        'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" '
        'accent5="accent5" accent6="accent6" hlink="hlink" '
        'folHlink="folHlink"/></p:clrMapOvr>'
        '</p:sld>'
    )


# Umumiy joylashuv o'lchamlari (EMU)
MARGIN_X = 686000
TITLE_Y = 360000
CONTENT_Y = 1500000
CONTENT_W = SLIDE_W - 2 * MARGIN_X


def header_shapes(next_id, title_text, part_label=None):
    """Har bir kontent slaydi uchun yuqori sarlavha bloki + urg'u chiziq."""
    shapes = []
    # Chap urg'u panel
    shapes.append(rect(next_id, "accentbar", MARGIN_X, TITLE_Y + 40000,
                       70000, 760000, C_ACCENT))
    next_id += 1
    paras = []
    if part_label:
        paras.append(para(run(part_label, 14, C_ACCENT, bold=True),
                          align="l", space_before=0))
    paras.append(para(run(title_text, 30, C_TEXT, bold=True),
                      align="l", space_before=200 if part_label else 0))
    shapes.append(textbox(next_id, "title", MARGIN_X + 170000, TITLE_Y,
                          CONTENT_W - 170000, 900000, paras, anchor="t"))
    next_id += 1
    return shapes, next_id


def caption_url(next_id, url):
    """Slayd pastidagi Unsplash URL manzili (ko'rinadigan izoh)."""
    paras = [
        para(run("Vizual manba (Unsplash): ", 12, C_ACCENT2, bold=True,
                 italic=True), align="l", space_before=0),
        para(run(url, 12, C_MUTED, italic=True), align="l", space_before=100),
    ]
    sp = textbox(next_id, "url_caption", MARGIN_X, SLIDE_H - 760000,
                 CONTENT_W, 640000, paras, fill=C_PANEL, anchor="ctr")
    return sp, next_id + 1


# ===========================================================================
# SLAYDLAR KONTENTI
# ===========================================================================
UNSPLASH = {
    1: "https://unsplash.com/photos/a-close-up-of-a-usb-flash-drive-on-a-table-9SoCnyQmkzI",
    2: "https://unsplash.com/photos/black-and-silver-usb-flash-drive-ZV_64LdGoao",
    3: "https://unsplash.com/photos/macro-shot-of-a-circuit-board-FO7JIlwjOtU",
    4: "https://unsplash.com/photos/green-computer-circuit-board-wtpVH-C58FA",
    5: "https://unsplash.com/photos/microchip-close-up-electronics-ZO02-30bFfs",
    6: "https://unsplash.com/photos/silicon-wafer-semiconductor-5fNmWej4tAA",
    7: "https://unsplash.com/photos/binary-code-digital-data-stream-xkBaqlcqeb4",
    8: "https://unsplash.com/photos/colorful-photo-gallery-thumbnails-on-screen-2JEWnQ6y1Ac",
    9: "https://unsplash.com/photos/modern-data-storage-technology-OqtafYT5kTw",
    10: "https://unsplash.com/photos/thank-you-note-on-digital-display-5QgIuuBxKwM",
}


def build_slides():
    slides = []

    # ------------------------------------------------------------------ 1
    # Titul varag'i
    sid = 2
    shapes = []
    shapes.append(rect(sid, "panel", 0, 0, SLIDE_W, SLIDE_H, C_TITLE_BG)); sid += 1
    shapes.append(rect(sid, "topbar", 0, 0, SLIDE_W, 90000, C_ACCENT)); sid += 1
    shapes.append(rect(sid, "botbar", 0, SLIDE_H - 90000, SLIDE_W, 90000, C_ACCENT2)); sid += 1

    inst = [
        para(run("Jizzax davlat pedagogika universiteti", 20, C_ACCENT2, bold=True),
             align="ctr", space_before=0),
        para(run("Axborot tizimlari fani", 16, C_MUTED, italic=True),
             align="ctr", space_before=200),
    ]
    shapes.append(textbox(sid, "univ", MARGIN_X, 520000, CONTENT_W, 900000,
                          inst, anchor="t")); sid += 1

    title = [
        para(run("Fleshkalar (USB) qanday qilib ma'lumot saqlaydi?", 36,
                 C_TEXT, bold=True), align="ctr", space_before=0),
        para(run("Harakatlanuvchi qismlarsiz minglab rasmlarning elektr "
                 "zaryadlari ko'rinishida saqlanishi", 18, C_ACCENT,
                 italic=True), align="ctr", space_before=400, line=120000),
    ]
    shapes.append(textbox(sid, "maintitle", MARGIN_X, 2300000, CONTENT_W,
                          2000000, title, anchor="ctr")); sid += 1

    presenter = [
        para(run("Taqdimotchi: Toshboʻriyeva Marjona", 18, C_TEXT, bold=True),
             align="ctr", space_before=0),
        para(run("242_26 - guruh", 16, C_MUTED), align="ctr", space_before=150),
    ]
    shapes.append(textbox(sid, "presenter", MARGIN_X, 4700000, CONTENT_W,
                          900000, presenter, anchor="ctr")); sid += 1

    cap, sid = caption_url(sid, UNSPLASH[1])
    shapes.append(cap)
    slides.append(slide_xml(shapes, bg=C_TITLE_BG))

    # ------------------------------------------------------------------ 2
    # Reja
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Taqdimot rejasi")
    shapes += hdr
    plan = [
        para([run("1-qism.  ", 20, C_ACCENT, bold=True),
              run("Fleshkalarning ishlash prinsipi va NAND xotira",
                  20, C_TEXT)], bullet=False, space_before=500),
        para([run("2-qism.  ", 20, C_ACCENT, bold=True),
              run("Ma'lumotlarni elektr zaryadi ko'rinishida saqlash",
                  20, C_TEXT)], bullet=False, space_before=500),
        para([run("3-qism.  ", 20, C_ACCENT, bold=True),
              run("Yozish, o'qish va o'chirish jarayonlari",
                  20, C_TEXT)], bullet=False, space_before=500),
        para([run("Xulosa.  ", 20, C_ACCENT2, bold=True),
              run("Fleshkalarning afzalliklari va texnologiya kelajagi",
                  20, C_TEXT)], bullet=False, space_before=500),
    ]
    shapes.append(textbox(sid, "plan", MARGIN_X, CONTENT_Y, CONTENT_W,
                          3200000, plan, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[2])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 3
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Fleshkaning ishlash prinsipi",
                             "1-QISM — NAND XOTIRA")
    shapes += hdr
    body = [
        para([run("Qattiq jism xotirasi (Solid-State). ", 18, C_ACCENT, bold=True),
              run("Fleshkalar ma'lumotni butunlay elektron tarzda, yarimo'tkazgichli "
                  "mikrosxemalarda saqlaydi.", 16, C_TEXT)], bullet=True),
        para([run("Harakatlanuvchi qismlar yo'q. ", 18, C_ACCENT, bold=True),
              run("Qattiq disklardan farqli ravishda fleshkalarda aylanuvchi plastinka "
                  "yoki o'qiydigan boshcha mavjud emas.", 16, C_TEXT)], bullet=True),
        para([run("NAND tuzilmasi. ", 18, C_ACCENT, bold=True),
              run("Xotira katakchalari NAND mantiqiy sxemasi asosida zich qatorlarga "
                  "birlashtirilib, kichik hajmda katta sig'imni ta'minlaydi.", 16, C_TEXT)],
             bullet=True),
        para([run("Energiyasiz saqlash. ", 18, C_ACCENT, bold=True),
              run("Ma'lumot elektr quvvati uzilgandan keyin ham saqlanib qoladi — bu "
                  "energiyaga bog'liq bo'lmagan (non-volatile) xotira.", 16, C_TEXT)],
             bullet=True),
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, CONTENT_W, 3400000,
                          body, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[3])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 4
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "NAND xotira katakchalari tuzilishi",
                             "1-QISM — NAND XOTIRA")
    shapes += hdr
    body = [
        para([run("Xotira katakchasi (cell). ", 18, C_ACCENT, bold=True),
              run("Ma'lumotning eng kichik birligi maxsus tranzistorda saqlanadi; "
                  "har bir katakcha bir yoki bir nechta bitni ifodalaydi.", 16, C_TEXT)],
             bullet=True),
        para([run("Qatorlar va ustunlar. ", 18, C_ACCENT, bold=True),
              run("Katakchalar o'zaro ulanib, so'z chiziqlari (word line) va bit "
                  "chiziqlari (bit line) to'rini hosil qiladi.", 16, C_TEXT)], bullet=True),
        para([run("Zichlik va sig'im. ", 18, C_ACCENT, bold=True),
              run("3D NAND texnologiyasida katakchalar vertikal qatlamlab joylashtirilib, "
                  "sig'im sezilarli oshiriladi.", 16, C_TEXT)], bullet=True),
        para([run("Ishonchlilik. ", 18, C_ACCENT, bold=True),
              run("Xatolarni tuzatuvchi kodlar (ECC) va kiyilishni taqsimlash (wear "
                  "leveling) ma'lumot butunligini saqlaydi.", 16, C_TEXT)], bullet=True),
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, CONTENT_W, 3400000,
                          body, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[4])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 5
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Suzuvchi zulfin (Floating Gate) tranzistori",
                             "2-QISM — ZARYAD KO'RINISHIDA SAQLASH")
    shapes += hdr
    body = [
        para([run("Suzuvchi zulfin. ", 18, C_ACCENT, bold=True),
              run("Har bir katakchada ikki qavatli zulfinli tranzistor bo'lib, ulardan "
                  "biri izolyator bilan to'liq o'ralgan suzuvchi zulfindir.", 16, C_TEXT)],
             bullet=True),
        para([run("Elektronlarni qamab qo'yish. ", 18, C_ACCENT, bold=True),
              run("Elektronlar suzuvchi zulfinga kiritilib, oksid qatlam ularning "
                  "chiqib ketishiga yo'l qo'ymaydi — zaryad shu yerda qamalib qoladi.",
                  16, C_TEXT)], bullet=True),
        para([run("Doimiy saqlash fizikasi. ", 18, C_ACCENT, bold=True),
              run("Izolyatsiya tufayli zaryad quvvat uzilsa ham yillar davomida saqlanadi, "
                  "bu esa ma'lumotning doimiyligini ta'minlaydi.", 16, C_TEXT)], bullet=True),
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, CONTENT_W, 3400000,
                          body, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[5])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 6
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Zaryad qanday qilib ma'lumotga aylanadi",
                             "2-QISM — ZARYAD KO'RINISHIDA SAQLASH")
    shapes += hdr
    body = [
        para([run("Zaryad miqdori — holat. ", 18, C_ACCENT, bold=True),
              run("Suzuvchi zulfindagi elektronlar mavjudligi yoki yo'qligi katakchaning "
                  "mantiqiy holatini belgilaydi.", 16, C_TEXT)], bullet=True),
        para([run("Chegaraviy kuchlanish. ", 18, C_ACCENT, bold=True),
              run("Qamalgan zaryad tranzistorning ochilish kuchlanishini o'zgartiradi; "
                  "aynan shu farq o'qish paytida aniqlanadi.", 16, C_TEXT)], bullet=True),
        para([run("Ko'p bitli katakchalar. ", 18, C_ACCENT, bold=True),
              run("MLC, TLC va QLC texnologiyalarida bitta katakcha zaryadning bir necha "
                  "darajasini saqlab, 2, 3 yoki 4 bitni ifodalaydi.", 16, C_TEXT)], bullet=True),
        para([run("Barqarorlik. ", 18, C_ACCENT, bold=True),
              run("Zaryad darajalari aniq o'lchanadigan oraliqlarga bo'linib, ma'lumot "
                  "ishonchli tiklanadi.", 16, C_TEXT)], bullet=True),
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, CONTENT_W, 3400000,
                          body, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[6])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 7
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Yozish, o'qish va o'chirish jarayonlari",
                             "3-QISM — JARAYONLAR")
    shapes += hdr
    body = [
        para([run("Yozish. ", 18, C_ACCENT, bold=True),
              run("Yuqori kuchlanish ta'sirida elektronlar oksid qatlamdan o'tib "
                  "(tunnel effekti) suzuvchi zulfinga joylashtiriladi.", 16, C_TEXT)],
             bullet=True),
        para([run("O'qish. ", 18, C_ACCENT, bold=True),
              run("Katakchaga kichik kuchlanish berilib, tok o'tishi tekshiriladi va "
                  "zaryad holati 0 yoki 1 sifatida aniqlanadi.", 16, C_TEXT)], bullet=True),
        para([run("O'chirish. ", 18, C_ACCENT, bold=True),
              run("Teskari kuchlanish zaryadni suzuvchi zulfindan chiqarib yuboradi; "
                  "o'chirish bloklar bo'yicha amalga oshiriladi.", 16, C_TEXT)], bullet=True),
        para([run("Ikkilik sanoq sistemasi. ", 18, C_ACCENT, bold=True),
              run("Zaryad bor holat — 0, zaryad yo'q holat — 1 (yoki aksincha) sifatida "
                  "talqin qilinadi.", 16, C_TEXT)], bullet=True),
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, CONTENT_W, 3400000,
                          body, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[7])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 8
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Minglab rasm elektron axborot sifatida",
                             "3-QISM — JARAYONLAR")
    shapes += hdr
    body = [
        para([run("Rasm — baytlar to'plami. ", 18, C_ACCENT, bold=True),
              run("Har bir rasm pikesllardan, piksellar esa 0 va 1 lardan iborat "
                  "baytlar ketma-ketligi sifatida kodlanadi.", 16, C_TEXT)], bullet=True),
        para([run("Baytlardan katakchalarga. ", 18, C_ACCENT, bold=True),
              run("Bu bitlar millionlab xotira katakchalaridagi zaryad holatlari orqali "
                  "fizik tarzda saqlanadi.", 16, C_TEXT)], bullet=True),
        para([run("Katta sig'im. ", 18, C_ACCENT, bold=True),
              run("Mikrosxemadagi katakchalar soni milliardlab bo'lgani uchun bitta "
                  "fleshka minglab yuqori sifatli rasmni saqlay oladi.", 16, C_TEXT)],
             bullet=True),
        para([run("Manzillash. ", 18, C_ACCENT, bold=True),
              run("Boshqaruvchi (controller) har bir faylni aniq manzillarga joylashtirib, "
                  "keyinchalik xatosiz tiklaydi.", 16, C_TEXT)], bullet=True),
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, CONTENT_W, 3400000,
                          body, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[8])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 9
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Xulosa", "4-QISM — XULOSA")
    shapes += hdr
    body = [
        para([run("Afzalliklari. ", 18, C_ACCENT2, bold=True),
              run("Fleshkalar ixcham, yengil, tez va energiya tejamkor bo'lib, kundalik "
                  "ma'lumot almashinuvini osonlashtiradi.", 16, C_TEXT)], bullet=True),
        para([run("Chidamliligi. ", 18, C_ACCENT2, bold=True),
              run("Harakatlanuvchi qismlar yo'qligi tufayli ular zarba va tebranishga "
                  "chidamli, mexanik buzilishlarga kam uchraydi.", 16, C_TEXT)], bullet=True),
        para([run("Texnologiya kelajagi. ", 18, C_ACCENT2, bold=True),
              run("3D NAND va yangi yarimo'tkazgich materiallari sig'imni oshirib, narxni "
                  "pasaytirmoqda — kelajakda yanada tezroq va hajmli xotira kutilmoqda.",
                  16, C_TEXT)], bullet=True),
        para([run("Umumiy xulosa. ", 18, C_ACCENT2, bold=True),
              run("Fleshka — bu harakatsiz, lekin elektr zaryadlari asosida ishlaydigan "
                  "ishonchli raqamli xotira yechimidir.", 16, C_TEXT)], bullet=True),
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, CONTENT_W, 3400000,
                          body, anchor="t")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[9])
    shapes.append(cap)
    slides.append(slide_xml(shapes))

    # ------------------------------------------------------------------ 10
    sid = 2
    shapes = []
    shapes.append(rect(sid, "panel", 0, 0, SLIDE_W, SLIDE_H, C_TITLE_BG)); sid += 1
    shapes.append(rect(sid, "topbar", 0, 0, SLIDE_W, 90000, C_ACCENT)); sid += 1
    shapes.append(rect(sid, "botbar", 0, SLIDE_H - 90000, SLIDE_W, 90000, C_ACCENT2)); sid += 1
    body = [
        para(run("E'tiboringiz uchun rahmat!", 40, C_ACCENT, bold=True),
             align="ctr", space_before=0),
        para(run("Savollaringiz uchun ochiq maydon", 22, C_TEXT, italic=True),
             align="ctr", space_before=600),
        para(run("Toshboʻriyeva Marjona — 242_26-guruh", 16, C_MUTED),
             align="ctr", space_before=800),
    ]
    shapes.append(textbox(sid, "thanks", MARGIN_X, 2300000, CONTENT_W, 2400000,
                          body, anchor="ctr")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[10])
    shapes.append(cap)
    slides.append(slide_xml(shapes, bg=C_TITLE_BG))

    return slides


# ===========================================================================
# OOXML paket qismlari (statik)
# ===========================================================================
def content_types(n):
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, n + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        f'{overrides}'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>'
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
)


def presentation_xml(n):
    sldids = "".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, n + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{sldids}</p:sldIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="screen16x9"/>'
        f'<p:notesSz cx="{SLIDE_H}" cy="{SLIDE_W}"/>'
        '</p:presentation>'
    )


def presentation_rels(n):
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    for i in range(1, n + 1):
        rels.append(
            f'<Relationship Id="rId{i + 1}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/slide{i}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{n + 2}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        f'Target="theme/theme1.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels) + '</Relationships>'
    )


THEME = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
    '<a:themeElements>'
    '<a:clrScheme name="Office">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="0B1E3B"/></a:dk2>'
    '<a:lt2><a:srgbClr val="EAF2FF"/></a:lt2>'
    '<a:accent1><a:srgbClr val="27E0C8"/></a:accent1>'
    '<a:accent2><a:srgbClr val="5AA9FF"/></a:accent2>'
    '<a:accent3><a:srgbClr val="12294A"/></a:accent3>'
    '<a:accent4><a:srgbClr val="9FB8D8"/></a:accent4>'
    '<a:accent5><a:srgbClr val="061329"/></a:accent5>'
    '<a:accent6><a:srgbClr val="2E86DE"/></a:accent6>'
    '<a:hlink><a:srgbClr val="5AA9FF"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="9FB8D8"/></a:folHlink>'
    '</a:clrScheme>'
    '<a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme>'
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '</a:fillStyleLst>'
    '<a:lnStyleLst>'
    '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '</a:lnStyleLst>'
    '<a:effectStyleLst>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '</a:effectStyleLst>'
    '<a:bgFillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '</a:bgFillStyleLst>'
    '</a:fmtScheme>'
    '</a:themeElements>'
    '</a:theme>'
)


SLIDE_MASTER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0B1E3B"/></a:solidFill>'
    '<a:effectLst/></p:bgPr></p:bg>'
    '<p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld>'
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
    'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
    'accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '<p:txStyles>'
    '<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
    '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:bodyStyle>'
    '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
    '</p:txStyles>'
    '</p:sldMaster>'
)

SLIDE_MASTER_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    '</Relationships>'
)

SLIDE_LAYOUT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'type="blank" preserve="1">'
    '<p:cSld name="Blank">'
    '<p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld>'
    '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
    'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
    'accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'
    '</p:sldLayout>'
)

SLIDE_LAYOUT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>'
)


def slide_rels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )


def core_xml():
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Fleshkalar (USB) qanday qilib ma&#39;lumot saqlaydi?</dc:title>'
        '<dc:creator>Toshbo\u02bbriyeva Marjona (242_26-guruh)</dc:creator>'
        '<cp:lastModifiedBy>Toshbo\u02bbriyeva Marjona</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        '</cp:coreProperties>'
    )


def app_xml(n):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Python OOXML Generator</Application>'
        f'<Slides>{n}</Slides>'
        '<Company>Jizzax davlat pedagogika universiteti</Company>'
        '</Properties>'
    )


# ===========================================================================
# Paketlash
# ===========================================================================
def build(output_path):
    slides = build_slides()
    n = len(slides)
    assert n == 10, f"Slaydlar soni 10 bo'lishi shart, hozir: {n}"

    parts = {
        "[Content_Types].xml": content_types(n),
        "_rels/.rels": ROOT_RELS,
        "ppt/presentation.xml": presentation_xml(n),
        "ppt/_rels/presentation.xml.rels": presentation_rels(n),
        "ppt/theme/theme1.xml": THEME,
        "ppt/slideMasters/slideMaster1.xml": SLIDE_MASTER,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": SLIDE_MASTER_RELS,
        "ppt/slideLayouts/slideLayout1.xml": SLIDE_LAYOUT,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": SLIDE_LAYOUT_RELS,
        "docProps/core.xml": core_xml(),
        "docProps/app.xml": app_xml(n),
    }
    for i, s in enumerate(slides, start=1):
        parts[f"ppt/slides/slide{i}.xml"] = s
        parts[f"ppt/slides/_rels/slide{i}.xml.rels"] = slide_rels()

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml birinchi bo'lib yozilishi tavsiya etiladi
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            z.writestr(name, data)
    return output_path, n


if __name__ == "__main__":
    out = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "Fleshkalar_USB_malumot_saqlashi.pptx",
    )
    path, n = build(out)
    print(f"Yaratildi: {path}")
    print(f"Slaydlar soni: {n}")
