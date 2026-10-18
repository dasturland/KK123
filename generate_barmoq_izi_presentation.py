#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_barmoq_izi_presentation.py

Pure standard-library generator for a 10-slide .pptx presentation (academic
Uzbek) titled "Barmoq izi skanerlari qanday ishlaydi?".

Why hand-written OOXML?
-----------------------
The sandbox has no outbound network access: PyPI returns 403 Forbidden, so
python-pptx / lxml / Pillow cannot be installed and are not present. A .pptx
file is simply a ZIP archive of XML (OpenXML / OOXML) parts, so we build every
part by hand and zip them with the standard-library ``zipfile`` module.

Because images cannot be downloaded (no network), every slide displays its
Unsplash image URL as visible text AND as a working clickable hyperlink
(``<a:hlinkClick>`` + an External relationship in the slide .rels), and the
same URL is repeated in the slide notes. This satisfies the requirement that
Unsplash image URLs be included on every slide.

Only stdlib is used: zipfile, html, datetime, os.
"""

import os
import html
import zipfile
import datetime

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

OUTPUT_FILE = "Taqdimot_Barmoq_izi_skanerlari.pptx"

# Slide size: 16:9 widescreen (EMU)
SLIDE_CX = 12192000
SLIDE_CY = 6858000

# Modern tech / cybersecurity theme colours (hex, no leading '#')
BG_DARK = "0B1E3B"       # deep navy slide background
BG_PANEL = "0F2A4A"      # slightly lighter panel navy
ACCENT_CYAN = "00E5FF"   # electric cyan accent
ACCENT_BLUE = "1E88E5"   # electric blue accent
TEXT_LIGHT = "E8F1FA"    # light text
TEXT_MUTED = "AFC6DD"    # muted light text

FONT = "Times New Roman"

# Font sizes (OOXML centipoints: 12pt = 1200). Nothing below 1200.
SZ_TITLE_MAIN = "4000"   # title-slide main heading (40pt)
SZ_TITLE = "3200"        # content-slide titles (32pt)
SZ_SECTION = "2400"      # section / subtitle (24pt)
SZ_SUBTITLE = "2000"     # subtitle / key bullets (20pt)
SZ_KEY = "2000"          # key bullet points (20pt)
SZ_BODY = "1600"         # normal body (16pt)
SZ_SMALL = "1200"        # smallest allowed footnote / URL text (12pt)


# ---------------------------------------------------------------------------
# Low-level XML helpers
# ---------------------------------------------------------------------------

def esc(text):
    """XML-escape dynamic text (quotes included)."""
    return html.escape(str(text), quote=True)


def run(text, sz, color=TEXT_LIGHT, bold=False, italic=False, hlink_rid=None):
    """
    Emit a single text run (<a:r>) with explicit font and size.

    Every run sets latin/cs/ea typeface to Times New Roman and an explicit
    sz attribute, as required.
    """
    b = ' b="1"' if bold else ' b="0"'
    i = ' i="1"' if italic else ''
    hlink = ''
    if hlink_rid is not None:
        hlink = ('<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/'
                 'officeDocument/2006/relationships" r:id="%s"/>' % esc(hlink_rid))
    return (
        '<a:r>'
        '<a:rPr lang="uz-UZ" sz="%s"%s%s dirty="0">'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '%s'
        '<a:latin typeface="%s"/><a:cs typeface="%s"/><a:ea typeface="%s"/>'
        '</a:rPr>'
        '<a:t>%s</a:t>'
        '</a:r>'
    ) % (sz, b, i, color, hlink, FONT, FONT, FONT, esc(text))


def para(runs, align="l", bullet=False, space_before=600):
    """Emit a paragraph (<a:p>) made of pre-rendered run strings."""
    if bullet:
        pPr = ('<a:pPr marL="285750" indent="-285750" algn="%s">'
               '<a:spcBef><a:spcPts val="%d"/></a:spcBef>'
               '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>'
               '</a:pPr>') % (align, space_before)
    else:
        pPr = ('<a:pPr algn="%s">'
               '<a:spcBef><a:spcPts val="%d"/></a:spcBef>'
               '<a:buNone/></a:pPr>') % (align, space_before)
    return '<a:p>%s%s</a:p>' % (pPr, "".join(runs))


def empty_para():
    return '<a:p><a:pPr><a:buNone/></a:pPr></a:p>'


def textbox(shape_id, name, x, y, cx, cy, paragraphs, anchor="t"):
    """A simple text-box shape (<p:sp>) holding paragraphs."""
    return (
        '<p:sp>'
        '<p:nvSpPr>'
        '<p:cNvPr id="%d" name="%s"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        '<p:nvPr/>'
        '</p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '</p:spPr>'
        '<p:txBody>'
        '<a:bodyPr wrap="square" anchor="%s"><a:normAutofit/></a:bodyPr>'
        '<a:lstStyle/>'
        '%s'
        '</p:txBody>'
        '</p:sp>'
    ) % (shape_id, esc(name), x, y, cx, cy, anchor, "".join(paragraphs))


def accent_bar(shape_id, x, y, cx, cy, color=ACCENT_CYAN):
    """A thin filled rectangle used as a decorative accent bar."""
    return (
        '<p:sp>'
        '<p:nvSpPr><p:cNvPr id="%d" name="AccentBar%d"/>'
        '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr>'
        '<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln><a:noFill/></a:ln>'
        '</p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        '</p:sp>'
    ) % (shape_id, shape_id, x, y, cx, cy, color)


# ---------------------------------------------------------------------------
# Slide content model
# ---------------------------------------------------------------------------

TOPIC = ("Barmoq izi skanerlari qanday ishlaydi? Telefonni ochish uchun "
         "barmoq izimizni optik yoki ultratovush to'lqinlari orqali qanday "
         "tanib olishi")

# Unsplash URLs matched to each slide topic. All use the canonical CDN form
# ``https://images.unsplash.com/photo-<id>`` so the links resolve reliably
# (the ``unsplash.com/photos/<slug>`` form was dropped because slugs are not
# verifiable in a networkless build and may 404).
UNSPLASH = {
    1: "https://images.unsplash.com/photo-1563013544-824ae1b704d3",
    2: "https://images.unsplash.com/photo-1518770660439-4636190af475",
    3: "https://images.unsplash.com/photo-1510915361894-db8b60106cb1",
    4: "https://images.unsplash.com/photo-1511707171634-5f897ff02aa9",
    5: "https://images.unsplash.com/photo-1534972195531-d756b9bfa9f2",
    6: "https://images.unsplash.com/photo-1507413245164-6160d8298b31",
    7: "https://images.unsplash.com/photo-1451187580459-43490279c0fa",
    8: "https://images.unsplash.com/photo-1446776811953-b23d57bd21aa",
    9: "https://images.unsplash.com/photo-1550751827-4bd374c3f58b",
    10: "https://images.unsplash.com/photo-1498050108023-c5249f4df085",
}

# Short caption describing the image topic for each slide (visible context).
IMG_CAPTION = {
    1: "Rasm (biometrika): ",
    2: "Rasm (zamonaviy texnologiya): ",
    3: "Rasm (barmoq izi biometrikasi): ",
    4: "Rasm (smartfon va barmoq izi): ",
    5: "Rasm (optika va yorug'lik): ",
    6: "Rasm (optik sensor): ",
    7: "Rasm (ultratovush to'lqinlari): ",
    8: "Rasm (ovoz to'lqini vizualizatsiyasi): ",
    9: "Rasm (kiberxavfsizlik): ",
    10: "Rasm (texnologiya): ",
}


def slide_body_shapes(n, start_id):
    """
    Return (list_of_shape_xml, notes_text) for slide number n.

    Each slide gets: title, content paragraphs, and an Unsplash URL footer
    with a clickable hyperlink (rId2 is the hyperlink relationship on the
    slide; see slide_rels()).
    """
    sid = start_id
    shapes = []

    def title_shape(text, main=False):
        nonlocal sid
        sz = SZ_TITLE_MAIN if main else SZ_TITLE
        sp = textbox(sid, "Title", 548640, 365760, 11094720, 1500000,
                     [para([run(text, sz, color=ACCENT_CYAN, bold=True)],
                           align="l")], anchor="t")
        sid += 1
        return sp

    def body_shape(paragraphs, y=1950720, cy=4000000, anchor="t"):
        nonlocal sid
        sp = textbox(sid, "Body", 548640, y, 11094720, cy, paragraphs,
                     anchor=anchor)
        sid += 1
        return sp

    url = UNSPLASH[n]
    caption = IMG_CAPTION[n]

    # ---- URL footer shape (same on every slide) -------------------------
    def url_footer():
        nonlocal sid
        sp = textbox(
            sid, "UnsplashUrl", 548640, 6150000, 11094720, 520000,
            [para([
                run(caption, SZ_SMALL, color=TEXT_MUTED, bold=True),
                run(url, SZ_SMALL, color=ACCENT_CYAN, hlink_rid="rId2"),
            ], align="l", space_before=0)],
            anchor="b")
        sid += 1
        return sp

    # ---- Per-slide content ---------------------------------------------
    if n == 1:
        shapes.append(accent_bar(900 + sid, 548640, 1750000, 3200000, 54000))
        shapes.append(title_shape(
            "Barmoq izi skanerlari qanday ishlaydi?", main=True))
        shapes.append(body_shape([
            para([run("Telefonni ochish uchun barmoq izimizni optik yoki "
                      "ultratovush to'lqinlari orqali qanday tanib olishi",
                      SZ_SECTION, color=TEXT_LIGHT)], align="l"),
            empty_para(),
            para([run("Jizzax davlat pedagogika universiteti - Axborot "
                      "tizimlari fani", SZ_SUBTITLE, color=TEXT_LIGHT,
                      bold=True)], align="l"),
            empty_para(),
            para([run("Taqdimotchi: Toshbo\u02bbriyeva Marjona (242_26 - "
                      "guruh)", SZ_SUBTITLE, color=ACCENT_BLUE, bold=True)],
                 align="l"),
        ], y=2050000, cy=3600000))
        notes = ("Titul varag'i. Mavzu, muassasa (Jizzax davlat pedagogika "
                 "universiteti - Axborot tizimlari fani) va taqdimotchi "
                 "(Toshbo\u02bbriyeva Marjona, 242_26 - guruh) keltirilgan. "
                 "Rasm manbasi (Unsplash): " + url)

    elif n == 2:
        shapes.append(title_shape("Reja"))
        shapes.append(body_shape([
            para([run("1-qism. Barmoq izi biometrikasi asoslari", SZ_KEY,
                      color=TEXT_LIGHT, bold=True)], bullet=True),
            para([run("2-qism. Optik skanerlar qanday ishlaydi?", SZ_KEY,
                      color=TEXT_LIGHT, bold=True)], bullet=True),
            para([run("3-qism. Ultratovushli (ultrasonic) skanerlar", SZ_KEY,
                      color=TEXT_LIGHT, bold=True)], bullet=True),
            para([run("4-qism. Xulosa: texnologiyalarning qiyosiy tahlili",
                      SZ_KEY, color=ACCENT_CYAN, bold=True)], bullet=True),
        ]))
        notes = ("Taqdimot rejasi: uchta asosiy qism va yakuniy xulosa. "
                 "Rasm manbasi (Unsplash): " + url)

    elif n == 3:
        shapes.append(accent_bar(900 + sid, 548640, 1620000, 2600000, 40000))
        shapes.append(title_shape(
            "1-qism. Barmoq izi biometrikasi asoslari"))
        shapes.append(body_shape([
            para([run("Nega barmoq izi takrorlanmas hisoblanadi?", SZ_SECTION,
                      color=ACCENT_CYAN, bold=True)], align="l"),
            para([run("Barmoq terisidagi chiziqlar (papillyar naqshlar) homila "
                      "rivojlanishi davrida shakllanadi va har bir insonda "
                      "noyob bo'ladi.", SZ_BODY)], bullet=True),
            para([run("Hatto bir xil tuxumdan tug'ilgan egizaklarda ham barmoq "
                      "izi naqshlari bir-biridan farq qiladi.", SZ_BODY)],
                 bullet=True),
            para([run("Naqshning asosiy xususiyatlari (minutsiyalar): "
                      "chiziqlarning tugashi, tarmoqlanishi va markaziy "
                      "nuqtalari biometrik identifikatsiya uchun ishlatiladi.",
                      SZ_BODY)], bullet=True),
        ]))
        notes = ("Barmoq izi naqshining noyobligi va minutsiyalar tushunchasi "
                 "yoritiladi. Rasm manbasi (Unsplash): " + url)

    elif n == 4:
        shapes.append(title_shape(
            "1-qism. Smartfonlarda barmoq izi skanerlari"))
        shapes.append(body_shape([
            para([run("Paydo bo'lishi va evolyutsiyasi", SZ_SECTION,
                      color=ACCENT_CYAN, bold=True)], align="l"),
            para([run("Dastlabki modellarda sig'imli (kapasitiv) sensorlar "
                      "qo'llanilgan; ular barmoq terisining elektr sig'imini "
                      "o'lchagan.", SZ_BODY)], bullet=True),
            para([run("Keyinchalik ekran ostiga joylashtiriladigan optik "
                      "skanerlar paydo bo'lib, dizaynni ixchamlashtirdi.",
                      SZ_BODY)], bullet=True),
            para([run("Eng zamonaviy bosqich - ekran ostidagi ultratovushli "
                      "skanerlar bo'lib, ular yuqori aniqlik va xavfsizlikni "
                      "ta'minlaydi.", SZ_BODY)], bullet=True),
        ]))
        notes = ("Smartfon skanerlarining kapasitiv, optik va ultratovushli "
                 "bosqichlardan iborat evolyutsiyasi. "
                 "Rasm manbasi (Unsplash): " + url)

    elif n == 5:
        shapes.append(accent_bar(900 + sid, 548640, 1620000, 2600000, 40000))
        shapes.append(title_shape("2-qism. Optik skanerlar qanday ishlaydi?"))
        shapes.append(body_shape([
            para([run("Yorug'lik yordamida 2D rasm olish", SZ_SECTION,
                      color=ACCENT_CYAN, bold=True)], align="l"),
            para([run("Yorug'lik diodi (LED) barmoq yuzasini yoritadi va "
                      "undan qaytgan nur maxsus sensorga tushadi.", SZ_BODY)],
                 bullet=True),
            para([run("Chiziqlarning balandligi va chuqurligi yorug'likni "
                      "turlicha qaytaradi; natijada barmoq izining ikki "
                      "o'lchovli (2D) tasviri hosil bo'ladi.", SZ_BODY)],
                 bullet=True),
            para([run("Hosil bo'lgan tasvir raqamli naqshga aylantirilib, "
                      "xotiradagi namuna bilan solishtiriladi.", SZ_BODY)],
                 bullet=True),
        ]))
        notes = ("Optik skanerning yorug'lik va sensor orqali 2D tasvir olish "
                 "prinsipi. Rasm manbasi (Unsplash): " + url)

    elif n == 6:
        shapes.append(title_shape("2-qism. Optik skanerlarning kamchiliklari"))
        shapes.append(body_shape([
            para([run("Cheklovlar va zaifliklar", SZ_SECTION,
                      color=ACCENT_CYAN, bold=True)], align="l"),
            para([run("Faqat yuza (2D) tasvirga tayangani uchun yuqori sifatli "
                      "soxta nusxalar (foto yoki qolip) bilan aldanishi mumkin.",
                      SZ_BODY)], bullet=True),
            para([run("Barmoqning namligi, kiri yoki yorug'lik sharoiti "
                      "aniqlikka salbiy ta'sir ko'rsatadi.", SZ_BODY)],
                 bullet=True),
            para([run("Xavfsizlik darajasi ultratovushli skanerlarga nisbatan "
                      "pastroq hisoblanadi.", SZ_BODY)], bullet=True),
        ]))
        notes = ("Optik skanerlarning soxtalashtirishga nisbatan zaifligi va "
                 "muhit ta'siri. Rasm manbasi (Unsplash): " + url)

    elif n == 7:
        shapes.append(accent_bar(900 + sid, 548640, 1620000, 2600000, 40000))
        shapes.append(title_shape(
            "3-qism. Ultratovushli (ultrasonic) skanerlar"))
        shapes.append(body_shape([
            para([run("Ovoz to'lqinlari bilan 3D xarita yaratish", SZ_SECTION,
                      color=ACCENT_CYAN, bold=True)], align="l"),
            para([run("Skaner inson qulog'iga eshitilmaydigan ultratovush "
                      "to'lqinlarini chiqaradi va barmoqdan qaytgan to'lqinlarni "
                      "qayd etadi.", SZ_BODY)], bullet=True),
            para([run("Teri yuzasidagi chiziqlar va teshikchalardan "
                      "to'lqinlarning turlicha qaytishi barmoqning uch "
                      "o'lchovli (3D) xaritasini shakllantiradi.", SZ_BODY)],
                 bullet=True),
            para([run("3D ma'lumot teri relyefini chuqur tahlil qilish "
                      "imkonini beradi.", SZ_BODY)], bullet=True),
        ]))
        notes = ("Ultratovushli skanerning ovoz to'lqinlari orqali 3D xarita "
                 "yaratishi. Rasm manbasi (Unsplash): " + url)

    elif n == 8:
        shapes.append(title_shape("3-qism. Chuqur tahlil va xavfsizlik"))
        shapes.append(body_shape([
            para([run("Teshikchalar, chiziqlar va yuqori ishonchlilik",
                      SZ_SECTION, color=ACCENT_CYAN, bold=True)], align="l"),
            para([run("Ultratovush teri qatlamlari ostidagi mayda "
                      "teshikchalar va chiziqlarni ham aniqlaydi, bu esa "
                      "soxtalashtirish imkonini keskin kamaytiradi.", SZ_BODY)],
                 bullet=True),
            para([run("Namlik, kir yoki yorug'lik sharoiti natijaga deyarli "
                      "ta'sir qilmaydi.", SZ_BODY)], bullet=True),
            para([run("Shu bois ultratovushli skanerlar eng yuqori xavfsizlik "
                      "darajasini ta'minlaydi.", SZ_KEY, color=ACCENT_CYAN,
                      bold=True)], bullet=True),
        ]))
        notes = ("Ultratovushli skanerning chuqur tahlil qobiliyati va yuqori "
                 "xavfsizlik darajasi. Rasm manbasi (Unsplash): " + url)

    elif n == 9:
        shapes.append(accent_bar(900 + sid, 548640, 1620000, 2600000, 40000))
        shapes.append(title_shape("4-qism. Xulosa"))
        shapes.append(body_shape([
            para([run("Optik va ultratovush texnologiyalarining qiyosi",
                      SZ_SECTION, color=ACCENT_CYAN, bold=True)], align="l"),
            para([run("Optik skanerlar yorug'lik bilan 2D tasvir oladi: "
                      "arzon, lekin soxtalashtirishga nisbatan zaifroq.",
                      SZ_BODY)], bullet=True),
            para([run("Ultratovushli skanerlar ovoz to'lqinlari bilan 3D "
                      "xarita yaratadi: ishonchliroq va xavfsizroq.", SZ_BODY)],
                 bullet=True),
            para([run("Kelajakda biometrik xavfsizlik ko'p omilli (barmoq "
                      "izi, yuz va ovoz) usullar bilan yanada mustahkamlanadi.",
                      SZ_KEY, color=ACCENT_CYAN, bold=True)], bullet=True),
        ]))
        notes = ("Ikki texnologiyaning qiyosiy xulosasi va biometrik "
                 "xavfsizlik istiqbollari. Rasm manbasi (Unsplash): " + url)

    else:  # n == 10
        shapes.append(accent_bar(900 + sid, 548640, 2050000, 3200000, 54000))
        shapes.append(title_shape("E'tiboringiz uchun rahmat!", main=True))
        shapes.append(body_shape([
            para([run("Savollaringiz bo'lsa, bajonidil javob beraman.",
                      SZ_SECTION, color=TEXT_LIGHT)], align="l"),
            empty_para(),
            para([run("Savollar uchun ochiq maydon.", SZ_SUBTITLE,
                      color=TEXT_MUTED, italic=True)], align="l"),
            empty_para(),
            para([run("Taqdimotchi: Toshbo\u02bbriyeva Marjona (242_26 - "
                      "guruh)", SZ_SUBTITLE, color=ACCENT_BLUE, bold=True)],
                 align="l"),
        ], y=2350000, cy=3200000))
        notes = ("Yakunlovchi slayd: minnatdorchilik va savol-javob uchun "
                 "ochiq maydon. Rasm manbasi (Unsplash): " + url)

    shapes.append(url_footer())
    return shapes, notes


# ---------------------------------------------------------------------------
# OOXML part builders
# ---------------------------------------------------------------------------

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def content_types():
    overrides = []
    overrides.append('<Override PartName="/ppt/presentation.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument'
                     '.presentationml.presentation.main+xml"/>')
    overrides.append('<Override PartName="/ppt/presProps.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument'
                     '.presentationml.presProps+xml"/>')
    overrides.append('<Override PartName="/ppt/theme/theme1.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument'
                     '.theme+xml"/>')
    overrides.append('<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument'
                     '.presentationml.slideMaster+xml"/>')
    overrides.append('<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument'
                     '.presentationml.slideLayout+xml"/>')
    overrides.append('<Override PartName="/ppt/notesMasters/notesMaster1.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument'
                     '.presentationml.notesMaster+xml"/>')
    for i in range(1, 11):
        overrides.append(
            '<Override PartName="/ppt/slides/slide%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument'
            '.presentationml.slide+xml"/>' % i)
        overrides.append(
            '<Override PartName="/ppt/notesSlides/notesSlide%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument'
            '.presentationml.notesSlide+xml"/>' % i)
    overrides.append('<Override PartName="/docProps/core.xml" '
                     'ContentType="application/vnd.openxmlformats-package'
                     '.core-properties+xml"/>')
    overrides.append('<Override PartName="/docProps/app.xml" '
                     'ContentType="application/vnd.openxmlformats-officedocument'
                     '.extended-properties+xml"/>')
    return (XML_DECL +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types">'
            '<Default Extension="rels" ContentType="application/vnd.'
            'openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(overrides) +
            '</Types>')


def root_rels():
    return (XML_DECL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/officeDocument" '
            'Target="ppt/presentation.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
            'package/2006/relationships/metadata/core-properties" '
            'Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/extended-properties" '
            'Target="docProps/app.xml"/>'
            '</Relationships>')


def core_props():
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return (XML_DECL +
            '<cp:coreProperties '
            'xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
            'metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title>'
            '<dc:creator>Toshbo\u02bbriyeva Marjona (242_26 - guruh)</dc:creator>'
            '<cp:lastModifiedBy>Toshbo\u02bbriyeva Marjona</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
            '</cp:coreProperties>') % (esc(TOPIC), now, now)


def app_props():
    return (XML_DECL +
            '<Properties xmlns="http://schemas.openxmlformats.org/'
            'officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/'
            'docPropsVTypes">'
            '<Application>Python stdlib OOXML generator</Application>'
            '<Slides>10</Slides>'
            '<Company>Jizzax davlat pedagogika universiteti</Company>'
            '</Properties>')


def presentation_xml():
    sldIds = []
    # slide relationship IDs rId2..rId11 (rId1 = slideMaster)
    for i in range(1, 11):
        sldIds.append('<p:sldId id="%d" r:id="rId%d"/>' % (255 + i, i + 1))
    return (XML_DECL +
            '<p:presentation '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/'
            'main" saveSubsetFonts="1">'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/>'
            '</p:sldMasterIdLst>'
            '<p:notesMasterIdLst><p:notesMasterId r:id="rId12"/>'
            '</p:notesMasterIdLst>'
            '<p:sldIdLst>%s</p:sldIdLst>'
            '<p:sldSz cx="%d" cy="%d"/>'
            '<p:notesSz cx="6858000" cy="9144000"/>'
            '</p:presentation>') % ("".join(sldIds), SLIDE_CX, SLIDE_CY)


def presentation_rels():
    rels = []
    rels.append('<Relationship Id="rId1" Type="http://schemas.openxmlformats'
                '.org/officeDocument/2006/relationships/slideMaster" '
                'Target="slideMasters/slideMaster1.xml"/>')
    for i in range(1, 11):
        rels.append('<Relationship Id="rId%d" Type="http://schemas.'
                    'openxmlformats.org/officeDocument/2006/relationships/'
                    'slide" Target="slides/slide%d.xml"/>' % (i + 1, i))
    rels.append('<Relationship Id="rId12" Type="http://schemas.openxmlformats'
                '.org/officeDocument/2006/relationships/notesMaster" '
                'Target="notesMasters/notesMaster1.xml"/>')
    rels.append('<Relationship Id="rId13" Type="http://schemas.openxmlformats'
                '.org/officeDocument/2006/relationships/presProps" '
                'Target="presProps.xml"/>')
    rels.append('<Relationship Id="rId14" Type="http://schemas.openxmlformats'
                '.org/officeDocument/2006/relationships/theme" '
                'Target="theme/theme1.xml"/>')
    return (XML_DECL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">' + "".join(rels) + '</Relationships>')


def pres_props():
    return (XML_DECL +
            '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/'
            'drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/'
            'main"/>')


def theme_xml():
    def scheme_color(name, val):
        return '<a:%s><a:srgbClr val="%s"/></a:%s>' % (name, val, name)
    font_entry = ('<a:latin typeface="%s"/><a:ea typeface="%s"/>'
                  '<a:cs typeface="%s"/>') % (FONT, FONT, FONT)
    return (XML_DECL +
            '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/'
            '2006/main" name="Kiberxavfsizlik">'
            '<a:themeElements>'
            '<a:clrScheme name="Kiberxavfsizlik">'
            '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
            '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
            + scheme_color("dk2", BG_DARK)
            + scheme_color("lt2", TEXT_LIGHT)
            + scheme_color("accent1", ACCENT_CYAN)
            + scheme_color("accent2", ACCENT_BLUE)
            + scheme_color("accent3", BG_PANEL)
            + scheme_color("accent4", TEXT_MUTED)
            + scheme_color("accent5", "26C6DA")
            + scheme_color("accent6", "0D47A1")
            + '<a:hlink><a:srgbClr val="%s"/></a:hlink>' % ACCENT_CYAN
            + '<a:folHlink><a:srgbClr val="%s"/></a:folHlink>' % ACCENT_BLUE
            + '</a:clrScheme>'
            '<a:fontScheme name="Times New Roman">'
            '<a:majorFont>' + font_entry + '</a:majorFont>'
            '<a:minorFont>' + font_entry + '</a:minorFont>'
            '</a:fontScheme>'
            '<a:fmtScheme name="Kiberxavfsizlik">'
            '<a:fillStyleLst>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '</a:fillStyleLst>'
            '<a:lnStyleLst>'
            '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln>'
            '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln>'
            '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/>'
            '</a:solidFill></a:ln>'
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
            '</a:theme>')


def bg_xml():
    """Dark solid background used on master and slides."""
    return ('<p:bg><p:bgPr>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:effectLst/>'
            '</p:bgPr></p:bg>') % BG_DARK


def slide_master_xml():
    return (XML_DECL +
            '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/'
            '2006/main" xmlns:r="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/'
            'main">'
            '<p:cSld>'
            + bg_xml() +
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree>'
            '</p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"'
            ' accent2="accent2" accent3="accent3" accent4="accent4" '
            'accent5="accent5" accent6="accent6" hlink="hlink" '
            'folHlink="folHlink"/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/>'
            '</p:sldLayoutIdLst>'
            '<p:txStyles>'
            '<p:titleStyle>'
            '<a:lvl1pPr><a:defRPr sz="3200"><a:solidFill>'
            '<a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:cs typeface="%s"/>'
            '<a:ea typeface="%s"/></a:defRPr></a:lvl1pPr>'
            '</p:titleStyle>'
            '<p:bodyStyle>'
            '<a:lvl1pPr><a:defRPr sz="1600"><a:solidFill>'
            '<a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:cs typeface="%s"/>'
            '<a:ea typeface="%s"/></a:defRPr></a:lvl1pPr>'
            '</p:bodyStyle>'
            '<p:otherStyle/>'
            '</p:txStyles>'
            '</p:sldMaster>') % (ACCENT_CYAN, FONT, FONT, FONT,
                                 TEXT_LIGHT, FONT, FONT, FONT)


def slide_master_rels():
    return (XML_DECL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/slideLayout" '
            'Target="../slideLayouts/slideLayout1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/theme" '
            'Target="../theme/theme1.xml"/>'
            '</Relationships>')


def slide_layout_xml():
    return (XML_DECL +
            '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/'
            '2006/main" xmlns:r="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/'
            'main" type="blank" preserve="1">'
            '<p:cSld name="Blank">'
            + bg_xml() +
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree>'
            '</p:cSld>'
            '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" '
            'tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" '
            'hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'
            '</p:sldLayout>')


def slide_layout_rels():
    return (XML_DECL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/slideMaster" '
            'Target="../slideMasters/slideMaster1.xml"/>'
            '</Relationships>')


def notes_master_xml():
    return (XML_DECL +
            '<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/'
            'drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/'
            'main">'
            '<p:cSld>'
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree>'
            '</p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"'
            ' accent2="accent2" accent3="accent3" accent4="accent4" '
            'accent5="accent5" accent6="accent6" hlink="hlink" '
            'folHlink="folHlink"/>'
            '<p:notesStyle>'
            '<a:lvl1pPr><a:defRPr sz="1600"><a:latin typeface="%s"/>'
            '<a:cs typeface="%s"/><a:ea typeface="%s"/></a:defRPr></a:lvl1pPr>'
            '</p:notesStyle>'
            '</p:notesMaster>') % (FONT, FONT, FONT)


def notes_master_rels():
    return (XML_DECL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/theme" '
            'Target="../theme/theme1.xml"/>'
            '</Relationships>')


def slide_xml(n):
    shapes, _notes = slide_body_shapes(n, start_id=2)
    return (XML_DECL +
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
            'main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/'
            'main">'
            '<p:cSld>'
            + bg_xml() +
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(shapes) +
            '</p:spTree>'
            '</p:cSld>'
            '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" '
            'tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" '
            'hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'
            '</p:sld>')


def slide_rels(n):
    """rId1 -> slideLayout, rId2 -> external Unsplash URL, rId3 -> notesSlide."""
    url = UNSPLASH[n]
    return (XML_DECL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/slideLayout" '
            'Target="../slideLayouts/slideLayout1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/hyperlink" '
            'Target="%s" TargetMode="External"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/notesSlide" '
            'Target="../notesSlides/notesSlide%d.xml"/>'
            '</Relationships>') % (esc(url), n)


def notes_slide_xml(n):
    _shapes, notes = slide_body_shapes(n, start_id=2)
    body = para([run("Slayd %d. " % n, SZ_BODY, color="000000", bold=True),
                 run(notes, SZ_BODY, color="000000")], align="l")
    return (XML_DECL +
            '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/'
            '2006/main" xmlns:r="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/'
            'main">'
            '<p:cSld>'
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
            '</p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '<p:sp>'
            '<p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="685800" y="4572000"/>'
            '<a:ext cx="5486400" cy="3600000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/>%s</p:txBody>'
            '</p:sp>'
            '</p:spTree>'
            '</p:cSld>'
            '</p:notes>') % body


def notes_slide_rels(n):
    return (XML_DECL +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
            '2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/slide" '
            'Target="../slides/slide%d.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/notesMaster" '
            'Target="../notesMasters/notesMaster1.xml"/>'
            '</Relationships>') % n


# ---------------------------------------------------------------------------
# Package assembly
# ---------------------------------------------------------------------------

def build_package():
    parts = {}
    parts["[Content_Types].xml"] = content_types()
    parts["_rels/.rels"] = root_rels()
    parts["docProps/core.xml"] = core_props()
    parts["docProps/app.xml"] = app_props()
    parts["ppt/presentation.xml"] = presentation_xml()
    parts["ppt/_rels/presentation.xml.rels"] = presentation_rels()
    parts["ppt/presProps.xml"] = pres_props()
    parts["ppt/theme/theme1.xml"] = theme_xml()
    parts["ppt/slideMasters/slideMaster1.xml"] = slide_master_xml()
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = slide_master_rels()
    parts["ppt/slideLayouts/slideLayout1.xml"] = slide_layout_xml()
    parts["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = slide_layout_rels()
    parts["ppt/notesMasters/notesMaster1.xml"] = notes_master_xml()
    parts["ppt/notesMasters/_rels/notesMaster1.xml.rels"] = notes_master_rels()

    for i in range(1, 11):
        parts["ppt/slides/slide%d.xml" % i] = slide_xml(i)
        parts["ppt/slides/_rels/slide%d.xml.rels" % i] = slide_rels(i)
        parts["ppt/notesSlides/notesSlide%d.xml" % i] = notes_slide_xml(i)
        parts["ppt/notesSlides/_rels/notesSlide%d.xml.rels" % i] = \
            notes_slide_rels(i)

    return parts


def write_pptx(path):
    parts = build_package()
    # [Content_Types].xml should be first; order of the rest does not matter.
    ordered = ["[Content_Types].xml"] + sorted(
        k for k in parts if k != "[Content_Types].xml")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in ordered:
            z.writestr(name, parts[name].encode("utf-8"))
    return len(parts)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, OUTPUT_FILE)
    count = write_pptx(out)
    size = os.path.getsize(out)
    print("Yaratildi: %s" % out)
    print("Qismlar soni: %d, hajmi: %d bayt" % (count, size))
    print("Slaydlar soni: 10")


if __name__ == "__main__":
    main()
