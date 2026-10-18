#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kompyuterlar nima uchun qiziydi va qanday soviydi? - 10 slaydli taqdimot generatori.

Ushbu skript faqat Python standart kutubxonasidan foydalanadi (zipfile, xml).
python-pptx va lxml talab QILINMAYDI (paket registri bloklangan). .pptx fayl
OPC/OpenXML ZIP paketini qo'lda yig'ish orqali yaratiladi.

Rasmlar tarmoq cheklovi tufayli yuklab olinmaydi/joylashtirilmaydi; har bir
slaydda Unsplash URL manzili ko'rinadigan matn qatori sifatida beriladi.
"""

import os
import zipfile
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# Slayd o'lchami (16:9) - EMU birligida
# ---------------------------------------------------------------------------
SLIDE_W = 12192000
SLIDE_H = 6858000

OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "Kompyuterlar_qizishi_va_sovishi.pptx")

# Shrift o'lchamlari (OpenXML yuzdan bir punkt birligida; 12pt = 1200)
SZ_MAIN_TITLE = 4400   # 44pt - 1-slayd asosiy sarlavhasi
SZ_SLIDE_TITLE = 3600  # 36pt - slayd sarlavhalari
SZ_HEADING = 2800      # 28pt - bo'lim/reja sarlavhalari
SZ_KEY = 2200          # 22pt - asosiy nuqtalar
SZ_BODY = 1600         # 16pt - tafsilot matni
SZ_SMALL = 1300        # 13pt - Unsplash URL (minimumdan yuqori)

NSMAP_SP = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
)


# ---------------------------------------------------------------------------
# Matn / paragraf yordamchilari
# ---------------------------------------------------------------------------
def run_xml(text, sz, bold=False, color=None, italic=False):
    """Bitta <a:r> run yaratadi; har doim Times New Roman + aniq o'lcham."""
    b = ' b="1"' if bold else ''
    i = ' i="1"' if italic else ''
    fill = ''
    if color:
        fill = '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % color
    return (
        '<a:r>'
        '<a:rPr lang="uz-UZ" sz="%d"%s%s dirty="0">%s'
        '<a:latin typeface="Times New Roman" panose="02020603050405020304" '
        'pitchFamily="18" charset="0"/>'
        '<a:cs typeface="Times New Roman" panose="02020603050405020304" '
        'pitchFamily="18" charset="0"/>'
        '</a:rPr>'
        '<a:t>%s</a:t>'
        '</a:r>'
    ) % (sz, b, i, fill, escape(text))


def para_xml(text, sz, bold=False, color=None, italic=False, bullet=False,
             align=None):
    """Bitta <a:p> paragrafini yaratadi."""
    ppr_bits = []
    if align:
        ppr_bits.append(' algn="%s"' % align)
    if bullet:
        ppr = ('<a:pPr marL="285750" indent="-285750"%s>'
               '<a:buFont typeface="Arial" pitchFamily="34" charset="0"/>'
               '<a:buChar char="&#8226;"/></a:pPr>' % (''.join(ppr_bits)))
    elif ppr_bits:
        ppr = '<a:pPr%s/>' % (''.join(ppr_bits))
    else:
        ppr = ''
    if text == "":
        # bo'sh ajratuvchi qator
        return '<a:p>%s<a:endParaRPr lang="uz-UZ" sz="%d"/></a:p>' % (ppr, sz)
    return '<a:p>%s%s</a:p>' % (ppr, run_xml(text, sz, bold, color, italic))


def textbox_sp(sp_id, name, x, y, cx, cy, paragraphs, anchor="t"):
    """Matn qutisi (<p:sp>) shaklini yaratadi."""
    body = ''.join(paragraphs)
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
        '<a:bodyPr wrap="square" rtlCol="0" anchor="%s"><a:normAutofit/></a:bodyPr>'
        '<a:lstStyle/>'
        '%s'
        '</p:txBody>'
        '</p:sp>'
    ) % (sp_id, escape(name), x, y, cx, cy, anchor, body)


def slide_xml(title_para, body_paras, title_anchor="t"):
    """To'liq slayd XML ni yig'adi: sarlavha qutisi + tana qutisi."""
    title_sp = textbox_sp(2, "Sarlavha", 685800, 365760,
                          SLIDE_W - 2 * 685800, 1143000,
                          title_para, anchor="t")
    body_sp = textbox_sp(3, "Matn", 685800, 1676400,
                        SLIDE_W - 2 * 685800, SLIDE_H - 1676400 - 365760,
                        body_paras, anchor=title_anchor)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sld %s>'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr>'
        '<p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/>'
        '<p:nvPr/>'
        '</p:nvGrpSpPr>'
        '<p:grpSpPr>'
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>'
        '</p:grpSpPr>'
        '%s%s'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:overrideClrMapping '
        'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" '
        'accent5="accent5" accent6="accent6" hlink="hlink" '
        'folHlink="folHlink"/></p:clrMapOvr>'
        '</p:sld>'
    ) % (NSMAP_SP, title_sp, body_sp)


def unsplash_line(url):
    return "Rasm (Unsplash): " + url


# ---------------------------------------------------------------------------
# 10 slaydning mazmuni
# ---------------------------------------------------------------------------
DARK = "0B2545"    # to'q ko'k
ACCENT = "0E6BA8"  # ko'k
CYAN = "0A9396"    # moviy-yashil


def build_slides():
    slides = []

    # --- 1-slayd: Titul varag'i ---
    title = [para_xml("Kompyuterlar nima uchun qiziydi va qanday soviydi?",
                      SZ_MAIN_TITLE, bold=True, color=DARK, align="ctr")]
    body = [
        para_xml("Protsessor ishlayotganda issiqlik ajralishi fizikasi "
                 "va kulerlar (ventilyatorlar) yordamida sovutish",
                 SZ_HEADING, color=ACCENT, align="ctr"),
        para_xml("", SZ_BODY),
        para_xml("Taqdimotchi: Toshbo\u02bbriyeva Marjona (242_26 - guruh)",
                 SZ_KEY, bold=True, align="ctr"),
        para_xml("Jizzax davlat pedagogika universiteti - Axborot tizimlari fani",
                 SZ_KEY, align="ctr"),
        para_xml("", SZ_BODY),
        para_xml(unsplash_line("https://unsplash.com/photos/a-close-up-of-a-computer-processor-fPkvU7RDmCo"),
                 SZ_SMALL, color=ACCENT, align="ctr"),
    ]
    slides.append(slide_xml(title, body, title_anchor="ctr"))

    # --- 2-slayd: Reja ---
    title = [para_xml("Reja", SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("1-qism. Protsessor nima uchun qiziydi?", SZ_KEY, bold=True,
                 color=ACCENT, bullet=True),
        para_xml("Tranzistorlar ishlashi, Joul-Lents qonuni va elektr "
                 "energiyasining issiqlikka aylanishi.", SZ_BODY, bullet=True),
        para_xml("2-qism. Issiqlikning oqibatlari.", SZ_KEY, bold=True,
                 color=ACCENT, bullet=True),
        para_xml("Kompyuter qiziganda nima yuz beradi: trottling (throttling) "
                 "va uning ishlash tezligiga ta'siri.", SZ_BODY, bullet=True),
        para_xml("3-qism. Sovutish fizikasi va mexanizmlari.", SZ_KEY, bold=True,
                 color=ACCENT, bullet=True),
        para_xml("Radiatorlar, termopastaning vazifasi, kulerlar yordamida "
                 "havoni so'rib olish va haydash.", SZ_BODY, bullet=True),
        para_xml("4-qism. Xulosa.", SZ_KEY, bold=True, color=ACCENT, bullet=True),
        para_xml("Issiqlik va sovutish balansining barqaror ishlashdagi "
                 "ahamiyati.", SZ_BODY, bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/green-and-black-computer-motherboard-FO7JIlwjOtU"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 3-slayd: 1-qism, tranzistorlar ---
    title = [para_xml("1-qism. Protsessor nima uchun qiziydi?",
                      SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("Tranzistorlar - protsessorning asosiy elementlari",
                 SZ_HEADING, bold=True, color=ACCENT),
        para_xml("Zamonaviy protsessor milliardlab mayda tranzistorlardan "
                 "iborat bo'lib, ular elektron kalitlar vazifasini bajaradi.",
                 SZ_BODY, bullet=True),
        para_xml("Har bir tranzistor soniyada millionlab marta ochilib-yopiladi "
                 "(0 va 1 holatlar), bu esa elektr tokining uzluksiz oqishini "
                 "talab qiladi.", SZ_BODY, bullet=True),
        para_xml("Tranzistorlarning ochilib-yopilishida yarimo'tkazgich "
                 "materialning qarshiligi tufayli elektr energiyasining bir "
                 "qismi issiqlikka aylanadi.", SZ_BODY, bullet=True),
        para_xml("Tranzistorlar soni va chastota (GHz) qancha yuqori bo'lsa, "
                 "ajralib chiqadigan issiqlik ham shuncha ko'p bo'ladi.",
                 SZ_BODY, bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/macro-photography-of-black-circuit-board-FO7JIlwjOtU"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 4-slayd: 1-qism, Joul-Lents ---
    title = [para_xml("1-qism. Joul-Lents qonuni va energiya aylanishi",
                      SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("Elektr energiyasining issiqlikka aylanish fizikasi",
                 SZ_HEADING, bold=True, color=ACCENT),
        para_xml("Joul-Lents qonuni: Q = I\u00b2 \u00b7 R \u00b7 t", SZ_KEY,
                 bold=True, color=CYAN, bullet=True),
        para_xml("Bu yerda Q - ajralib chiqqan issiqlik miqdori, I - tok "
                 "kuchi, R - o'tkazgich qarshiligi, t - vaqt.", SZ_BODY,
                 bullet=True),
        para_xml("O'tkazgichdan tok o'tganda zaryadlangan zarralar atom "
                 "panjarasi bilan to'qnashadi va energiyasining bir qismini "
                 "issiqlik ko'rinishida yo'qotadi.", SZ_BODY, bullet=True),
        para_xml("Shu sababli protsessor qanchalik ko'p hisob-kitob bajarsa, "
                 "tok kuchi (I) ortadi va ajralib chiqadigan issiqlik "
                 "kvadratik ravishda ko'payadi.", SZ_BODY, bullet=True),
        para_xml("Ajralgan issiqlik vaqtida chiqarilmasa, kristall harorati "
                 "xavfli darajaga ko'tariladi.", SZ_BODY, bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/a-close-up-of-a-blue-and-black-electronic-device-ZV_64LdGoao"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 5-slayd: 2-qism, oqibatlar ---
    title = [para_xml("2-qism. Issiqlikning oqibatlari",
                      SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("Kompyuter qiziganda nima yuz beradi?",
                 SZ_HEADING, bold=True, color=ACCENT),
        para_xml("Harorat belgilangan chegaradan oshsa, tizim barqarorligi "
                 "buziladi va xatoliklar yuzaga keladi.", SZ_BODY, bullet=True),
        para_xml("Ortiqcha issiqlik tranzistorlar va boshqa elektron "
                 "komponentlarning tez eskirishiga olib keladi.", SZ_BODY,
                 bullet=True),
        para_xml("Haddan tashqari qizish natijasida kompyuter o'z-o'zidan "
                 "o'chib qolishi yoki qotib qolishi mumkin.", SZ_BODY,
                 bullet=True),
        para_xml("Shuning uchun protsessorda harorat datchiklari bo'lib, "
                 "ular haroratni doimiy nazorat qilib turadi.", SZ_BODY,
                 bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/red-and-black-abstract-illustration-mcSDtbWXUZU"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 6-slayd: 2-qism, throttling ---
    title = [para_xml("2-qism. Trottling (Throttling) jarayoni",
                      SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("Himoya mexanizmi va uning tezlikka ta'siri",
                 SZ_HEADING, bold=True, color=ACCENT),
        para_xml("Trottling - harorat xavfli chegaraga yetganda protsessor "
                 "o'z ish chastotasini (tezligini) atayin pasaytirishi.",
                 SZ_BODY, bullet=True),
        para_xml("Chastotaning pasayishi bilan tok kuchi kamayadi, natijada "
                 "Joul-Lents qonuniga ko'ra ajralib chiqadigan issiqlik ham "
                 "kamayadi.", SZ_BODY, bullet=True),
        para_xml("Bu mexanizm protsessorni kuyib ketishdan saqlaydi, biroq "
                 "kompyuterning umumiy ishlash tezligi sezilarli pasayadi.",
                 SZ_BODY, bullet=True),
        para_xml("Natijada o'yinlar va og'ir dasturlarda \"sekinlashish\" "
                 "(lag) seziladi. Yaxshi sovutish trottlingni oldini oladi.",
                 SZ_BODY, bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/a-computer-processor-with-the-letters-amd-on-it-Dawk-Nx-R3o"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 7-slayd: 3-qism, radiator + termopasta ---
    title = [para_xml("3-qism. Sovutish fizikasi: radiator va termopasta",
                      SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("Issiqlikni protsessordan uzatish",
                 SZ_HEADING, bold=True, color=ACCENT),
        para_xml("Radiator (issiqlik tarqatgich) - ko'p sonli metall "
                 "qanotlardan iborat bo'lib, issiqlik almashinuv yuzasini "
                 "kengaytiradi.", SZ_BODY, bullet=True),
        para_xml("Mis va alyuminiy yuqori issiqlik o'tkazuvchanligi tufayli "
                 "radiatorlar uchun asosiy material hisoblanadi.", SZ_BODY,
                 bullet=True),
        para_xml("Termopasta protsessor va radiator orasidagi mikroskopik "
                 "havo bo'shliqlarini to'ldiradi; havo issiqlikni yomon "
                 "o'tkazadi, pasta esa aloqani yaxshilaydi.", SZ_BODY,
                 bullet=True),
        para_xml("Shunday qilib issiqlik issiqlik o'tkazuvchanlik orqali "
                 "kristalldan radiatorga samarali uzatiladi.", SZ_BODY,
                 bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/a-cpu-cooler-with-copper-heat-pipes-and-fins-9SoCnyQmkzI"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 8-slayd: 3-qism, kulerlar ---
    title = [para_xml("3-qism. Kulerlar (ventilyatorlar) va havo oqimi",
                      SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("Havoni so'rib olish va haydash jarayoni",
                 SZ_HEADING, bold=True, color=ACCENT),
        para_xml("Kuler (ventilyator) radiator qanotlari orqali havo oqimini "
                 "hosil qiladi va qizigan havoni korpusdan tashqariga haydaydi.",
                 SZ_BODY, bullet=True),
        para_xml("Bu jarayon konveksiya deb ataladi: harakatlanayotgan havo "
                 "issiqlikni o'zi bilan olib ketadi.", SZ_BODY, bullet=True),
        para_xml("Korpusda havo balansi muhim: oldingi ventilyatorlar salqin "
                 "havoni so'rib oladi, orqadagilar esa issiq havoni haydaydi.",
                 SZ_BODY, bullet=True),
        para_xml("Suvli (suyuqlikli) sovutish tizimlarida issiqlik suyuqlik "
                 "orqali radiatorga olib boriladi va u yerda ventilyatorlar "
                 "bilan sovutiladi.", SZ_BODY, bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/black-computer-cooling-fan-RLy1pDgoYGw"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 9-slayd: 4-qism, xulosa ---
    title = [para_xml("4-qism. Xulosa", SZ_SLIDE_TITLE, bold=True, color=DARK)]
    body = [
        para_xml("Issiqlik va sovutish balansining ahamiyati",
                 SZ_HEADING, bold=True, color=ACCENT),
        para_xml("Protsessordagi issiqlik elektr energiyasining issiqlikka "
                 "aylanishi (Joul-Lents qonuni) tufayli muqarrar ravishda "
                 "ajralib chiqadi.", SZ_BODY, bullet=True),
        para_xml("Ajralgan issiqlik va chiqarilgan issiqlik o'rtasidagi "
                 "muvozanat tizim barqaror va samarali ishlashini ta'minlaydi.",
                 SZ_BODY, bullet=True),
        para_xml("Sifatli radiator, termopasta va to'g'ri havo oqimi "
                 "trottlingni oldini oladi va protsessor umrini uzaytiradi.",
                 SZ_BODY, bullet=True),
        para_xml("Demak, kompyuterning yuqori unumdorligi to'g'ri tashkil "
                 "etilgan sovutish tizimiga bevosita bog'liqdir.", SZ_BODY,
                 bullet=True),
        para_xml(unsplash_line("https://unsplash.com/photos/a-computer-case-with-rgb-cooling-fans-mAcDFkmc5fU"),
                 SZ_SMALL, color=ACCENT),
    ]
    slides.append(slide_xml(title, body))

    # --- 10-slayd: Yakunlovchi ---
    title = [para_xml("E'tiboringiz uchun rahmat!", SZ_SLIDE_TITLE, bold=True,
                      color=DARK, align="ctr")]
    body = [
        para_xml("Savollaringiz uchun maydon ochiq.", SZ_HEADING,
                 color=ACCENT, align="ctr"),
        para_xml("", SZ_BODY),
        para_xml("Toshbo\u02bbriyeva Marjona (242_26 - guruh)", SZ_KEY,
                 bold=True, align="ctr"),
        para_xml("Jizzax davlat pedagogika universiteti - Axborot tizimlari fani",
                 SZ_BODY, align="ctr"),
        para_xml("", SZ_BODY),
        para_xml(unsplash_line("https://unsplash.com/photos/a-glowing-computer-cpu-on-a-dark-background-mcSDtbWXUZU"),
                 SZ_SMALL, color=ACCENT, align="ctr"),
    ]
    slides.append(slide_xml(title, body, title_anchor="ctr"))

    return slides


# ---------------------------------------------------------------------------
# OPC paketining qolgan qismlari
# ---------------------------------------------------------------------------
def content_types_xml(n_slides):
    overrides = [
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/presProps.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'presentationml.presProps+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'theme+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'presentationml.slideLayout+xml"/>',
    ]
    for i in range(1, n_slides + 1):
        overrides.append(
            '<Override PartName="/ppt/slides/slide%d.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'presentationml.slide+xml"/>' % i)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types">'
        '<Default Extension="rels" ContentType="application/vnd.'
        'openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '%s'
        '</Types>'
    ) % ''.join(overrides)


def root_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
        '2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/officeDocument" '
        'Target="ppt/presentation.xml"/>'
        '</Relationships>'
    )


def presentation_xml(n_slides):
    # relationship ID lari: rId1 = master, rId2 = presProps, rId3 = theme,
    # rId4..rId(3+n) = slaydlar
    sld_ids = []
    for i in range(1, n_slides + 1):
        rid = 3 + i
        sld_ids.append('<p:sldId id="%d" r:id="rId%d"/>' % (255 + i, rid))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:presentation %s saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/>'
        '</p:sldMasterIdLst>'
        '<p:sldIdLst>%s</p:sldIdLst>'
        '<p:sldSz cx="%d" cy="%d"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '</p:presentation>'
    ) % (NSMAP_SP, ''.join(sld_ids), SLIDE_W, SLIDE_H)


def presentation_rels_xml(n_slides):
    rels = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/slideMaster" '
        'Target="slideMasters/slideMaster1.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/presProps" '
        'Target="presProps.xml"/>',
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/theme" '
        'Target="theme/theme1.xml"/>',
    ]
    for i in range(1, n_slides + 1):
        rid = 3 + i
        rels.append(
            '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/slide" '
            'Target="slides/slide%d.xml"/>' % (rid, i))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
        '2006/relationships">%s</Relationships>'
    ) % ''.join(rels)


def pres_props_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/'
        'drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships" xmlns:p="http://schemas.'
        'openxmlformats.org/presentationml/2006/main"/>'
    )


def theme_xml():
    # Kompyuter-texnologiya / fizika palitrasi: to'q ko'k va moviy aksentlar
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
        'main" name="Kompyuter texnologiyalari">'
        '<a:themeElements>'
        '<a:clrScheme name="Kompyuter">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="0B2545"/></a:dk2>'
        '<a:lt2><a:srgbClr val="EEF4FB"/></a:lt2>'
        '<a:accent1><a:srgbClr val="0E6BA8"/></a:accent1>'
        '<a:accent2><a:srgbClr val="0A9396"/></a:accent2>'
        '<a:accent3><a:srgbClr val="13293D"/></a:accent3>'
        '<a:accent4><a:srgbClr val="2A9D8F"/></a:accent4>'
        '<a:accent5><a:srgbClr val="48CAE4"/></a:accent5>'
        '<a:accent6><a:srgbClr val="005377"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0E6BA8"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="005377"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="Times New Roman">'
        '<a:majorFont>'
        '<a:latin typeface="Times New Roman"/><a:ea typeface=""/>'
        '<a:cs typeface=""/>'
        '</a:majorFont>'
        '<a:minorFont>'
        '<a:latin typeface="Times New Roman"/><a:ea typeface=""/>'
        '<a:cs typeface=""/>'
        '</a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="Office">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/>'
        '<a:satMod val="300000"/></a:schemeClr></a:gs>'
        '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/>'
        '<a:satMod val="300000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/>'
        '<a:satMod val="350000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/>'
        '<a:satMod val="130000"/></a:schemeClr></a:gs>'
        '<a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/>'
        '<a:satMod val="130000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/>'
        '<a:satMod val="135000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/>'
        '<a:satMod val="105000"/></a:schemeClr></a:solidFill>'
        '<a:prstDash val="solid"/></a:ln>'
        '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/></a:ln>'
        '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/>'
        '<a:satMod val="170000"/></a:schemeClr></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/>'
        '<a:satMod val="150000"/><a:shade val="98000"/>'
        '<a:lumMod val="102000"/></a:schemeClr></a:gs>'
        '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/>'
        '<a:satMod val="130000"/><a:shade val="90000"/>'
        '<a:lumMod val="103000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/>'
        '<a:satMod val="120000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
        '</a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '<a:objectDefaults/>'
        '<a:extraClrSchemeLst/>'
        '</a:theme>'
    )


def slide_master_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sldMaster %s>'
        '<p:cSld>'
        '<p:bg><p:bgRef idx="1001"><a:schemeClr val="lt1"/></p:bgRef></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree>'
        '</p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" '
        'accent5="accent5" accent6="accent6" hlink="hlink" '
        'folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/>'
        '</p:sldLayoutIdLst>'
        '<p:txStyles>'
        '<p:titleStyle>'
        '<a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" '
        'latinLnBrk="0" hangingPunct="1"><a:defRPr sz="4400" kern="1200">'
        '<a:solidFill><a:schemeClr val="dk2"/></a:solidFill>'
        '<a:latin typeface="Times New Roman"/><a:ea typeface="+mj-ea"/>'
        '<a:cs typeface="Times New Roman"/></a:defRPr></a:lvl1pPr>'
        '</p:titleStyle>'
        '<p:bodyStyle>'
        '<a:lvl1pPr marL="285750" indent="-285750"><a:defRPr sz="1800">'
        '<a:latin typeface="Times New Roman"/>'
        '<a:cs typeface="Times New Roman"/></a:defRPr></a:lvl1pPr>'
        '</p:bodyStyle>'
        '<p:otherStyle>'
        '<a:defPPr><a:defRPr lang="uz-UZ"/></a:defPPr>'
        '</p:otherStyle>'
        '</p:txStyles>'
        '</p:sldMaster>'
    ) % NSMAP_SP


def slide_master_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
        '2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/theme" '
        'Target="../theme/theme1.xml"/>'
        '</Relationships>'
    )


def slide_layout_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sldLayout %s type="blank" preserve="1">'
        '<p:cSld name="Bo\'sh">'
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
        '</p:sldLayout>'
    ) % NSMAP_SP


def slide_layout_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
        '2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/slideMaster" '
        'Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>'
    )


def slide_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
        '2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )


# ---------------------------------------------------------------------------
# Paketni yozish
# ---------------------------------------------------------------------------
def main():
    slides = build_slides()
    n = len(slides)
    assert n == 10, "Slaydlar soni aniq 10 ta bo'lishi kerak, hozir: %d" % n

    with zipfile.ZipFile(OUTPUT_PATH, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml(n))
        z.writestr("_rels/.rels", root_rels_xml())
        z.writestr("ppt/presentation.xml", presentation_xml(n))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(n))
        z.writestr("ppt/presProps.xml", pres_props_xml())
        z.writestr("ppt/theme/theme1.xml", theme_xml())
        z.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",
                   slide_master_rels_xml())
        z.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",
                   slide_layout_rels_xml())
        for i, s in enumerate(slides, start=1):
            z.writestr("ppt/slides/slide%d.xml" % i, s)
            z.writestr("ppt/slides/_rels/slide%d.xml.rels" % i,
                       slide_rels_xml())

    print("Taqdimot yaratildi: %s" % OUTPUT_PATH)
    print("Slaydlar soni: %d" % n)


if __name__ == "__main__":
    main()
