#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Kompyuterlar_qizishi_va_sovishi.pptx uchun o'z-o'zini tekshiruvchi skript.

Ushbu skript FAQAT Python standart kutubxonasidan foydalanadi (zipfile, re,
sys). python-pptx yoki lxml talab qilinmaydi (paket registri bloklangan).

U yaratilgan .pptx faylni ZIP sifatida ochadi va qat'iy talablarni TEKSHIRADI.
Agar biror shart buzilsa, skript xato xabari bilan NOL BO'LMAGAN kod qaytaradi.
Barcha shartlar bajarilsa, PASS xulosasi chop etiladi va 0 qaytariladi.

Tekshiriladigan shartlar:
  * Aniq 10 ta slayd (ppt/slides/slideN.xml, N=1..10).
  * Har bir <a:r> run Times New Roman shriftidan foydalanadi (<a:latin>).
  * Barcha sz="NNNN" qiymatlari >= 1200 (12pt).
  * 1-slaydda: 'Toshbo\u02bbriyeva Marjona', '242_26',
    'Jizzax davlat pedagogika universiteti' mavjud.
  * Har bir slaydda 'unsplash' (katta-kichik harfga bog'liq emas) mavjud
    va barcha 10 ta Unsplash rasm identifikatori (ID) noyob.

This verifier is written with the Python standard library only (no python-pptx /
lxml, since the package registry is blocked). It genuinely fails (non-zero exit)
when any invariant is violated.
"""

import os
import re
import sys
import zipfile

PPTX_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "Kompyuterlar_qizishi_va_sovishi.pptx")

MIN_SZ = 1200  # 12pt OpenXML yuzdan bir punkt birligida
EXPECTED_SLIDES = 10
TIMES = "Times New Roman"

SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
# Bitta <a:r> ... </a:r> run bloki (eng kichik mos kelish).
RUN_RE = re.compile(r"<a:r>.*?</a:r>", re.DOTALL)
SZ_RE = re.compile(r'sz="(\d+)"')
# Unsplash havolasidagi rasm identifikatori: .../photos/<slug>-<ID>
# Unsplash ID doimo yo'l qismining oxirgi 11 ta belgisidir
# ([A-Za-z0-9_-], URL oxirida, matn tirnoq/bo'sh joy bilan tugaydi).
UNSPLASH_ID_RE = re.compile(
    r"unsplash\.com/photos/[A-Za-z0-9_-]+?([A-Za-z0-9_-]{11})(?=[^A-Za-z0-9_-]|$)")

REQUIRED_SLIDE1 = [
    "Toshbo\u02bbriyeva Marjona",  # U+02BB okina
    "242_26",
    "Jizzax davlat pedagogika universiteti",
]


class VerificationError(Exception):
    """Qat'iy shart buzilganda ko'tariladi."""


def fail(msg):
    raise VerificationError(msg)


def load_slides(zf):
    """ppt/slides/slideN.xml qismlarini {N: xml_matn} ko'rinishida qaytaradi."""
    slides = {}
    for name in zf.namelist():
        m = SLIDE_RE.match(name)
        if m:
            slides[int(m.group(1))] = zf.read(name).decode("utf-8")
    return slides


def check_slide_count(slides):
    nums = sorted(slides)
    if len(nums) != EXPECTED_SLIDES:
        fail("Slaydlar soni noto'g'ri: %d ta topildi, %d ta kutilgan edi "
             "(%s)" % (len(nums), EXPECTED_SLIDES, nums))
    expected = list(range(1, EXPECTED_SLIDES + 1))
    if nums != expected:
        fail("Slayd raqamlari ketma-ketligi noto'g'ri: %s, kutilgan: %s"
             % (nums, expected))


def check_fonts_and_sizes(slides):
    """Har bir run Times New Roman + har bir sz >= 1200 ekanini tekshiradi."""
    min_sz = None
    max_sz = None
    total_runs = 0

    for n in sorted(slides):
        xml = slides[n]

        # 1) Har bir <a:r> run ichida Times New Roman latin shrifti bo'lsin.
        runs = RUN_RE.findall(xml)
        if not runs:
            fail("slide%d.xml da hech qanday <a:r> run topilmadi" % n)
        for idx, run in enumerate(runs, 1):
            total_runs += 1
            if '<a:latin typeface="%s"' % TIMES not in run:
                snippet = run[:120].replace("\n", " ")
                fail("slide%d.xml dagi %d-run Times New Roman latin "
                     "shriftidan foydalanmaydi: %s" % (n, idx, snippet))

        # 2) Butun slayd XML bo'ylab har bir sz="NNNN" >= 1200 bo'lsin.
        sizes = [int(s) for s in SZ_RE.findall(xml)]
        if not sizes:
            fail("slide%d.xml da hech qanday sz atributi topilmadi" % n)
        for sz in sizes:
            if sz < MIN_SZ:
                fail("slide%d.xml da shrift o'lchami minimumdan kichik: "
                     "sz=%d (minimum %d = 12pt)" % (n, sz, MIN_SZ))
            if min_sz is None or sz < min_sz:
                min_sz = sz
            if max_sz is None or sz > max_sz:
                max_sz = sz

    return min_sz, max_sz, total_runs


def check_slide1_strings(slides):
    xml = slides[1]
    for needle in REQUIRED_SLIDE1:
        if needle not in xml:
            fail("1-slaydda majburiy matn topilmadi: %r" % needle)


def check_unsplash(slides):
    """Har bir slaydda Unsplash havolasi bor va 10 ta rasm ID si noyob."""
    ids = {}
    for n in sorted(slides):
        xml = slides[n]
        if "unsplash" not in xml.lower():
            fail("slide%d.xml da Unsplash havolasi topilmadi" % n)
        found = UNSPLASH_ID_RE.findall(xml)
        if not found:
            fail("slide%d.xml dagi Unsplash havolasida rasm ID si "
                 "topilmadi" % n)
        for pid in found:
            if pid in ids:
                fail("Unsplash rasm ID si takrorlanmoqda: %r (slide%d va "
                     "slide%d). Har bir ID bitta rasmga mos kelishi kerak."
                     % (pid, ids[pid], n))
            ids[pid] = n
    return ids


def main():
    if not os.path.exists(PPTX_PATH):
        print("XATO: fayl topilmadi: %s" % PPTX_PATH, file=sys.stderr)
        print("Avval 'python3 generate_presentation.py' ni ishga tushiring.",
              file=sys.stderr)
        return 2

    try:
        with zipfile.ZipFile(PPTX_PATH) as zf:
            bad = zf.testzip()
            if bad is not None:
                fail("ZIP paketi buzilgan qism: %s" % bad)
            slides = load_slides(zf)
            check_slide_count(slides)
            min_sz, max_sz, total_runs = check_fonts_and_sizes(slides)
            check_slide1_strings(slides)
            unsplash_ids = check_unsplash(slides)
    except VerificationError as exc:
        print("TEKSHIRUV MUVAFFAQIYATSIZ (FAIL): %s" % exc, file=sys.stderr)
        return 1
    except zipfile.BadZipFile as exc:
        print("TEKSHIRUV MUVAFFAQIYATSIZ (FAIL): .pptx yaroqli ZIP emas: %s"
              % exc, file=sys.stderr)
        return 1

    print("PASS: %d ta slayd, %d ta run - barchasi Times New Roman; "
          "shrift o'lchami min=%d (%.0fpt), max=%d (%.0fpt), minimum >= %d "
          "(12pt); 1-slayd muallif/universitet matnlari mavjud; har bir "
          "slaydda Unsplash havolasi bor va %d ta rasm ID si noyob."
          % (len(slides), total_runs, min_sz, min_sz / 100.0,
             max_sz, max_sz / 100.0, MIN_SZ, len(unsplash_ids)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
