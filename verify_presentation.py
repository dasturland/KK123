#!/usr/bin/env python3
"""Verify Taqdimot_Barmoq_izi_skanerlari.pptx against the project's hard requirements.

This is the project's "test" command. It reopens the generated .pptx (an OOXML
ZIP package) and asserts, using only the Python standard library:

  1. Exactly 10 slides match ^ppt/slides/slide\\d+\\.xml$ (slide1..slide10).
  2. For every slide: every run-property font size (sz="NNNN") is >= 1200
     (12pt), and every text run uses the "Times New Roman" latin typeface.
  3. Each of the 10 slides contains an Unsplash URL ('unsplash.com').
  4. The title slide carries the required identity text (university, subject,
     author surname, group) and the deck shows a font hierarchy (a title
     sz >= 3000 plus smaller body sizes).

Exit status is 0 only when every assertion passes; any failure prints a FAIL
line describing the problem and exits with status 1.

Stdlib only: no network, no pip, no python-pptx / lxml.
"""

import re
import sys
import zipfile

PPTX = "Taqdimot_Barmoq_izi_skanerlari.pptx"

SLIDE_RE = re.compile(r"^ppt/slides/slide\d+\.xml$")
SLIDE_NUM_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
MIN_SZ = 1200  # 12pt in OOXML centipoints
TITLE_SZ = 3000  # a title run must be at least 30pt somewhere

# A single text run: <a:r>...<a:rPr ...>...</a:rPr>...</a:r>
RUN_RE = re.compile(r"<a:r>(.*?)</a:r>", re.DOTALL)
# Run properties element and its sz attribute.
RPR_RE = re.compile(r"<a:rPr\b[^>]*>.*?</a:rPr>|<a:rPr\b[^>]*/>", re.DOTALL)
SZ_RE = re.compile(r'\bsz="(\d+)"')
LATIN_TNR_RE = re.compile(r'<a:latin\b[^>]*typeface="Times New Roman"')

TITLE_REQUIRED = [
    "Jizzax davlat pedagogika universiteti",
    "Axborot tizimlari",
    "Toshbo",          # author surname start (apostrophe may be XML-escaped)
    "riyeva Marjona",  # author surname remainder
    "242_26",          # group
]


def fail(msg):
    print(f"FAIL: {msg}")
    sys.exit(1)


def main():
    try:
        zf = zipfile.ZipFile(PPTX)
    except FileNotFoundError:
        fail(f"{PPTX} not found. Run: python3 generate_presentation.py")
    except zipfile.BadZipFile:
        fail(f"{PPTX} is not a valid ZIP/OOXML package")

    with zf:
        names = zf.namelist()
        slide_names = sorted(
            (n for n in names if SLIDE_RE.match(n)),
            key=lambda n: int(SLIDE_NUM_RE.match(n).group(1)),
        )

        # --- Assertion 1: exactly 10 slides, slide1..slide10 ---
        if len(slide_names) != 10:
            fail(f"expected exactly 10 slides, found {len(slide_names)}: {slide_names}")
        found_nums = sorted(int(SLIDE_NUM_RE.match(n).group(1)) for n in slide_names)
        if found_nums != list(range(1, 11)):
            fail(f"slides must be numbered 1..10, found {found_nums}")

        all_sizes = []
        unsplash_slide_count = 0

        # --- Assertions 2 & 3 over every slide ---
        for name in slide_names:
            xml = zf.read(name).decode("utf-8")

            runs = RUN_RE.findall(xml)
            if not runs:
                fail(f"{name} contains no text runs")

            for run in runs:
                rpr_match = RPR_RE.search(run)
                if not rpr_match:
                    fail(f"{name} has a text run with no <a:rPr> run properties")
                rpr = rpr_match.group(0)

                sz_match = SZ_RE.search(rpr)
                if not sz_match:
                    fail(f"{name} has a run with no sz (font size) attribute")
                sz = int(sz_match.group(1))
                if sz < MIN_SZ:
                    fail(f"{name} has font size sz={sz} below minimum {MIN_SZ} (12pt)")
                all_sizes.append(sz)

                if not LATIN_TNR_RE.search(rpr):
                    fail(f"{name} has a run not using Times New Roman")

            if "unsplash.com" in xml:
                unsplash_slide_count += 1
            else:
                fail(f"{name} is missing an Unsplash URL ('unsplash.com')")

        # --- Assertion 4a: title slide identity text ---
        title_xml = zf.read("ppt/slides/slide1.xml").decode("utf-8")
        for needle in TITLE_REQUIRED:
            if needle not in title_xml:
                fail(f"title slide (slide1.xml) is missing required text: {needle!r}")

        # --- Assertion 4b: font hierarchy present ---
        if not all_sizes:
            fail("no font sizes found across slides")
        max_sz = max(all_sizes)
        min_sz = min(all_sizes)
        if max_sz < TITLE_SZ:
            fail(f"no title-sized run found (max sz={max_sz}, need >= {TITLE_SZ})")
        if not any(s < max_sz for s in all_sizes):
            fail("no body text smaller than the largest title (no hierarchy)")

        # --- PASS summary ---
        print("PASS: Taqdimot_Barmoq_izi_skanerlari.pptx meets all requirements")
        print(f"  slides                : {len(slide_names)} (slide1..slide10)")
        print(f"  font sizes (centipt)  : min={min_sz} ({min_sz/100:g}pt)  "
              f"max={max_sz} ({max_sz/100:g}pt)")
        print(f"  font / size rule      : all runs Times New Roman, all sz >= {MIN_SZ}")
        print(f"  slides with Unsplash  : {unsplash_slide_count}/{len(slide_names)}")
        print(f"  hierarchy             : title run >= {TITLE_SZ} + smaller body sizes")
        sys.exit(0)


if __name__ == "__main__":
    main()
