#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Yaratilgan .pptx faylni tekshiruvchi skript (standart kutubxona).

Tekshiradi:
  (a) ZIP/OOXML paketi yaroqli va barcha XML qismlar well-formed;
  (b) Qat'iy 10 ta slayd (presentation.xml va [Content_Types].xml bo'yicha);
  (c) Faqat "Times New Roman" shrifti ishlatilgan;
  (d) Hech bir shrift 12 pt dan kichik emas;
  (e) Titul slaydda universitet / fan / taqdimotchi ma'lumotlari aynan mavjud;
  (f) Har bir slaydda Unsplash URL manzili bor.
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def main(path):
    ok = True
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        assert bad is None, f"Buzuq ZIP qismi: {bad}"
        names = z.namelist()

        # (a) barcha xml qismlar well-formed
        for nm in names:
            if nm.endswith(".xml") or nm.endswith(".rels"):
                ET.fromstring(z.read(nm))

        # (b) slaydlar soni
        slide_files = [n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)]
        count = len(slide_files)
        pres = z.read("ppt/presentation.xml").decode("utf-8")
        sldid_count = len(re.findall(r"<p:sldId ", pres))

        print(f"(b) slayd fayllari: {count}, presentation.xml sldId: {sldid_count}")
        assert count == 10, f"Slaydlar soni 10 emas: {count}"
        assert sldid_count == 10, f"sldId soni 10 emas: {sldid_count}"

        # (c) + (d) shrift va o'lcham
        all_sizes = []
        non_tnr = set()
        unsplash_per_slide = {}
        slide_texts = {}
        for nm in sorted(slide_files, key=lambda s: int(re.search(r"\d+", s.split("/")[-1]).group())):
            xml = z.read(nm).decode("utf-8")
            idx = int(re.search(r"slide(\d+)\.xml", nm).group(1))
            # typeface'lar
            for tf in re.findall(r'typeface="([^"]+)"', xml):
                if tf and tf != "Times New Roman":
                    non_tnr.add(tf)
            # sz atributlari (yuzdan birlarda)
            for sz in re.findall(r'sz="(\d+)"', xml):
                all_sizes.append(int(sz))
            # Unsplash URL
            unsplash_per_slide[idx] = "unsplash.com" in xml
            # matn
            texts = re.findall(r"<a:t>(.*?)</a:t>", xml, re.S)
            slide_texts[idx] = " ".join(texts)

        min_pt = min(all_sizes) / 100.0
        print(f"(c) 'Times New Roman'dan boshqa shriftlar: {non_tnr or 'yoʻq'}")
        print(f"(d) eng kichik shrift: {min_pt} pt (jami {len(all_sizes)} ta run)")
        assert not non_tnr, f"Boshqa shrift topildi: {non_tnr}"
        assert min_pt >= 12, f"12 pt dan kichik shrift bor: {min_pt}"

        # (e) titul slayd ma'lumotlari
        t1 = slide_texts[1]
        for needle in [
            "Jizzax davlat pedagogika universiteti",
            "Axborot tizimlari fani",
            "Toshbo\u02bbriyeva Marjona",
            "242_26",
        ]:
            assert needle in t1, f"Titul slaydda yoʻq: {needle!r}"
        print("(e) titul slayd ma'lumotlari: barchasi mavjud")

        # (f) har slaydda unsplash
        missing = [i for i, v in unsplash_per_slide.items() if not v]
        print(f"(f) Unsplash URL yoʻq slaydlar: {missing or 'yoʻq'}")
        assert not missing, f"Unsplash URL yoʻq slaydlar: {missing}"

        # 4 qismli tuzilma reja slaydida
        plan = slide_texts[2]
        for lbl in ["1-qism", "2-qism", "3-qism", "Xulosa"]:
            assert lbl in plan, f"Reja slaydida yoʻq: {lbl}"
        print("(b2) 4 qismli tuzilma (1,2,3-qism + Xulosa): reja slaydida mavjud")

    print("\nBARCHA TEKSHIRUVLAR MUVAFFAQIYATLI OʻTDI ✓")
    return ok


if __name__ == "__main__":
    p = sys.argv[1] if len(sys.argv) > 1 else "Fleshkalar_USB_malumot_saqlashi.pptx"
    sys.exit(0 if main(p) else 1)
