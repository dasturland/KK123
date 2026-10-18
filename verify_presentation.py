#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Yaratilgan .pptx faylni tekshiruvchi skript (standart kutubxona).

Tekshiradi:
  (a) ZIP/OOXML paketi yaroqli va barcha XML qismlar well-formed;
  (b) Qat'iy 10 ta slayd (presentation.xml va [Content_Types].xml bo'yicha);
  (c) Faqat "Times New Roman" shrifti ishlatilgan;
  (d) Hech bir shrift 12 pt dan kichik emas;
  (e) Titul slaydda universitet / fan / taqdimotchi ma'lumotlari aynan mavjud;
  (f) Har bir slaydda Unsplash URL manzili bor;
  (g) Har bir slaydda JOYLANGAN PNG rasm bor (<p:pic> + media + rels),
      PNG imzolari yaroqli va [Content_Types].xml png ni qamrab olgan.
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

PNG_SIG = b"\x89PNG\r\n\x1a\n"


def main(path):
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
        pic_per_slide = {}
        slide_texts = {}
        for nm in sorted(slide_files, key=lambda s: int(re.search(r"\d+", s.split("/")[-1]).group())):
            xml = z.read(nm).decode("utf-8")
            idx = int(re.search(r"slide(\d+)\.xml", nm).group(1))
            for tf in re.findall(r'typeface="([^"]+)"', xml):
                if tf and tf != "Times New Roman":
                    non_tnr.add(tf)
            for sz in re.findall(r'sz="(\d+)"', xml):
                all_sizes.append(int(sz))
            unsplash_per_slide[idx] = "unsplash.com" in xml
            pic_per_slide[idx] = "<p:pic>" in xml and 'r:embed="rId2"' in xml
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

        # (g) joylangan rasmlar
        no_pic = [i for i, v in pic_per_slide.items() if not v]
        assert not no_pic, f"<p:pic> rasm yoʻq slaydlar: {no_pic}"
        media = [n for n in names if re.match(r"ppt/media/image\d+\.png$", n)]
        print(f"(g) har slaydda <p:pic> bor; media fayllari: {len(media)} ta")
        assert len(media) == 10, f"Media PNG soni 10 emas: {len(media)}"
        for m in media:
            data = z.read(m)
            assert data.startswith(PNG_SIG), f"PNG imzosi buzuk: {m}"
            # IHDR dan o'lchamni o'qiymiz
            w, h = int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
            assert w > 100 and h > 50, f"PNG o'lchami kichik: {m} {w}x{h}"
        ct = z.read("[Content_Types].xml").decode("utf-8")
        assert 'Extension="png"' in ct, "[Content_Types].xml da png yoʻq"
        # har bir slayd rels faylida rasm munosabati bor
        for i in range(1, 11):
            rels = z.read(f"ppt/slides/_rels/slide{i}.xml.rels").decode("utf-8")
            assert "relationships/image" in rels, f"slide{i} rels da image yoʻq"
        print("(g) barcha PNG lar yaroqli, rels va content-types to'g'ri")

    print("\nBARCHA TEKSHIRUVLAR MUVAFFAQIYATLI OʻTDI ✓")
    return True


if __name__ == "__main__":
    p = sys.argv[1] if len(sys.argv) > 1 else "Fleshkalar_USB_malumot_saqlashi.pptx"
    sys.exit(0 if main(p) else 1)
