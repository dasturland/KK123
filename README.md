# Taqdimot: Fleshkalar (USB) qanday qilib ma'lumot saqlaydi?

Ushbu repozitoriyada "Axborot tizimlari" fani uchun tayyorlangan akademik
taqdimot va uni dasturiy yo'l bilan generatsiya qiluvchi skript joylashgan.

- **Oliy ta'lim muassasasi:** Jizzax davlat pedagogika universiteti
- **Fan:** Axborot tizimlari
- **Taqdimotchi:** Toshboʻriyeva Marjona (242_26-guruh)
- **Mavzu:** Fleshkalar (USB) ichida harakatlanuvchi qismlar bo'lmasa ham,
  ma'lumot elektr zaryadlari ko'rinishida qanday saqlanishi.

## Fayllar

| Fayl | Tavsif |
|------|--------|
| `Fleshkalar_USB_malumot_saqlashi.pptx` | Tayyor taqdimot (qat'iy 10 ta slayd, har birida joylangan diagramma). |
| `generate_presentation.py` | Taqdimotni va uning ichidagi barcha diagrammalarni qaytadan yaratuvchi skript. |
| `verify_presentation.py` | Barcha talablarni tekshiruvchi skript. |

## Taqdimot tuzilmasi (10 slayd)

1. Titul varag'i (+ dekorativ fleshka banneri)
2. Reja (3 asosiy qism + xulosa) (+ banner)
3. 1-qism: Fleshkaning ishlash prinsipi (Solid-State) — *fleshka platasi anatomiyasi diagrammasi*
4. 1-qism: NAND xotira katakchalari tuzilishi — *word line × bit line matritsasi*
5. 2-qism: Suzuvchi zulfin (Floating Gate) tranzistori — *tranzistor kesimi*
6. 2-qism: Zaryad qanday qilib ma'lumotga aylanadi — *SLC/MLC Vth taqsimoti grafiklari*
7. 3-qism: Yozish, o'qish va o'chirish jarayonlari (0 va 1) — *uch panelli sxema*
8. 3-qism: Minglab rasm elektron axborot sifatida — *rasm → piksel → bayt → katakcha zanjiri*
9. 4-qism: Xulosa (afzalliklar, chidamlilik, kelajak) — *3D NAND va sig'im o'sishi grafigi*
10. Yakunlovchi slayd (E'tiboringiz uchun rahmat) (+ banner)

## Texnik talablar (bajarilgan)

- Qat'iy **10 ta slayd**.
- Faqat **Times New Roman** shrifti.
- Eng kichik shrift **12 pt** (sarlavha va muhim nuqtalar ierarxik kattaroq).
- **Har bir slaydda joylangan (embedded) PNG diagramma** — jami 10 ta rasm
  `ppt/media/` ichida saqlanadi.
- Har bir slaydda mavzuga mos **Unsplash URL** manzili ko'rinadigan izoh sifatida.
- Raqamli texnologiya / mikroelektronika uslubidagi rang palitrasi.

## Qaytadan yaratish va tekshirish

```bash
python3 generate_presentation.py   # .pptx faylni (rasmlari bilan) yaratadi
python3 verify_presentation.py     # barcha talablarni tekshiradi
```

## Eslatma: nega `python-pptx` va `Pillow` ishlatilmagan?

`.pptx` fayli — bu OpenXML (OOXML) spetsifikatsiyasiga mos XML qismlardan
iborat ZIP arxividir, PNG esa `zlib` bilan kodlanadigan oddiy formatdir.
Ushbu muhitda PyPI ga tarmoq orqali ulanish cheklangan edi, shu bois:

- taqdimot fayli Python standart kutubxonasi (`zipfile`) yordamida,
  to'liq spetsifikatsiyaga rioya qilgan holda qo'lda yaratiladi;
- barcha diagrammalar ham sof Python'dagi mini-grafika dvijok
  (`Canvas` + 5×7 bitmap shrift + PNG koder) bilan pikselma-piksel chiziladi
  — tashqi rasm yuklab olish talab qilinmaydi.

Natijaviy fayl Microsoft PowerPoint, LibreOffice Impress va Google Slides'da
ochiladi.
