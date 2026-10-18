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
| `Fleshkalar_USB_malumot_saqlashi.pptx` | Tayyor taqdimot (qat'iy 10 ta slayd). |
| `generate_presentation.py` | Taqdimotni qaytadan yaratuvchi skript. |
| `verify_presentation.py` | Barcha talablarni tekshiruvchi skript. |

## Taqdimot tuzilmasi (10 slayd)

1. Titul varag'i
2. Reja (3 asosiy qism + xulosa)
3. 1-qism: Fleshkaning ishlash prinsipi (Solid-State)
4. 1-qism: NAND xotira katakchalari tuzilishi
5. 2-qism: Suzuvchi zulfin (Floating Gate) tranzistori
6. 2-qism: Zaryad qanday qilib ma'lumotga aylanadi
7. 3-qism: Yozish, o'qish va o'chirish jarayonlari (0 va 1)
8. 3-qism: Minglab rasm elektron axborot sifatida
9. 4-qism: Xulosa (afzalliklar, chidamlilik, kelajak)
10. Yakunlovchi slayd (E'tiboringiz uchun rahmat)

## Texnik talablar (bajarilgan)

- Qat'iy **10 ta slayd**.
- Faqat **Times New Roman** shrifti.
- Eng kichik shrift **12 pt** (sarlavha va muhim nuqtalar ierarxik kattaroq).
- Har bir slaydda mavzuga mos **Unsplash URL** manzili ko'rinadigan izoh sifatida.
- Raqamli texnologiya / mikroelektronika uslubidagi rang palitrasi.

## Qaytadan yaratish va tekshirish

```bash
python3 generate_presentation.py   # .pptx faylni yaratadi
python3 verify_presentation.py      # barcha talablarni tekshiradi
```

## Eslatma: nega `python-pptx` ishlatilmagan?

`.pptx` fayli — bu OpenXML (OOXML) spetsifikatsiyasiga mos XML qismlardan
iborat ZIP arxividir. Ushbu muhitda PyPI (python-pptx paketi) ga tarmoq
orqali ulanish cheklangan edi, shu bois fayl Python standart kutubxonasi
(`zipfile`) yordamida, to'liq spetsifikatsiyaga rioya qilgan holda qo'lda
yaratilgan. Natijaviy fayl Microsoft PowerPoint, LibreOffice Impress va
Google Slides'da ochiladi.
