# Barmoq izi skanerlari — taqdimot (10 slayd)

Akademik oʻzbek tilidagi 10 slaydli PowerPoint taqdimoti: **"Barmoq izi
skanerlari qanday ishlaydi?"** — telefonni ochish uchun barmoq izini optik
yoki ultratovush toʻlqinlari orqali tanib olish mavzusida.

Taqdimot va uni tekshiruvchi skript **faqat Python standart kutubxonasidan**
(stdlib) foydalanadi. Internet yoki `pip` talab qilinmaydi.

## Loyihaning maqsadi

- `generate_barmoq_izi_presentation.py` — `.pptx` faylini (OOXML paketi) qoʻlda yozib,
  `zipfile` yordamida yigʻadi va `Taqdimot_Barmoq_izi_skanerlari.pptx` faylini
  loyiha ildizida hosil qiladi.
- `verify_barmoq_izi_presentation.py` — hosil qilingan `.pptx` ni qayta ochib, barcha
  qatʼiy talablarni tekshiradi (loyihaning "test" buyrugʻi).
- `Taqdimot_Barmoq_izi_skanerlari.pptx` — hosil qilingan taqdimot (10 slayd).

## Qatʼiy talablar (tekshiriladi)

- **Qatʼiy 10 ta slayd** (`ppt/slides/slide1.xml` .. `slide10.xml`).
- Barcha matnlar uchun **faqat "Times New Roman"** shrifti.
- Eng kichik shrift oʻlchami **12pt** (OOXML da `sz="1200"`); hech bir matn
  bundan kichik emas. Sarlavhalar kattaroq (ierarxiya mavjud, titul `sz=4000`).
- Tuzilma: **3 ta asosiy reja qismi + 1 ta xulosa**.
- Titul varagʻi (1-slayd): *Jizzax davlat pedagogika universiteti — Axborot
  tizimlari fani*; taqdimotchi *Toshboʻriyeva Marjona (242_26 - guruh)*.
- Har bir slaydda mavzuga mos **Unsplash rasm URL manzili**.

## Faqat stdlib — internet/`pip` ishlatilmaydi

Ushbu muhitda tarmoqqa ruxsat yoʻq (PyPI 403 Forbidden), shuning uchun
`python-pptx`, `lxml` yoki `Pillow` **ishlatilmaydi va oʻrnatilmaydi**. `.pptx`
fayli OOXML XML qismlarini qoʻlda yozib, Python ning `zipfile` moduli bilan
arxivlash orqali yaratiladi.

### Unsplash rasmlari nega URL koʻrinishida?

Tarmoq yoʻqligi sababli rasm fayllarini (binary) yuklab olib paketga
joylashtirib boʻlmaydi. Shu bois har bir slaydda mavzuga mos **Unsplash URL
manzili** quyidagicha beriladi:

- slaydda **koʻrinadigan matn** sifatida,
- **bosiladigan havola** (OOXML `<a:hlinkClick>` + slayd `.rels` dagi tashqi
  `TargetMode="External"` bogʻlanishi) sifatida,
- shu slaydning **eslatmalar (notes)** qismida takrorlanadi.

Shunday qilib "Unsplash URL manzillari albatta qoʻshib ketilishi shart"
talabi internetsiz muhitda ham bajariladi.

## Qayta yaratish (generate)

```bash
python3 generate_barmoq_izi_presentation.py
```

Bu buyruq loyiha ildizida `Taqdimot_Barmoq_izi_skanerlari.pptx` faylini
qaytadan hosil qiladi.

## Tekshirish (verify)

```bash
python3 verify_barmoq_izi_presentation.py; echo exit=$?
```

Skript PASS xulosasini chop etadi va barcha talablar bajarilganda `0` kod
bilan tugaydi. Agar slaydlar soni 10 ta boʻlmasa, biror shrift "Times New
Roman" boʻlmasa, biror `sz` 1200 dan kichik boʻlsa yoki biror slaydda Unsplash
URL boʻlmasa — skript xato (FAIL) xabarini chiqarib, nol boʻlmagan kod bilan
tugaydi.

## Talablar

- Python 3.11+ (faqat standart kutubxona). Qoʻshimcha paket shart emas.
