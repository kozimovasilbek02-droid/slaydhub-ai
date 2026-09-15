---
name: image-to-pptx
description: Rasmlarni (.png, .jpg, skrinshotlar) to'liq 1:1 slayd taqdimotiga yoki 100% tahrirlanadigan vektor shakllar (shapes) va matnlardan iborat PowerPoint (.pptx) fayllariga aylantiruvchi Antigravity ko'nikmasi.
---

# Image to PPTX Skill for Antigravity

Ushbu ko'nikma foydalanuvchi taqdim etgan rasmlarni (.png, .jpg, .webp) ikki xil rejimda PowerPoint (.pptx) taqdimotiga aylantirish uchun xizmat qiladi:

## 1. Rejim A: 1:1 To'liq Slayd-Shou (Full Bleed Background)
Skrinshotlar yoki tayyor grafik slayd rasmlarini o'lchamini buzmasdan 16:9 widescreen formatida ketma-ketlikda PPTX ga aylantirish:
```bash
python convert_images_to_pptx.py "rasmlar_papkasi" -o "output/presentation.pptx"
```

## 2. Rejim B: AI Vision Tahrirlanadigan Shakllar (100% Editable Shapes & Texts)
Rasmdagi doiralar, yo'llar, strelkalar, ranglar va matnlarni ajratib, ularni Microsoft PowerPointning haqiqiy `AutoShape`, `BlockArc`, `Freeform` va `TextBox` elementlari sifatida qayta chizish:
```bash
python cli.py "rasmlar_papkasi/" -o "output/editable_presentation.pptx"
```

## 3. Web UI Studio:
```bash
streamlit run app.py
```
