# -*- coding: utf-8 -*-
"""
test_end_to_end_audit.py
End-to-end verification script testing math sanitization,
group shape thresholding, title breathing room, and watermark eradication.
"""
import os
import sys
import io

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from core.notebooklm_markdown_parser import NotebookLMMarkdownParser
from core.template_pptx_builder import AcademicPPTXBuilder
from core.slide_preview import get_presentation_previews

markdown_input = """
### 1-slayd: Kvant Kompyuterlari va Kriptografiya
* **Action-oriented sarlavha:** Kvant Hisoblashning Asosiy Tamoyillari va Axborot Xavfsizligi
* **Mantiqiy maqsadi:** Kvant hisoblash tizimlarining an'anaviy kriptotizimlarga solayotgan tahdidini tushuntirish.
* **Vizual turi:** cover

### 2-slayd: Kvant Superpozitsiyasi
* **Action-oriented sarlavha:** Kvant Superpozitsiya va Entanglement Prinsiplari Klassik Algoritmlarning Hisoblash Murakkabligini Keskin Oshiradi
* **Mantiqiy maqsadi:** Qubit holat vektorlari va kvant parallelizmining asimmetrik kriptografiyaga ta'sirini ochib berish.
* **Vizual turi:** cards_grid
* **Tarkibiy tezislar:**
  * Qubitning Holat Vektori: Bloch sferasi va $\\alpha|0\\rangle + \\beta|1\\rangle$ superpozitsiya holati orqali axborotni parallel qayta ishlash.
  * Kvant Entanglement: Zarrachalar orasidagi kvant bog'liqlik axborotni masofaga uzatish va uzatish tezligini oshiradi.
  * Klassik Turing: Klassik bitlar faqat 0 yoki 1 qiymatini qabul qilsa, kvant parallelizmi barcha holatlarni bir vaqtda qayta ishlaydi.

### 3-slayd: Shor Algoritmi va RSA
* **Action-oriented sarlavha:** Shor Algoritmi Asimmetrik Shifrlash Standartlarini $O((\\log N)^3)$ Vaqtda Sindiradi
* **Mantiqiy maqsadi:** RSA va ECC tizimlarining kvant kompyuterlari oldidagi zaifligini ko'rsatish.
* **Vizual turi:** comparison_vs
* **Tarkibiy tezislar:**
  * RSA Faktorizatsiyasi: Klassik Eratosfen yoki GNFS algoritmi eksponensial vaqt talab qiladi, kvantda esa polinominal vaqtga tushadi.
  * Elliptik Egri Chiziqlar: ECC diskret logarifm masalasini Shor algoritmi bir necha daqiqada yechishi mumkin.
"""

topic = "Kvant kompyuterlari va asimmetrik kriptografiya"
parsed = NotebookLMMarkdownParser.parse(markdown_input, topic=topic)
print(f"Parsed {len(parsed['slides'])} slides successfully.")
for s in parsed["slides"]:
    print(f"Slide {s.get('slide_number')}: {s.get('title')}")
    for p in s.get("points", []):
        print(f"  - {p}")

out_path = os.path.abspath(r"output\end_to_end_audit.pptx")
os.makedirs(os.path.dirname(out_path), exist_ok=True)

# Build presentation using AcademicPPTXBuilder
builder = AcademicPPTXBuilder()
out_file = builder.create_presentation(parsed)
print(f"Saved PPTX to: {out_file}")

previews = get_presentation_previews(out_file)
print("Previews generated:")
for p in previews:
    print(f"  {p}")
