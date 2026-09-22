# -*- coding: utf-8 -*-
"""
Dual Mode Comparison Test:
Runs presentation generation for both:
1. Rejim 1: Slayd Studiyasi (NotebookLM & Native PPTX)
2. Rejim 2: 5,170+ PPTX Shablonlar Bazasi (Direct Template Customizer)
Using the exact same topic and structured academic content.
"""

import os
import sys
import time
import json
import shutil
from pathlib import Path

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from core.config import config
from core.academic_catalog import get_academic_catalog
from core.academic_matcher import get_academic_matcher
from core.notebooklm_markdown_parser import NotebookLMMarkdownParser
from core.template_pptx_builder import get_academic_pptx_builder
from core.pptx_text_replacer import replace_presentation_content
from core.slide_preview import get_presentation_previews, _get_file_hash, PREVIEW_CACHE_DIR

TOPIC = "Kvant kompyuterlari va asimmetrik kriptografiya"

REALISTIC_MD = f"""### 1-slayd: Titul va Kirish
* **Action-oriented sarlavha:** {TOPIC}
* **Mantiqiy maqsadi:** Titul slaydi va tadqiqot tafsilotlari
* **Tarkibiy tezislar:**
  * Kvant algoritmlari davrida axborot xavfsizligi muammolari
  * Post-kvant kriptografiyasining dolzarbligi
  * Tadqiqotchi: Alisher Vohidov | O'zMU 2026

### 2-slayd: Tadqiqotning Mantiqiy Tuzilishi
* **Action-oriented sarlavha:** Tadqiqot Rejasi va Strukturasi
* **Mantiqiy maqsadi:** Taqdimot yo'l xaritasi
* **Tarkibiy tezislar:**
  * 1. Kvant hisoblashlari va Shor algoritmi xavfi
  * 2. Asimmetrik kriptografiya tahlili: RSA va ECC zaifliklari
  * 3. NIST PQC standartlari: ML-KEM va ML-DSA
  * 4. Gibrid xavfsizlik arxitekturasi va tavsiyalar

### 3-slayd: Kvant Tahdidlari va Shor Algoritmi
* **Action-oriented sarlavha:** Shor Algoritmi Asimmetrik Kriptotizimlarni Ishdan Chiqaradi
* **Mantiqiy maqsadi:** Nazariy asos va matematik muammo
* **Tarkibiy tezislar:**
  * Shor algoritmi diskret logarifm va ko'paytuvchilarga ajratishni polinomiyal vaqtda O((log N)³) yechadi.
  * RSA-2048 va ECC-256 kvant quvvati qarshisida o'z kuchini to'liq yo'qotadi.
  * Ma'lumotlarni hozir saqlab, kelajakda shifrdan chiqarish ("Harvest Now, Decrypt Later") tahdidi mavjud.

### 4-slayd: Tadqiqot Obyekti va Vazifalari
* **Action-oriented sarlavha:** Ilmiy Izlanishning Asosiy Vazifalari
* **Mantiqiy maqsadi:** Obyekt, predmet va maqsadlarni belgilash
* **Tarkibiy tezislar:**
  * Obyekt: Korporativ tarmoqlardagi TLS 1.3 va IPsec protokollari.
  * Predmet: Post-kvant KEM mexanizmlarining tezkorligi va xavfsizligi.
  * Maqsad: Mavjud PKI infratuzilmasiga minimal ortiqcha yuklama bilan gibrid shifrlashni joriy etish.

### 5-slayd: Asosiy Miqdoriy Ko'rsatkichlar
* **Action-oriented sarlavha:** Kvant Xavfsizligining Miqdoriy Metriklari
* **Mantiqiy maqsadi:** Empirik ko'rsatkichlar
* **Tarkibiy tezislar:**
  * 99.9% Barqarorlik: NIST FIPS 203 standarti asosidagi panjarali shifrlash kafolati.
  * 2,500+ Mantiqiy Kubit: RSA kalitini sindirish uchun zarur bo'lgan minimal kvant hisoblash quvvati.
  * 1.2 ms Handshake: Gibrid TLS ulanishida yuzaga keladigan o'rtacha kechikish vaqti.

### 6-slayd: Qiyosiy Tahlil: Klassik vs Post-Kvant
* **Action-oriented sarlavha:** Klassik RSA va Post-Kvant ML-KEM Qiyosi
* **Mantiqiy maqsadi:** Model A va Model B taqqoslash
* **Tarkibiy tezislar:**
  * Klassik RSA-2048: Kalit hajmi 256 bayt, kvantga chidamlilik 0%, shifrlash tezligi past.
  * Post-Kvant ML-KEM: Kalit hajmi 1184 bayt, kvantga chidamlilik 100%, panjarali SVP matematik murakkablik.

### 7-slayd: Post-Kvantga O'tish Xronologiyasi
* **Action-oriented sarlavha:** Tizimli Migratsiya Bosqichlari (2026-2030)
* **Mantiqiy maqsadi:** Bosqichma-bosqich yo'l xaritasi
* **Tarkibiy tezislar:**
  * 1-Bosqich (Audit): Barcha mavjud sertifikatlar va kriptografik kalitlar inventarizatsiyasi.
  * 2-Bosqich (Gibrid Rejim): Klassik ECDH va ML-KEM algoritmlarini parallel qo'llash.
  * 3-Bosqich (To'liq O'tish): Infratuzilmani 100% PQC standartlariga ko'chirish.

### 8-slayd: Tizim Arxitekturasi va Tarkibi
* **Action-oriented sarlavha:** Uch Bosqichli Himoyalangan Arxitektura
* **Mantiqiy maqsadi:** Blokli tizim tarkibi
* **Tarkibiy tezislar:**
  * Kalit Almashinuv Qatlami: ML-KEM-768 protokoli asosidagi kapsulyatsiya.
  * Raqamli Imzo Qatlami: ML-DSA standarti yordamida soxtalashtirishdan himoya.
  * Kripto-Moslashuvchanlik: Yangi standartlar paydo bo'lganda algoritmlarni to'xtovsiz almashtirish.

### 9-slayd: Xulosalar va Tavsiyalar
* **Action-oriented sarlavha:** Ilmiy Xulosalar va Amaliy Tavsiyalar
* **Mantiqiy maqsadi:** Tadqiqot xulosalari
* **Tarkibiy tezislar:**
  * Post-kvant kriptografiyasiga o'tish kechiktirib bo'lmas strategik xavfsizlik talabidir.
  * Gibrid kalit almashinuvi klassik tizimlar bilan to'liq orqaga moslikni (backward compatibility) ta'minlaydi.
  * Xalqaro NIST FIPS 203, 204 standartlarini milliy qonunchilikka tatbiq etish tavsiya etiladi.

### 10-slayd: Foydalanilgan Manbalar
* **Action-oriented sarlavha:** Ilmiy Adabiyotlar va Standartlar Ro'yxati
* **Mantiqiy maqsadi:** Bibliografiya
* **Tarkibiy tezislar:**
  * 1. NIST FIPS 203: Module-Lattice-Based Key-Encapsulation Mechanism Standard (2024).
  * 2. Shor, P. W. Algorithms for Quantum Computation: Discrete Logarithms and Factoring (IEEE, 1994).
  * 3. Bernstein, D. J. Post-quantum cryptography: dealing with the inevitable (Springer, 2009).
"""


def main():
    print("=" * 70)
    print("IKKALA REJIMDA BIR XIL MAVZUDAGI SLAYDLARNI GENERATSIYA QILISH TESTI")
    print("=" * 70)

    catalog = get_academic_catalog()
    matcher = get_academic_matcher()
    builder = get_academic_pptx_builder()

    # Parse content
    parsed_content = NotebookLMMarkdownParser.parse(
        REALISTIC_MD,
        topic=TOPIC,
        default_metadata={
            "author_name": "Alisher Vohidov",
            "organization": "O'zbekiston Milliy Universiteti",
            "supervisor": "DSc., prof. A. Karimov"
        }
    )

    out_dir = Path("output/academic_studio")
    out_dir.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------------------
    # REJIM 1: Slayd Studiyasi (NotebookLM & Native PPTX)
    # -------------------------------------------------------------
    print("\n[REJIM 1] Slayd Studiyasi (Splicing & Blueprint Engine)...")
    bp1 = matcher.build_blueprint(
        topic=TOPIC,
        slide_count=10,
        language="uz",
        presentation_mode="scientific_defense_bmi",
        presenter_info={
            "author_name": "Alisher Vohidov",
            "organization": "O'zbekiston Milliy Universiteti",
            "supervisor": "DSc., prof. A. Karimov"
        }
    )
    # Update blueprint archetypes from parsed content
    for idx, sp in enumerate(parsed_content.get("slides", [])):
        if idx < len(bp1["slides"]):
            arch = sp.get("layout_type", "cards_grid")
            bp1["slides"][idx]["layout_type"] = arch
            bp1["slides"][idx]["title_hint"] = sp.get("title", "")
            matches = catalog.search_slides_by_archetype(arch, category=bp1.get("category"), topic=TOPIC, limit=1)
            if matches:
                bp1["slides"][idx]["matched_slide_ref"] = matches[0]
                bp1["slides"][idx]["pptx_path"] = matches[0]["pptx_path"]

    out1_path = str(out_dir / "Rejim1_Slayd_Studiyasi.pptx")
    builder.create_presentation(parsed_content, bp1, engine_mode="harmonized")
    # Move/copy to distinct file
    latest_cust = str(out_dir / "Kvant_kompyuterlari_va_asimmetrik_kripto_Customized.pptx")
    shutil.copy2(latest_cust, out1_path)
    print(f"  [OK] Rejim 1 fayli: {out1_path} ({os.path.getsize(out1_path):,} bytes)")

    # Render Rejim 1 Previews
    h1 = _get_file_hash(out1_path)
    d1 = PREVIEW_CACHE_DIR / h1
    if d1.exists():
        shutil.rmtree(d1)
    prev1 = get_presentation_previews(out1_path, max_slides=10)
    print(f"  [OK] Rejim 1 prevyulari: {len(prev1)} ta rasm keshlandi ({d1})")

    # -------------------------------------------------------------
    # REJIM 2: 5,170+ PPTX Shablonlar Bazasi (Direct Template In-Place Customizer)
    # -------------------------------------------------------------
    print("\n[REJIM 2] 5,170+ PPTX Shablonlar Bazasi (Direct Template In-Place)...")
    # Find a great authentic template for Quantum/Academic from catalog
    template_candidates = catalog.search_templates("Quantum", limit=5)
    if not template_candidates:
        template_candidates = catalog.search_templates("3D", limit=5)
    if not template_candidates:
        template_candidates = catalog.templates[:5]
    
    # Use template
    target_tmpl = template_candidates[0]
    tmpl_path = target_tmpl["pptx_path"]
    tmpl_title = target_tmpl["title"]
    print(f"  Tanlangan shablon: '{tmpl_title}' ({os.path.basename(tmpl_path)})")

    bp2 = matcher.build_blueprint(
        topic=TOPIC,
        slide_count=10,
        language="uz"
    )
    bp2["primary_pptx_path"] = tmpl_path
    bp2["primary_deck_title"] = tmpl_title
    # In Rejim 2, all slides are customized directly in-place on the selected template (no cross-deck splicing)
    bp2["slides"] = []  # Clear cross-deck references to test single template in-place replacer

    out2_path = str(out_dir / "Rejim2_Shablonlar_Bazasi.pptx")
    replace_presentation_content(tmpl_path, parsed_content, out2_path, blueprint=bp2)
    print(f"  [OK] Rejim 2 fayli: {out2_path} ({os.path.getsize(out2_path):,} bytes)")

    # Render Rejim 2 Previews
    h2 = _get_file_hash(out2_path)
    d2 = PREVIEW_CACHE_DIR / h2
    if d2.exists():
        shutil.rmtree(d2)
    prev2 = get_presentation_previews(out2_path, max_slides=10)
    print(f"  [OK] Rejim 2 prevyulari: {len(prev2)} ta rasm keshlandi ({d2})")

    print("\n" + "=" * 70)
    print("MUVAFFAQIYATLI GENERATSIYA QILINDI!")
    print(f"Rejim 1 Prevyu papkasi: {d1}")
    print(f"Rejim 2 Prevyu papkasi: {d2}")
    print("=" * 70)


if __name__ == "__main__":
    main()
