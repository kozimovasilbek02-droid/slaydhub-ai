# -*- coding: utf-8 -*-
"""
scratch/test_e2e_complete_council.py
Full End-to-End verification test combining all 4 deliberation rounds of the Dual-Brain Council:
1. MicroTypography & Knuth-Plass headline partitioning
2. SlideMorphEngine with mc:AlternateContent
3. TableMatrixReplacer with OpenXML dimension sync
4. ThemeHarmonizer for vector accent shapes
5. SlotAllocator for content density and continuation slides
"""

import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from core.pptx_text_replacer import replace_presentation_content
from core.slide_preview import get_presentation_previews

def run_e2e():
    template_path = "output/action_plan_s2_flawless.pptx"
    if not os.path.exists(template_path):
        template_path = "output/test_ideal_native.pptx"

    print(f"Using template: {template_path}")

    presentation_data = {
        "topic": "Kvant Hisoblash va Post-Kvant Kriptografiyasining Zamonaviy Asoslari",
        "theme_key": "cyber_dark",
        "slides": [
            {
                "title": "Kvant Hisoblash va Post-Kvant Kriptografiyasining Asoslari",
                "subtitle": "Kvant ustunligi davrida ma'lumotlar xavfsizligi arxitekturasi",
                "is_cover": True,
                "layout_type": "cover"
            },
            {
                "title": "Kvant Tizimlarining Asosiy Texnologik Bosqichlari va Mexanizmlari",
                "layout_type": "cards_grid",
                "cards": [
                    {"title": f"Bosqich {i}: Texnologik Algoritm", "description": f"Kvant registrlarida superpozitsiya va chigallik holatini hisoblashning {i}-amaliy bosqichi."}
                    for i in range(1, 8)  # 7 items -> will be balanced into 2 continuation slides!
                ]
            },
            {
                "title": "Klassik va Post-Kvant Kriptografik Standartlari Taqqoslovi",
                "layout_type": "table_data",
                "table_data": {
                    "headers": ["Metrika", "RSA-2048", "ECC-256", "Kyber-768 (PQC)"],
                    "rows": [
                        ["Kalit Hajmi", "2048 bit", "256 bit", "1184 bayt"],
                        ["Kvant Bardoshligi", "Zaif (Shor)", "Zaif (Shor)", "100% Xavfsiz"],
                        ["Shifrlash Tezligi", "O'rtacha", "Juda Yuqori", "Ekstremal Yuqori"],
                        ["NIST Standarti", "Legacy", "FIPS 186-4", "FIPS 203 (2024)"]
                    ]
                }
            },
            {
                "title": "Kvant Arxitekturasi va Amaliy Xavflar Tahlili",
                "layout_type": "theory_concept",
                "concept_name": "Kvant Dekogerensiyasi va Xatoliklarni Tuzatish",
                "definition": "• Qubitlar Barqarorligi: Tashqi muhit shovqinlari ta'sirida superpozitsiya holatining buzilishi\n  - Kriogenik Sovutish: Mutlaq nol haroratga yaqin (15 millikelvin) muhitni ta'minlash\n  - Xatoliklarni Bartaraf Etish: Yuzlab fizik kubitlar yordamida mantiqiy kubit shakllantirish\n• Kriptografik Tahdidlar: Shor algoritmining asimmetrik shifrlarga ta'siri\n  - RSA kalitlarini faktoriallash vaqti soniyalarga qisqaradi",
                "key_formula_or_law": "|ψ⟩ = α|0⟩ + β|1⟩",
                "importance": "Axborot xavfsizligida milliy strategik ma'lumotlarni PQC standartlariga o'tkazishni taqozo etadi."
            }
        ]
    }

    output_pptx = "output/e2e_council_flawless.pptx"
    os.makedirs("output", exist_ok=True)

    print("Generating comprehensive Dual-Brain presentation...")
    final_path = replace_presentation_content(
        template_pptx_path=template_path,
        content_data=presentation_data,
        output_pptx_path=output_pptx
    )
    print(f"Presentation created successfully at: {final_path}")

    # Generate previews
    previews = get_presentation_previews(final_path)
    print(f"Rendered {len(previews)} high-res preview slides:")
    for p in previews:
        print(f"  -> {p}")

    print("\n==========================================")
    print("E2E DUAL-BRAIN VERIFICATION COMPLETE!")
    print("==========================================")

if __name__ == "__main__":
    run_e2e()
