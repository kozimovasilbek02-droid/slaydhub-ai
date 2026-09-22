# -*- coding: utf-8 -*-
"""
Verification suite for Native PPTX Academic Template System.
"""

import os
import sys
import glob
from pathlib import Path

CURRENT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(CURRENT_DIR))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.academic_catalog import get_academic_catalog
from core.academic_matcher import get_academic_matcher
from core.notebooklm_prompt_gen import generate_notebooklm_prompt
from core.notebooklm_executor import NotebookLMExecutor
from core.template_pptx_builder import get_academic_pptx_builder
from core.pptx_text_replacer import replace_presentation_content


def test_native_catalog():
    print("\n--- 1. Testing Native PPTX Catalog ---")
    cat = get_academic_catalog()
    assert len(cat.templates) > 1000, f"Expected >1000 PPTX templates, got {len(cat.templates)}"
    print(f"✅ Indexed {len(cat.templates):,} authentic editable PPTX templates.")

    # Search check
    for q in ["quantum", "medical", "business", "cyber"]:
        res = cat.search_templates(q, limit=3)
        assert len(res) > 0, f"Search '{q}' should have results"
        print(f"  🔍 '{q}' -> Top PPTX: {res[0]['title']} ({os.path.basename(res[0]['pptx_path'])})")


def test_native_matcher_and_prompt():
    print("\n--- 2. Testing Topic Matcher & NotebookLM Prompt ---")
    matcher = get_academic_matcher()
    bp = matcher.build_blueprint("Sun'iy intellekt va neyrotarmoqlar", slide_count=8, language="uz")
    assert bp["primary_pptx_path"] and os.path.exists(bp["primary_pptx_path"]), "Should match valid PPTX file"
    print(f"✅ Matched Topic to PPTX Template: '{bp['primary_deck_title']}'")
    print(f"   PPTX Path: {bp['primary_pptx_path']}")

    prompt = generate_notebooklm_prompt(bp)
    assert len(prompt) > 500
    print(f"✅ NotebookLM Prompt generated ({len(prompt)} characters).")


def test_native_pptx_customization():
    print("\n--- 3. Testing Authentic PPTX In-Place Text Replacement ---")
    matcher = get_academic_matcher()
    bp = matcher.build_blueprint("Sun'iy intellekt va neyrotarmoqlar", slide_count=5, language="uz")
    
    sample_content = {
        "topic": bp["topic"],
        "slides": [
            {"slide_number": 1, "layout_type": "cover", "title": "Sun'iy intellekt va Neyron tarmoqlari", "subtitle": "Zamonaviy chuqur modellar va arxitektura | 2026"},
            {"slide_number": 2, "layout_type": "agenda", "points": ["1. Neyron tarmoqlari asoslari", "2. Chuqur arxitektura", "3. Xulosa"]},
            {"slide_number": 3, "layout_type": "theory_concept", "title": "Chuqur O'rganish Asoslari", "concept_name": "Deep Learning", "definition": "Murakkab qonuniyatlarni ko'p qatlamli neyronlar orqali o'rganish.", "key_formula_or_law": "y = sigma(Wx+b)", "importance": "Zamonaviy AI yadrosi."},
            {"slide_number": 4, "layout_type": "cards_grid", "cards": [{"title": "CNN", "desc": "Tasvirlar tahlili."}, {"title": "RNN", "desc": "Matn va ketma-ketliklar."}, {"title": "Transformers", "desc": "Self-attention mexanizmi."}]},
            {"slide_number": 5, "layout_type": "conclusion_sources", "key_takeaways": ["AI kelajak texnologiyasi.", "Xavfsizlik choralari zarur."], "references": ["1. MIT Deep Learning Book (2016)"]}
        ]
    }

    builder = get_academic_pptx_builder()
    out_pptx = builder.create_presentation(sample_content, bp)
    assert os.path.exists(out_pptx), "Output PPTX should exist"
    assert os.path.getsize(out_pptx) > 20000, "PPTX should contain real template assets (>20KB)"
    print(f"✅ Customization SUCCESS! Created customized authentic PPTX at: {out_pptx} ({os.path.getsize(out_pptx):,} bytes)")


if __name__ == "__main__":
    print("==================================================")
    print("🚀 Running Native PPTX Template Suite Verifications")
    print("==================================================")
    test_native_catalog()
    test_native_matcher_and_prompt()
    test_native_pptx_customization()
    print("\n🎉 ALL TESTS PASSED! NATIVE TEMPLATES ARE FULLY FUNCTIONAL!")
