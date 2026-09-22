# -*- coding: utf-8 -*-
"""
scratch/test_authentic_template_bullet.py
Tests preserving authentic designer Wingdings checkmark bullets on a real PPTX template.
"""

import sys
import os
import io
from pathlib import Path
from pptx import Presentation

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from core.pptx_text_replacer import PPTXTextReplacer
from core.bullet_harmonizer import BulletHarmonizer, A_NS
from core.slide_preview import get_presentation_previews

template_path = "output/uploaded_templates/Abstract Dynamic Geometric PowerPoint Templates.pptx"
prs = Presentation(template_path)

# Slide 16 has the Wingdings checkmark bullet list
slide = prs.slides[16]

# Find the bullet shape
target_shape = None
for shape in slide.shapes:
    if shape.has_text_frame and shape.name == "TextBox 5":
        target_shape = shape
        break

assert target_shape is not None, "TextBox 5 not found on slide 16!"

print(f"Original text in {target_shape.name}: {target_shape.text_frame.text!r}")

# Check original bullet
orig_style = BulletHarmonizer.extract_paragraph_bullet_style(target_shape.text_frame.paragraphs[0])
print(f"Extracted original style: Font={orig_style.font_typeface}, Char={orig_style.char}, marL={orig_style.mar_l}, indent={orig_style.indent}")

# Replace text using PPTXTextReplacer.set_shape_text_preserving_style
new_text = """
• Birinchi Asosiy Vazifa: Kvant algoritmlarini loyihalash va optimallashtirish
• Ikkinchi Strategik Bosqich: NIST xavfsizlik standartlariga to'liq moslashish
• Uchinchi Amaliy Qadam: Milliy axborot infratuzilmasini himoyalash
"""

PPTXTextReplacer.set_shape_text_preserving_style(
    shape=target_shape,
    new_text=new_text,
    default_font_size=16.0,
    is_title=False,
    is_card_title=False,
    theme_key="tech_minimal"
)

# Verify that all 3 new paragraphs preserved the Wingdings checkmark!
for idx, p in enumerate(target_shape.text_frame.paragraphs):
    pPr = p._p.find(f"{{{A_NS}}}pPr")
    buChar = pPr.find(f"{{{A_NS}}}buChar")
    buFont = pPr.find(f"{{{A_NS}}}buFont")
    print(f"Paragraph {idx}: Text={p.text[:35]!r} | buChar={buChar.attrib.get('char') if buChar is not None else None} | buFont={buFont.attrib.get('typeface') if buFont is not None else None}")
    assert buChar is not None and buChar.attrib.get("char") == "v", f"Paragraph {idx} did not preserve 'v' char!"
    assert buFont is not None and buFont.attrib.get("typeface") == "Wingdings", f"Paragraph {idx} did not preserve Wingdings font!"

out_path = "output/test_authentic_wingdings_preserved.pptx"
prs.save(out_path)
print(f"Saved authentic template test to: {out_path}")

previews = get_presentation_previews(out_path)
print(f"Rendered {len(previews)} preview slides: {previews}")
print("🎉 AUTHENTIC WINGDINGS BULLET PRESERVATION VERIFIED 100%!")
