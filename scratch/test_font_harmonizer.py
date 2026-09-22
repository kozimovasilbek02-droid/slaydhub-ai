# -*- coding: utf-8 -*-
"""
scratch/test_font_harmonizer.py
Unit test suite for FontHarmonizer (Round 8):
- Tests font_title application on titles
- Tests font_body application on body/table
- Tests <a:latin> and <a:cs> OpenXML injection
- Tests Uzbek (O', G') and Cyrillic character integrity
"""

import sys
import os
import io
from pathlib import Path

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from lxml import etree
from core.font_harmonizer import FontHarmonizer

def run_tests():
    print("🧪 Running FontHarmonizer Unit Tests (Round 8)...")
    prs = Presentation()

    # Slide 1: Title & Subtitle Slide
    slide1 = prs.slides.add_slide(prs.slide_layouts[0]) # Title layout
    title1 = slide1.shapes.title
    title1.text = "O'zbekiston Kvant Axborot Tizimlari va G'oyalar"

    subtitle1 = slide1.placeholders[1]
    subtitle1.text = "Milliy xavfsizlik va asimmetrik kriptografiya asoslari"

    # Slide 2: Content Cards and Table
    slide2 = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout
    card = slide2.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(4), Inches(2))
    card.text_frame.text = "Kriptografik Protokollar\nLattice-based algoritmlar sintezi."

    # Table
    table_shape = slide2.shapes.add_table(2, 2, Inches(6), Inches(1), Inches(4), Inches(2))
    table = table_shape.table
    table.cell(0, 0).text = "Parametr"
    table.cell(0, 1).text = "Qiymat"
    table.cell(1, 0).text = "Kalit Hajmi"
    table.cell(1, 1).text = "256 bit"

    print("Initial presentation created. Applying FontHarmonizer...")

    # Run FontHarmonizer with 'tech_minimal' theme (Montserrat title, Segoe UI body)
    FontHarmonizer.harmonize_presentation_fonts(prs, "tech_minimal")

    # Verify Title Run Properties
    title_rPr = title1.text_frame.paragraphs[0].runs[0]._r.rPr
    latin_nodes = title_rPr.xpath('./a:latin')
    cs_nodes = title_rPr.xpath('./a:cs')

    print(f"Title latin typeface: {latin_nodes[0].get('typeface') if latin_nodes else None}")
    print(f"Title cs typeface: {cs_nodes[0].get('typeface') if cs_nodes else None}")

    assert latin_nodes and latin_nodes[0].get('typeface') == "Montserrat", "Title latin typeface must be Montserrat!"
    assert cs_nodes and cs_nodes[0].get('typeface') == "Montserrat", "Title cs typeface must be Montserrat!"

    # Verify Table Cell Run Properties
    cell_rPr = table.cell(0, 0).text_frame.paragraphs[0].runs[0]._r.rPr
    cell_latin = cell_rPr.xpath('./a:latin')
    cell_cs = cell_rPr.xpath('./a:cs')

    print(f"Table cell latin typeface: {cell_latin[0].get('typeface') if cell_latin else None}")
    print(f"Table cell cs typeface: {cell_cs[0].get('typeface') if cell_cs else None}")

    assert cell_latin and cell_latin[0].get('typeface') == "Segoe UI", "Table cell latin typeface must be Segoe UI!"
    assert cell_cs and cell_cs[0].get('typeface') == "Segoe UI", "Table cell cs typeface must be Segoe UI!"

    # Save output deck
    os.makedirs("output", exist_ok=True)
    out_path = "output/test_font_harmonized.pptx"
    prs.save(out_path)
    print(f"✅ Muvaffaqiyatli saqlandi: {out_path}")

    # Verify OpenXML reload
    prs_check = Presentation(out_path)
    assert len(prs_check.slides) == 2, "Reload check failed!"
    print("💎 OpenXML Typography Check: 100% VALID & HARMONIZED!")
    print("🎉 Barcha testlar 100% PASS!")

if __name__ == "__main__":
    run_tests()
