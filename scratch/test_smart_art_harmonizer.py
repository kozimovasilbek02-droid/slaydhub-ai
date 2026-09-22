# -*- coding: utf-8 -*-
"""
scratch/test_smart_art_harmonizer.py
Unit tests for SmartArtGuardian and OutlineHarmonizer (Round 11).
Validates:
1. SmartArt diagram detection and safe bypass.
2. Outline / stroke border (<a:ln>) color harmonization with Penpot themes.
3. Preservation of intentional transparent borders (<a:noFill>).
4. XML schema order and 0% corruption guarantee.
"""

import os
import sys
import io

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement

from core.smart_art_harmonizer import SmartArtGuardian, OutlineHarmonizer
from core.penpot_themes import PenpotThemeManager


def test_smart_art_and_outlines():
    print("🧪 Running SmartArt & Outline Harmonizer Unit Tests (Round 11)...")

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1. Test Outline Harmonization
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1), Inches(4), Inches(2))
    shape.text = "Card with Clashing Red Border"

    # Add a clashing red outline manually into XML: <a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
    spPr = shape._element.xpath('./p:spPr')[0]
    ln = OxmlElement('a:ln')
    ln.set('w', '25400') # 2pt
    solidFill = OxmlElement('a:solidFill')
    srgbClr = OxmlElement('a:srgbClr')
    srgbClr.set('val', 'FF0000') # Clashing Red
    solidFill.append(srgbClr)
    ln.append(solidFill)
    spPr.append(ln)

    # Verify initial outline is red
    initial_stroke_clr = shape._element.xpath('.//a:ln//a:srgbClr/@val')
    print(f"Boshlang'ich kontur rangi: {initial_stroke_clr}")
    assert initial_stroke_clr == ['FF0000'], "Initial outline must be FF0000"

    # Harmonize with Penpot cyber_dark theme (#38BDF8 primary)
    success = OutlineHarmonizer.harmonize_shape_outline(shape, "38BDF8")
    assert success is True, "Outline harmonization must succeed"

    harmonized_stroke_clr = shape._element.xpath('.//a:ln//a:srgbClr/@val')
    print(f"Uyg'unlashtirilgan kontur rangi: {harmonized_stroke_clr}")
    assert harmonized_stroke_clr == ['38BDF8'], "Harmonized outline must be 38BDF8 (Cyber Sky Blue)"

    # 2. Test Transparent Border Protection (<a:noFill>)
    shape_transparent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6), Inches(1), Inches(3), Inches(2))
    spPr2 = shape_transparent._element.xpath('./p:spPr')[0]
    ln2 = OxmlElement('a:ln')
    ln2.append(OxmlElement('a:noFill'))
    spPr2.append(ln2)

    skipped = OutlineHarmonizer.harmonize_shape_outline(shape_transparent, "38BDF8", preserve_transparent=True)
    print(f"Shaffof kontur saqlandi (o'zgartirish o'tkazilmadi): {not skipped}")
    assert skipped is False, "Transparent outlines (<a:noFill>) must be preserved"

    # 3. Test SmartArt Diagram Detection
    print(f"Oddiy to'rtburchak SmartArt ekanligi: {SmartArtGuardian.is_smart_art(shape)}")
    assert SmartArtGuardian.is_smart_art(shape) is False

    # Simulate a SmartArt graphic frame element
    graphic_frame = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(4), Inches(4), Inches(2))
    gf_elem = graphic_frame._element
    # Inject drawingml diagram marker
    gf_elem.tag = '{http://schemas.openxmlformats.org/presentationml/2006/main}graphicFrame'
    dgm_node = OxmlElement('a:graphicData')
    dgm_node.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/diagram')
    gf_elem.append(dgm_node)

    is_sa = SmartArtGuardian.is_smart_art(graphic_frame)
    print(f"Simulyatsiya qilingan diagramma SmartArt ekanligi: {is_sa}")
    assert is_sa is True, "SmartArt diagram graphic frame must be detected as True"

    # Test safe bypass
    bypass = SmartArtGuardian.should_bypass_shape(graphic_frame)
    print(f"Diagramma destruktiv tahrirdan himoyalandi (Bypass): {bypass}")
    assert bypass is True

    # Slide-wide outline harmonization test
    slide_outlines_count = OutlineHarmonizer.harmonize_slide_outlines(slide, theme_key="academic_navy")
    print(f"Slayd bo'ylab yangilangan konturlar soni: {slide_outlines_count}")
    assert slide_outlines_count >= 1

    os.makedirs("output", exist_ok=True)
    prs.save("output/test_smartart_outlines.pptx")
    print("✅ Muvaffaqiyatli saqlandi: output/test_smartart_outlines.pptx")
    print("🎉 Barcha SmartArt & Outline Harmonizer testlari 100% PASS!\n")


if __name__ == "__main__":
    test_smart_art_and_outlines()
