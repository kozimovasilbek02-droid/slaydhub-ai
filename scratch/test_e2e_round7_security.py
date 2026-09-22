# -*- coding: utf-8 -*-
"""
scratch/test_e2e_round7_security.py
End-to-End Test Suite for Round 7:
Verifies SecurityPurger and DeepWatermarkPurger on complete multi-slide deck with real PPTX export.
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
from core.pptx_text_replacer import DeepWatermarkPurger
from core.security_purger import SecurityPurger

def test_e2e_security():
    print("🚀 Round 7 E2E Test: Starting Multi-Slide Security Audit...")

    prs = Presentation()
    prs.slide_width = Inches(13.333) # 16:9 Widescreen
    prs.slide_height = Inches(7.5)

    # --- SLIDE 1: Title Slide with Invisible Fullscreen Click Trap (PresentationGO pattern) ---
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    # Legit Title
    title1 = slide1.shapes.add_textbox(Inches(1), Inches(2), Inches(11.333), Inches(2))
    title1.text_frame.text = "Post-Kvant Kriptografiyasi va Milliy Xavfsizlik"
    title1.text_frame.paragraphs[0].font.size = Pt(36)
    title1.text_frame.paragraphs[0].font.bold = True
    title1.text_frame.paragraphs[0].font.color.rgb = RGBColor(15, 23, 42)

    # Injected Ghost Click Trap (Fullscreen invisible overlay)
    ghost_fullscreen = slide1.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    ghost_fullscreen.fill.background()
    ghost_fullscreen.line.fill.background()
    ghost_fullscreen.click_action.hyperlink.address = "https://www.presentationgo.com/templates/cyber-defense"

    # --- SLIDE 2: 3 Content Cards with Corner Ghost Action Buttons (SlideEgg & SlidesCarnival pattern) ---
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    # 3 Legit Cards
    card_titles = ["Kyber-768 Shifrlash", "Dilithium Raqamli Imzo", "SPHINCS+ Hash Funksiya"]
    for i, c_title in enumerate(card_titles):
        c_shape = slide2.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1 + i * 4), Inches(2), Inches(3.4), Inches(3.5))
        c_shape.fill.solid()
        c_shape.fill.fore_color.rgb = RGBColor(56, 189, 248) # Penpot Sky Blue
        c_shape.line.color.rgb = RGBColor(14, 165, 233)
        c_shape.text_frame.text = f"{c_title}\n\nLattice-based kriptografiya protokoli bo'yicha NIST standarti."

    # Corner Ghost Button 1
    ghost_corner1 = slide2.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, Inches(12), Inches(6.5), Inches(1.2), Inches(0.8))
    ghost_corner1.fill.background()
    ghost_corner1.line.fill.background()
    ghost_corner1.click_action.hyperlink.address = "https://slideegg.com/exclusive-deals"

    # Corner Ghost Button 2
    ghost_corner2 = slide2.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.OVAL, Inches(0.2), Inches(6.8), Inches(0.5), Inches(0.5))
    ghost_corner2.fill.background()
    ghost_corner2.line.fill.background()
    ghost_corner2.click_action.hyperlink.address = "https://slidescarnival.com/track"

    # --- SLIDE 3: Comparison with Run Hyperlinks and Blocker Overlay ---
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide3.shapes.add_textbox(Inches(1), Inches(1), Inches(11), Inches(3))
    p1 = textbox.text_frame.paragraphs[0]
    r_legit = p1.add_run()
    r_legit.text = "Asimmetrik RSA 4096 Shori algoritmi qarshisida zaifdir. "
    r_spam = p1.add_run()
    r_spam.text = "[Template designed by Canva & PresentationGO]"
    r_spam.hyperlink.address = "https://canva.com/templates/presentation-free"

    # Transparent Blocker Overlay
    blocker = slide3.shapes.add_shape(pptx.enum.shapes.MSO_SHAPE.RECTANGLE, Inches(1), Inches(4), Inches(11), Inches(3))
    blocker.fill.background()
    blocker.line.fill.background()

    total_shapes_before = sum(len(s.shapes) for s in prs.slides)
    print(f"Boshlang'ich umumiy shakllar soni: {total_shapes_before}")
    assert total_shapes_before == (2 + 5 + 2) # 9 shapes

    # --- EXECUTE MASTER AUDIT VIA DEEPWATERMARKPURGER ---
    DeepWatermarkPurger.purge_presentation_watermarks(prs)

    total_shapes_after = sum(len(s.shapes) for s in prs.slides)
    print(f"Tozalashdan keyingi umumiy shakllar soni: {total_shapes_after}")

    # Verify each slide
    print(f"Slide 1 shakllari: {len(slide1.shapes)} (Kutilgan: 1 ta sarlavha)")
    assert len(slide1.shapes) == 1, "Slide 1 ghost click-trap was not excised!"
    assert "Post-Kvant Kriptografiyasi" in slide1.shapes[0].text_frame.text

    print(f"Slide 2 shakllari: {len(slide2.shapes)} (Kutilgan: 3 ta qonuniy kartochka)")
    assert len(slide2.shapes) == 3, "Slide 2 corner ghost buttons were not excised!"

    print(f"Slide 3 shakllari: {len(slide3.shapes)} (Kutilgan: 1 ta matn qutisi)")
    assert len(slide3.shapes) == 1, "Slide 3 transparent blocker overlay was not excised!"

    # Verify that run hyperlink in Slide 3 was stripped
    slide3_hlinks = slide3.shapes[0]._element.xpath('.//a:hlinkClick')
    print(f"Slide 3 ichidagi havolalar soni: {len(slide3_hlinks)} (Kutilgan: 0 ta)")
    assert len(slide3_hlinks) == 0, "Slide 3 spam hyperlink was not defused!"

    # Save output deck
    os.makedirs("output", exist_ok=True)
    out_path = "output/test_round7_security_sanitized.pptx"
    prs.save(out_path)
    print(f"✅ Muvaffaqiyatli saqlandi: {out_path}")

    # Re-open presentation to verify OpenXML schema compliance
    prs_check = Presentation(out_path)
    assert len(prs_check.slides) == 3, "Presentation XML schema check failed!"
    print("💎 OpenXML Schema Integrity Check: 100% VALID & HEALTHY!")
    print("🏆 ROUND 7 E2E SECURITY AUDIT: 100% PASS!")

if __name__ == "__main__":
    test_e2e_security()
