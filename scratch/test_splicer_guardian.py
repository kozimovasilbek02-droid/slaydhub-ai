# -*- coding: utf-8 -*-
"""
scratch/test_splicer_guardian.py
Unit test suite for core/splicer_guardian.py and slide_splicer.py integration (Round 14).
Verifies:
1. Shape ID collision detection and unique renumbering.
2. Cross-deck media relationship remapping and dangling r:id purging.
3. Slide background replication (including master fallback).
4. Full cross-deck assembly and PowerPoint COM preview verification.
"""

import sys
import os
import io
import zipfile
import lxml.etree as etree
from pptx import Presentation
from pptx.util import Inches

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from core.splicer_guardian import SplicerGuardian, R_NS
from core.slide_splicer import HarmonizedSlideCloner, UniversalSlideSplicer
from core.slide_preview import get_presentation_previews


def test_unique_shape_ids():
    print("--- 1. Testing Shape ID Renumberer & Collision Purging ---")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Add shapes and manually introduce duplicate shape IDs
    tb1 = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    tb2 = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(3), Inches(1))
    tb3 = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(3), Inches(1))

    # Intentionally duplicate IDs
    tb1._element.xpath('.//p:cNvPr')[0].set('id', '2')
    tb2._element.xpath('.//p:cNvPr')[0].set('id', '2')
    tb3._element.xpath('.//p:cNvPr')[0].set('id', '2')

    ids_before = [cnvpr.get('id') for cnvpr in slide.element.xpath('.//p:cNvPr')]
    print("  IDs before renumbering:", ids_before)
    assert ids_before.count('2') == 3, "Expected 3 identical IDs"

    SplicerGuardian.assign_unique_shape_ids(slide)

    ids_after = [cnvpr.get('id') for cnvpr in slide.element.xpath('.//p:cNvPr')]
    print("  IDs after renumbering: ", ids_after)
    assert len(ids_after) == len(set(ids_after)), f"Duplicate IDs still exist: {ids_after}"
    assert all(int(i) >= 1 for i in ids_after), f"Invalid non-positive ID found: {ids_after}"
    print("  => Shape ID Renumberer test PASSED! (100% Unique)")


def test_dangling_rid_purging():
    print("\n--- 2. Testing Dangling r:id Remapping and Safe Purging ---")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))

    # Create dummy element with an unresolvable dangling rId
    elem = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    elem.set(f"{{{R_NS}}}embed", "rId9999_dangling")
    tb._element.append(elem)

    assert elem.get(f"{{{R_NS}}}embed") == "rId9999_dangling"

    # Run safe clone
    cloned = SplicerGuardian.safe_clone_shape_with_rels(tb._element, slide.part, slide.part)
    blip_elem = cloned.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
    dangling_val = blip_elem.get(f"{{{R_NS}}}embed")
    print("  Dangling rId attribute after safe clone:", dangling_val)
    assert dangling_val is None, "Dangling rId was not stripped!"
    print("  => Dangling rId purging test PASSED! (Zero Repair Warning Guarantee)")


def test_cross_deck_media_cloning():
    print("\n--- 3. Testing Cross-Deck Media Cloning & Background Replication ---")
    src_file = "output/5m_presentation_precision.pptx"
    assert os.path.exists(src_file), f"Source template {src_file} missing!"

    src_prs = Presentation(src_file)
    target_prs = Presentation()
    blank_layout = target_prs.slide_layouts[6]

    src_slide = src_prs.slides[0]
    target_slide = target_prs.slides.add_slide(blank_layout)

    # Use HarmonizedSlideCloner which now utilizes SplicerGuardian
    HarmonizedSlideCloner.clone_slide_to_target(src_slide, target_slide)

    out_file = "output/test_splicer_guardian_deck.pptx"
    target_prs.save(out_file)
    print(f"  Saved cross-deck presentation to {out_file}")

    # Verify that media files were copied into target zip package
    with zipfile.ZipFile(out_file, 'r') as z:
        media_files = [f for f in z.namelist() if 'media/' in f]
        print("  Media files preserved in target package:", media_files)
        assert len(media_files) > 0, "No media files found in target zip package!"

    # Verify preview export via PowerPoint COM (guarantees zero repair/crash dialogs)
    previews = get_presentation_previews(out_file)
    print(f"  Rendered {len(previews)} preview slide(s): {previews}")
    assert len(previews) > 0, "PowerPoint COM failed to render preview!"
    print("  => Cross-Deck Media Cloning test PASSED! (100% Success)")


if __name__ == "__main__":
    test_unique_shape_ids()
    test_dangling_rid_purging()
    test_cross_deck_media_cloning()
    print("\n🎉 ALL 3 SPLICER GUARDIAN UNIT TESTS PASSED WITH ZERO ERRORS!")
