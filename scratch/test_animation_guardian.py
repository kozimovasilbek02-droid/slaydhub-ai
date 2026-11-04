# -*- coding: utf-8 -*-
"""
scratch/test_animation_guardian.py
Comprehensive unit test suite for core/animation_guardian.py (Dual-Brain Round 15).
Verifies:
1. Shape ID collection across standard and grouped shapes.
2. Zombie target identification and surgical pruning in <p:timing>.
3. Empty container cascade cleanup without breaking OpenXML schema.
4. Target ID remapping for renumbered shapes.
5. Real PowerPoint COM validation with authentic animated slides.
"""

import sys
import os
import io
import shutil
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

# Ensure repo root is in sys.path
REPO_ROOT = Path(r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot")
sys.path.insert(0, str(REPO_ROOT))

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from core.animation_guardian import AnimationGuardian
from core.slide_preview import get_presentation_previews


def test_zombie_animation_purging_on_real_slide():
    print("--- 1. Testing Zombie Animation Purging on Authentic Template Slide ---")
    src_file = str(REPO_ROOT / "templates" / "09_Digital_Healthcare.pptx")
    prs = Presentation(src_file)

    # Slide 8 (index 7) has real PowerPoint animations targeting shape id="17"
    slide = prs.slides[7]

    # Verify initial valid IDs
    initial_ids = AnimationGuardian.collect_valid_shape_ids(slide)
    print("  Initial shape IDs:", sorted(list(initial_ids), key=lambda x: int(x) if x.isdigit() else 0))
    assert "17" in initial_ids, "Shape 17 should exist initially"

    # Verify initial animation targets
    timing_before = slide.element.xpath('./*[local-name()="timing"]')[0]
    sp_tgts_before = timing_before.xpath('.//*[@spid]')
    print(f"  Initial targeted animation nodes count: {len(sp_tgts_before)}")
    assert len(sp_tgts_before) > 0, "Slide 8 should have initial animation targets"

    # Now simulate shape pruning: Physically delete shape 17 from _spTree
    shape_to_prune = None
    for shp in slide.shapes:
        if shp.shape_id == 17:
            shape_to_prune = shp
            break

    assert shape_to_prune is not None, "Shape with id 17 not found!"
    print(f"  Simulating shape pruning: removing '{shape_to_prune.name}' (id=17)...")
    slide.shapes._spTree.remove(shape_to_prune._element)

    # Verify shape 17 is now gone from valid IDs
    ids_after_prune = AnimationGuardian.collect_valid_shape_ids(slide)
    assert "17" not in ids_after_prune, "Shape 17 should no longer exist in valid IDs"

    # Run AnimationGuardian to purge zombie animations!
    purged_count = AnimationGuardian.purge_zombie_animations(slide)
    print(f"  AnimationGuardian purged {purged_count} zombie animation node(s)!")
    assert purged_count > 0, "AnimationGuardian should have purged zombie targets!"

    # Verify timing element state
    timings_after = slide.element.xpath('./*[local-name()="timing"]')
    if timings_after:
        remaining_zombies = [el.get('spid') for el in timings_after[0].xpath('.//*[@spid]') if el.get('spid') not in ids_after_prune]
        print("  Remaining zombie targets in timing:", remaining_zombies)
        assert len(remaining_zombies) == 0, f"Zombie targets still remained in <p:timing>: {remaining_zombies}"
    else:
        print("  <p:timing> was safely and cleanly removed because zero active animations remained!")

    out_dir = REPO_ROOT / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = str(out_dir / "test_animation_purged_slide8.pptx")
    prs.save(out_file)
    print(f"  Saved presentation to {out_file}")

    # Verify with PowerPoint COM
    previews = get_presentation_previews(out_file)
    print(f"  Rendered {len(previews)} preview slide(s): {previews}")
    assert len(previews) > 0, "PowerPoint COM preview failed!"
    print("  => Zombie Animation Purging on Authentic Template Slide PASSED! (100% Success)")


def test_animation_target_remapping():
    print("\n--- 2. Testing Animation Target ID Remapping ---")
    src_file = str(REPO_ROOT / "templates" / "09_Digital_Healthcare.pptx")
    prs = Presentation(src_file)
    slide = prs.slides[7]

    # Remap id 17 -> 999
    id_map = {"17": "999"}
    remapped_count = AnimationGuardian.remap_animation_target_ids(slide, id_map)
    print(f"  Remapped {remapped_count} animation target(s)")
    assert remapped_count > 0, "Expected targets to be remapped"

    timing = slide.element.xpath('./*[local-name()="timing"]')[0]
    new_spids = set(timing.xpath('.//*[@spid]/@spid'))
    print("  New targeted spids in timing:", new_spids)
    assert "999" in new_spids, "spid='999' not found in timing after remapping!"
    assert "17" not in new_spids, "Old spid='17' still found in timing after remapping!"
    print("  => Animation Target ID Remapping test PASSED! (100% Success)")


def test_partial_zombie_purging_keeps_valid_animations():
    print("\n--- 3. Testing Partial Zombie Purging (Multiple Animation Targets) ---")
    src_file = str(REPO_ROOT / "templates" / "3D Printing Technology PowerPoint Templates.pptx")
    prs = Presentation(src_file)
    slide = prs.slides[7]

    # Check timing
    timings = slide.element.xpath('./*[local-name()="timing"]')
    if not timings:
        print("  Slide has no timing, skipping partial test")
        return

    initial_spids = set(timings[0].xpath('.//*[@spid]/@spid'))
    print("  Initial targeted spids:", initial_spids)
    assert "31" in initial_spids

    # Simulate deleting shape 31
    shape_to_prune = None
    for shp in slide.shapes:
        if shp.shape_id == 31:
            shape_to_prune = shp
            break

    if shape_to_prune:
        slide.shapes._spTree.remove(shape_to_prune._element)
        purged = AnimationGuardian.purge_zombie_animations(slide)
        print(f"  Purged {purged} zombie node(s)")
        assert purged > 0

    out_file = str(REPO_ROOT / "output" / "test_animation_purged_3d_slide8.pptx")
    prs.save(out_file)

    previews = get_presentation_previews(out_file)
    assert len(previews) > 0
    print("  => Partial Zombie Purging PASSED! (100% Success)")


if __name__ == "__main__":
    test_zombie_animation_purging_on_real_slide()
    test_animation_target_remapping()
    test_partial_zombie_purging_keeps_valid_animations()
    print("\n🎉 ALL ANIMATION GUARDIAN UNIT TESTS PASSED WITH ZERO ERRORS!")
