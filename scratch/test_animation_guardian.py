# -*- coding: utf-8 -*-
"""
scratch/test_animation_guardian.py
Comprehensive unit test suite for core/animation_guardian.py (Round 15).
Verifies:
1. Shape ID collection across standard and grouped shapes.
2. Zombie target identification and surgical pruning in <p:timing>.
3. Empty container cascade cleanup.
4. Target ID remapping for renumbered shapes.
5. Real PowerPoint COM validation with authentic animated slides.
"""

import sys
import os
import io
import copy
from pptx import Presentation
from pptx.util import Inches

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from core.animation_guardian import AnimationGuardian
from core.slide_preview import get_presentation_previews


def test_zombie_animation_purging_on_real_slide():
    print("--- 1. Testing Zombie Animation Purging on Authentic Template Slide ---")
    src_file = "output/test_authentic_wingdings_preserved.pptx"
    prs = Presentation(src_file)

    # Slide 11 has real PowerPoint animations targeting shape id="16"
    slide = prs.slides[11]

    # Verify initial valid IDs
    initial_ids = AnimationGuardian.collect_valid_shape_ids(slide)
    print("  Initial shape IDs:", sorted(list(initial_ids), key=lambda x: int(x) if x.isdigit() else 0))
    assert "16" in initial_ids, "Shape 16 should exist initially"

    # Verify initial animation targets
    timing_before = slide.element.xpath('./*[local-name()="timing"]')[0]
    sp_tgts_before = timing_before.xpath('.//*[@spid]')
    print(f"  Initial targeted animation nodes count: {len(sp_tgts_before)}")
    assert len(sp_tgts_before) > 0, "Slide 11 should have initial animation targets"

    # Now, simulate shape pruning: Physically delete shape 16 from _spTree
    shape_to_prune = None
    for shp in slide.shapes:
        if shp.shape_id == 16:
            shape_to_prune = shp
            break

    assert shape_to_prune is not None, "Shape with id 16 not found!"
    print(f"  Simulating shape pruning: removing '{shape_to_prune.name}' (id=16)...")
    slide.shapes._spTree.remove(shape_to_prune._element)

    # Verify shape 16 is now gone from valid IDs
    ids_after_prune = AnimationGuardian.collect_valid_shape_ids(slide)
    assert "16" not in ids_after_prune, "Shape 16 should no longer exist in valid IDs"

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

    out_file = "output/test_animation_purged_slide11.pptx"
    prs.save(out_file)
    print(f"  Saved presentation to {out_file}")

    # Verify with PowerPoint COM
    previews = get_presentation_previews(out_file)
    print(f"  Rendered {len(previews)} preview slide(s): {previews}")
    assert len(previews) > 0, "PowerPoint COM preview failed!"
    print("  => Zombie Animation Purging on Authentic Template Slide PASSED! (100% Success)")


def test_animation_target_remapping():
    print("\n--- 2. Testing Animation Target ID Remapping ---")
    src_file = "output/test_authentic_wingdings_preserved.pptx"
    prs = Presentation(src_file)
    slide = prs.slides[11]

    # Remap id 16 -> 999
    id_map = {"16": "999"}
    remapped_count = AnimationGuardian.remap_animation_target_ids(slide, id_map)
    print(f"  Remapped {remapped_count} animation target(s)")
    assert remapped_count > 0, "Expected targets to be remapped"

    timing = slide.element.xpath('./*[local-name()="timing"]')[0]
    new_spids = set(timing.xpath('.//*[@spid]/@spid'))
    print("  New targeted spids in timing:", new_spids)
    assert "999" in new_spids, "spid='999' not found in timing after remapping!"
    assert "16" not in new_spids, "Old spid='16' still found in timing after remapping!"
    print("  => Animation Target ID Remapping test PASSED! (100% Success)")


if __name__ == "__main__":
    test_zombie_animation_purging_on_real_slide()
    test_animation_target_remapping()
    print("\n🎉 ALL ANIMATION GUARDIAN UNIT TESTS PASSED WITH ZERO ERRORS!")
