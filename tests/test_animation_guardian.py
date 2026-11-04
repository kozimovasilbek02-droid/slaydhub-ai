# -*- coding: utf-8 -*-
"""
tests/test_animation_guardian.py
Unit tests for AnimationGuardian: Timing Tree Purger & Zombie Animation Destroyer (Round 15).
"""

import unittest
import sys
from pathlib import Path
from pptx import Presentation

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from core.animation_guardian import AnimationGuardian
from core.slide_preview import get_presentation_previews


class TestAnimationGuardian(unittest.TestCase):
    """Verifies that AnimationGuardian eliminates zombie animation targets without corruption."""

    def test_zombie_animation_purging_real_template(self):
        src_file = str(REPO_ROOT / "templates" / "09_Digital_Healthcare.pptx")
        prs = Presentation(src_file)
        slide = prs.slides[7]  # Slide 8 has animation targeting shape id 17

        initial_ids = AnimationGuardian.collect_valid_shape_ids(slide)
        self.assertIn("17", initial_ids)

        # Delete shape 17
        shape_to_remove = None
        for s in slide.shapes:
            if s.shape_id == 17:
                shape_to_remove = s
                break
        self.assertIsNotNone(shape_to_remove)
        slide.shapes._spTree.remove(shape_to_remove._element)

        # Purge zombie animations
        purged = AnimationGuardian.purge_zombie_animations(slide)
        self.assertGreater(purged, 0)

        # Verify no remaining zombies
        timings = slide.element.xpath('./*[local-name()="timing"]')
        if timings:
            remaining_spids = set(timings[0].xpath('.//*[@spid]/@spid'))
            self.assertNotIn("17", remaining_spids)

        # Save and verify with PowerPoint COM
        out_file = str(REPO_ROOT / "output" / "test_guardian_ci_slide8.pptx")
        prs.save(out_file)
        previews = get_presentation_previews(out_file)
        self.assertGreater(len(previews), 0)

    def test_animation_target_remapping(self):
        src_file = str(REPO_ROOT / "templates" / "09_Digital_Healthcare.pptx")
        prs = Presentation(src_file)
        slide = prs.slides[7]

        id_map = {"17": "888"}
        remapped = AnimationGuardian.remap_animation_target_ids(slide, id_map)
        self.assertGreater(remapped, 0)

        timing = slide.element.xpath('./*[local-name()="timing"]')[0]
        spids = set(timing.xpath('.//*[@spid]/@spid'))
        self.assertIn("888", spids)
        self.assertNotIn("17", spids)


if __name__ == "__main__":
    unittest.main()
