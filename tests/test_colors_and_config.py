# -*- coding: utf-8 -*-
"""
tests/test_colors_and_config.py
Unit tests for core.config and core.utils.colors.
"""

import unittest
from pathlib import Path
from pptx.dml.color import RGBColor

from core.config import config, is_folder_blacklisted
from core.utils.colors import (
    hex_to_rgb,
    rgb_to_hex,
    get_luminance,
    get_contrast_text_color,
    ACADEMIC_PALETTES,
)


class TestConfig(unittest.TestCase):
    """Tests for core.config."""

    def test_base_dir_resolution(self):
        self.assertTrue(config.BASE_DIR.exists())
        self.assertTrue((config.BASE_DIR / "core").is_dir())

    def test_output_dirs(self):
        self.assertTrue(config.OUTPUT_STUDIO_DIR.exists())

    def test_blacklist_filter(self):
        # Blacklisted folders
        bad_path1 = "C:/Users/user/Desktop/slayd/Tarjima_Qilingan_Top50/deck.pptx"
        bad_path2 = "G:/Shared/slayd/yuklab_olingan_top50/presentation.pptx"
        self.assertTrue(is_folder_blacklisted(bad_path1))
        self.assertTrue(is_folder_blacklisted(bad_path2))

        # Allowed folders
        good_path1 = "C:/Users/user/Desktop/slayd/Slidenest/deck.pptx"
        good_path2 = "G:/Мой диск/Powepoint/PresentationGO/Prezentatsiyalar/good.pptx"
        self.assertFalse(is_folder_blacklisted(good_path1))
        self.assertFalse(is_folder_blacklisted(good_path2))


class TestColors(unittest.TestCase):
    """Tests for core.utils.colors."""

    def test_hex_to_rgb_with_hash(self):
        c = hex_to_rgb("#1A365D")
        self.assertEqual((c[0], c[1], c[2]), (26, 54, 93))

    def test_hex_to_rgb_without_hash(self):
        c = hex_to_rgb("FFFFFF")
        self.assertEqual((c[0], c[1], c[2]), (255, 255, 255))

    def test_hex_to_rgb_short(self):
        c = hex_to_rgb("#FFF")
        self.assertEqual((c[0], c[1], c[2]), (255, 255, 255))

    def test_hex_to_rgb_invalid_fallback(self):
        c = hex_to_rgb("invalid-hex", default=(0, 0, 0))
        self.assertEqual((c[0], c[1], c[2]), (0, 0, 0))
        c_default = hex_to_rgb("invalid-hex")
        self.assertEqual((c_default[0], c_default[1], c_default[2]), (30, 41, 59))

    def test_rgb_to_hex(self):
        self.assertEqual(rgb_to_hex(RGBColor(26, 54, 93)), "#1A365D")
        self.assertEqual(rgb_to_hex((255, 255, 255)), "#FFFFFF")

    def test_contrast_text_color(self):
        # Dark blue -> light text (pure white RGBColor)
        white_c = get_contrast_text_color("#1A365D")
        self.assertEqual((white_c[0], white_c[1], white_c[2]), (255, 255, 255))
        # Bright yellow -> dark text (dark slate RGBColor)
        dark_c = get_contrast_text_color("#FFFF00")
        self.assertEqual((dark_c[0], dark_c[1], dark_c[2]), (15, 23, 42))

    def test_palettes_presence(self):
        self.assertIn("deep_navy", ACADEMIC_PALETTES)
        self.assertIn("emerald_academic", ACADEMIC_PALETTES)
        self.assertIn("royal_indigo", ACADEMIC_PALETTES)
        self.assertIn("ruby_crimson", ACADEMIC_PALETTES)
        self.assertIn("corporate_slate", ACADEMIC_PALETTES)


if __name__ == "__main__":
    unittest.main()
