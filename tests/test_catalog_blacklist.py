# -*- coding: utf-8 -*-
"""
tests/test_catalog_blacklist.py
Tests that low quality scraped folders (Tarjima_Qilingan_Top50 and Yuklab_Olingan_Top50)
are strictly blacklisted from the template catalog, archetype searches, and index files.
"""

import unittest
from core.academic_catalog import AcademicCatalog, get_academic_catalog


class TestCatalogBlacklist(unittest.TestCase):

    def test_is_valid_base_template_rejects_blacklisted_folders(self):
        cat = AcademicCatalog
        # Should reject Tarjima_Qilingan_Top50
        p1 = r"C:\Users\user\Desktop\slayd\Tarjima_Qilingan_Top50\01_test\slide.pptx"
        self.assertFalse(cat.is_valid_base_template(p1))

        # Should reject Yuklab_Olingan_Top50
        p2 = r"C:\Users\user\Desktop\slayd\Yuklab_Olingan_Top50\01_test\slide.pptx"
        self.assertFalse(cat.is_valid_base_template(p2))

        # Should accept legitimate templates
        p3 = r"C:\Users\user\Desktop\slayd\FreePPT7\General\Quantum.pptx"
        self.assertTrue(cat.is_valid_base_template(p3))

    def test_catalog_contains_zero_blacklisted_templates(self):
        cat = get_academic_catalog()
        for tmpl in cat.templates:
            p = tmpl.get("pptx_path", "").lower()
            self.assertNotIn("tarjima_qilingan", p, f"Found blacklisted template: {p}")
            self.assertNotIn("yuklab_olingan", p, f"Found blacklisted template: {p}")

    def test_search_archetypes_returns_no_blacklisted_slides(self):
        cat = get_academic_catalog()
        for arch in ["cover", "theory_concept", "comparison_vs", "timeline_steps", "diagram_anatomy"]:
            slides = cat.search_slides_by_archetype(arch, limit=3)
            for s in slides:
                p = s.get("pptx_path", "").lower()
                self.assertNotIn("tarjima_qilingan", p)
                self.assertNotIn("yuklab_olingan", p)


if __name__ == "__main__":
    unittest.main()
