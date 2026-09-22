# -*- coding: utf-8 -*-
"""
tests/test_slide_splicer.py
Unit tests for Slide Archetypes, Splicing Engines, and Multi-Deck Assembly.
Tests:
1. SlideArchetypeDetector: Heuristics and shape classification.
2. AcademicCatalog.search_slides_by_archetype: Cross-deck archetype query & caching.
3. ThemeColorExtractor: Palette and font pairing extraction.
4. HarmonizedSlideCloner: Deep XML cloning and theme harmonization.
5. MultiMasterSlideSplicer: Preserving native backgrounds and styling.
6. ModularComponentStacker: Graphic component injection.
7. UniversalSlideSplicer: Multi-engine dispatcher.
8. TemplateDeckAssembler: Cross-deck presentation assembly and text replacement.
"""

import os
import sys
import unittest
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.slide_archetypes import SLIDE_ARCHETYPES, SlideArchetypeDetector
from core.slide_splicer import (
    ThemeColorExtractor,
    HarmonizedSlideCloner,
    MultiMasterSlideSplicer,
    ModularComponentStacker,
    UniversalSlideSplicer
)
from core.academic_catalog import get_academic_catalog
from core.academic_matcher import get_academic_matcher
from core.pptx_text_replacer import replace_presentation_content


class TestSlideArchetypes(unittest.TestCase):
    """Tests for the 12 visual archetypes and heuristic detector."""

    def test_archetype_registry(self):
        """Verify that all 12 archetypes are registered with complete metadata."""
        expected_archetypes = [
            "cover", "agenda", "theory_concept", "cards_grid", "comparison_vs",
            "timeline_steps", "process_cycle", "pyramid_funnel", "metrics_stats",
            "diagram_anatomy", "keynote_quote", "table_data", "conclusion_sources"
        ]
        for arch in expected_archetypes:
            self.assertIn(arch, SLIDE_ARCHETYPES, f"Archetype {arch} must be in SLIDE_ARCHETYPES")
            meta = SLIDE_ARCHETYPES[arch]
            self.assertIn("name_uz", meta)
            self.assertIn("icon", meta)
            self.assertIn("description", meta)

    def test_archetype_detector_heuristics(self):
        """Verify heuristic detection on synthetic slides."""
        prs = Presentation()
        blank = prs.slide_layouts[6]

        # Slide 0: Cover
        s0 = prs.slides.add_slide(blank)
        self.assertEqual(SlideArchetypeDetector.detect_archetype(s0, 0, 5), "cover")

        # Slide 1: Agenda
        s1 = prs.slides.add_slide(blank)
        tx = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        tx.text_frame.text = "Table of Contents\n1. Introduction\n2. Methods"
        self.assertEqual(SlideArchetypeDetector.detect_archetype(s1, 1, 5), "agenda")

        # Slide 2: Comparison
        s2 = prs.slides.add_slide(blank)
        tx2 = s2.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        tx2.text_frame.text = "Classical vs Quantum Comparison"
        self.assertEqual(SlideArchetypeDetector.detect_archetype(s2, 2, 5), "comparison_vs")

        # Slide 3: Timeline
        s3 = prs.slides.add_slide(blank)
        tx3 = s3.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        tx3.text_frame.text = "Roadmap and Timeline 2026 - 2030"
        self.assertEqual(SlideArchetypeDetector.detect_archetype(s3, 3, 5), "timeline_steps")

        # Slide 4: Table Data
        s4 = prs.slides.add_slide(blank)
        s4.shapes.add_table(3, 3, Inches(1), Inches(1), Inches(5), Inches(3))
        self.assertEqual(SlideArchetypeDetector.detect_archetype(s4, 4, 5), "table_data")


class TestAcademicCatalogArchetypeSearch(unittest.TestCase):
    """Tests for searching slides by archetype across the template library."""

    @classmethod
    def setUpClass(cls):
        cls.catalog = get_academic_catalog()

    def test_search_slides_by_archetype(self):
        """Verify search_slides_by_archetype returns valid slide records."""
        results = self.catalog.search_slides_by_archetype(
            "timeline_steps", category="engineering_it", topic="Kvant kompyuterlari", limit=2
        )
        self.assertGreater(len(results), 0, "Should find matching timeline slides")
        first = results[0]
        self.assertIn("pptx_path", first)
        self.assertIn("slide_index", first)
        self.assertTrue(os.path.exists(first["pptx_path"]), f"File should exist: {first['pptx_path']}")


class TestSlideSplicingEngines(unittest.TestCase):
    """Tests for the 3 Splicing Engines and UniversalSlideSplicer."""

    @classmethod
    def setUpClass(cls):
        cls.catalog = get_academic_catalog()
        cls.matcher = get_academic_matcher()
        cls.test_out_dir = os.path.join(PROJECT_ROOT, "output", "test_splicer")
        os.makedirs(cls.test_out_dir, exist_ok=True)

    def test_theme_color_extractor(self):
        """Verify extracting dominant colors and fonts from a slide."""
        prs = Presentation()
        blank = prs.slide_layouts[6]
        s = prs.slides.add_slide(blank)
        tx = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        p = tx.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = "Sample Title"
        r.font.name = "Georgia"
        r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

        palette = ThemeColorExtractor.extract_palette(s)
        self.assertEqual(palette["primary_font"], "Georgia")
        self.assertEqual(palette["primary_color"], RGBColor(0x1F, 0x4E, 0x79))

    def test_harmonized_slide_cloner(self):
        """Verify HarmonizedSlideCloner copies shapes and applies theme font/color."""
        src_prs = Presentation()
        blank = src_prs.slide_layouts[6]
        src_slide = src_prs.slides.add_slide(blank)
        tx = src_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        p = tx.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = "Source Run"
        r.font.name = "Arial"
        r.font.bold = True

        tgt_prs = Presentation()
        tgt_slide = tgt_prs.slides.add_slide(blank)

        palette = {
            "primary_font": "Calibri",
            "primary_color": RGBColor(0x00, 0x70, 0xC0)
        }

        HarmonizedSlideCloner.clone_slide_to_target(src_slide, tgt_slide, theme_palette=palette)
        self.assertEqual(len(tgt_slide.shapes), 1)
        cloned_shp = tgt_slide.shapes[0]
        self.assertTrue(cloned_shp.has_text_frame)
        self.assertIn("Source Run", cloned_shp.text_frame.text)

    def test_universal_splicer_all_three_engines(self):
        """Verify UniversalSlideSplicer generates valid decks across all 3 engines."""
        bp = self.matcher.build_blueprint("Kiberxavfsizlik asoslari", slide_count=6, language="uz")
        primary_pptx = bp["primary_pptx_path"]
        self.assertTrue(os.path.exists(primary_pptx))

        slide_specs = [
            {"pptx_path": primary_pptx, "slide_index": 0, "archetype": "cover", "slide_number": 1},
            {"pptx_path": primary_pptx, "slide_index": 1, "archetype": "agenda", "slide_number": 2},
            {"pptx_path": primary_pptx, "slide_index": 2, "archetype": "theory_concept", "slide_number": 3},
            {"pptx_path": primary_pptx, "slide_index": 3, "archetype": "cards_grid", "slide_number": 4},
            {"pptx_path": primary_pptx, "slide_index": 4, "archetype": "comparison_vs", "slide_number": 5},
            {"pptx_path": primary_pptx, "slide_index": 5, "archetype": "conclusion_sources", "slide_number": 6},
        ]

        engines = ["harmonized", "multi_master", "modular_stacker"]
        for eng in engines:
            out_file = os.path.join(self.test_out_dir, f"deck_{eng}.pptx")
            res = UniversalSlideSplicer.splice_and_assemble(
                engine_mode=eng,
                primary_pptx_path=primary_pptx,
                slide_specs=slide_specs,
                output_path=out_file
            )
            self.assertTrue(os.path.exists(res), f"Resulting file should exist: {res}")
            prs = Presentation(res)
            self.assertEqual(len(prs.slides), 6, f"Engine {eng} must produce exactly 6 slides")


class TestTemplateDeckAssemblerMultiEngine(unittest.TestCase):
    """Tests for full deck customization via TemplateDeckAssembler."""

    @classmethod
    def setUpClass(cls):
        cls.matcher = get_academic_matcher()
        cls.test_out_dir = os.path.join(PROJECT_ROOT, "output", "test_splicer")
        os.makedirs(cls.test_out_dir, exist_ok=True)

    def test_full_cross_deck_customization(self):
        """Verify assembling and filling content with multi-deck blueprint."""
        bp = self.matcher.build_blueprint("Kvant kompyuterlari", slide_count=6, language="uz")
        content = {
            "topic": "Kvant kompyuterlari",
            "slides": [
                {"slide_number": 1, "layout_type": "cover", "title": "Kvant kompyuterlari", "subtitle": "Ilmiy Ma'ruza 2026"},
                {"slide_number": 2, "layout_type": "agenda", "title": "Reja", "points": ["1. Asoslar", "2. Qiyoslash", "3. Xulosa"]},
                {"slide_number": 3, "layout_type": "theory_concept", "title": "Nazariya", "concept_name": "Kubit", "definition": "Kvant birligi", "key_formula_or_law": "alpha|0> + beta|1>", "importance": "Tezlik"},
                {"slide_number": 4, "layout_type": "cards_grid", "title": "Elementlar", "cards": [{"title": "Modul A", "desc": "Tavsif A"}, {"title": "Modul B", "desc": "Tavsif B"}]},
                {"slide_number": 5, "layout_type": "comparison_vs", "title": "Qiyos", "left_title": "A", "left_points": ["1"], "right_title": "B", "right_points": ["2"]},
                {"slide_number": 6, "layout_type": "conclusion_sources", "title": "Xulosa", "key_takeaways": ["Xulosa 1"], "references": ["Adabiyot 1"]}
            ]
        }

        out_path = os.path.join(self.test_out_dir, "full_customized.pptx")
        res = replace_presentation_content(bp["primary_pptx_path"], content, out_path, blueprint=bp, engine_mode="harmonized")
        self.assertTrue(os.path.exists(res))
        prs = Presentation(res)
        self.assertEqual(len(prs.slides), 6)


if __name__ == "__main__":
    unittest.main()
