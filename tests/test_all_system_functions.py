# -*- coding: utf-8 -*-
"""
Unittest wrapper for the complete SlaydHub AI automated function test suite.
Can be executed with standard test runners:
    python -m unittest discover tests
    python tests/test_all_system_functions.py
"""

import unittest
import sys
import os

# Add root directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from test_all_functions import (
    TestReporter,
    test_academic_catalog,
    test_academic_matcher,
    test_prompt_generator,
    test_json_parser_and_repair,
    test_topic_logic_validator,
    test_watermark_and_ad_purging,
    test_visual_archetype_renderers,
    test_spatial_geometry_and_bounds,
    test_slide_preview_engine,
    test_end_to_end_pipeline,
)


class TestAllSystemFunctions(unittest.TestCase):
    """Executes each module test suite and asserts all tests pass."""

    def setUp(self):
        self.reporter = TestReporter()

    def _assert_no_failures(self, module_name: str):
        mod_results = [r for r in self.reporter.results if r["module"] == module_name]
        self.assertGreater(len(mod_results), 0, f"No tests were run for {module_name}")
        failures = [r for r in mod_results if not r["passed"]]
        if failures:
            err_msg = f"{module_name} had {len(failures)} failures:\n" + "\n".join(
                f"- {r['test']}: {r['message']}" for r in failures
            )
            self.assertEqual(len(failures), 0, err_msg)

    def test_01_academic_catalog(self):
        test_academic_catalog(self.reporter)
        self._assert_no_failures("Academic Catalog")

    def test_02_academic_matcher(self):
        test_academic_matcher(self.reporter)
        self._assert_no_failures("Academic Matcher")

    def test_03_prompt_generator(self):
        test_prompt_generator(self.reporter)
        self._assert_no_failures("Prompt Generator")

    def test_04_json_parser_and_repair(self):
        test_json_parser_and_repair(self.reporter)
        self._assert_no_failures("JSON Parser & Repair")

    def test_05_topic_logic_validator(self):
        test_topic_logic_validator(self.reporter)
        self._assert_no_failures("Topic Logic Validator")

    def test_06_watermark_and_ad_purging(self):
        test_watermark_and_ad_purging(self.reporter)
        self._assert_no_failures("Watermark & Ad Purger")

    def test_07_visual_archetype_renderers(self):
        test_visual_archetype_renderers(self.reporter)
        self._assert_no_failures("Visual Card Renderers")

    def test_08_spatial_geometry_and_bounds(self):
        test_spatial_geometry_and_bounds(self.reporter)
        self._assert_no_failures("Spatial Geometry & Bounds")

    def test_09_slide_preview_engine(self):
        test_slide_preview_engine(self.reporter)
        self._assert_no_failures("Slide Preview Engine")

    def test_10_end_to_end_pipeline(self):
        test_end_to_end_pipeline(self.reporter)
        self._assert_no_failures("End-to-End Pipeline")


if __name__ == "__main__":
    unittest.main()
