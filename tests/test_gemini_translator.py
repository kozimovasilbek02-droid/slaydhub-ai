# -*- coding: utf-8 -*-
"""
tests/test_gemini_translator.py
GeminiTranslator modulining tozalash, guruhlash va xatoliklarni qayta ishlash testlari.
"""
import sys, pathlib, unittest, os
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

from backend.core.gemini_translator import sanitize_text, GeminiTranslator, detect_domain

class TestSanitizeText(unittest.TestCase):
    def test_empty(self):
        self.assertEqual(sanitize_text(""), "")
        self.assertEqual(sanitize_text(None), "")

    def test_soft_break_char(self):
        # \x0b or _x000B_ should be replaced with space so words don't stick
        self.assertEqual(sanitize_text("Leonardo\x0bDa Vinci"), "Leonardo Da Vinci")
        self.assertEqual(sanitize_text("Leonardo_x000B_Da Vinci"), "Leonardo Da Vinci")

    def test_extra_spaces_collapsed(self):
        self.assertEqual(sanitize_text("  Hello    World  "), "Hello World")

class TestDetectDomain(unittest.TestCase):
    def test_domain_detection(self):
        self.assertIn("Tibbiyot", detect_domain("Klinik anatomiya va shifoxona terapiyasi"))
        self.assertIn("Axborot", detect_domain("Sun'iy intellekt va dasturlash algoritmlari"))
        self.assertEqual(detect_domain(""), "general")

class TestGeminiTranslatorMocked(unittest.TestCase):
    @patch.dict(os.environ, {"GOOGLE_API_KEY": "", "GEMINI_API_KEY": "", "GROQ_API_KEY": ""}, clear=False)
    def test_mock_batch_translation_gemini(self):
        translator = GeminiTranslator(api_key="test_dummy_key")
        
        # Mocking genai response
        mock_response = MagicMock()
        mock_response.text = '[{"id": "item1", "translated": "Renessans dahosi"}]'
        translator.client.models.generate_content = MagicMock(return_value=mock_response)
        
        items = [{"id": "item1", "original_text": "Renaissance Genius"}]
        stats = {}
        res = translator.translate_items_batch(items, target_script="latin", stats=stats)
        
        self.assertEqual(len(res), 1)
        self.assertEqual(res[0]["id"], "item1")
        self.assertEqual(res[0]["translated_text"], "Renessans dahosi")
        self.assertEqual(stats["total"], 1)
        self.assertEqual(stats["failed"], 0)

    @patch.dict(os.environ, {"GROQ_API_KEY": "dummy_groq_key"})
    def test_mock_batch_translation_groq(self):
        translator = GeminiTranslator(api_key="test_dummy_key")
        translator._translate_via_groq = MagicMock(return_value={"item1": "Renessans dahosi"})
        
        items = [{"id": "item1", "original_text": "Renaissance Genius"}]
        stats = {}
        res = translator.translate_items_batch(items, target_script="latin", stats=stats)
        
        self.assertEqual(len(res), 1)
        self.assertEqual(res[0]["id"], "item1")
        self.assertEqual(res[0]["translated_text"], "Renessans dahosi")
        self.assertEqual(stats["total"], 1)
        self.assertEqual(stats["failed"], 0)

    @patch.dict(os.environ, {"GOOGLE_API_KEY": "", "GEMINI_API_KEY": ""}, clear=False)
    def test_empty_items(self):
        translator = GeminiTranslator(api_key="test_dummy_key")
        stats = {}
        res = translator.translate_items_batch([], target_script="latin", stats=stats)
        self.assertEqual(res, [])
        self.assertEqual(stats["total"], 0)

if __name__ == '__main__':
    unittest.main()

