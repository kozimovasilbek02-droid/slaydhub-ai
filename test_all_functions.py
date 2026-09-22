# -*- coding: utf-8 -*-
"""
SlaydHub AI - All-in-One Automated System Function Test Suite
Tests every major module and function across the entire presentation generation pipeline:
1. Academic Catalog & Multilingual Indexing
2. Academic Matcher & Blueprint Generation
3. Prompt Generator (Structured JSON & Compact Markdown)
4. JSON Parsing & Automatic Syntax Repair
5. Topic Logic & Content Sanitizer (Step normalization, length bounds)
6. Deep Watermark & Advertisement Purging
7. Multi-Archetype Visual Card Renderers (KPI, Timeline, Vertical Columns, Wide Cards)
8. Spatial Collision & Safe Zone Geometry
9. Slide Splicer & Theme Color Extraction
10. Slide Previews & Ultra-Fast Zip Thumbnail Extraction
11. Full End-to-End Presentation Generation Pipeline
"""

import os
import sys
import time
import json
import traceback
from pathlib import Path
from typing import Dict, Any, List

# Project Root Setup
PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# Core Modules
from core.academic_catalog import get_academic_catalog, AcademicCatalog, ACADEMIC_CATEGORIES
from core.academic_matcher import get_academic_matcher, AcademicMatcher
from core.notebooklm_prompt_gen import NotebookLMPromptGenerator, generate_notebooklm_prompt
from core.notebooklm_executor import NotebookLMExecutor
from core.topic_logic_validator import TopicLogicValidator
from core.pptx_text_replacer import (
    PPTXTextReplacer,
    DeepWatermarkPurger,
    AdSlideFilter,
    SafeShapeGeometry,
    SafeContentZoneValidator,
    SpatialCollisionEngine,
    SemanticShapeCleaner
)
from core.slide_splicer import ThemeColorExtractor, UniversalSlideSplicer
from core.slide_preview import get_template_thumbnail, get_presentation_previews
from core.template_pptx_builder import get_academic_pptx_builder


class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


class TestReporter:
    def __init__(self):
        self.results = []
        self.start_time = time.time()

    def record(self, module_name: str, test_name: str, passed: bool, message: str = "", elapsed: float = 0.0):
        self.results.append({
            "module": module_name,
            "test": test_name,
            "passed": passed,
            "message": message,
            "elapsed": elapsed
        })
        status_str = f"{Colors.GREEN}[PASS]{Colors.ENDC}" if passed else f"{Colors.FAIL}[FAIL]{Colors.ENDC}"
        time_str = f"({elapsed:.3f}s)" if elapsed > 0 else ""
        print(f"  {status_str} {test_name} {time_str}")
        if not passed and message:
            print(f"         {Colors.FAIL}Error: {message}{Colors.ENDC}")

    def print_summary(self):
        total_time = time.time() - self.start_time
        total_tests = len(self.results)
        passed_tests = sum(1 for r in self.results if r["passed"])
        failed_tests = total_tests - passed_tests

        print("\n" + "=" * 70)
        print(f"{Colors.BOLD}SLAYDHUB AI - TEST SUITE NATIJALARI{Colors.ENDC}")
        print("=" * 70)

        # Group by module
        modules = {}
        for r in self.results:
            modules.setdefault(r["module"], []).append(r)

        for mod, items in modules.items():
            mod_pass = sum(1 for i in items if i["passed"])
            mod_total = len(items)
            icon = "✅" if mod_pass == mod_total else "❌"
            print(f"{icon} {Colors.BOLD}{mod:<36}{Colors.ENDC} {mod_pass}/{mod_total} muvaffaqiyatli")

        print("-" * 70)
        print(f"Jami testlar soni:  {total_tests}")
        print(f"Muvaffaqiyatli:     {Colors.GREEN}{passed_tests}{Colors.ENDC}")
        print(f"Xatoliklar:         {Colors.FAIL if failed_tests > 0 else Colors.GREEN}{failed_tests}{Colors.ENDC}")
        print(f"Ketgan umumiy vaqt: {total_time:.2f} soniya")
        print("=" * 70)

        if failed_tests == 0:
            print(f"{Colors.GREEN}{Colors.BOLD}🎉 BARCHA FUNKSIYALAR VA MODULLAR 100% TO'G'RI ISHLAMOQDA!{Colors.ENDC}\n")
            return True
        else:
            print(f"{Colors.FAIL}{Colors.BOLD}⚠️ AYRIM FUNKSIYALARDA XATOLIKLAR ANIQLANDI!{Colors.ENDC}\n")
            return False


# ==============================================================================
# TEST SUITE IMPLEMENTATION
# ==============================================================================

def test_academic_catalog(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}1. Academic Catalog & Multilingual Indexing Testing...{Colors.ENDC}")
    mod = "Academic Catalog"
    
    # Test 1.1: Catalog instance and template loading
    t0 = time.time()
    try:
        catalog = get_academic_catalog()
        assert len(catalog.templates) >= 10, f"Kamida 10 ta shablon kutilgan, topildi: {len(catalog.templates)}"
        reporter.record(mod, "Shablonlar bazasini indekslash", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Shablonlar bazasini indekslash", False, str(e), elapsed=time.time() - t0)

    # Test 1.2: Multilingual Uzbek-English query expansion
    t0 = time.time()
    try:
        terms_kvant = catalog.expand_query_terms("kvant")
        assert "quantum" in terms_kvant, "Kvant so'zi 'quantum' ga kengaytirilmadi"
        terms_tib = catalog.expand_query_terms("tibbiyot")
        assert "medical" in terms_tib, "Tibbiyot so'zi 'medical' ga kengaytirilmadi"
        terms_ai = catalog.expand_query_terms("dasturlash")
        assert any(t in ["programming", "coding", "software"] for t in terms_ai)
        reporter.record(mod, "Ko'p tilli (O'zbek-Ingliz) so'rov kengaytirish", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Ko'p tilli (O'zbek-Ingliz) so'rov kengaytirish", False, str(e), elapsed=time.time() - t0)

    # Test 1.3: Topic classification
    t0 = time.time()
    try:
        cat_kvant = catalog.classify_topic("Kvant hisoblashlari va post-kvant xavfsizlik")
        assert cat_kvant in ["exact_sciences", "engineering_it", "stem_science"], f"Noto'g'ri kategoriya: {cat_kvant}"
        cat_med = catalog.classify_topic("Yurak qon tomir kasalliklari profilaktikasi")
        assert cat_med in ["medicine_health", "natural_sciences"], f"Tibbiyot noto'g'ri klassifikatsiya qilindi: {cat_med}"
        reporter.record(mod, "Mavzu bo'yicha ilmiy sohani aniqlash (classify_topic)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Mavzu bo'yicha ilmiy sohani aniqlash (classify_topic)", False, str(e), elapsed=time.time() - t0)

    # Test 1.4: Search accuracy and physical file resolution
    t0 = time.time()
    try:
        results = catalog.search_templates("quantum", limit=3)
        assert len(results) > 0, "Qidiruv natija bermadi"
        for r in results:
            assert os.path.exists(r["pptx_path"]), f"Shablon fayli topilmadi: {r['pptx_path']}"
        reporter.record(mod, "Shablonlarni semantik qidirish va fayl tekshiruvi", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Shablonlarni semantik qidirish va fayl tekshiruvi", False, str(e), elapsed=time.time() - t0)


def test_academic_matcher(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}2. Academic Matcher & Blueprint Planning Testing...{Colors.ENDC}")
    mod = "Academic Matcher"

    # Test 2.1: Blueprint Generation with custom presenter info
    t0 = time.time()
    try:
        matcher = get_academic_matcher()
        presenter = {
            "author_name": "Alisher Vohidov",
            "organization": "O'zMU",
            "supervisor": "Prof. A. Karimov",
            "event_name": "Himoya 2026"
        }
        bp = matcher.build_blueprint(
            topic="Kvant Kriptografiyasi",
            slide_count=7,
            language="uz",
            presenter_info=presenter
        )
        assert len(bp["slides"]) == 7, f"Kutilgan 7 slayd, amalda: {len(bp['slides'])}"
        assert bp["slides"][0]["layout_type"] == "cover", "1-slayd cover bo'lishi shart"
        assert bp["presenter_info"]["author_name"] == "Alisher Vohidov"
        reporter.record(mod, "Reja va slaydlar tuzilmasini (Blueprint) generatsiya qilish", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Reja va slaydlar tuzilmasini (Blueprint) generatsiya qilish", False, str(e), elapsed=time.time() - t0)

    # Test 2.2: Layout diversity in blueprint
    t0 = time.time()
    try:
        layouts = [s["layout_type"] for s in bp["slides"]]
        assert "cover" in layouts
        assert "agenda" in layouts or len(layouts) < 6
        assert len(set(layouts)) >= 4, f"Slaydlar turlari xilma-xil emas: {layouts}"
        reporter.record(mod, "Rejada rang-barang arxetiplar taqsimoti", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Rejada rang-barang arxetiplar taqsimoti", False, str(e), elapsed=time.time() - t0)


def test_prompt_generator(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}3. Prompt Generator (Boy Mazmun Standartlari) Testing...{Colors.ENDC}")
    mod = "Prompt Generator"

    # Test 3.1: Full structured JSON prompt generation
    t0 = time.time()
    try:
        matcher = get_academic_matcher()
        bp = matcher.build_blueprint("Sun'iy Intellekt va Neyrotarmoqlar", slide_count=6, language="uz")
        full_prompt = NotebookLMPromptGenerator.generate_prompt(bp)
        
        # Verify obsolete rules are purged
        assert "BIR SLAYD — BIR FIKR" not in full_prompt, "Eskirgan 'Bir slayd - bir fikr' qoidasi hali ham promptda mavjud!"
        assert "3-4 BLOK MEZONI" not in full_prompt, "Eskirgan 3-4 blok cheklovi hali ham promptda mavjud!"
        # Verify rich multi-archetype instructions
        assert "BOY VA KO'P QIRRALI MAZMUN" in full_prompt, "Yangi boy mazmun standarti kiritilmagan!"
        assert "RANG-BARANG VIZUAL STRUKTURALAR" in full_prompt
        assert "metrics_stats" in full_prompt
        assert "comparison_vs" in full_prompt
        reporter.record(mod, "To'liq Promptda boy mazmun va yangi standartlar", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "To'liq Promptda boy mazmun va yangi standartlar", False, str(e), elapsed=time.time() - t0)

    # Test 3.2: Compact Markdown prompt generation
    t0 = time.time()
    try:
        compact_prompt = NotebookLMPromptGenerator.generate_compact_markdown_prompt(
            topic="Biologik Membranalarning Fizik-Kimyoviy Xossalari",
            domain="Biologiya / Kimyo"
        )
        assert len(compact_prompt) > 300
        assert "BOY VA KO'P QIRRALI MAZMUN" in compact_prompt.upper()
        reporter.record(mod, "Ixcham (Compact Markdown) prompt generatsiyasi", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Ixcham (Compact Markdown) prompt generatsiyasi", False, str(e), elapsed=time.time() - t0)


def test_json_parser_and_repair(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}4. JSON Parsing & Automatic Syntax Repair Testing...{Colors.ENDC}")
    mod = "JSON Parser & Repair"

    # Test 4.1: Clean JSON parsing
    t0 = time.time()
    try:
        clean_json_str = '{"topic": "Test", "slides": [{"slide_number": 1, "layout_type": "cover"}]}'
        parsed = NotebookLMExecutor.parse_or_repair_json(clean_json_str)
        assert parsed["topic"] == "Test"
        reporter.record(mod, "Toza JSON matnini o'qish", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Toza JSON matnini o'qish", False, str(e), elapsed=time.time() - t0)

    # Test 4.2: Markdown code block wrapped JSON
    t0 = time.time()
    try:
        wrapped_json = "Quyida taqdimot kontenti keltirilgan:\n```json\n{\"topic\": \"AI\", \"slides\": []}\n```\nOmad tilayman!"
        parsed = NotebookLMExecutor.parse_or_repair_json(wrapped_json)
        assert parsed["topic"] == "AI"
        reporter.record(mod, "Markdown blokidagi (```json) matndan ajratib olish", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Markdown blokidagi (```json) matndan ajratib olish", False, str(e), elapsed=time.time() - t0)

    # Test 4.3: Syntax repair (trailing commas & unclosed braces)
    t0 = time.time()
    try:
        broken_json = '{"topic": "Repair Test", "slides": [{"title": "S1",}, {"title": "S2",},]'
        parsed = NotebookLMExecutor.parse_or_repair_json(broken_json)
        assert len(parsed["slides"]) == 2
        assert parsed["slides"][0]["title"] == "S1"
        reporter.record(mod, "Sintaktik xatolarni (trailing comma) avtomatik tuzatish", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Sintaktik xatolarni (trailing comma) avtomatik tuzatish", False, str(e), elapsed=time.time() - t0)


def test_topic_logic_validator(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}5. Topic Logic & Semantic Content Validator Testing...{Colors.ENDC}")
    mod = "Topic Logic Validator"

    # Test 5.1: Watermark and dummy text sanitization
    t0 = time.time()
    try:
        cleaned = TopicLogicValidator.sanitize_text_field("Visit slidescarnival.com for more templates")
        assert cleaned == "", f"Suvbelgi matni filtrlanmadi: {cleaned}"
        valid_text = TopicLogicValidator.sanitize_text_field("Kvant kompyuterlari arxitekturasi")
        assert valid_text == "Kvant kompyuterlari arxitekturasi"
        reporter.record(mod, "Reklama va suvbelgi matnlarini tozalash (sanitize_text_field)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Reklama va suvbelgi matnlarini tozalash (sanitize_text_field)", False, str(e), elapsed=time.time() - t0)

    # Test 5.2: Step numbering in timeline_steps (01, 02, 03 preserved)
    t0 = time.time()
    try:
        sample_deck = {
            "topic": "Post-Kvant Kripto",
            "slides": [
                {
                    "slide_number": 4,
                    "layout_type": "timeline_steps",
                    "title": "Bosqichlar",
                    "steps": [
                        {"step_number": "01", "title": "Audit", "desc": "Tahlil qilish."},
                        {"step_number": "02", "title": "Gibrid", "desc": "Sinash va ulanish."},
                        {"step_number": "03", "title": "Migratsiya", "desc": "To'liq o'tish."}
                    ]
                }
            ]
        }
        refined = TopicLogicValidator.validate_and_refine_presentation(sample_deck)
        ref_steps = refined["slides"][0]["steps"]
        assert ref_steps[0]["step_number"] == "01"
        assert ref_steps[1]["step_number"] == "02", f"2-bosqich '01' ga aylanib ketgan: {ref_steps[1]}"
        assert ref_steps[2]["step_number"] == "03", f"3-bosqich '01' ga aylanib ketgan: {ref_steps[2]}"
        reporter.record(mod, "Ketma-ket bosqichlarni to'g'ri nishonlash (01, 02, 03)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Ketma-ket bosqichlarni to'g'ri nishonlash (01, 02, 03)", False, str(e), elapsed=time.time() - t0)

    # Test 5.3: Bullet length optimization
    t0 = time.time()
    try:
        long_sentence = "Bu juda ham uzun ilmiy gap bo'lib uning ichida o'nlab so'zlar bor va slayd maydoniga sig'may ketishi xavfi mavjud bo'lgani uchun qisqartirilishi talab etiladi."
        shortened = TopicLogicValidator.optimize_bullet_length(long_sentence, max_words=10)
        assert len(shortened.split()) <= 12
        reporter.record(mod, "Matn uzunligini slaydga mos optimallashtirish (optimize_bullet_length)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Matn uzunligini slaydga mos optimallashtirish (optimize_bullet_length)", False, str(e), elapsed=time.time() - t0)


def test_watermark_and_ad_purging(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}6. Deep Watermark & Ad Purging Testing...{Colors.ENDC}")
    mod = "Watermark & Ad Purger"

    # Test 6.1: Watermark text detection
    t0 = time.time()
    try:
        assert DeepWatermarkPurger.is_watermark_text("Designed with SlidesCarnival") is True
        assert DeepWatermarkPurger.is_watermark_text("free powerpoint templates by PresentationGO") is True
        assert DeepWatermarkPurger.is_watermark_text("Kvant Kriptografiyasi Asoslari") is False
        reporter.record(mod, "Suvbelgi kalit so'zlarini aniqlash (is_watermark_text)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Suvbelgi kalit so'zlarini aniqlash (is_watermark_text)", False, str(e), elapsed=time.time() - t0)

    # Test 6.2: Ad slide filter
    t0 = time.time()
    try:
        prs = Presentation()
        s_ad = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s_ad.shapes.add_textbox(0, 0, 500000, 500000)
        tb.text_frame.text = "Credits & Resources: This presentation template is free thanks to SlidesCarnival and Pexels."
        assert AdSlideFilter.is_ad_or_credit_slide(s_ad) is True

        s_norm = prs.slides.add_slide(prs.slide_layouts[6])
        tb2 = s_norm.shapes.add_textbox(0, 0, 500000, 500000)
        tb2.text_frame.text = "Tadqiqotning amaliy ahamiyati va ilmiy yangiligi"
        assert AdSlideFilter.is_ad_or_credit_slide(s_norm) is False
        reporter.record(mod, "Reklama va minnatdorchilik slaydlarini aniqlash (AdSlideFilter)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Reklama va minnatdorchilik slaydlarini aniqlash (AdSlideFilter)", False, str(e), elapsed=time.time() - t0)

    # Test 6.3: Presentation watermarks purge
    t0 = time.time()
    try:
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb_bad = s.shapes.add_textbox(0, 0, 100000, 100000)
        tb_bad.text_frame.text = "presentationgo.com"
        tb_good = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tb_good.text_frame.text = "Asosiy Ilmiy Xulosa"

        DeepWatermarkPurger.purge_presentation_watermarks(prs)
        remaining_texts = [shp.text_frame.text for shp in s.shapes if shp.has_text_frame]
        assert "presentationgo.com" not in remaining_texts
        assert "Asosiy Ilmiy Xulosa" in remaining_texts
        reporter.record(mod, "Taqdimotdan barcha reklama shakllarini tozalash (purge_presentation_watermarks)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Taqdimotdan barcha reklama shakllarini tozalash (purge_presentation_watermarks)", False, str(e), elapsed=time.time() - t0)


def test_visual_archetype_renderers(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}7. Visual Archetype Card Renderers Testing...{Colors.ENDC}")
    mod = "Visual Card Renderers"

    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    sw_pt = 720.0
    sh_pt = 405.0
    scale_factor = 1.0
    left_m_pt = 45.0
    content_w_pt = sw_pt - (left_m_pt * 2)
    top_start_pt = 95.0

    # Test 7.1: KPI / Metric Stat Cards
    t0 = time.time()
    try:
        s_metric = prs.slides.add_slide(prs.slide_layouts[6])
        metric_data = {
            "layout_type": "metrics_stats",
            "metrics": [
                {"value": "99.9%", "label": "Barqarorlik", "desc": "NIST FIPS 203 standarti."},
                {"value": "2,500+", "label": "Kubit", "desc": "Minimal zarur quvvat."},
                {"value": "1.2 ms", "label": "Tezlik", "desc": "Gibrid ulanish vaqti."}
            ]
        }
        metric_ids = PPTXTextReplacer.render_metric_stat_cards(
            slide=s_metric,
            slide_data=metric_data,
            content_items=[],
            sw_pt=sw_pt,
            sh_pt=sh_pt,
            scale_factor=scale_factor,
            left_m_pt=left_m_pt,
            content_w_pt=content_w_pt,
            top_start_pt=top_start_pt
        )
        assert len(metric_ids) == 3, f"3 ta KPI karta kutilgan edi, amalda: {len(metric_ids)}"
        # Check that numbers are big and bold
        for shp in s_metric.shapes:
            if shp.has_text_frame:
                p0 = shp.text_frame.paragraphs[0]
                assert p0.font.bold is True
                assert p0.font.size.pt >= 28.0
        reporter.record(mod, "KPI Metrika kartalari (render_metric_stat_cards)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "KPI Metrika kartalari (render_metric_stat_cards)", False, str(e), elapsed=time.time() - t0)

    # Test 7.2: Process Step Cards (timeline_steps)
    t0 = time.time()
    try:
        s_steps = prs.slides.add_slide(prs.slide_layouts[6])
        step_data = {
            "layout_type": "timeline_steps",
            "steps": [
                {"step_number": "01", "title": "Audit", "desc": "Kalitlarni tekshirish."},
                {"step_number": "02", "title": "Gibrid", "desc": "TLS 1.3 sinovlari."},
                {"step_number": "03", "title": "To'liq Migratsiya", "desc": "PQC tizimiga o'tish."}
            ]
        }
        step_ids = PPTXTextReplacer.render_process_step_cards(
            slide=s_steps,
            slide_data=step_data,
            content_items=[],
            sw_pt=sw_pt,
            sh_pt=sh_pt,
            scale_factor=scale_factor,
            left_m_pt=left_m_pt,
            content_w_pt=content_w_pt,
            top_start_pt=top_start_pt
        )
        assert len(step_ids) == 3
        # Check badge numbers
        badges = [shp.text_frame.paragraphs[0].text for shp in s_steps.shapes if shp.has_text_frame]
        assert "STEP 01" in badges
        assert "STEP 02" in badges
        assert "STEP 03" in badges
        reporter.record(mod, "Bosqichma-bosqich jarayon kartalari (render_process_step_cards)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Bosqichma-bosqich jarayon kartalari (render_process_step_cards)", False, str(e), elapsed=time.time() - t0)

    # Test 7.3: Vertical Column Cards (cards_grid)
    t0 = time.time()
    try:
        s_cols = prs.slides.add_slide(prs.slide_layouts[6])
        items = [
            ("Lattice Kripto", "Shortest Vector Problem."),
            ("Raqamli Imzo", "ML-DSA standarti."),
            ("Infratuzilma", "PKI integratsiyasi.")
        ]
        col_ids = PPTXTextReplacer.render_vertical_column_cards(
            slide=s_cols,
            content_items=items,
            sw_pt=sw_pt,
            sh_pt=sh_pt,
            scale_factor=scale_factor,
            left_m_pt=left_m_pt,
            content_w_pt=content_w_pt,
            top_start_pt=top_start_pt
        )
        assert len(col_ids) == 3
        reporter.record(mod, "Vertikal ustunli kartalar (render_vertical_column_cards)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Vertikal ustunli kartalar (render_vertical_column_cards)", False, str(e), elapsed=time.time() - t0)

    # Test 7.4: Wide Rectangular Cards (agenda & conclusion_sources)
    t0 = time.time()
    try:
        s_wide = prs.slides.add_slide(prs.slide_layouts[6])
        agenda_items = [
            ("1. Kriptografiya asoslari", "Klassik algoritmlar"),
            ("2. Kvant xatarlari", "Shor algoritmi"),
            ("3. Post-kvant standartlari", "NIST ML-KEM"),
            ("4. Ilmiy xulosalar", "Tavsiyalar")
        ]
        wide_ids = PPTXTextReplacer.render_wide_rectangular_cards(
            slide=s_wide,
            content_items=agenda_items,
            sw_pt=sw_pt,
            sh_pt=sh_pt,
            scale_factor=scale_factor,
            left_m_pt=left_m_pt,
            content_w_pt=content_w_pt,
            top_start_pt=top_start_pt
        )
        assert len(wide_ids) == 4
        reporter.record(mod, "Keng to'g'ri to'rtburchak kartalar (render_wide_rectangular_cards)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Keng to'g'ri to'rtburchak kartalar (render_wide_rectangular_cards)", False, str(e), elapsed=time.time() - t0)

    # Test 7.5: Central Dispatcher (render_cards_by_archetype)
    t0 = time.time()
    try:
        s_dispatch = prs.slides.add_slide(prs.slide_layouts[6])
        d_ids = PPTXTextReplacer.render_cards_by_archetype(
            slide=s_dispatch,
            slide_data={"layout_type": "metrics_stats", "metrics": [{"value": "100%", "label": "Xavfsizlik"}]},
            layout_type="metrics_stats",
            content_items=[("100%", "Xavfsizlik")],
            sw_pt=sw_pt,
            sh_pt=sh_pt,
            scale_factor=scale_factor,
            left_m_pt=left_m_pt,
            content_w_pt=content_w_pt,
            top_start_pt=top_start_pt
        )
        assert len(d_ids) == 1
        reporter.record(mod, "Arxetip bo'yicha dinamik yo'naltirish (render_cards_by_archetype)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Arxetip bo'yicha dinamik yo'naltirish (render_cards_by_archetype)", False, str(e), elapsed=time.time() - t0)


def test_spatial_geometry_and_bounds(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}8. Spatial Collision & Safe Zone Geometry Testing...{Colors.ENDC}")
    mod = "Spatial Geometry & Bounds"

    # Test 8.1: Safe unit conversion to PT
    t0 = time.time()
    try:
        pt_val = SafeShapeGeometry.to_pt(Inches(2.0), 0.0)
        assert abs(pt_val - 144.0) < 0.1, f"2 dyuym 144 pt bo'lishi kerak, amalda: {pt_val}"
        reporter.record(mod, "EMU dan Point (pt) ga xavfsiz konvertatsiya", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "EMU dan Point (pt) ga xavfsiz konvertatsiya", False, str(e), elapsed=time.time() - t0)

    # Test 8.2: Safe content zone bounds check
    t0 = time.time()
    try:
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        in_shape = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        assert SafeContentZoneValidator.is_in_safe_zone(in_shape, 720.0, 405.0) is True

        out_shape = s.shapes.add_textbox(Inches(15), Inches(10), Inches(4), Inches(2))
        assert SafeContentZoneValidator.is_in_safe_zone(out_shape, 720.0, 405.0) is False
        reporter.record(mod, "Slayd maydonidan chiqib ketishni aniqlash (SafeContentZoneValidator)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Slayd maydonidan chiqib ketishni aniqlash (SafeContentZoneValidator)", False, str(e), elapsed=time.time() - t0)

    # Test 8.3: Spatial collision detection & resolution
    t0 = time.time()
    try:
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        # Two overlapping boxes
        box1 = s.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        box2 = s.shapes.add_textbox(Inches(1.5), Inches(1.2), Inches(3), Inches(1))
        
        collisions = SpatialCollisionEngine.detect_collisions([box1, box2])
        assert len(collisions) > 0, "To'qnashuv aniqlanmadi!"

        resolved = SpatialCollisionEngine.resolve_collisions([box1, box2])
        assert resolved is True or len(SpatialCollisionEngine.detect_collisions([box1, box2])) == 0
        reporter.record(mod, "To'qnashuvlarni aniqlash va ajratish (SpatialCollisionEngine)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "To'qnashuvlarni aniqlash va ajratish (SpatialCollisionEngine)", False, str(e), elapsed=time.time() - t0)


def test_slide_preview_engine(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}9. Slide Previews & Fast Caching Engine Testing...{Colors.ENDC}")
    mod = "Slide Preview Engine"

    # Test 9.1: Fast ZIP thumbnail extraction (< 20ms)
    t0 = time.time()
    try:
        catalog = get_academic_catalog()
        template_path = catalog.templates[0]["pptx_path"]
        thumb_path = get_template_thumbnail(template_path)
        assert thumb_path is not None, "Shablon thumbnaillari olinmadi"
        assert os.path.exists(thumb_path), f"Keshdagi thumbnail topilmadi: {thumb_path}"
        assert os.path.getsize(thumb_path) > 200, "Thumbnail fayli bo'sh"
        reporter.record(mod, "ZIP ichidan thumbnail ajratib olish (get_template_thumbnail)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "ZIP ichidan thumbnail ajratib olish (get_template_thumbnail)", False, str(e), elapsed=time.time() - t0)

    # Test 9.2: Full slide deck preview export via PowerPoint COM
    t0 = time.time()
    try:
        # Check on an existing sample presentation or create a fast 2-slide PPTX
        prs = Presentation()
        s1 = prs.slides.add_slide(prs.slide_layouts[6])
        s1.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text_frame.text = "Preview Test 1"
        s2 = prs.slides.add_slide(prs.slide_layouts[6])
        s2.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1)).text_frame.text = "Preview Test 2"
        
        test_pptx = str(PROJECT_ROOT / "output" / "test_preview_export.pptx")
        os.makedirs(os.path.dirname(test_pptx), exist_ok=True)
        prs.save(test_pptx)

        previews = get_presentation_previews(test_pptx, max_slides=2)
        assert len(previews) == 2, f"2 ta slayd rasmi kutilgan edi, topildi: {len(previews)}"
        for p in previews:
            assert os.path.exists(p)
            assert os.path.getsize(p) > 1000
        reporter.record(mod, "Taqdimot slaydlarini JPG formatida renderlash (get_presentation_previews)", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Taqdimot slaydlarini JPG formatida renderlash (get_presentation_previews)", False, str(e), elapsed=time.time() - t0)


def test_end_to_end_pipeline(reporter: TestReporter):
    print(f"\n{Colors.CYAN}{Colors.BOLD}10. Complete End-to-End Presentation Generation Testing...{Colors.ENDC}")
    mod = "End-to-End Pipeline"

    t0 = time.time()
    try:
        matcher = get_academic_matcher()
        bp = matcher.build_blueprint("Kvant Hisoblashlari va Post-Kvant Xavfsizlik", slide_count=6, language="uz")

        sample_content = {
            "topic": bp["topic"],
            "presentation_mode": "academic",
            "presenter_info": {
                "author_name": "Alisher Vohidov",
                "organization": "O'zMU / Axborot xavfsizligi",
                "supervisor": "DSc., prof. A. Karimov"
            },
            "slides": [
                {
                    "slide_number": 1,
                    "layout_type": "cover",
                    "title": "Kvant Hisoblashlari va Post-Kvant Xavfsizlik",
                    "subtitle": "Kvant kompyuterlari davrida axborot xavfsizligi arxitekturasi | 2026",
                    "author_name": "Alisher Vohidov",
                    "organization": "O'zbekiston Milliy Universiteti"
                },
                {
                    "slide_number": 2,
                    "layout_type": "agenda",
                    "title": "Tadqiqot Rejasi",
                    "points": [
                        "1. Kvant tahdidlari va Shor algoritmi",
                        "2. NIST PQC miqdoriy ko'rsatkichlari",
                        "3. ML-KEM arxitekturasining afzalliklari",
                        "4. Xulosalar va tavsiyalar"
                    ]
                },
                {
                    "slide_number": 3,
                    "layout_type": "metrics_stats",
                    "title": "Asosiy Miqdoriy Ko'rsatkichlar",
                    "metrics": [
                        {"value": "99.9%", "label": "Kripto-Barqarorlik", "desc": "NIST FIPS 203 standarti bo'yicha."},
                        {"value": "2,500+", "label": "Kubitlar", "desc": "Shor algoritmining zarur quvvati."},
                        {"value": "1.2 ms", "label": "Handshake", "desc": "TLS 1.3 protokoli bilan tezligi."}
                    ]
                },
                {
                    "slide_number": 4,
                    "layout_type": "timeline_steps",
                    "title": "Post-Kvant Migratsiya Bosqichlari",
                    "steps": [
                        {"step_number": "01", "title": "Audit", "desc": "Mavjud kalitlarni xatlovdan o'tkazish."},
                        {"step_number": "02", "title": "Gibrid Ulanish", "desc": "Klassik va PQC algoritmlarini birgalikda sinash."},
                        {"step_number": "03", "title": "To'liq O'tish", "desc": "Infratuzilmani yangi standartga o'tkazish."}
                    ]
                },
                {
                    "slide_number": 5,
                    "layout_type": "cards_grid",
                    "title": "Tizimning Uch Asosiy Ustuni",
                    "cards": [
                        {"title": "Panjarali Shifrlash", "desc": "SVP matematik murakkabligi."},
                        {"title": "Raqamli Imzo", "desc": "ML-DSA standarti himoyasi."},
                        {"title": "Kripto-Moslashuvchanlik", "desc": "Tezkor konfiguratsiya imkoniyati."}
                    ]
                },
                {
                    "slide_number": 6,
                    "layout_type": "conclusion_sources",
                    "title": "Xulosalar va Ilmiy Manbalar",
                    "key_takeaways": [
                        "Post-kvantga o'tish 2026-2030 yillar uchun strategik talabdir.",
                        "Gibrid xavfsizlik uzluksiz himoya beradi."
                    ],
                    "references": [
                        "1. NIST FIPS 203: Module-Lattice-Based KEM Standard (2024).",
                        "2. Shor, P. W. Quantum Factorization Algorithms (1997)."
                    ]
                }
            ]
        }

        builder = get_academic_pptx_builder()
        out_pptx = builder.create_presentation(sample_content, bp)

        assert os.path.exists(out_pptx), f"Generatsiya qilingan fayl mavjud emas: {out_pptx}"
        assert os.path.getsize(out_pptx) > 20000, f"Fayl hajmi juda kichik: {os.path.getsize(out_pptx)} bayt"

        prs = Presentation(out_pptx)
        assert len(prs.slides) == 6, f"Kutilgan 6 slayd, amalda: {len(prs.slides)}"

        # Verify Slide 1 has Title
        s1_texts = [shp.text_frame.text for shp in prs.slides[0].shapes if shp.has_text_frame]
        assert any("Kvant Hisoblashlari" in t for t in s1_texts)

        # Verify Slide 3 has 3 KPI Metric Cards
        s3_boxes = [shp for shp in prs.slides[2].shapes if shp.has_text_frame and len(shp.text_frame.text.strip()) > 0]
        assert len(s3_boxes) >= 3, f"3-slaydda metrika kartalari yetarli emas: {len(s3_boxes)}"

        # Verify Slide 4 has Process Steps (including nested group shapes)
        def _get_all_slide_text(slide):
            texts = []
            def _rec(shapes):
                for shp in shapes:
                    if shp.shape_type == 6:  # GroupShape
                        _rec(shp.shapes)
                    elif shp.has_text_frame and shp.text_frame.text.strip():
                        texts.append(shp.text_frame.text.strip())
            _rec(slide.shapes)
            return " ".join(texts)

        s4_all_text = _get_all_slide_text(prs.slides[3])
        assert "Audit" in s4_all_text, f"'Audit' bosqichi topilmadi: {s4_all_text}"
        assert ("Gibrid" in s4_all_text or "Ulanish" in s4_all_text), f"'Gibrid' bosqichi topilmadi: {s4_all_text}"
        assert ("To'liq" in s4_all_text or "O'tish" in s4_all_text), f"'To\\'liq O\\'tish' bosqichi topilmadi: {s4_all_text}"

        reporter.record(mod, "Haqiqiy shablon asosida to'liq 6-slaydli taqdimot yaratish", True, elapsed=time.time() - t0)
    except Exception as e:
        reporter.record(mod, "Haqiqiy shablon asosida to'liq 6-slaydli taqdimot yaratish", False, f"{e}\n{traceback.format_exc()}", elapsed=time.time() - t0)


def main():
    print("=" * 70)
    print(f"{Colors.BOLD}{Colors.HEADER}SLAYDHUB AI - AVTOMATLASHGAN TIZIM TEST RUNNERI{Colors.ENDC}")
    print("Har bir modul va funksiya avtomatik tekshiruvdan o'tkazilmoqda...")
    print("=" * 70)

    reporter = TestReporter()

    test_academic_catalog(reporter)
    test_academic_matcher(reporter)
    test_prompt_generator(reporter)
    test_json_parser_and_repair(reporter)
    test_topic_logic_validator(reporter)
    test_watermark_and_ad_purging(reporter)
    test_visual_archetype_renderers(reporter)
    test_spatial_geometry_and_bounds(reporter)
    test_slide_preview_engine(reporter)
    test_end_to_end_pipeline(reporter)

    success = reporter.print_summary()
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
