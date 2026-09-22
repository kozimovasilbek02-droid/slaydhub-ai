# -*- coding: utf-8 -*-
"""
Academic Topic Matcher & Blueprint Builder (Native PPTX Edition)
Analyzes academic & scientific defense topics (BMI, ВКР, Dissertatsiya, Ma'ruza)
and maps them directly to authentic, editable PPTX template files with:
1. Strict academic structure (Dolzarblik, Obyekt, Predmet, Maqsad/Vazifalar, Gipoteza, Yangilik, Aprobatsiya, Xulosalar).
2. "Bir slayd — bir fikr" cognitive load rule.
3. Action-oriented / Xulosaviy sarlavhalar.
4. Titul metama'lumotlari (Muallif, Tashkilot, Ilmiy rahbar, E-mail).
"""

import os
import re
import logging
from typing import List, Dict, Any, Optional

from core.academic_catalog import (
    get_academic_catalog,
    ACADEMIC_CATEGORIES,
    SLIDE_LAYOUT_TYPES,
    AcademicCatalog
)
from core.slide_archetypes import SLIDE_ARCHETYPES

logger = logging.getLogger("AcademicMatcher")

PRESENTATION_MODES = {
    "general_academic": {
        "name_uz": "Akademik Ma'ruza va O'quv Taqdimoti",
        "name_ru": "Академическая Лекция и Учебная Презентация",
        "name_en": "General Academic & Lecture Deck"
    },
    "scientific_defense_bmi": {
        "name_uz": "Ilmiy Himoya / BMI / ВКР / Dissertatsiya",
        "name_ru": "Научная Защита / ВКР / Диссертация",
        "name_en": "Scientific Thesis Defense (BSc/MSc/PhD)"
    },
    "business_project": {
        "name_uz": "Biznes Reja va Startap Loyihasi",
        "name_ru": "Бизнес-План и Стартап Проект",
        "name_en": "Business Plan & Startup Pitch"
    }
}


class AcademicMatcher:
    """
    Topic analysis and Academic Blueprint engine for Native PPTX templates.
    """
    def __init__(self, catalog: Optional[AcademicCatalog] = None):
        self.catalog = catalog or get_academic_catalog()

    def detect_category(self, topic: str) -> str:
        """Detects the primary academic category for a given topic."""
        topic_clean = topic.lower().strip()
        expanded_terms = self.catalog.expand_query_terms(topic)
        
        cat_scores: Dict[str, int] = {cat: 0 for cat in ACADEMIC_CATEGORIES.keys()}
        
        for term in expanded_terms:
            term_l = term.lower()
            for cat_key, cat_data in ACADEMIC_CATEGORIES.items():
                for kw in cat_data["keywords"]:
                    if kw in term_l or term_l in kw:
                        cat_scores[cat_key] += 3
        
        for cat_key, cat_data in ACADEMIC_CATEGORIES.items():
            for kw in cat_data["keywords"]:
                if re.search(r'\b' + re.escape(kw) + r'\b', topic_clean):
                    cat_scores[cat_key] += 5

        sorted_cats = sorted(cat_scores.items(), key=lambda x: x[1], reverse=True)
        if sorted_cats and sorted_cats[0][1] > 0:
            return sorted_cats[0][0]
        return "general_academic"

    def get_narrative_layout_sequence(self, slide_count: int, category: str, presentation_mode: str = "general_academic") -> List[str]:
        """Determines the optimal sequence of slide layouts based on presentation mode and topic category."""
        if presentation_mode == "scientific_defense_bmi":
            # Strict Thesis / Defense Structure
            if slide_count <= 6:
                return ["cover", "agenda", "theory_concept", "cards_grid", "comparison_vs", "conclusion_sources"][:slide_count]
            if slide_count == 8:
                return [
                    "cover",            # 1. Titul
                    "agenda",           # 2. Strukturasi
                    "theory_concept",   # 3. Dolzarblik & Muammo
                    "cards_grid",       # 4. Obyekt, Predmet, Maqsad va Vazifalar
                    "diagram_anatomy",  # 5. Nazariy Asoslar & Gipoteza
                    "timeline_steps",   # 6. Tadqiqot Metodologiyasi & Aprobatsiya
                    "metrics_stats",    # 7. Empirik Natijalar
                    "conclusion_sources"# 8. Xulosalar & Adabiyotlar
                ]
            if slide_count == 10:
                return [
                    "cover",            # 1. Titul
                    "agenda",           # 2. Ishning Tuzilishi
                    "theory_concept",   # 3. Mavzuning Dolzarbligi & Muammoning Qo'yilishi
                    "cards_grid",       # 4. Tadqiqot Obyekti, Predmeti, Maqsad va Vazifalari
                    "diagram_anatomy",  # 5. Ilmiy Gipoteza & Nazariy-Uslubiy Asoslar
                    "process_cycle",    # 6. Algoritmik / Siklik Tadqiqot Modeli
                    "timeline_steps",   # 7. Tadqiqot Bosqichlari va Aprobatsiya Jarayoni
                    "metrics_stats",    # 8. Eksperimental Natijalar va Empirik Tahlil
                    "comparison_vs",    # 9. Mavjud Yondashuvlar Bilan Qiyosiy Tahlil
                    "conclusion_sources"# 10. Asosiy Xulosalar, Amaliy Tavsiyalar va Adabiyotlar
                ]
            # 12-15 slides
            base_bmi = [
                "cover", "agenda", "theory_concept", "cards_grid", "pyramid_funnel",
                "diagram_anatomy", "process_cycle", "timeline_steps", "metrics_stats",
                "table_data", "comparison_vs", "conclusion_sources"
            ]
            if slide_count <= len(base_bmi):
                return base_bmi[:slide_count]
            diff = slide_count - len(base_bmi)
            extra_pool = ["keynote_quote", "cards_grid", "case_study"]
            return (base_bmi[:-1] + [extra_pool[i % len(extra_pool)] for i in range(diff)] + [base_bmi[-1]])[:slide_count]

        elif presentation_mode == "business_project":
            if slide_count <= 6:
                return ["cover", "agenda", "theory_concept", "process_cycle", "metrics_stats", "conclusion_sources"][:slide_count]
            if slide_count == 8:
                return [
                    "cover", "agenda", "theory_concept", "cards_grid",
                    "process_cycle", "timeline_steps", "metrics_stats", "conclusion_sources"
                ]
            if slide_count == 10:
                return [
                    "cover", "agenda", "theory_concept", "cards_grid",
                    "diagram_anatomy", "process_cycle", "timeline_steps",
                    "metrics_stats", "comparison_vs", "conclusion_sources"
                ]
            base_biz = [
                "cover", "agenda", "theory_concept", "cards_grid", "pyramid_funnel",
                "diagram_anatomy", "process_cycle", "timeline_steps", "metrics_stats",
                "comparison_vs", "keynote_quote", "conclusion_sources"
            ]
            if slide_count <= len(base_biz):
                return base_biz[:slide_count]
            diff = slide_count - len(base_biz)
            return (base_biz[:-1] + ["table_data" for _ in range(diff)] + [base_biz[-1]])[:slide_count]

        # General Academic Mode
        if slide_count <= 5:
            return ["cover", "agenda", "theory_concept", "comparison_vs", "conclusion_sources"][:slide_count]
        
        if slide_count == 8:
            return [
                "cover",
                "agenda",
                "theory_concept",
                "cards_grid",
                "diagram_anatomy" if category in ["exact_sciences", "medicine_health", "engineering_it"] else "process_cycle",
                "timeline_steps" if category in ["humanities_arts", "social_sciences_economy"] else "metrics_stats",
                "comparison_vs",
                "conclusion_sources"
            ]
            
        if slide_count == 10:
            return [
                "cover",
                "agenda",
                "theory_concept",
                "cards_grid",
                "diagram_anatomy" if category in ["exact_sciences", "medicine_health", "engineering_it"] else "pyramid_funnel",
                "process_cycle",
                "timeline_steps",
                "metrics_stats",
                "comparison_vs",
                "conclusion_sources"
            ]

        # 12-15 slides
        base = [
            "cover", "agenda", "theory_concept", "cards_grid", "pyramid_funnel",
            "diagram_anatomy", "process_cycle", "timeline_steps", "metrics_stats",
            "table_data", "comparison_vs", "conclusion_sources"
        ]
        if slide_count <= len(base):
            return base[:slide_count]
        
        diff = slide_count - len(base)
        cycle = ["keynote_quote", "cards_grid", "case_study"]
        extended = base[:-1] + [cycle[i % len(cycle)] for i in range(diff)] + [base[-1]]
        return extended[:slide_count]

    def build_blueprint(
        self,
        topic: str,
        slide_count: int = 10,
        language: str = "uz",
        selected_category: Optional[str] = None,
        presentation_mode: str = "general_academic",
        presenter_info: Optional[Dict[str, str]] = None
    ) -> Dict[str, Any]:
        """
        Builds a structured academic blueprint mapped to authentic PPTX templates using
        cross-deck visual archetype matching.
        """
        category = selected_category or self.detect_category(topic)
        cat_info = ACADEMIC_CATEGORIES.get(category, ACADEMIC_CATEGORIES["general_academic"])
        mode_info = PRESENTATION_MODES.get(presentation_mode, PRESENTATION_MODES["general_academic"])
        
        # Search matching authentic PPTX templates for overall theme
        matching_decks = self.catalog.search_templates(topic, category=category, limit=8)
        
        if not matching_decks:
            matching_decks = self.catalog.search_templates("", category=category, limit=8)
        if not matching_decks:
            matching_decks = self.catalog.templates[:8]

        primary_deck = matching_decks[0] if matching_decks else None
        pptx_path = primary_deck["pptx_path"] if primary_deck else ""
        cover_image = primary_deck.get("cover_image", "") if primary_deck else ""

        layout_sequence = self.get_narrative_layout_sequence(slide_count, category, presentation_mode=presentation_mode)
        slides_blueprint = []
        deck_slides = primary_deck.get("slides", []) if primary_deck else []

        for idx, layout_type in enumerate(layout_sequence, start=1):
            arch_meta = SLIDE_ARCHETYPES.get(layout_type, {})
            layout_meta = SLIDE_LAYOUT_TYPES.get(layout_type, SLIDE_LAYOUT_TYPES.get("theory_concept", {}))
            l_name = arch_meta.get(f"name_{language}") or layout_meta.get(f"name_{language}", layout_type)
            title_hint = self.generate_slide_title_hint(idx, layout_type, topic, language, presentation_mode=presentation_mode)

            # Match authentic slide possessing this visual archetype
            if idx == 1 and primary_deck:
                # Cover slide always comes from primary deck slide 0
                slide_ref = {
                    "pptx_path": primary_deck["pptx_path"],
                    "slide_index": 0,
                    "archetype": "cover",
                    "template_id": primary_deck["id"],
                    "template_title": primary_deck["title"],
                    "image_path": primary_deck.get("cover_image", "")
                }
            else:
                matches = self.catalog.search_slides_by_archetype(layout_type, category=category, topic=topic, limit=2)
                if matches:
                    slide_ref = matches[0]
                else:
                    # fallback to primary deck
                    img_p = ""
                    if idx - 1 < len(deck_slides):
                        img_p = deck_slides[idx - 1].get("image_path", "")
                    elif cover_image:
                        img_p = cover_image
                    slide_ref = {
                        "pptx_path": pptx_path,
                        "slide_index": min(idx - 1, max(0, len(deck_slides) - 1)),
                        "archetype": layout_type,
                        "template_id": primary_deck["id"] if primary_deck else "generic",
                        "template_title": primary_deck["title"] if primary_deck else "Academic Template",
                        "image_path": img_p
                    }

            slide_item = {
                "slide_number": idx,
                "layout_type": layout_type,
                "layout_name": l_name,
                "title_hint": title_hint,
                "template_id": slide_ref.get("template_id", "generic"),
                "template_title": slide_ref.get("template_title", "Academic Template"),
                "template_slide_number": slide_ref.get("slide_index", idx - 1) + 1,
                "image_path": slide_ref.get("image_path", ""),
                "pptx_path": slide_ref.get("pptx_path", pptx_path),
                "matched_slide_ref": slide_ref
            }
            slides_blueprint.append(slide_item)

        p_info = presenter_info or {
            "author_name": "Tadqiqotchi / Talaba",
            "organization": "Oliy Ta'lim Muassasasi / Kafedra",
            "supervisor": "Ilmiy rahbar: PhD, dotsent",
            "email": "tadqiqotchi@edu.uz",
            "event_name": "Ilmiy-amaliy anjuman / Himoya 2026"
        }

        blueprint = {
            "topic": topic,
            "category": category,
            "category_name": cat_info.get(f"name_{language}", cat_info["name_uz"]),
            "presentation_mode": presentation_mode,
            "presentation_mode_name": mode_info.get(f"name_{language}", mode_info["name_uz"]),
            "presenter_info": p_info,
            "slide_count": slide_count,
            "language": language,
            "engine_mode": "harmonized",
            "primary_deck_id": primary_deck["id"] if primary_deck else "",
            "primary_deck_title": primary_deck["title"] if primary_deck else "Academic Presentation",
            "primary_pptx_path": pptx_path,
            "matching_decks": matching_decks[:6],
            "slides": slides_blueprint
        }
        return blueprint

    def generate_slide_title_hint(
        self,
        slide_num: int,
        layout_type: str,
        topic: str,
        lang: str,
        presentation_mode: str = "general_academic"
    ) -> str:
        """Generates localized action-oriented title suggestions."""
        if presentation_mode == "scientific_defense_bmi":
            if lang == "uz":
                bmi_titles = {
                    1: f"{topic}",
                    2: "Tadqiqotning Tuzilishi va Asosiy Rejasi",
                    3: "Mavzuning Dolzarbligi va Muammoning Qo'yilishi",
                    4: "Tadqiqot Obyekti, Predmeti, Maqsad va Vazifalari",
                    5: "Ilmiy Gipoteza va Nazariy-Uslubiy Asoslar",
                    6: "Ilmiy Yangilik va Algoritmik Jarayon Modeli",
                    7: "Tadqiqot Bosqichlari va Aprobatsiya Jarayoni",
                    8: "Eksperimental Natijalar va Empirik Tahlil",
                    9: "Mavjud Yechimlar Bilan Qiyosiy Tahlil",
                    10: "Asosiy Ilmiy Xulosalar va Foydalanilgan Adabiyotlar"
                }
                return bmi_titles.get(slide_num, f"Slayd {slide_num}: {topic}")
            elif lang == "ru":
                bmi_titles = {
                    1: f"{topic}",
                    2: "Структура Исследования и План Работы",
                    3: "Актуальность Темы и Постановка Проблемы",
                    4: "Объект, Предмет, Цель и Задачи Исследования",
                    5: "Научная Гипотеза и Теоретико-Методологическая База",
                    6: "Научная Новизна и Модель Процесса",
                    7: "Этапы Исследования и Апробация Результатов",
                    8: "Экспериментальные Данные и Эмпирический Анализ",
                    9: "Сравнительный Анализ с Существующими Аналогами",
                    10: "Основные Научные Выводы и Список Литературы"
                }
                return bmi_titles.get(slide_num, f"Слайд {slide_num}: {topic}")
            else: # en
                bmi_titles = {
                    1: f"{topic}",
                    2: "Research Structure & Presentation Agenda",
                    3: "Relevance of the Topic & Problem Statement",
                    4: "Research Object, Subject, Goal & Objectives",
                    5: "Scientific Hypothesis & Theoretical Foundations",
                    6: "Scientific Novelty & Algorithmic Workflow Model",
                    7: "Research Methodology & Approbation Stages",
                    8: "Experimental Results & Empirical Evaluation",
                    9: "Comparative Analysis with Baseline Approaches",
                    10: "Key Academic Conclusions & References"
                }
                return bmi_titles.get(slide_num, f"Slide {slide_num}: {topic}")

        # General Academic
        if lang == "uz":
            titles = {
                "cover": f"{topic}",
                "agenda": "Mundarija va Asosiy Yo'nalishlar",
                "theory_concept": "Nazariy Asoslar va Fundamental Tushunchalar",
                "cards_grid": "Klassifikatsiya va Asosiy Ustunlar",
                "diagram_anatomy": "Tizimli Tuzilma va Komponentlar Tahlili",
                "comparison_vs": "Qiyosiy va Differensial Tahlil",
                "timeline_steps": "Rivojlanish Bosqichlari va Ketma-ketlik",
                "process_cycle": "Uzluksiz Sikl va Algoritmik Jarayon Modeli",
                "pyramid_funnel": "Ierarxik Tuzilma va Bosqichma-bosqich Ierarxiya",
                "metrics_stats": "Empirik Ko'rsatkichlar va Metrik Tahlil",
                "table_data": "Taqqoslama Parametrlar va Ko'rsatkichlar Jadvali",
                "keynote_quote": "Asosiy Ilmiy Iqtibos va Fundamental Tezis",
                "case_study": "Amaliy Tajriba va Keys Tadqiqoti",
                "qa_discussion": "Ilmiy Munozaralar va Dolzarb Savollar",
                "conclusion_sources": "Ilmiy Xulosa va Foydalanilgan Adabiyotlar"
            }
        elif lang == "ru":
            titles = {
                "cover": f"{topic}",
                "agenda": "Оглавление и Ключевые Темы",
                "theory_concept": "Теоретические Основы и Определения",
                "cards_grid": "Классификация и Основные Направления",
                "diagram_anatomy": "Системная Структура и Архитектура",
                "comparison_vs": "Сравнительный Анализ",
                "timeline_steps": "Этапы Развития и Последовательность",
                "process_cycle": "Круговой Цикл и Модель Процесса",
                "pyramid_funnel": "Иерархическая Структура и Пирамида Уровней",
                "metrics_stats": "Эмпирические Данные и Статистика",
                "table_data": "Таблица Параметров и Сравнительных Показателей",
                "keynote_quote": "Ключевая Цитата и Фундаментальный Тезис",
                "case_study": "Практический Кейс-Анализ",
                "qa_discussion": "Научные Дискуссии и Вопросы",
                "conclusion_sources": "Научные Выводы и Список Литературы"
            }
        else: # en
            titles = {
                "cover": f"{topic}",
                "agenda": "Table of Contents & Key Topics",
                "theory_concept": "Theoretical Foundations & Core Concepts",
                "cards_grid": "Classification & Key Dimensions",
                "diagram_anatomy": "System Architecture & Structural Analysis",
                "comparison_vs": "Comparative & Differential Analysis",
                "timeline_steps": "Evolutionary Stages & Chronology",
                "process_cycle": "Iterative Cycle & Process Workflow",
                "pyramid_funnel": "Hierarchical Architecture & Pyramid Levels",
                "metrics_stats": "Empirical Metrics & Statistical Analysis",
                "table_data": "Structured Parameters & Data Matrix",
                "keynote_quote": "Keynote Takeaway & Theoretical Quote",
                "case_study": "Practical Application & Case Study",
                "qa_discussion": "Scientific Inquiries & Q&A Discussion",
                "conclusion_sources": "Academic Conclusions & References"
            }
        return titles.get(layout_type, f"Slide {slide_num}: {topic}")


# Singleton instance
_matcher_instance: Optional[AcademicMatcher] = None

def get_academic_matcher() -> AcademicMatcher:
    global _matcher_instance
    if _matcher_instance is None:
        _matcher_instance = AcademicMatcher()
    return _matcher_instance
