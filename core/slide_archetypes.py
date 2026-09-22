# -*- coding: utf-8 -*-
"""
core/slide_archetypes.py
Defines the 12 Visual Slide Archetypes for SlaydHub AI.
Provides structural metadata, shape layout heuristics, detection rules,
and JSON schema requirements for each visual archetype.
"""
import re
from typing import Dict, Any, List, Optional
from pptx.enum.shapes import MSO_SHAPE_TYPE

SLIDE_ARCHETYPES = {
    "cover": {
        "id": "cover",
        "name_uz": "Titul Slaydi",
        "name_ru": "Титульный слайд",
        "name_en": "Cover / Title",
        "icon": "🎓",
        "description": "Taqdimot nomi, muallif, tashkilot va ilmiy rahbar",
        "min_shapes": 2,
        "max_shapes": 6
    },
    "agenda": {
        "id": "agenda",
        "name_uz": "Mantiqiy Reja",
        "name_ru": "План доклада",
        "name_en": "Agenda / Table of Contents",
        "icon": "📋",
        "description": "Tadqiqotning 3-5 ta asosiy bo'limlari ketma-ketligi",
        "min_shapes": 3,
        "max_shapes": 10
    },
    "theory_concept": {
        "id": "theory_concept",
        "name_uz": "Nazariy Asos & Gipoteza",
        "name_ru": "Теоретические основы",
        "name_en": "Theoretical Concept",
        "icon": "💡",
        "description": "Asosiy qonuniyat, ilmiy gipoteza, ta'rif va formulalar",
        "min_shapes": 2,
        "max_shapes": 6
    },
    "cards_grid": {
        "id": "cards_grid",
        "name_uz": "Kartalar To'plami",
        "name_ru": "Сетка карточек",
        "name_en": "Cards Grid",
        "icon": "🗂️",
        "description": "3-4 ta parallel ustunli bloklar va asosiy jihatlar",
        "min_shapes": 3,
        "max_shapes": 12
    },
    "comparison_vs": {
        "id": "comparison_vs",
        "name_uz": "Qiyosiy Tahlil (A vs B)",
        "name_ru": "Сравнительный анализ",
        "name_en": "Comparison / Versus",
        "icon": "⚖️",
        "description": "Ikki yondashuv yoki modelning tomonma-tomon taqqoslanishi",
        "min_shapes": 2,
        "max_shapes": 10
    },
    "timeline_steps": {
        "id": "timeline_steps",
        "name_uz": "Vaqt Chizig'i / Bosqichlar",
        "name_ru": "Временная шкала / Этапы",
        "name_en": "Timeline / Roadmap",
        "icon": "⏳",
        "description": "Xronologik rivojlanish, bosqichlar yoki yo'l xaritasi",
        "min_shapes": 3,
        "max_shapes": 12
    },
    "process_cycle": {
        "id": "process_cycle",
        "name_uz": "Doiraviy Sikl / Jarayon",
        "name_ru": "Круговой процесс",
        "name_en": "Process Cycle",
        "icon": "🔄",
        "description": "Uzluksiz aylanma jarayon, takroriy algoritm yoki sikl",
        "min_shapes": 3,
        "max_shapes": 12
    },
    "pyramid_funnel": {
        "id": "pyramid_funnel",
        "name_uz": "Ierarxiya & Piramida",
        "name_ru": "Иерархия и пирамида",
        "name_en": "Pyramid / Hierarchy",
        "icon": "🔺",
        "description": "Bosqichma-bosqich ierarxiya, voronka yoki darajalar",
        "min_shapes": 3,
        "max_shapes": 8
    },
    "metrics_stats": {
        "id": "metrics_stats",
        "name_uz": "KPI & Statistik Metrikalar",
        "name_ru": "Метрики и показатели",
        "name_en": "KPI & Statistics",
        "icon": "📊",
        "description": "Katta raqamlar, foizlar, o'sish ko'rsatkichlari",
        "min_shapes": 3,
        "max_shapes": 10
    },
    "diagram_anatomy": {
        "id": "diagram_anatomy",
        "name_uz": "Anatomik / Tizim Tarkibi",
        "name_ru": "Анатомия и архитектура",
        "name_en": "Anatomy / Breakdown",
        "icon": "🧩",
        "description": "Markaziy modelning ichki komponentlari va vazifalari",
        "min_shapes": 3,
        "max_shapes": 10
    },
    "keynote_quote": {
        "id": "keynote_quote",
        "name_uz": "Asosiy Iqtibos & Qonuniyat",
        "name_ru": "Ключевая цитата",
        "name_en": "Keynote / Quote",
        "icon": "💬",
        "description": "Katta ilmiy iqtibos, ta'kid yoki fundamental xulosa",
        "min_shapes": 1,
        "max_shapes": 4
    },
    "table_data": {
        "id": "table_data",
        "name_uz": "Strukturali Jadval",
        "name_ru": "Табличные данные",
        "name_en": "Data Table",
        "icon": "📅",
        "description": "Parametrlar, mezonlar va taqqoslash ko'rsatkichlari jadvali",
        "min_shapes": 1,
        "max_shapes": 6
    },
    "conclusion_sources": {
        "id": "conclusion_sources",
        "name_uz": "Xulosalar & Adabiyotlar",
        "name_ru": "Выводы и литература",
        "name_en": "Conclusions & References",
        "icon": "🎯",
        "description": "Asosiy ilmiy xulosalar, tavsiyalar va foydalanilgan adabiyotlar",
        "min_shapes": 2,
        "max_shapes": 8
    }
}


class SlideArchetypeDetector:
    """
    Analyzes shape geometry, count, and keywords on a slide to classify
    its visual archetype.
    """

    @classmethod
    def detect_archetype(cls, slide, slide_idx: int = 0, total_slides: int = 1) -> str:
        """
        Determines the most accurate visual archetype for a given slide.
        """
        if slide_idx == 0:
            return "cover"

        text_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        all_text = " ".join([s.text_frame.text.lower() for s in text_shapes])
        has_table = any(s.has_table for s in slide.shapes)

        if has_table or "table slide" in all_text:
            return "table_data"

        if any(k in all_text for k in ["table of contents", "agenda", "contents", "index", "reja"]):
            return "agenda"

        if any(k in all_text for k in ["credits", "sources", "references", "thank you", "thanks", "adabiyotlar", "xulosa"]):
            return "conclusion_sources"

        if any(k in all_text for k in ["vs", "versus", "comparison", "compare", "taqqoslash"]):
            return "comparison_vs"

        if any(k in all_text for k in ["pyramid", "hierarchy", "funnel", "ierarxiya"]):
            return "pyramid_funnel"

        if any(k in all_text for k in ["cycle", "circular", "loop", "sikl", "jarayon"]):
            return "process_cycle"

        if any(k in all_text for k in ["timeline", "roadmap", "milestone", "2026", "2027", "2028"]):
            return "timeline_steps"

        if any(k in all_text for k in ["%", "$", "sales", "growth", "kpi", "metrika", "stats"]):
            return "metrics_stats"

        if any(k in all_text for k in ["quote", "iqtibos", "saying", "keynote"]) or (len(text_shapes) == 2 and any(len(s.text_frame.text.split()) > 15 for s in text_shapes)):
            return "keynote_quote"

        # Check geometry of non-text vector illustrations
        group_shapes = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.GROUP]
        if group_shapes and len(text_shapes) in [3, 4, 5, 6]:
            return "diagram_anatomy"

        # Step badges presence check (e.g. 01, 02, 03)
        step_badges = [s for s in text_shapes if re.match(r'^(?:step\s*)?0?[1-9][\.\:\)]?$', s.text_frame.text.strip(), re.I)]
        if len(step_badges) >= 3:
            # Check horizontal distribution
            xs = [s.left for s in step_badges]
            if max(xs) - min(xs) > 3000000: # wide distribution
                return "timeline_steps"

        # Column clustering check
        if len(text_shapes) in [6, 7, 8, 9, 10, 11, 12]:
            return "cards_grid"

        if len(text_shapes) in [2, 3, 4]:
            return "theory_concept"

        return "cards_grid"
