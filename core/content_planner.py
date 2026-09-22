# -*- coding: utf-8 -*-
"""
core/content_planner.py
AppFlowy-Compatible Presentation Structure & Content Planner for SlaydHub AI.

Organizes topics into balanced, academic & professional slide blueprints
compatible with AppFlowy Kanban boards and markdown document exports.
"""

import json
from pathlib import Path
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, asdict

from core.config import config


@dataclass
class PlannedSlide:
    slide_number: int
    archetype: str        # 'title' | 'agenda' | 'concept' | 'two_column' | 'stats' | 'timeline' | 'qa'
    title: str
    subtitle: str
    bullets: List[str]
    visual_cue: str       # Fooocus / Penpot prompt hint
    speaker_notes: str


class AppFlowyContentPlanner:
    """
    Builds and exports structured slide outlines matching AppFlowy project boards.
    """

    @classmethod
    def generate_plan(cls, topic: str, slide_count: int = 8) -> List[PlannedSlide]:
        """
        Generates a balanced academic / business slide deck structure.
        """
        slides: List[PlannedSlide] = []

        # 1. Title Slide
        slides.append(PlannedSlide(
            slide_number=1,
            archetype="title",
            title=topic,
            subtitle="Zamonaviy ilmiy va amaliy tahlil",
            bullets=[],
            visual_cue="Minimalist abstract dark blue gradient with subtle geometric shapes",
            speaker_notes="Taqdimotga kirish, mavzuning dolzarbligi va maqsadlari."
        ))

        # 2. Agenda / Reja
        slides.append(PlannedSlide(
            slide_number=2,
            archetype="agenda",
            title="Taqdimot Rejasi",
            subtitle="Asosiy ko'rib chiqiladigan masalalar",
            bullets=[
                "1. Mavzuning mohiyati va nazariy asoslari",
                "2. Asosiy tamoyillar va arxitektura",
                "3. Amaliy tahlil va ko'rsatkichlar",
                "4. Muammolar, yechimlar va xulosalar"
            ],
            visual_cue="Numbered list cards with clean vertical step indicators",
            speaker_notes="Taqdimot davomida qamrab olinadigan asosiy bosqichlar."
        ))

        # 3. Asosiy tushunchalar (Concept)
        slides.append(PlannedSlide(
            slide_number=3,
            archetype="concept",
            title=f"{topic}: Nazariy Asoslar",
            subtitle="Tushunchalar va ilmiy tamoyillar",
            bullets=[
                "Tizimning asosiy tarkibiy elementlari va ularning o'zaro bog'liqligi.",
                "Zamonaviy texnologik va ilmiy yondashuvlarning o'rni.",
                "Nazariy modellarning amaliyotga tatbiqi."
            ],
            visual_cue="High-contrast infographic diagram illustrating core concept",
            speaker_notes="Asosiy qonuniyatlar va nazariy asoslarga to'xtalamiz."
        ))

        # 4. Solishtirma / Arxitektura (Two column)
        slides.append(PlannedSlide(
            slide_number=4,
            archetype="two_column",
            title="Tizim Tuzilishi va Arxitektura",
            subtitle="Funksional bloklar va imkoniyatlar",
            bullets=[
                "Chap ustun: Asosiy komponentlar va ularning vazifalari.",
                "O'ng ustun: Integratsiya jarayonlari va xavfsizlik mexanizmlari.",
                "Natija: Tizimli samaradorlik va barqarorlik."
            ],
            visual_cue="Two-column balanced card layout with accent borders",
            speaker_notes="Tizimning ichki tuzilishi va bloklarining ishlash mexanizmi."
        ))

        # 5. Natijalar va Metrikalar (Stats)
        slides.append(PlannedSlide(
            slide_number=5,
            archetype="stats",
            title="Amaliy Natijalar va Metrikalar",
            subtitle="Samaradorlik va asosiy ko'rsatkichlar",
            bullets=[
                "85%+ Samaradorlik o'sishi va optimallashuv.",
                "2.5x Tezkorlik va resurslarni tejash ko'rsatkichi.",
                "100% Jarayonlar avtomatizatsiyasi va ishonchlilik."
            ],
            visual_cue="Large prominent metric numbers with supporting micro-cards",
            speaker_notes="Olingan amaliy natijalar va statistik ko'rsatkichlar tahlili."
        ))

        # 6. Xulosa va Q&A
        slides.append(PlannedSlide(
            slide_number=6,
            archetype="qa",
            title="Xulosa va Savol-Javoblar",
            subtitle="E'tiboringiz uchun rahmat!",
            bullets=[
                "Asosiy xulosalar umumlashtirildi.",
                "Kelajakdagi rivojlanish istiqbollari belgilandi.",
                "Savollaringizni berishingiz mumkin."
            ],
            visual_cue="Centered elegant conclusion card with contact information",
            speaker_notes="Taqdimotni yakunlash va auditoriya bilan muloqot."
        ))

        return slides

    @classmethod
    def export_to_appflowy_markdown(cls, topic: str, slides: List[PlannedSlide]) -> Path:
        """
        Exports the presentation plan as an AppFlowy-compatible Markdown board file.
        """
        export_dir = config.OUTPUT_ROOT / "appflowy_boards"
        export_dir.mkdir(parents=True, exist_ok=True)
        filename = export_dir / f"plan_{topic.replace(' ', '_')[:30]}.md"

        content = [
            f"# 📋 AppFlowy Board: {topic}",
            f"**Turi:** Taqdimot Loyihasi | **Slaydlar:** {len(slides)} ta | **Holat:** Tayyor\n",
            "## 📌 Slaydlar Ketma-ketligi\n"
        ]

        for s in slides:
            content.append(f"### Slayd {s.slide_number}: {s.title} ({s.archetype.upper()})")
            content.append(f"- **Subtitle:** {s.subtitle}")
            content.append(f"- **Dizayn/Prompt (Fooocus & Penpot):** `{s.visual_cue}`")
            content.append("- **Punktlar:**")
            for b in s.bullets:
                content.append(f"  - {b}")
            content.append(f"- **Ma'ruzachi Eslatmasi:** *{s.speaker_notes}*\n")

        with open(filename, "w", encoding="utf-8") as f:
            f.write("\n".join(content))

        return filename
