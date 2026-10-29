# -*- coding: utf-8 -*-
import re
import fitz
from typing import List, Dict, Any, Tuple
from core.gemini_client import call_gemini_cascade, SOFF_CATEGORIES_PROMPT_NOTE

CHAPTER_HEADER_RE = re.compile(
    r"^\s*(?:BOB|FASL|GLAVA|CHAPTER|PARAGRAF|\u00A7|\d+[\.\)]|\bI{1,4}\b|\bV\b|\bVI{0,3}\b|\bX\b)\s*[\.\:\-]?\s*(.+)$",
    re.IGNORECASE | re.MULTILINE,
)

TOC_MARKERS = [
    "mundarija",
    "o'qlavleniye",
    "oglavleniye",
    "soderjaniye",
    "table of contents",
    "contents",
    "fihrist",
]


def extract_smart_pdf_text(pdf_path: str, max_chars: int = 27500) -> str:
    """
    Katta PDF darsliklardan (8-50 MB) aqlli TOC-yo'naltirilgan matn ajratib oluvchi funksiya.
    Mundarija (TOC) sahifalarini aniqlab, boblar boshlanishini va Xulosa qismini proporsional oladi.
    Har doim 'with fitz.open(...) as doc:' kontekst menejeri orqali C-xotirani yopadi.
    """
    with fitz.open(pdf_path) as doc:
        total_pages = len(doc)
        if total_pages == 0:
            return ""

        if total_pages <= 15:
            text_parts = [page.get_text("text") for page in doc]
            return "\n".join(text_parts)[:max_chars]

        # 1. Ichki PDF TOC mavjudligini tekshirish
        toc = doc.get_toc()
        target_pages = set()

        if toc:
            for item in toc:
                if len(item) >= 3 and isinstance(item[2], int):
                    p_idx = item[2] - 1
                    if 0 <= p_idx < total_pages:
                        target_pages.add(p_idx)
                        if p_idx + 1 < total_pages:
                            target_pages.add(p_idx + 1)

        # 2. Matnli mundarija sahifalarini qidirish (boshlang'ich 25 va oxirgi 15 sahifa)
        scan_indices = sorted(
            set(
                list(range(min(25, total_pages)))
                + list(range(max(0, total_pages - 15), total_pages))
            )
        )

        for p_idx in scan_indices:
            page_txt = doc[p_idx].get_text("text")
            first_lines = page_txt.lower()[:400]
            if any(marker in first_lines for marker in TOC_MARKERS):
                target_pages.add(p_idx)
                if p_idx + 1 < total_pages:
                    target_pages.add(p_idx + 1)
                if p_idx + 2 < total_pages:
                    target_pages.add(p_idx + 2)

        # 3. Agar TOC topilmasa, Regex orqali bob sarlavhalarini qidirish
        if len(target_pages) < 3:
            for p_idx in range(0, total_pages, max(1, total_pages // 30)):
                page_txt = doc[p_idx].get_text("text")
                if CHAPTER_HEADER_RE.search(page_txt):
                    target_pages.add(p_idx)
                    if p_idx + 1 < total_pages:
                        target_pages.add(p_idx + 1)

        # 4. Majburiy: Kirish (0..4 sahifalar) va Xulosa (oxirgi 4 sahifa)
        for p in range(min(4, total_pages)):
            target_pages.add(p)
        for p in range(max(0, total_pages - 4), total_pages):
            target_pages.add(p)

        sorted_pages = sorted(target_pages)
        budget_per_page = max(350, max_chars // max(1, len(sorted_pages)))

        collected_chunks = []
        total_len = 0

        for p_idx in sorted_pages:
            page_txt = doc[p_idx].get_text("text").strip()
            if not page_txt:
                continue
            trimmed = page_txt[:budget_per_page]
            collected_chunks.append(f"--- [SAHIFA {p_idx+1}] ---\n{trimmed}")
            total_len += len(trimmed)
            if total_len >= max_chars:
                break

        return "\n\n".join(collected_chunks)[:max_chars]


BOOK_DECOMPOSITION_SCHEMA = {
    "type": "ARRAY",
    "items": {
        "type": "OBJECT",
        "required": [
            "title",
            "subtitle",
            "category_id",
            "category_name",
            "price_uzs",
            "description",
            "tags",
            "slides",
        ],
        "properties": {
            "title": {"type": "STRING"},
            "subtitle": {"type": "STRING"},
            "category_id": {"type": "INTEGER"},
            "category_name": {"type": "STRING"},
            "price_uzs": {"type": "INTEGER"},
            "description": {"type": "STRING"},
            "tags": {"type": "STRING"},
            "slides": {
                "type": "ARRAY",
                "items": {
                    "type": "OBJECT",
                    "required": ["heading", "points", "archetype", "speaker_note"],
                    "properties": {
                        "heading": {"type": "STRING"},
                        "points": {"type": "ARRAY", "items": {"type": "STRING"}},
                        "archetype": {
                            "type": "STRING",
                            "enum": [
                                "cards_grid",
                                "comparison_vs",
                                "timeline_steps",
                                "theory_concept",
                            ],
                        },
                        "speaker_note": {"type": "STRING"},
                    },
                },
            },
        },
    },
}


async def decompose_book_into_presentations(
    sample_text: str, filename: str
) -> Tuple[List[Dict[str, Any]], str]:
    """Katta kitob matnini mustaqil taqdimot spetsifikatsiyalariga ajratadi."""
    if len(sample_text.strip()) < 80:
        return [], "Kitobdan matn ajratib bo'lmadi (ehtimol skanerlangan rasm yoki himoyalangan PDF)"

    prompt = f"""Sen professional taqdimotlar arxitektori va ilmiy muharrirsan.
Quyida «{filename}» nomli darslik/qo'llanmadan olingan tizimli mundarija va asosiy boblar matni keltirilgan.

VAZIFA:
Ushbu kitobni talabalar, o'qituvchilar va mutaxassislar uchun mustaqil taqdimotlarga (.pptx) ajrat:
1. Kitobning asosiy bo'limlari yoki yirik boblari bo'yicha 2 tadan 4 tagacha to'liq, mustaqil taqdimot tuz.
2. Har bir taqdimot ichida 6 tadan 8 tagacha to'liq slayd ('slides') bo'lsin (har birida 'archetype', 3-4 ta aniq faktik nuqta va 'speaker_note').
3. Har bir taqdimot uchun Soff.uz talablariga mos:
   - 'price_uzs': Slaydlar soni <= 6 bo'lsa 3000, 7-12 bo'lsa 7000, >= 13 bo'lsa 10000 UZS.
   - {SOFF_CATEGORIES_PROMPT_NOTE}
   - 'description': Kamida 350 belgidan iborat boy ilmiy annotatsiya va slaydlar ro'yxati.
   - 'tags': Kamida 8 ta vergul bilan ajratilgan o'zbekcha qidiruv kalit so'zlari.

Kitobdan olingan matn:
{sample_text[:27500]}
"""
    result, err = await call_gemini_cascade(
        prompt, schema=BOOK_DECOMPOSITION_SCHEMA, max_output_tokens=8192
    )
    if result and isinstance(result, list):
        return result, ""
    return [], err or "Kitobni taqdimotlarga ajratishda xatolik yuz berdi"
