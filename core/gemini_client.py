# -*- coding: utf-8 -*-
import os
import re
import json
import time
import random
import asyncio
from typing import List, Dict, Any, Tuple, Optional
import aiohttp

# ── 1. MODELLAR KASKADI VA CIRCUIT BREAKER ────────────────────────────────────
GEMINI_MODELS_CASCADE = [
    "gemini-3.1-flash-lite",
    "gemini-flash-lite-latest",
    "gemini-3.8-flash",
    "gemini-flash-latest",
    "gemini-4-argon",
]

# 404 olgan eskirgan modellarni 1 soat davomida so'rov yubormaslik uchun kesh
DEAD_MODELS_CACHE: Dict[str, float] = {}
DEAD_MODEL_TTL_SECONDS = 3600

# ── 2. JSON REPAIR FUNKSIYASI ─────────────────────────────────────────────────
def repair_truncated_json(broken_json: str) -> str:
    """Uzilib qolgan yoki noto'g'ri yopilgan JSON javoblarni tiklaydi."""
    s = broken_json.strip()
    if not s:
        return "{}"
    if s.endswith(","):
        s = s[:-1].strip()

    # Ochiq qolgan qavslar hisobi
    open_curly = s.count("{") - s.count("}")
    open_square = s.count("[") - s.count("]")

    if s.count('"') % 2 != 0:
        s += '"'

    s += "]" * max(0, open_square)
    s += "}" * max(0, open_curly)
    return s


# ── 3. STRUCTURED OUTPUT SCHEMAS ──────────────────────────────────────────────
STAGE1_OUTLINE_SCHEMA = {
    "type": "OBJECT",
    "required": ["title", "subtitle", "category_id", "recommended_theme", "agenda"],
    "properties": {
        "title": {"type": "STRING"},
        "subtitle": {"type": "STRING"},
        "category_id": {"type": "INTEGER"},
        "recommended_theme": {
            "type": "STRING",
            "enum": ["academic_navy", "emerald_science", "tech_minimal", "warm_philosophy", "cyber_dark"],
        },
        "agenda": {
            "type": "ARRAY",
            "items": {"type": "STRING"},
        },
    },
}

STAGE2_SLIDES_SCHEMA = {
    "type": "OBJECT",
    "required": ["slides", "tags", "description"],
    "properties": {
        "slides": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "required": ["heading", "archetype", "points", "speaker_note"],
                "properties": {
                    "heading": {"type": "STRING"},
                    "archetype": {
                        "type": "STRING",
                        "enum": ["cards_grid", "comparison_vs", "timeline_steps", "theory_concept"],
                    },
                    "points": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "speaker_note": {"type": "STRING"},
                },
            },
        },
        "tags": {"type": "STRING"},
        "description": {"type": "STRING"},
    },
}

DOCX_SCHEMA = {
    "type": "OBJECT",
    "required": ["title", "meta", "introduction", "chapters", "conclusion", "references"],
    "properties": {
        "title": {"type": "STRING"},
        "meta": {
            "type": "OBJECT",
            "properties": {
                "doc_type": {"type": "STRING"},
                "ministry": {"type": "STRING"},
                "university": {"type": "STRING"},
                "department": {"type": "STRING"},
                "student": {"type": "STRING"},
                "group": {"type": "STRING"},
                "advisor": {"type": "STRING"},
                "city_year": {"type": "STRING"},
            },
        },
        "introduction": {"type": "ARRAY", "items": {"type": "STRING"}},
        "chapters": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "required": ["heading", "paragraphs"],
                "properties": {
                    "heading": {"type": "STRING"},
                    "paragraphs": {"type": "ARRAY", "items": {"type": "STRING"}},
                },
            },
        },
        "conclusion": {"type": "ARRAY", "items": {"type": "STRING"}},
        "references": {"type": "ARRAY", "items": {"type": "STRING"}},
    },
}

DTM_QUIZ_SCHEMA = {
    "type": "OBJECT",
    "required": ["title", "questions"],
    "properties": {
        "title": {"type": "STRING"},
        "questions": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "required": ["question", "options", "correct_index", "explanation"],
                "properties": {
                    "question": {"type": "STRING"},
                    "options": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "correct_index": {"type": "INTEGER"},
                    "explanation": {"type": "STRING"},
                },
            },
        },
    },
}


# ── 4. ASYNCHRONOUS GEMINI CASCADE CLIENT ─────────────────────────────────────
async def call_gemini_cascade(
    prompt: str,
    api_key: Optional[str] = None,
    schema: Optional[Dict[str, Any]] = None,
    max_output_tokens: int = 4096,
    temperature: float = 0.3,
    system_instruction: Optional[str] = None,
) -> Tuple[Optional[Dict[str, Any]], str]:
    """
    Gemini API ga asinxron so'rov yuboruvchi kaskadli funksiya.
    404 da modelni zaxiraga oladi, 429/503 da exponential backoff bilan kutadi.
    """
    if not api_key:
        api_key = os.getenv("GEMINI_API_KEY", "")
    if not api_key:
        return None, "GEMINI_API_KEY o'rnatilmagan"

    now = time.time()
    available_models = [m for m in GEMINI_MODELS_CASCADE if now - DEAD_MODELS_CACHE.get(m, 0) > DEAD_MODEL_TTL_SECONDS]
    if not available_models:
        available_models = ["gemini-3.1-flash-lite", "gemini-flash-latest"]

    gen_config: Dict[str, Any] = {
        "temperature": temperature,
        "maxOutputTokens": max_output_tokens,
        "responseMimeType": "application/json",
    }
    if schema:
        gen_config["responseSchema"] = schema

    payload: Dict[str, Any] = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": gen_config,
    }
    if system_instruction:
        payload["systemInstruction"] = {"parts": [{"text": system_instruction}]}

    last_error = ""

    async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=75)) as session:
        for model in available_models:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"

            for attempt in range(1, 3):
                try:
                    async with session.post(url, json=payload) as resp:
                        status = resp.status

                        if status == 200:
                            data = await resp.json()
                            candidates = data.get("candidates", [])
                            if not candidates:
                                last_error = f"{model}: bo'sh nomzod (xavfsizlik filtri)"
                                break

                            first_cand = candidates[0]
                            content = first_cand.get("content", {})
                            parts = content.get("parts", [])
                            if not parts:
                                last_error = f"{model}: bo'sh part"
                                break

                            raw_text = parts[0].get("text", "").strip()
                            raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.MULTILINE)
                            raw_text = re.sub(r"\s*```$", "", raw_text, flags=re.MULTILINE).strip()

                            try:
                                return json.loads(raw_text), ""
                            except json.JSONDecodeError:
                                repaired = repair_truncated_json(raw_text)
                                return json.loads(repaired), ""

                        elif status == 404:
                            DEAD_MODELS_CACHE[model] = time.time()
                            last_error = f"Model {model} topilmadi (404), zaxiraga olindi"
                            break  # Keyingi modelga o'tish

                        elif status in (429, 503):
                            delay = (2 ** attempt) + random.uniform(0.5, 1.5)
                            last_error = f"{model}: HTTP {status} (Limit/Server band), {delay:.1f}s kutilmoqda"
                            await asyncio.sleep(delay)
                            continue

                        else:
                            txt = await resp.text()
                            last_error = f"{model}: HTTP {status} — {txt[:200]}"
                            break

                except asyncio.TimeoutError:
                    last_error = f"{model}: 75s vaqt tugadi (Timeout)"
                    break
                except Exception as ex:
                    last_error = f"{model}: Ulanish xatosi: {str(ex)}"
                    break

    return None, last_error


# ── 5. PROMPT BUILDERS ────────────────────────────────────────────────────────
def build_stage1_outline_prompt(topic_or_excerpt: str, slide_count: int = 10, doc_mode: str = "pptx") -> str:
    """Bosqich 1: Faqat sarlavha va reja bandlarini generatsiya qilish (~200 token)."""
    return (
        f"Vazifa: Quyidagi mavzu/matn asosida {doc_mode.upper()} uchun aniq {slide_count} ta "
        f"mantiqiy reja bandini (agenda) o'zbek adabiy tilida (lotin yozuvida) tuzib ber.\n"
        f"DIQQAT: Slayd matnlarini yozma, faqat qisqa va aniq reja sarlavhalarini qaytar!\n\n"
        f"Manba: {topic_or_excerpt[:2500]}"
    )


def build_stage2_locked_expansion_prompt(
    title: str,
    locked_agenda: List[str],
    context_excerpt: str = "",
) -> str:
    """
    Bosqich 2: Foydalanuvchi tahrirlab tasdiqlagan LOCKED_AGENDA ro'yxatiga 100% bo'ysunib,
    har bir band uchun aniq bittadan slayd generatsiya qilish.
    """
    numbered_agenda = "\n".join(f"{i}. {heading}" for i, heading in enumerate(locked_agenda, 1))
    return (
        f"Taqdimot sarlavhasi: {title}\n"
        f"Foydalanuvchi tomonidan tasdiqlangan QAT'IY REJA ({len(locked_agenda)} ta band):\n"
        f"{numbered_agenda}\n\n"
        f"Qo'shimcha kontekst:\n{context_excerpt[:12000]}\n\n"
        f"QAT'IY QOIDALAR:\n"
        f"1. 'slides' massivida aniq {len(locked_agenda)} ta slayd bo'lishi shart — har bir reja bandiga 1 tadan!\n"
        f"2. Reja sarlavhalarini ('heading') ma'nosini o'zgartirma.\n"
        f"3. Har bir slaydda 3-4 tadan ilmiy asoslangan punkt ('points'), mos 'archetype' va qisqa 'speaker_note' yoz.\n"
        f"4. Har bir slaydda Penpot va Knuth-Plass qoidalariga rioya qiling."
    )


def build_academic_docx_prompt(topic: str, doc_type: str = "REFERAT", pages: int = 12) -> str:
    """OTM standartidagi Word akademik hujjati uchun chuqur matn generatsiyasi prompti."""
    return (
        f"Sen O'zbekiston Oliy ta'limi bo'yicha ilmiy ekspert va dotsentsan.\n"
        f"Quyidagi mavzuda {doc_type.upper()} ({pages} sahifaga mo'ljallangan to'liq ilmiy ish) yozib ber:\n"
        f"Mavzu: «{topic}»\n\n"
        f"QAT'IY TALABLAR:\n"
        f"1. 'meta' maydoniga rasmiy OTM rekvizitlarini kiriting (Vazirlik, Universitet, Kafedra, Talaba, Rahbar).\n"
        f"2. 'introduction' (Kirish): Mavzuning dolzarbligi, maqsadi, vazifalari va ob'ekti (kamida 3-4 ta to'liq abzas).\n"
        f"3. 'chapters': Mavzu bo'yicha kamida 3 ta mustaqil bob. Har bir bobda kamida 4-6 tadan batafsil tahliliy abzas bo'lsin.\n"
        f"4. 'conclusion': Ilmiy xulosa va amaliy tavsiyalar (kamida 3 ta abzas).\n"
        f"5. 'references': Kamida 8-12 ta rasmiy ilmiy adabiyot, monografiya, darslik va O'zbekiston qonunchilik hujjatlari (GOST 7.1-2003 formati)."
    )


def build_dtm_quiz_prompt(topic_or_text: str, question_count: int = 10) -> str:
    """DTM standarti bo'yicha 4 variantli testlar tuzish prompti."""
    return (
        f"Vazifa: Quyidagi mavzu yoki matn asosida O'zbekiston DTM (Bilimni baholash agentligi) "
        f"standartiga mos aniq {question_count} ta 4 variantli (A, B, C, D) test savoli tuzib ber.\n"
        f"Har bir savol uchun:\n"
        f"- 'question': Aniq savol matni\n"
        f"- 'options': 4 ta variant (aniq 4 ta element)\n"
        f"- 'correct_index': To'g'ri javob indeksi (0, 1, 2 yoki 3)\n"
        f"- 'explanation': Nega shu javob to'g'riligi haqida qisqa ilmiy tushuntirish\n\n"
        f"Manba:\n{topic_or_text[:10000]}"
    )
