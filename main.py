# -*- coding: utf-8 -*-
"""
SlaydHub AI & Soff.uz 24/7 Autonomous Telegram Bot
Production-ready for Render.com (512MB RAM & High-Resilience Architecture)
Refactored and Audited by Gemini 4 Argon
"""

import os
import re
import sys
import gc
import json
import time
import base64
import random
import asyncio
import tempfile
import aiohttp
import fitz  # PyMuPDF

# Ensure UTF-8 output on all operating systems
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from aiohttp import web
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

# ── 1. CONFIGURATION & SECRETS ───────────────────────────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SELLER_TOKEN = os.getenv("SELLER_TOKEN")
SOFF_REFRESH_TOKEN = os.getenv("SOFF_REFRESH_TOKEN")
SOFF_PHONE = os.getenv("SOFF_PHONE")
SOFF_PASSWORD = os.getenv("SOFF_PASSWORD")
PORT = int(os.getenv("PORT", 10000))

# Local fallbacks for development/testing
if not SELLER_TOKEN:
    local_token_file = r"C:\Users\user\Desktop\Ish\Soff.uz\scripts\token.txt"
    if os.path.exists(local_token_file):
        try:
            with open(local_token_file, "r", encoding="utf-8") as f:
                SELLER_TOKEN = f.read().strip()
        except Exception:
            pass

if not GEMINI_API_KEY or not BOT_TOKEN:
    local_env_file = r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot\.env"
    if os.path.exists(local_env_file):
        try:
            with open(local_env_file, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("GEMINI_API_KEY=") and not GEMINI_API_KEY:
                        GEMINI_API_KEY = line.split("=", 1)[1].strip()
                    elif line.startswith("BOT_TOKEN=") and not BOT_TOKEN:
                        BOT_TOKEN = line.split("=", 1)[1].strip()
        except Exception:
            pass

if not BOT_TOKEN:
    print("[WARNING] BOT_TOKEN environment variable is not set!")

bot = Bot(token=BOT_TOKEN or "DUMMY_TOKEN")
dp = Dispatcher()

# ── 2. MEMORY AND CONCURRENCY CONTROLS (RENDER 512MB RAM PROTECTION) ─────────
# Limit to 1 heavy PDF extraction/generation at a time to prevent OOM
HEAVY_TASK_SEMAPHORE = asyncio.Semaphore(1)

# Cache for generated presentations waiting for upload
UPLOAD_CACHE_TTL_SECONDS = 1800  # 30 minutes
PENDING_UPLOADS: dict[str, dict] = {}


def release_memory() -> None:
    """PyMuPDF internal C-store shrink and Python Garbage Collector."""
    try:
        fitz.TOOLS.store_shrink(100)
    except Exception:
        pass
    gc.collect()


async def cleanup_expired_uploads_loop() -> None:
    """Periodically cleans up old PPTX files and expired PENDING_UPLOADS items."""
    while True:
        await asyncio.sleep(300)
        now = time.time()
        expired_keys = [
            uid for uid, meta in list(PENDING_UPLOADS.items())
            if now - meta.get("created_at", now) > UPLOAD_CACHE_TTL_SECONDS
        ]
        for uid in expired_keys:
            meta = PENDING_UPLOADS.pop(uid, None)
            if meta:
                fp = meta.get("file_path")
                if fp and os.path.exists(fp):
                    try:
                        os.remove(fp)
                    except OSError:
                        pass
        if expired_keys:
            release_memory()


# ── 3. SMART PDF CHUNKING V2 (TOC-GUIDED & BALANCED BUDGET) ───────────────────
CHAPTER_HEADER_RE = re.compile(
    r"(?:^\s*(?:I{1,3}|IV|V|VI{0,3}|IX|X|\d+)\s*[-\.]?\s*(?:BOB|BO['’]LIM|MAVZU|QISM|GLAVA|ГЛАВА|РАЗДЕЛ|CHAPTER|PART)\b)"
    r"|(?:^\s*§\s*\d+)"
    r"|(?:^\s*MUNDARIJA|ОГЛАВЛЕНИЕ|CONTENTS\b)",
    re.IGNORECASE | re.MULTILINE
)


def extract_smart_pdf_text_v2(pdf_path: str, max_total_chars: int = 27500) -> tuple[str, int]:
    """
    Extracts structured representation of large (8-50 MB) PDFs with guaranteed
    budget allocation so that Introduction, Chapters, and Conclusion are never lost.
    """
    BUDGET_TOC = 3500
    BUDGET_INTRO = 5500
    BUDGET_CHAPTERS = 14500
    BUDGET_CONCLUSION = 4000

    sections: list[str] = []

    with fitz.open(pdf_path) as doc:
        total_pages = len(doc)
        if total_pages == 0:
            return "", 0

        # 1. Table of Contents
        toc = doc.get_toc()
        toc_chapter_pages: list[int] = []

        if toc:
            toc_lines = []
            for lvl, title, p_num in toc[:60]:
                indent = "  " * max(0, lvl - 1)
                toc_lines.append(f"{indent}- {title.strip()} (bet {p_num})")
                if lvl in (1, 2) and 1 <= p_num <= total_pages:
                    toc_chapter_pages.append(p_num - 1)
            sections.append("[MUNDARIJA VA TUZILMA]:\n" + "\n".join(toc_lines)[:BUDGET_TOC])
        else:
            # Check beginning and end of book for text TOC
            toc_candidates = list(range(min(15, total_pages))) + list(range(max(0, total_pages - 8), total_pages))
            found_toc_texts = []
            for p_no in sorted(set(toc_candidates)):
                txt = doc[p_no].get_text("text").strip()
                if re.search(r"(MUNDARIJA|ОГЛАВЛЕНИЕ|CONTENTS|\.{4,}\s*\d+)", txt, re.IGNORECASE):
                    found_toc_texts.append(f"--- Mundarija (bet {p_no + 1}) ---\n{txt[:1500]}")
            if found_toc_texts:
                sections.append("[MATNLI MUNDARIJA]:\n" + "\n".join(found_toc_texts)[:BUDGET_TOC])

        # 2. Introduction & Abstract
        intro_chunks = []
        intro_chars = 0
        for p_no in range(min(12, total_pages)):
            txt = doc[p_no].get_text("text").strip()
            if len(txt) < 60:
                continue
            snippet = f"[Bet {p_no + 1}] {txt[:1200]}"
            if intro_chars + len(snippet) > BUDGET_INTRO:
                break
            intro_chunks.append(snippet)
            intro_chars += len(snippet)
        if intro_chunks:
            sections.append("[KITOB ANNOTATSIYASI VA KIRISH]:\n" + "\n\n".join(intro_chunks))

        # 3. Main Chapters
        target_pages: list[int] = []
        if toc_chapter_pages:
            unique_toc_pages = sorted(set(toc_chapter_pages))
            step = max(1, len(unique_toc_pages) // 12)
            target_pages = unique_toc_pages[::step][:12]
        else:
            scan_step = 1 if total_pages <= 150 else max(1, total_pages // 120)
            detected_headers = []
            for p_no in range(10, max(11, total_pages - 6), scan_step):
                head_txt = doc[p_no].get_text("text", clip=fitz.Rect(0, 0, 600, 260)).strip()
                if CHAPTER_HEADER_RE.search(head_txt):
                    detected_headers.append(p_no)
            if len(detected_headers) >= 4:
                h_step = max(1, len(detected_headers) // 10)
                target_pages = detected_headers[::h_step][:10]
            else:
                mid_start = min(12, total_pages - 1)
                mid_end = max(mid_start + 1, total_pages - 6)
                stride = max(1, (mid_end - mid_start) // 10)
                target_pages = list(range(mid_start, mid_end, stride))[:10]

        per_chapter_budget = max(900, BUDGET_CHAPTERS // max(1, len(target_pages)))
        chapter_chunks = []
        for p_no in target_pages:
            txt = doc[p_no].get_text("text").strip()
            if len(txt) < 500 and (p_no + 1) < total_pages:
                txt += "\n" + doc[p_no + 1].get_text("text").strip()
            if txt:
                chapter_chunks.append(f"--- Bob namunasi (Sahifa {p_no + 1}) ---\n{txt[:per_chapter_budget]}")

        if chapter_chunks:
            sections.append("[ASOSIY BOBLAR VA MAVZULAR MAZMUNI]:\n" + "\n\n".join(chapter_chunks)[:BUDGET_CHAPTERS])

        # 4. Conclusion & Summary
        if total_pages > 20:
            end_chunks = []
            for p_no in range(max(12, total_pages - 8), total_pages):
                txt = doc[p_no].get_text("text").strip()
                if len(txt) > 80:
                    end_chunks.append(f"[Bet {p_no + 1}] {txt[:1000]}")
            if end_chunks:
                sections.append("[XULOSA VA ADABIYOTLAR]:\n" + "\n\n".join(end_chunks)[:BUDGET_CONCLUSION])

    release_memory()
    combined = "\n\n".join(sections)[:max_total_chars]
    return combined, total_pages


# ── 4. GEMINI MULTI-MODEL FALLBACK ENGINE ────────────────────────────────────
GEMINI_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-flash-lite-latest",
    "gemini-3.8-flash",
    "gemini-flash-latest",
    "gemini-4-argon",
]

DEAD_MODELS_CACHE: dict[str, float] = {}
DEAD_MODEL_TTL = 3600

PRESENTATION_RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "required": ["presentations"],
    "properties": {
        "presentations": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "required": ["title", "category_id", "subtitle", "theme", "slides", "tags", "description"],
                "properties": {
                    "title": {"type": "STRING"},
                    "category_id": {"type": "INTEGER"},
                    "subtitle": {"type": "STRING"},
                    "theme": {
                        "type": "STRING",
                        "enum": ["academic_navy", "emerald_science", "tech_minimal", "cyber_dark", "warm_philosophy"]
                    },
                    "slides": {
                        "type": "ARRAY",
                        "items": {
                            "type": "OBJECT",
                            "required": ["heading", "archetype", "points"],
                            "properties": {
                                "heading": {"type": "STRING"},
                                "archetype": {
                                    "type": "STRING",
                                    "enum": ["cards_grid", "comparison_vs", "timeline_steps", "theory_concept"]
                                },
                                "points": {"type": "ARRAY", "items": {"type": "STRING"}},
                                "speaker_note": {"type": "STRING"}
                            }
                        }
                    },
                    "tags": {"type": "STRING"},
                    "description": {"type": "STRING"}
                }
            }
        }
    }
}


def repair_truncated_json(raw_text: str) -> dict:
    """Attempts to cleanly parse or repair JSON truncated by token limits."""
    clean = raw_text.strip()
    match = re.search(r"(\{[\s\S]*\})", clean)
    if match:
        clean = match.group(1)
    try:
        return json.loads(clean)
    except json.JSONDecodeError:
        last_brace = clean.rfind("}")
        if last_brace != -1:
            for suffix in ["]}", "]}]}", "]}}"]:
                try:
                    return json.loads(clean[:last_brace + 1] + suffix)
                except json.JSONDecodeError:
                    continue
        raise


async def analyze_book_with_gemini_async(
    sample_text: str,
    filename: str,
    api_key: str
) -> tuple[list[dict], str]:
    """
    Fully async, status-aware Gemini Multi-Model Fallback with circuit-breaker
    and exponential backoff + jitter for 429/503.
    """
    if not api_key:
        return [], "GEMINI_API_KEY o'rnatilmagan"

    prompt = (
        f"Sen professional taqdimotlar arxitektori va ilmiy muharrirsan.\n"
        f"Quyidagi kitob/hujjatdan olingan matnni chuqur tahlil qilib, uni Soff.uz ta'limiy platformasi uchun "
        f"2 tadan 5 tagacha mustaqil, to'liq ochib berilgan PowerPoint taqdimotlariga ajratib ber.\n\n"
        f"Fayl nomi: {filename}\n"
        f"Kitobdan olingan mundarija va matn parchalari:\n\"\"\"\n{sample_text}\n\"\"\"\n\n"
        f"TALABLAR:\n"
        f"1. Har bir taqdimot toza o'zbek adabiy tilida (lotin yozuvida) bo'lsin.\n"
        f"2. Har bir taqdimotda 7 tadan 10 tagacha mustaqil, chuqur tahliliy slayd bo'lsin.\n"
        f"3. Har bir slaydda 3-4 tadan aniq, mazmunli asosiy punktlar (tushunchalar, formulalar, faktlar) bo'lsin.\n"
        f"4. Har bir slayd uchun mazmuniga mos 'archetype' ('cards_grid', 'comparison_vs', 'timeline_steps', 'theory_concept') "
        f"va qisqa spiker izohi ('speaker_note') berilsin.\n"
        f"5. Kategoriya ID lari: Tibbiyot=1478, Pedagogika=1463, Fizika=1433, Tarix=1475, Iqtisodiyot=1442, "
        f"Informatika=1440, Adabiyot=1418, Kimyo=1445, Biologiya=1423, Matematika=3215, Huquq=1439."
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "responseMimeType": "application/json",
            "responseSchema": PRESENTATION_RESPONSE_SCHEMA,
            "temperature": 0.2,
            "maxOutputTokens": 8192
        }
    }

    last_error = ""
    now = time.time()
    timeout = aiohttp.ClientTimeout(total=65, connect=10)

    async with aiohttp.ClientSession(timeout=timeout) as session:
        for model in GEMINI_MODELS:
            if model in DEAD_MODELS_CACHE and (now - DEAD_MODELS_CACHE[model]) < DEAD_MODEL_TTL:
                continue

            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"

            for attempt in range(1, 3):
                try:
                    async with session.post(url, json=payload) as resp:
                        status = resp.status
                        body_text = await resp.text()

                        if status == 200:
                            resp_json = json.loads(body_text)
                            candidates = resp_json.get("candidates") or []
                            if not candidates:
                                last_error = f"{model}: Bo'sh javob"
                                break

                            cand = candidates[0]
                            finish_reason = cand.get("finishReason", "STOP")
                            if finish_reason in ("SAFETY", "RECITATION", "PROHIBITED_CONTENT"):
                                last_error = f"{model}: Kontent filtrlandi ({finish_reason})"
                                break

                            parts = cand.get("content", {}).get("parts") or []
                            if not parts or "text" not in parts[0]:
                                last_error = f"{model}: Javob tarkibida matn topilmadi"
                                break

                            data = repair_truncated_json(parts[0]["text"])
                            presentations = data.get("presentations", [])
                            if presentations:
                                print(f"[SUCCESS] {model} orqali {len(presentations)} ta taqdimot ajratildi!")
                                return presentations, f"OK ({model})"
                            last_error = f"{model}: presentations ro'yxati bo'sh"
                            break

                        elif status == 404:
                            DEAD_MODELS_CACHE[model] = time.time()
                            last_error = f"{model} (404 Not Found)"
                            break

                        elif status in (429, 503):
                            retry_after = float(resp.headers.get("Retry-After", 0))
                            backoff = max(retry_after, (2 ** attempt) + random.uniform(0.4, 1.4))
                            last_error = f"{model} ({status}), {backoff:.1f}s kutilmoqda..."
                            await asyncio.sleep(backoff)
                            continue

                        elif status in (400, 401, 403):
                            if status == 400 and "responseSchema" in payload["generationConfig"]:
                                fallback_payload = {
                                    "contents": payload["contents"],
                                    "generationConfig": {"responseMimeType": "application/json", "temperature": 0.2}
                                }
                                async with session.post(url, json=fallback_payload) as resp2:
                                    if resp2.status == 200:
                                        rj2 = await resp2.json()
                                        raw2 = rj2["candidates"][0]["content"]["parts"][0]["text"]
                                        pres2 = repair_truncated_json(raw2).get("presentations", [])
                                        if pres2:
                                            return pres2, f"OK ({model}-no-schema)"
                            if status in (401, 403):
                                return [], f"Kritik API kalit xatosi ({status})"
                            last_error = f"{model} ({status})"
                            break

                        else:
                            last_error = f"{model} ({status})"
                            break

                except asyncio.TimeoutError:
                    last_error = f"{model}: Timeout"
                except Exception as exc:
                    last_error = f"{model} xatosi: {exc}"
                    break

    return [], last_error


# ── 5. PPTX GENERATION ENGINE (MULTI-THEME & MULTI-ARCHETYPE) ────────────────
UZ_ORPHANS = {
    "va", "ham", "yoki", "ammo", "lekin", "biroq", "chunki", "uchun", "bilan",
    "kabi", "ko'ra", "qarab", "oid", "bu", "shu", "har", "bir", "eng", "juda", "agar"
}


def bind_uzbek_orphans(text: str) -> str:
    if not text:
        return ""
    words = text.replace("--", " — ").split()
    out = []
    for i, w in enumerate(words):
        clean_w = re.sub(r"[^a-zA-Z0-9\']", "", w).lower()
        if clean_w in UZ_ORPHANS and i < len(words) - 1:
            out.append(w + "\u00A0")
        else:
            out.append(w + " ")
    return "".join(out).strip()


def apply_anti_word_break(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("latinLnBrk", "0")
    pPr.set("eaLnBrk", "0")
    pPr.set("hangingPunct", "1")


DESIGN_THEMES = {
    "academic_navy": {
        "bg": (248, 250, 252),
        "cover_bg": (15, 23, 42),
        "primary": (30, 58, 138),
        "accent": (217, 119, 6),
        "card_bg": (255, 255, 255),
        "border": (203, 213, 225),
        "text_main": (15, 23, 42),
        "text_body": (51, 65, 85),
        "text_muted": (100, 116, 139),
    },
    "emerald_science": {
        "bg": (240, 253, 244),
        "cover_bg": (6, 78, 59),
        "primary": (5, 150, 105),
        "accent": (13, 148, 136),
        "card_bg": (255, 255, 255),
        "border": (167, 243, 208),
        "text_main": (6, 78, 59),
        "text_body": (31, 41, 55),
        "text_muted": (75, 85, 99),
    },
    "tech_minimal": {
        "bg": (250, 250, 252),
        "cover_bg": (30, 27, 75),
        "primary": (37, 99, 235),
        "accent": (124, 58, 237),
        "card_bg": (255, 255, 255),
        "border": (226, 232, 240),
        "text_main": (15, 23, 42),
        "text_body": (51, 65, 85),
        "text_muted": (100, 116, 139),
    },
    "warm_philosophy": {
        "bg": (255, 251, 235),
        "cover_bg": (69, 26, 3),
        "primary": (180, 83, 9),
        "accent": (217, 119, 6),
        "card_bg": (255, 255, 255),
        "border": (253, 230, 138),
        "text_main": (69, 26, 3),
        "text_body": (68, 64, 60),
        "text_muted": (120, 113, 108),
    },
    "cyber_dark": {
        "bg": (15, 23, 42),
        "cover_bg": (2, 6, 23),
        "primary": (56, 189, 248),
        "accent": (168, 85, 247),
        "card_bg": (30, 41, 59),
        "border": (51, 65, 85),
        "text_main": (248, 250, 252),
        "text_body": (226, 232, 240),
        "text_muted": (148, 163, 184),
    }
}


def auto_contrast_color(rgb_tuple: tuple[int, int, int]) -> RGBColor:
    """Luminance contrast formula Y = 0.299R + 0.587G + 0.114B."""
    r, g, b = rgb_tuple
    luminance = 0.299 * r + 0.587 * g + 0.114 * b
    return RGBColor(255, 255, 255) if luminance < 128 else RGBColor(15, 23, 42)


def inject_morph_transition(slide) -> None:
    """Office 365 Native Morph (p16:morph) with Fade fallback."""
    MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    P16_NS = "http://schemas.microsoft.com/office/powerpoint/2015/main"
    trans_xml = (
        f'<p:transition {nsdecls("p")} xmlns:mc="{MC_NS}" spd="med" advClick="1">'
        f'  <mc:AlternateContent>'
        f'    <mc:Choice Requires="p16" xmlns:p16="{P16_NS}">'
        f'      <p16:morph option="byObject"/>'
        f'    </mc:Choice>'
        f'    <mc:Fallback><p:fade/></mc:Fallback>'
        f'  </mc:AlternateContent>'
        f'</p:transition>'
    )
    try:
        slide.element.append(parse_xml(trans_xml))
    except Exception:
        pass


def set_slide_background(slide, rgb_tuple: tuple[int, int, int]) -> None:
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb_tuple)


def generate_rich_pptx_file(spec: dict, out_path: str, deck_index: int = 0) -> int:
    """
    Creates high-aesthetic 16:9 presentation with Cover + Smart Agenda +
    Multi-archetype content slides + Conclusion slide.
    """
    theme_keys = list(DESIGN_THEMES.keys())
    requested_theme = spec.get("theme")
    if requested_theme not in DESIGN_THEMES:
        requested_theme = theme_keys[deck_index % len(theme_keys)]
    th = DESIGN_THEMES[requested_theme]

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    title_str = bind_uzbek_orphans(spec.get("title", "Ilmiy Taqdimot"))
    subtitle_str = bind_uzbek_orphans(spec.get("subtitle", "Ilmiy-amaliy ta'limiy taqdimot"))
    slides_info = spec.get("slides", [])

    # 1. Dark Hero Cover Slide
    s_cover = prs.slides.add_slide(blank)
    set_slide_background(s_cover, th["cover_bg"])
    cover_text_rgb = auto_contrast_color(th["cover_bg"])

    bar = s_cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(1.6), Inches(0.12), Inches(4.2))
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(*th["accent"])
    bar.line.fill.background()

    tb1 = s_cover.shapes.add_textbox(Inches(1.3), Inches(1.6), Inches(10.8), Inches(4.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_badge = tf1.paragraphs[0]
    p_badge.text = "SLAYDHUB AI & SOFF.UZ  |  AKADEMIK TAQDIMOT"
    p_badge.font.size = Pt(12)
    p_badge.font.bold = True
    p_badge.font.color.rgb = RGBColor(*th["accent"])

    p_title = tf1.add_paragraph()
    p_title.text = title_str
    p_title.font.size = Pt(36 if len(title_str) < 65 else 30)
    p_title.font.bold = True
    p_title.font.color.rgb = cover_text_rgb
    p_title.space_before = Pt(14)
    apply_anti_word_break(p_title)

    p_sub = tf1.add_paragraph()
    p_sub.text = subtitle_str
    p_sub.font.size = Pt(18)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)
    p_sub.space_before = Pt(18)
    apply_anti_word_break(p_sub)

    # 2. Smart Agenda Slide
    if len(slides_info) >= 3:
        s_agenda = prs.slides.add_slide(blank)
        set_slide_background(s_agenda, th["bg"])
        inject_morph_transition(s_agenda)

        tb_ag = s_agenda.shapes.add_textbox(Inches(0.9), Inches(0.5), Inches(11.5), Inches(0.9))
        p_ag = tb_ag.text_frame.paragraphs[0]
        p_ag.text = "Taqdimotning Mantiqiy Rejasi (Mundarija)"
        p_ag.font.size = Pt(26)
        p_ag.font.bold = True
        p_ag.font.color.rgb = RGBColor(*th["text_main"])

        max_agenda_items = min(8, len(slides_info))
        for idx, s_item in enumerate(slides_info[:max_agenda_items], 1):
            col = 0 if idx <= 4 else 1
            row = (idx - 1) % 4
            x_pos = Inches(0.9 + col * 5.85)
            y_pos = Inches(1.65 + row * 1.25)

            badge = s_agenda.shapes.add_textbox(x_pos, y_pos, Inches(0.75), Inches(0.85))
            bp = badge.text_frame.paragraphs[0]
            bp.text = f"{idx:02d}"
            bp.font.size = Pt(22)
            bp.font.bold = True
            bp.font.color.rgb = RGBColor(*th["primary"])

            item_tb = s_agenda.shapes.add_textbox(x_pos + Inches(0.8), y_pos + Inches(0.05), Inches(4.8), Inches(0.95))
            item_tf = item_tb.text_frame
            item_tf.word_wrap = True
            ip = item_tf.paragraphs[0]
            ip.text = bind_uzbek_orphans(s_item.get("heading", f"Mavzu {idx}"))
            ip.font.size = Pt(16)
            ip.font.bold = True
            ip.font.color.rgb = RGBColor(*th["text_body"])
            apply_anti_word_break(ip)

    # 3. Content Slides with Archetype Rotation
    archetypes_cycle = ["cards_grid", "timeline_steps", "comparison_vs", "theory_concept"]

    for idx, s_info in enumerate(slides_info, 1):
        sl = prs.slides.add_slide(blank)
        set_slide_background(sl, th["bg"])
        inject_morph_transition(sl)

        num_tb = sl.shapes.add_textbox(Inches(0.9), Inches(0.4), Inches(4.5), Inches(0.35))
        np = num_tb.text_frame.paragraphs[0]
        np.text = f"BO'LIM {idx:02d} / {len(slides_info):02d}"
        np.font.size = Pt(11)
        np.font.bold = True
        np.font.color.rgb = RGBColor(*th["primary"])

        head_tb = sl.shapes.add_textbox(Inches(0.9), Inches(0.75), Inches(11.5), Inches(1.0))
        head_tf = head_tb.text_frame
        head_tf.word_wrap = True
        hp = head_tf.paragraphs[0]
        hp.text = bind_uzbek_orphans(s_info.get("heading", f"Mavzu {idx}"))
        hp.font.size = Pt(25)
        hp.font.bold = True
        hp.font.color.rgb = RGBColor(*th["text_main"])
        apply_anti_word_break(hp)

        div = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(1.78), Inches(11.5), Inches(0.03))
        div.fill.solid()
        div.fill.fore_color.rgb = RGBColor(*th["border"])
        div.line.fill.background()

        points = [bind_uzbek_orphans(pt) for pt in s_info.get("points", []) if str(pt).strip()]
        archetype = s_info.get("archetype") or archetypes_cycle[(idx - 1) % len(archetypes_cycle)]

        if archetype == "comparison_vs" and len(points) >= 2:
            mid = max(1, len(points) // 2)
            cols_data = [("I. Asosiy Jihatlar va Nazariya", points[:mid]), ("II. Amaliy Natijalar va Tahlil", points[mid:])]
            for c_idx, (c_title, c_pts) in enumerate(cols_data):
                cx = Inches(0.9 + c_idx * 5.9)
                accent_line = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, Inches(2.05), Inches(5.55), Inches(0.06))
                accent_line.fill.solid()
                accent_line.fill.fore_color.rgb = RGBColor(*(th["primary"] if c_idx == 0 else th["accent"]))
                accent_line.line.fill.background()

                ctb = sl.shapes.add_textbox(cx, Inches(2.2), Inches(5.55), Inches(4.6))
                ctf = ctb.text_frame
                ctf.word_wrap = True
                cp0 = ctf.paragraphs[0]
                cp0.text = c_title
                cp0.font.size = Pt(18)
                cp0.font.bold = True
                cp0.font.color.rgb = RGBColor(*(th["primary"] if c_idx == 0 else th["accent"]))
                cp0.space_after = Pt(12)

                for pt in c_pts:
                    pp = ctf.add_paragraph()
                    pp.text = f"▸  {pt}"
                    pp.font.size = Pt(16)
                    pp.font.color.rgb = RGBColor(*th["text_body"])
                    pp.space_after = Pt(10)
                    apply_anti_word_break(pp)

        elif archetype == "timeline_steps" and len(points) >= 3:
            n_steps = min(4, len(points))
            col_w = 11.5 / n_steps
            for s_i, pt in enumerate(points[:n_steps], 1):
                sx = Inches(0.9 + (s_i - 1) * col_w)
                step_bar = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, sx, Inches(2.15), Inches(col_w - 0.25), Inches(0.08))
                step_bar.fill.solid()
                step_bar.fill.fore_color.rgb = RGBColor(*th["primary"])
                step_bar.line.fill.background()

                stb = sl.shapes.add_textbox(sx, Inches(2.35), Inches(col_w - 0.25), Inches(4.4))
                stf = stb.text_frame
                stf.word_wrap = True
                sp_num = stf.paragraphs[0]
                sp_num.text = f"BOSQICH 0{s_i}"
                sp_num.font.size = Pt(13)
                sp_num.font.bold = True
                sp_num.font.color.rgb = RGBColor(*th["accent"])
                sp_num.space_after = Pt(8)

                sp_txt = stf.add_paragraph()
                sp_txt.text = pt
                sp_txt.font.size = Pt(15.5)
                sp_txt.font.color.rgb = RGBColor(*th["text_body"])
                apply_anti_word_break(sp_txt)

        else:
            n_pts = min(4, max(1, len(points)))
            row_h = 4.8 / n_pts
            for p_i, pt in enumerate(points[:4]):
                ry = Inches(2.05 + p_i * row_h)
                left_bar = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), ry + Inches(0.08), Inches(0.08), Inches(row_h - 0.25))
                left_bar.fill.solid()
                left_bar.fill.fore_color.rgb = RGBColor(*(th["primary"] if p_i % 2 == 0 else th["accent"]))
                left_bar.line.fill.background()

                ptb = sl.shapes.add_textbox(Inches(1.15), ry, Inches(11.2), Inches(row_h - 0.1))
                ptf = ptb.text_frame
                ptf.word_wrap = True
                pp = ptf.paragraphs[0]
                pp.text = pt
                pp.font.size = Pt(17.5)
                pp.font.color.rgb = RGBColor(*th["text_body"])
                apply_anti_word_break(pp)

        # Speaker notes
        try:
            notes_slide = sl.notes_slide
            speaker_note = s_info.get("speaker_note") or (
                f"🎙️ Spiker nutqi ({s_info.get('heading', '')}): " + " ".join(points[:2])
            )
            notes_slide.notes_text_frame.text = speaker_note
        except Exception:
            pass

    # 4. Conclusion Slide
    s_end = prs.slides.add_slide(blank)
    set_slide_background(s_end, th["cover_bg"])
    inject_morph_transition(s_end)

    tb_end = s_end.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(11.333), Inches(2.8))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True
    pe1 = tf_end.paragraphs[0]
    pe1.text = "E'tiboringiz uchun rahmat!"
    pe1.font.size = Pt(38)
    pe1.font.bold = True
    pe1.font.color.rgb = auto_contrast_color(th["cover_bg"])

    pe2 = tf_end.add_paragraph()
    pe2.text = f"«{title_str}» mavzusi bo'yicha savol va mulohazalar"
    pe2.font.size = Pt(18)
    pe2.font.color.rgb = RGBColor(*th["accent"])
    pe2.space_before = Pt(16)

    prs.save(out_path)
    return len(prs.slides)


# ── 6. SOFF.UZ ASYNC API CLIENT WITH JWT AUTO-REFRESH & IDEMPOTENCY ──────────
class SoffAPIClient:
    """
    Soff.uz 2-step API client with JWT expiration check, auto-refresh,
    and idempotency locking to prevent double-upload duplicates.
    """
    BASE_URL = "https://api.soff.uz/api/v1"

    def __init__(
        self,
        access_token: str | None = None,
        refresh_token: str | None = None,
        login_phone: str | None = None,
        login_password: str | None = None
    ):
        self.access_token = access_token or os.getenv("SELLER_TOKEN", "")
        self.refresh_token = refresh_token or os.getenv("SOFF_REFRESH_TOKEN", "")
        self.login_phone = login_phone or os.getenv("SOFF_PHONE", "")
        self.login_password = login_password or os.getenv("SOFF_PASSWORD", "")
        self._auth_lock = asyncio.Lock()
        self._active_uploads: set[str] = set()

    @staticmethod
    def _is_jwt_expiring_soon(token: str, buffer_seconds: int = 300) -> bool:
        if not token or token.count(".") != 2:
            return False
        try:
            payload_b64 = token.split(".")[1]
            padding = "=" * (-len(payload_b64) % 4)
            payload = json.loads(base64.urlsafe_b64decode(payload_b64 + padding).decode("utf-8"))
            exp = payload.get("exp")
            if exp and (exp - time.time()) < buffer_seconds:
                return True
        except Exception:
            pass
        return False

    async def ensure_valid_token(self, session: aiohttp.ClientSession, force: bool = False) -> bool:
        if not force and self.access_token and not self._is_jwt_expiring_soon(self.access_token):
            return True

        async with self._auth_lock:
            if not force and self.access_token and not self._is_jwt_expiring_soon(self.access_token):
                return True

            if self.refresh_token:
                try:
                    async with session.post(
                        f"{self.BASE_URL}/auth/token/refresh/",
                        json={"refresh": self.refresh_token},
                        timeout=aiohttp.ClientTimeout(total=15)
                    ) as resp:
                        if resp.status in (200, 201):
                            data = await resp.json()
                            new_access = data.get("access") or data.get("token")
                            if new_access:
                                self.access_token = new_access
                                if data.get("refresh"):
                                    self.refresh_token = data["refresh"]
                                return True
                except Exception as exc:
                    print(f"[SOFF AUTH] Refresh token xatosi: {exc}")

            if self.login_phone and self.login_password:
                try:
                    async with session.post(
                        f"{self.BASE_URL}/auth/login/",
                        json={"phone": self.login_phone, "password": self.login_password},
                        timeout=aiohttp.ClientTimeout(total=15)
                    ) as resp:
                        if resp.status in (200, 201):
                            data = await resp.json()
                            new_access = data.get("access") or data.get("token")
                            if new_access:
                                self.access_token = new_access
                                self.refresh_token = data.get("refresh", self.refresh_token)
                                return True
                except Exception as exc:
                    print(f"[SOFF AUTH] Login xatosi: {exc}")

            return bool(self.access_token)

    async def upload_two_step(self, upload_id: str, item: dict) -> dict:
        if item.get("product_id"):
            return {"ok": True, "product_id": item["product_id"], "already_uploaded": True}

        if upload_id in self._active_uploads:
            return {"ok": False, "busy": True, "error": "Ushbu fayl hozirda yuklanmoqda, iltimos kuting..."}

        self._active_uploads.add(upload_id)
        try:
            fp = item.get("file_path", "")
            if not fp or not os.path.exists(fp):
                return {"ok": False, "error": "Fayl server xotirasidan topilmadi."}

            timeout = aiohttp.ClientTimeout(total=60, connect=10)
            async with aiohttp.ClientSession(timeout=timeout) as session:
                await self.ensure_valid_token(session)

                # STEP 1: Upload presentation file (reuse doc_id if already cached)
                doc_id = item.get("doc_id")
                if not doc_id:
                    fn = os.path.basename(fp)
                    for attempt in range(1, 4):
                        headers = {
                            "Authorization": f"Bearer {self.access_token}",
                            "User-Agent": "SlaydHubBot/3.0"
                        }
                        form1 = aiohttp.FormData()
                        form1.add_field("content_type", "file")
                        with open(fp, "rb") as fh:
                            file_bytes = fh.read()
                        form1.add_field(
                            "file",
                            file_bytes,
                            filename=fn,
                            content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                        try:
                            async with session.post(
                                f"{self.BASE_URL}/seller/product-create-first/",
                                data=form1,
                                headers=headers
                            ) as r1:
                                if r1.status in (200, 201):
                                    rj1 = await r1.json()
                                    doc_id = rj1.get("id")
                                    if doc_id:
                                        item["doc_id"] = doc_id
                                        break
                                elif r1.status == 401 and attempt == 1:
                                    await self.ensure_valid_token(session, force=True)
                                    continue
                                elif r1.status in (429, 500, 502, 503, 504):
                                    await asyncio.sleep(2 ** attempt)
                                    continue
                                else:
                                    err_txt = await r1.text()
                                    return {"ok": False, "error": f"Step-1 xatosi ({r1.status}): {err_txt[:120]}"}
                        except Exception as exc:
                            if attempt == 3:
                                return {"ok": False, "error": f"Step-1 ulanish xatosi: {exc}"}
                            await asyncio.sleep(2 ** attempt)

                if not doc_id:
                    return {"ok": False, "error": "Step-1 dan document_id olinmadi."}

                # STEP 2: Create Product Metadata
                for attempt in range(1, 4):
                    headers = {
                        "Authorization": f"Bearer {self.access_token}",
                        "User-Agent": "SlaydHubBot/3.0"
                    }
                    form2 = aiohttp.FormData()
                    form2.add_field("title", str(item["title"]))
                    form2.add_field("category", str(item.get("category_id", 1463)))
                    form2.add_field("price", str(item.get("price", 7000)))
                    form2.add_field("document", str(doc_id))
                    form2.add_field("description", str(item.get("description", "")))

                    try:
                        async with session.post(
                            f"{self.BASE_URL}/seller/product-create-second/",
                            data=form2,
                            headers=headers
                        ) as r2:
                            if r2.status in (200, 201):
                                rj2 = await r2.json()
                                new_id = rj2.get("id")
                                item["product_id"] = new_id
                                try:
                                    if os.path.exists(fp):
                                        os.remove(fp)
                                except OSError:
                                    pass
                                return {"ok": True, "product_id": new_id}
                            elif r2.status == 401 and attempt == 1:
                                await self.ensure_valid_token(session, force=True)
                                continue
                            elif r2.status in (429, 500, 502, 503, 504):
                                await asyncio.sleep(2 ** attempt)
                                continue
                            else:
                                err_txt = await r2.text()
                                return {"ok": False, "error": f"Step-2 xatosi ({r2.status}): {err_txt[:120]}"}
                    except Exception as exc:
                        if attempt == 3:
                            return {"ok": False, "error": f"Step-2 ulanish xatosi: {exc}"}
                        await asyncio.sleep(2 ** attempt)

            return {"ok": False, "error": "Step-2 maksimal urinishlardan so'ng ham bajarilmadi."}
        finally:
            self._active_uploads.discard(upload_id)


soff_client = SoffAPIClient(
    access_token=SELLER_TOKEN,
    refresh_token=SOFF_REFRESH_TOKEN,
    login_phone=SOFF_PHONE,
    login_password=SOFF_PASSWORD
)


# ── 7. TELEGRAM BOT HANDLERS ─────────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: types.Message):
    await message.answer(
        "👋 **Assalomu alaykum! SlaydHub & Soff.uz 24/7 avtonom botiga xush kelibsiz!**\n\n"
        "📚 Menga istalgan **PDF kitob yoki o'quv qo'llanmani** yuboring.\n\n"
        "⚡ **Yangi imkoniyatlar:**\n"
        "1. 8–50 MB kitoblarni tahlil qilish va xulosasini yo'qotmaslik (Smart Chunking v2);\n"
        "2. 5 xil akademik mavzu (Navy, Emerald, Tech, Warm, Cyber) va 4 xil slayd maketi;\n"
        "3. Har bir slaydga mantiqiy reja (Agenda), Morph animatsiyasi va spiker nutqi;\n"
        "4. Bitta tugma bilan **Soff.uz** ga to'g'ridan-to'g'ri joylash (Idempotent 2-bosqich)!\n\n"
        "📥 *Boshlash uchun PDF fayl yuboring!*",
        parse_mode="Markdown"
    )


@dp.message(F.document)
async def handle_document(message: types.Message):
    doc = message.document
    if not (doc.file_name and doc.file_name.lower().endswith(".pdf")):
        await message.answer("⚠️ Iltimos, faqat **PDF** formatidagi kitob yoki hujjat yuboring.")
        return

    file_size_bytes = doc.file_size or 0
    file_size_mb = round(file_size_bytes / (1024 * 1024), 1)

    # Pre-check Telegram 20MB limit
    if file_size_bytes > 20 * 1024 * 1024 and not os.getenv("LOCAL_BOT_API_URL"):
        await message.answer(
            f"⚠️ **«{doc.file_name}»** hajmi **{file_size_mb} MB**.\n"
            f"Standart Telegram Bot API maksimal 20 MB gacha fayl qabul qiladi. "
            f"Iltimos, PDF faylni biroz siqilgan holda yuboring."
        )
        return

    status_msg = await message.answer(
        f"📥 **«{doc.file_name}»** ({file_size_mb} MB) navbatga olindi.\n"
        f"⏳ 1/4: Server resurslari tayyorlanmoqda..."
    )

    async with HEAVY_TASK_SEMAPHORE:
        with tempfile.TemporaryDirectory() as tmp_dir:
            pdf_path = os.path.join(tmp_dir, "input_book.pdf")

            # 1. Download
            try:
                await status_msg.edit_text(f"📥 1/4: **«{doc.file_name}»** ({file_size_mb} MB) yuklab olinmoqda...")
                file_info = await bot.get_file(doc.file_id)
                await bot.download_file(file_info.file_path, pdf_path)
            except Exception as exc:
                await status_msg.edit_text(f"❌ Faylni yuklab olishda xatolik: {exc}")
                return

            # 2. Extract smart text (in thread to keep event loop free)
            await status_msg.edit_text("📖 2/4: Kitob mundarijasi va asosiy boblari o'qilmoqda...")
            try:
                sample_text, total_pages = await asyncio.to_thread(extract_smart_pdf_text_v2, pdf_path)
                if len(sample_text.strip()) < 80:
                    await status_msg.edit_text(
                        "⚠️ Ushbu PDF faylda matn qatlami topilmadi (ehtimol skanerlangan rasm). "
                        "Iltimos, raqamli PDF yuboring."
                    )
                    return
            except Exception as exc:
                await status_msg.edit_text(f"❌ PDF faylni o'qishda xatolik: {exc}")
                return

            # Immediately delete input PDF to free RAM/tmpfs
            try:
                os.remove(pdf_path)
            except OSError:
                pass

            # 3. Gemini Multi-Model Fallback
            await status_msg.edit_text(
                f"🧠 3/4: **Gemini AI** kitobni ({total_pages} sahifa) tahlil qilib, taqdimotlarga ajratmoqda..."
            )
            specs, status_info = await analyze_book_with_gemini_async(sample_text, doc.file_name, GEMINI_API_KEY)
            if not specs:
                await status_msg.edit_text(
                    f"❌ Kitobni mavzularga ajratishda xatolik: {status_info}\nIltimos, qayta urinib ko'ring."
                )
                return

            # 4. Generate PPTX and send
            await status_msg.edit_text(
                f"🎨 4/4: Kitobdan **{len(specs)} ta mustaqil taqdimot** aniqlandi! "
                f"Xilma-xil dizayndagi slaydlar (.pptx) shakllantirilmoqda..."
            )

            perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
            os.makedirs(perm_dir, exist_ok=True)

            user_id = message.from_user.id if message.from_user else 0
            for idx, spec in enumerate(specs, 1):
                title = spec.get("title", f"Taqdimot {idx}")
                safe_title = "".join(c for c in title if c.isalnum() or c in (" ", "_", "-")).strip()[:80] or f"Taqdimot_{idx}"
                upload_id = f"up_{user_id}_{idx}_{int(time.time())}"
                perm_pptx = os.path.join(perm_dir, f"{upload_id}_{safe_title}.pptx")

                slide_count = await asyncio.to_thread(generate_rich_pptx_file, spec, perm_pptx, idx - 1)
                price = 3000 if slide_count <= 6 else 7000 if slide_count <= 12 else 10000

                full_desc = (
                    f"{spec.get('description', '')}\n"
                    f"<p><strong>Jami slaydlar soni:</strong> {slide_count} ta slayd</p>\n"
                    f"<p>{spec.get('tags', '')}</p>"
                )

                PENDING_UPLOADS[upload_id] = {
                    "file_path": perm_pptx,
                    "title": title,
                    "category_id": spec.get("category_id", 1463),
                    "price": price,
                    "description": full_desc,
                    "created_at": time.time(),
                    "doc_id": None,
                    "product_id": None,
                }

                kb = InlineKeyboardMarkup(inline_keyboard=[
                    [InlineKeyboardButton(text="🚀 Soff.uz ga yuklash", callback_data=f"soff_{upload_id}")]
                ])

                caption = (
                    f"📊 **{idx}-Taqdimot: {title}**\n\n"
                    f"📄 Slaydlar soni: **{slide_count} ta** (Cover + Reja + Mazmun + Xulosa)\n"
                    f"🎨 Dizayn mavzusi: `{spec.get('theme', 'auto')}`\n"
                    f"💰 Tavsiya etilgan narx: **{price:,} so'm**\n"
                    f"🏷️ Teglar: {spec.get('tags', '')}\n\n"
                    f"👇 Quyidagi tugma orqali Soff.uz kabinetingizga 1-klikda joylang:"
                )

                await message.answer_document(
                    FSInputFile(perm_pptx, filename=f"{safe_title}.pptx"),
                    caption=caption,
                    reply_markup=kb,
                    parse_mode="Markdown"
                )

            release_memory()
            await status_msg.delete()


@dp.callback_query(F.data.startswith("soff_"))
async def handle_soff_upload(call: types.CallbackQuery):
    upload_id = call.data.replace("soff_", "")
    item = PENDING_UPLOADS.get(upload_id)
    if not item:
        await call.answer("⚠️ Ushbu fayl ma'lumoti eskirgan yoki topilmadi.", show_alert=True)
        return

    await call.answer("🚀 Soff.uz ga yuklanmoqda...", show_alert=False)

    res = await soff_client.upload_two_step(upload_id, item)
    if res.get("ok"):
        new_id = res["product_id"]
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Soff.uz da ko'rish", url=f"https://soff.uz/seller/product-update/{new_id}/")]
        ])
        try:
            await call.message.edit_reply_markup(reply_markup=kb)
        except Exception:
            pass
        await call.message.reply(
            f"🎉 **Muvaffaqiyatli yuklandi!**\nSoff.uz Mahsulot ID: `{new_id}`",
            parse_mode="Markdown"
        )
    else:
        err = res.get("error", "Noma'lum xatolik")
        await call.message.reply(f"❌ Soff.uz ga yuklashda xatolik: {err}")


# ── 8. WEB HEALTHCHECK SERVER (RENDER KEEPALIVE) ─────────────────────────────
async def handle_ping(request):
    return web.json_response({
        "status": "ok",
        "service": "SlaydHub AI & Soff.uz Bot",
        "version": "3.0-Argon-Audited",
        "uptime": "24/7"
    })


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", handle_ping)
    app.router.add_get("/health", handle_ping)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    print(f"[WEB] Health-Check server port {PORT} da ishga tushdi.")


# ── 9. MAIN RUNNER ───────────────────────────────────────────────────────────
async def main():
    print("[BOT] SlaydHub Telegram bot ishga tushmoqda...")
    await start_web_server()
    # Launch background cache cleanup
    asyncio.create_task(cleanup_expired_uploads_loop())
    print("[BOT] Polling muvaffaqiyatli boshlandi.")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
