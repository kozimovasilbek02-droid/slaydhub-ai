# -*- coding: utf-8 -*-
import os
import sys
import time
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable
from pptx import Presentation

current_dir = Path(__file__).resolve().parent
bot_dir = current_dir.parent
project_root = bot_dir.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from backend.core.gemini_translator import GeminiTranslator
from backend.core.pptx_processor import PPTXProcessor
from bot.config import GEMINI_API_KEY


class TranslatorService:
    """
    SlideTranslate AI (https://github.com/kozimovasilbek02-droid/slidetranslate-ai)
    repozitoriyasidagi mukammal dvigatel asosida ishlovchi tarjima xizmati.
    Reklama slaydlarini tozalash, auto-fit va yuqori sifatli kontekstli tarjimani ta'minlaydi.
    """
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or GEMINI_API_KEY
        self.translator = GeminiTranslator(api_key=self.api_key)

    def get_clean_presentation_title(self, raw_filename: str, target_script: str = "latin") -> str:
        import re
        base, _ = os.path.splitext(raw_filename)
        base = re.sub(r"[\(\[\{]\d+[\)\]\}]", "", base)
        base = re.sub(r"[-_]?\s*(powerpoint\s*templates?|templates?|shablon(lar)?|tarjima(si)?|ozbekcha|uz|translated|translation)", "", base, flags=re.IGNORECASE).strip()

        try:
            translated = self.translator.translate_single_text(f"Translate presentation title to Uzbek: {base}", target_script=target_script)
            if translated and translated.strip():
                clean = re.sub(r"^(taqdimot\s*(nomi|sarlavhasi)?\s*:?|title\s*:?)", "", translated.strip(), flags=re.IGNORECASE).strip()
                if clean:
                    base = clean
        except Exception:
            pass

        base = re.sub(r'[/\\:*?"<>|_]', " ", base)
        base = re.sub(r"\s+", " ", base).strip()
        return base if base else "Taqdimot"

    async def translate_presentation(
        self,
        input_pptx: str,
        output_pptx: str,
        target_script: str = "latin",
        progress_callback: Optional[Callable[[str], Any]] = None,
        presentation_title: str = ""
    ) -> Dict[str, Any]:
        if not os.path.exists(input_pptx):
            raise FileNotFoundError(f"Fayl topilmadi: {input_pptx}")

        if progress_callback:
            await progress_callback("🔍 Slayd ichidagi barcha matnlar va tuzilma tahlil qilinmoqda...")

        # 1. Matnlarni ajratish
        data = PPTXProcessor.extract_presentation_data(input_pptx)
        all_items = []
        for slide in data.get("slides", []):
            all_items.extend(slide.get("items", []))

        total_items = len(all_items)
        initial_slide_count = data.get("slides_count", len(data.get("slides", [])))

        if total_items == 0:
            if progress_callback:
                await progress_callback("⚠️ Slaydda tarjima qilinadigan matn topilmadi.")
            return {"success": False, "error": "Slaydda matn topilmadi", "output_path": None}

        if progress_callback:
            await progress_callback(
                f"🌐 Jami {initial_slide_count} ta slayd ({total_items} ta matn bloki).\n"
                f"⚡ SlideTranslate AI dvigateli orqali parallel tarjima qilinmoqda..."
            )

        # 2. Parallel tarjima qilish (SlideTranslate AI dvigateli)
        stats = {}
        translated_results = self.translator.translate_items_batch(
            items=all_items,
            target_script=target_script,
            domain="Taqdimot, fan, ta'lim va ilmiy tahlil",
            stats=stats,
            batch_size=35,
            max_workers=4
        )

        translations_map = {
            r["id"]: r.get("translated_text") or r.get("text", "")
            for r in translated_results
        }

        # 3. Reklamalarni tozalash va tarjimani slaydga qayta kiritish
        if progress_callback:
            await progress_callback("🧹 Reklama slaydlari (SlidesCarnival, Slidesgo va b.) tozalanmoqda va matnlar joylanmoqda...")

        final_path = PPTXProcessor.apply_translations_and_export(
            original_pptx_path=input_pptx,
            translations_map=translations_map,
            output_pptx_path=output_pptx,
            auto_fit=True,
            target_script=target_script,
            clean_watermarks=True,
            presentation_title=presentation_title
        )

        # Haqiqiy yakuniy slaydlar sonini tekshirish
        try:
            final_prs = Presentation(final_path)
            final_slide_count = len(final_prs.slides)
            cleaned_ad_slides = max(0, initial_slide_count - final_slide_count)
        except Exception:
            final_slide_count = initial_slide_count
            cleaned_ad_slides = 0

        return {
            "success": True,
            "initial_slides": initial_slide_count,
            "final_slides": final_slide_count,
            "cleaned_ad_slides": cleaned_ad_slides,
            "total_items": total_items,
            "output_path": final_path
        }
