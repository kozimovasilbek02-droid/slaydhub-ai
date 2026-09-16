# -*- coding: utf-8 -*-
"""
backend.core.image_translator
-----------------------------
Gemini Vision (Nano/Flash) orqali rasmli slaydlar, diagrammalar va infografikalar
ichidagi ruscha/inglizcha matnlarni aniqlab, o'zbek tiliga tarjima qilib,
rasm grafikasini buzmasdan qayta render qiluvchi modul.
"""
from __future__ import annotations

import os
import io
import json
import re
import logging
from typing import Dict, Any, List, Optional, Tuple
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

from google import genai
from google.genai import types

from backend.core.transliteration import ensure_script

logger = logging.getLogger(__name__)


class ImageTranslator:
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY", "")
        )
        self.client = genai.Client(api_key=self.api_key) if self.api_key else genai.Client()
        self.model_name = "gemini-3.1-flash-lite"

    def analyze_and_translate_image(self, image_bytes: bytes, target_script: str = "latin") -> bytes:
        """
        Rasm ichidagi matnlarni aniqlab, ularni o'zbekchaga tarjima qiladi va
        rasmning qolgan barcha grafik qismlarini saqlagan holda matnni qayta chizadi.
        Agar rasmda matn bo'lmasa yoki xato yuz bersa, asl rasm baytlarini qaytaradi.
        """
        try:
            pil_img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
            w, h = pil_img.size

            # Juda kichik ikonkalarni (masalan < 70px) tarjima qilish shart emas
            if w < 70 or h < 70:
                return image_bytes

            prompt = (
                "You are an elite slide, diagram, flowchart, and infographic OCR translator.\n"
                "Analyze this presentation image carefully.\n"
                "Does this image contain text, diagram labels, flowchart boxes, infocards, titles, or annotations in Russian, English, or non-Uzbek text?\n\n"
                "If it has NO text (pure photo without text, pattern, icon without words), return EXACTLY:\n"
                '{"has_text": false}\n\n'
                "If it HAS text/labels to translate into Uzbek (Latin script):\n"
                "Extract each distinct text phrase. The bounding box [ymin, xmin, ymax, xmax] must be 0-1000 scale and MUST FULLY AND GENEROUSLY COVER the original text phrase and its underlying box.\n"
                "Provide the natural Uzbek translation (Latin script, uppercase if original is uppercase), background hex color 'bg_hex' of that specific box/label, and text hex color 'text_hex'.\n\n"
                "Return JSON in this format:\n"
                "{\n"
                '  "has_text": true,\n'
                '  "items": [\n'
                '    {\n'
                '      "box_2d": [ymin, xmin, ymax, xmax],\n'
                '      "original_text": "...",\n'
                '      "translated_text": "...",\n'
                '      "bg_hex": "#E0E0E0",\n'
                '      "text_hex": "#000000",\n'
                '      "is_bold": true\n'
                '    }\n'
                '  ]\n'
                "}"
            )

            # Convert image to JPEG for Gemini API
            img_buf = io.BytesIO()
            pil_img.convert("RGB").save(img_buf, format="JPEG", quality=90)
            img_jpeg_bytes = img_buf.getvalue()

            response = self.client.models.generate_content(
                model=self.model_name,
                contents=[
                    types.Part.from_bytes(data=img_jpeg_bytes, mime_type="image/jpeg"),
                    prompt
                ],
                config=types.GenerateContentConfig(
                    temperature=0.1,
                    response_mime_type="application/json"
                )
            )

            resp_text = response.text.strip()
            if resp_text.startswith("```"):
                resp_text = re.sub(r"^```[a-zA-Z]*\n", "", resp_text)
                resp_text = re.sub(r"\n```$", "", resp_text).strip()

            data = json.loads(resp_text)
            if not data.get("has_text") or not data.get("items"):
                return image_bytes

            items = data.get("items", [])
            logger.info(f"ImageTranslator: {len(items)} ta matnli blok topildi va tarjima qilinmoqda...")

            draw = ImageDraw.Draw(pil_img)
            
            # Local font discovery
            fonts_dir = Path(__file__).resolve().parent.parent / "fonts"
            roboto_path = str(fonts_dir / "Roboto.ttf")
            font_path = roboto_path if os.path.exists(roboto_path) else "arial.ttf"

            for item in items:
                box = item.get("box_2d")
                if not box or len(box) != 4:
                    continue

                ymin, xmin, ymax, xmax = box
                left = int((xmin / 1000.0) * w)
                top = int((ymin / 1000.0) * h)
                right = int((xmax / 1000.0) * w)
                bottom = int((ymax / 1000.0) * h)

                box_w = max(10, right - left)
                box_h = max(10, bottom - top)

                trans_text = item.get("translated_text", "")
                if not trans_text:
                    continue

                trans_text = ensure_script(trans_text, target_script)

                bg_hex = item.get("bg_hex", "#FFFFFF").strip()
                if not bg_hex.startswith("#") or len(bg_hex) not in (4, 7):
                    bg_hex = "#FFFFFF"
                
                text_hex = item.get("text_hex", "#000000").strip()
                if not text_hex.startswith("#") or len(text_hex) not in (4, 7):
                    text_hex = "#000000"

                # 1. Bounding box foni (Inpainting / Clean Patch)
                pad_x = 4
                pad_y = 4
                draw.rectangle(
                    [max(0, left - pad_x), max(0, top - pad_y), min(w, right + pad_x), min(h, bottom + pad_y)],
                    fill=bg_hex
                )

                # 2. Matn o'lchamini qutiga moslash
                font_size = max(11, int(box_h * 0.65))

                try:
                    font = ImageFont.truetype(font_path, font_size)
                except Exception:
                    try:
                        font = ImageFont.truetype("arial.ttf", font_size)
                    except Exception:
                        font = ImageFont.load_default()

                try:
                    words = trans_text.split()
                    lines = []
                    cur_line = []
                    
                    for word in words:
                        test_line = " ".join(cur_line + [word])
                        bbox = draw.textbbox((0, 0), test_line, font=font)
                        if (bbox[2] - bbox[0]) <= box_w or not cur_line:
                            cur_line.append(word)
                        else:
                            lines.append(" ".join(cur_line))
                            cur_line = [word]
                    if cur_line:
                        lines.append(" ".join(cur_line))

                    total_text_h = len(lines) * (font_size + 4)
                    if total_text_h > box_h and font_size > 11:
                        font_size = max(9, int(font_size * (box_h / total_text_h) * 0.90))
                        try:
                            font = ImageFont.truetype(font_path, font_size)
                        except Exception:
                            pass

                    line_h = font_size + 3
                    start_y = top + max(0, (box_h - (len(lines) * line_h)) // 2)

                    for l_idx, line_str in enumerate(lines):
                        line_bbox = draw.textbbox((0, 0), line_str, font=font)
                        line_w = line_bbox[2] - line_bbox[0]
                        line_x = left + max(0, (box_w - line_w) // 2)
                        line_y = start_y + l_idx * line_h
                        draw.text((line_x, line_y), line_str, fill=text_hex, font=font)

                except Exception as draw_err:
                    draw.text((left, top), trans_text, fill=text_hex, font=font)

            out_buf = io.BytesIO()
            pil_img.convert("RGB").save(out_buf, format="PNG")
            return out_buf.getvalue()

        except Exception as e:
            logger.error(f"ImageTranslator xatolik: {e}")
            return image_bytes