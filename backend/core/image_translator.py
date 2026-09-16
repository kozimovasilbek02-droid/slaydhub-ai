# -*- coding: utf-8 -*-
"""
backend.core.image_translator
-----------------------------
Gemini Vision orqali rasmli slaydlar, diagrammalar, infografikalar va
arxitektura sxemalari ichidagi ruscha/inglizcha matnlarni aniqlab,
o'zbek tiliga tarjima qilib, rasm fonini va grafikasini buzmasdan
(adaptive pixel inpainting) yuqori sifatda qayta chizuvchi modul.
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
import numpy as np

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
        self.model_candidates = [
            "gemini-3.5-flash-lite",
            "gemini-3.5-flash",
            "gemini-3.6-flash",
            "gemini-3.7-flash",
            "gemini-3.8-flash",
            "gemini-3.1-flash-lite",
            "gemini-2.5-flash-lite",
        ]
        self.model_name = self.model_candidates[0]

    @staticmethod
    def _sample_background_color(rgb_img: Image.Image, left: int, top: int, right: int, bottom: int) -> Tuple[int, int, int]:
        """
        Bounding box atrofidagi perimetr piksellarini tahlil qilib,
        haqiqiy fon rangini (median RGB) aniqlaydi.
        Oq to'rtburchak (white sticker patch) xatolarini butunlay yo'qotadi.
        """
        w, h = rgb_img.size
        pad = 6
        samples: List[Tuple[int, int, int]] = []
        
        # Yuqori chiziq
        y_top = max(0, top - pad)
        for x in range(max(0, left - pad), min(w, right + pad)):
            samples.append(rgb_img.getpixel((x, y_top))[:3])
            
        # Pastki chiziq
        y_bot = min(h - 1, bottom + pad)
        for x in range(max(0, left - pad), min(w, right + pad)):
            samples.append(rgb_img.getpixel((x, y_bot))[:3])
            
        # Chap chiziq
        x_left = max(0, left - pad)
        for y in range(max(0, top - pad), min(h, bottom + pad)):
            samples.append(rgb_img.getpixel((x_left, y))[:3])
            
        # O'ng chiziq
        x_right = min(w - 1, right + pad)
        for y in range(max(0, top - pad), min(h, bottom + pad)):
            samples.append(rgb_img.getpixel((x_right, y))[:3])
            
        if not samples:
            return (220, 215, 210)
            
        r_vals = [s[0] for s in samples]
        g_vals = [s[1] for s in samples]
        b_vals = [s[2] for s in samples]
        
        return (int(np.median(r_vals)), int(np.median(g_vals)), int(np.median(b_vals)))

    @staticmethod
    def _get_contrasting_text_color(bg_rgb: Tuple[int, int, int]) -> Tuple[int, int, int]:
        """
        Fon rangining yorqinligiga (luminance) qarab to'q kulrang/qora
        yoki tiniq oq matn rangini avtomatik tanlaydi.
        """
        lum = 0.299 * bg_rgb[0] + 0.587 * bg_rgb[1] + 0.114 * bg_rgb[2]
        return (22, 22, 22) if lum > 130 else (255, 255, 255)

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
                "Analyze this presentation image carefully and extract ALL text blocks.\n\n"
                "RULES:\n"
                "1. If the image is a GEOGRAPHICAL MAP with dozens of tiny village/town labels, "
                "do NOT extract individual town/village names (maps must remain authentic). "
                "Only extract main slide title or large diagram legends if present.\n"
                "2. If the image has NO text (pure photo, artwork without text), return EXACTLY:\n"
                '{"has_text": false}\n\n'
                "3. For diagrams, flowcharts, 3D architectural callouts, infocards, and labeled illustrations:\n"
                "Extract EACH distinct text unit or callout.\n"
                "CRITICAL BOUNDING BOX & GROUPING RULES:\n"
                "- If a callout consists of a header/label (e.g. 'Центральная площадь', 'Жилище', 'Внутренний двор', 'Ров', 'Стена', 'Печь') "
                "and a descriptive subtext/paragraph underneath, group the header and description into ONE unified bounding box [ymin, xmin, ymax, xmax] "
                "covering from the very top of the header word down to the very bottom of the last line of description.\n"
                "- The bounding box (0-1000 scale) MUST GENEROUSLY enclose all letters, words, and lines of that text unit.\n"
                "- Provide natural, grammatically correct Uzbek translation in Latin script (e.g. 'Turar joy\\nTurar joy maydoni 110 dan 180 m² gacha...').\n\n"
                "Return JSON in this format:\n"
                "{\n"
                '  "has_text": true,\n'
                '  "items": [\n'
                '    {\n'
                '      "box_2d": [ymin, xmin, ymax, xmax],\n'
                '      "original_text": "...",\n'
                '      "translated_text": "...",\n'
                '      "is_header": false\n'
                '    }\n'
                '  ]\n'
                "}"
            )

            # Convert image to JPEG for Gemini API
            img_buf = io.BytesIO()
            pil_img.convert("RGB").save(img_buf, format="JPEG", quality=92)
            img_jpeg_bytes = img_buf.getvalue()

            response = None
            for model_name in self.model_candidates:
                try:
                    response = self.client.models.generate_content(
                        model=model_name,
                        contents=[
                            types.Part.from_bytes(data=img_jpeg_bytes, mime_type="image/jpeg"),
                            prompt
                        ],
                        config=types.GenerateContentConfig(
                            temperature=0.1,
                            response_mime_type="application/json"
                        )
                    )
                    if response and response.text:
                        break
                except Exception as model_err:
                    logger.warning(f"ImageTranslator model {model_name} failed: {model_err}")
                    continue

            if not response or not response.text:
                return image_bytes

            resp_text = response.text.strip()
            if resp_text.startswith("```"):
                resp_text = re.sub(r"^```[a-zA-Z]*\n", "", resp_text)
                resp_text = re.sub(r"\n```$", "", resp_text).strip()

            data = json.loads(resp_text)
            if not data.get("has_text") or not data.get("items"):
                return image_bytes

            items = data.get("items", [])
            
            # Xaritalar himoyasi: agar 14 tadan ko'p juda qisqa (1-so'zli) punktlar bo'lsa, xarita rasmini buzmaslik
            if len(items) > 14 and any(len(it.get("original_text", "").split()) <= 2 for it in items):
                long_items = [it for it in items if len(it.get("original_text", "").split()) > 2]
                if len(long_items) < len(items) * 0.35:
                    items = [it for it in items if (it.get("box_2d", [0, 0, 0, 0])[2] - it.get("box_2d", [0, 0, 0, 0])[0]) > 40]

            if not items:
                return image_bytes

            logger.info(f"ImageTranslator: {len(items)} ta matnli blok topildi va tarjima qilinmoqda...")

            draw = ImageDraw.Draw(pil_img)
            
            # Local font discovery
            fonts_dir = Path(__file__).resolve().parent.parent / "fonts"
            roboto_path = str(fonts_dir / "Roboto.ttf")
            font_path = roboto_path if os.path.exists(roboto_path) else "arial.ttf"

            # PASS 1: Erase with adaptive background color and smart multi-line padding
            for item in items:
                box = item.get("box_2d")
                if not box or len(box) != 4:
                    continue

                ymin, xmin, ymax, xmax = box
                orig_text = item.get("original_text", "")
                orig_lines = [l.strip() for l in orig_text.split("\n") if l.strip()]
                num_orig_lines = max(1, len(orig_lines))

                left = int((xmin / 1000.0) * w)
                top = int((ymin / 1000.0) * h)
                right = int((xmax / 1000.0) * w)
                bottom = int((ymax / 1000.0) * h)

                raw_w = max(10, right - left)
                raw_h = max(10, bottom - top)

                # Ko'p qatorli matnlarning pastki qatorlari qirqilib qolmasligi uchun kengaytirish
                if num_orig_lines >= 3 and raw_h < num_orig_lines * 13:
                    extra_h = (num_orig_lines * 14) - raw_h
                    bottom = min(h, bottom + extra_h)

                pad_x = max(8, int(raw_w * 0.06))
                pad_y = max(5, int(raw_h * 0.07))

                erase_left = max(0, left - pad_x)
                erase_top = max(0, top - pad_y)
                erase_right = min(w, right + pad_x)
                erase_bottom = min(h, bottom + pad_y)

                bg_rgb = self._sample_background_color(pil_img, erase_left, erase_top, erase_right, erase_bottom)
                draw.rectangle([erase_left, erase_top, erase_right, erase_bottom], fill=bg_rgb)
                
                item["_erase_box"] = [erase_left, erase_top, erase_right, erase_bottom]
                item["_bg_rgb"] = bg_rgb

            # Sarlavha va uning ostidagi tavsif qutilari to'qnashuvini oldini olish
            for i, item_a in enumerate(items):
                if "_erase_box" not in item_a:
                    continue
                box_a = item_a["_erase_box"]
                for j, item_b in enumerate(items):
                    if i >= j or "_erase_box" not in item_b:
                        continue
                    box_b = item_b["_erase_box"]
                    if max(box_a[0], box_b[0]) < min(box_a[2], box_b[2]):
                        if box_a[1] < box_b[1] and box_a[3] >= box_b[1]:
                            item_b["_erase_box"][1] = box_a[3] + 2

            # PASS 2: Render translated text cleanly onto the inpainted regions
            for item in items:
                if "_erase_box" not in item:
                    continue

                erase_left, erase_top, erase_right, erase_bottom = item["_erase_box"]
                bg_rgb = item["_bg_rgb"]
                text_rgb = self._get_contrasting_text_color(bg_rgb)

                trans_text = item.get("translated_text", "").strip()
                if not trans_text:
                    continue

                trans_text = ensure_script(trans_text, target_script)

                box_w = max(10, erase_right - erase_left)
                box_h = max(10, erase_bottom - erase_top)
                is_header = item.get("is_header", False) or len(trans_text.split()) <= 2

                # Adaptiv shrift o'lchami
                if is_header:
                    font_size = max(10, min(16, int(box_h * 0.65)))
                else:
                    word_count = len(trans_text.split())
                    if word_count > 15:
                        font_size = max(8, min(11, int(box_h * 0.12)))
                    elif word_count > 6:
                        font_size = max(8, min(12, int(box_h * 0.22)))
                    else:
                        font_size = max(9, min(13, int(box_h * 0.45)))

                try:
                    font = ImageFont.truetype(font_path, font_size)
                except Exception:
                    font = ImageFont.load_default()

                # Matnni qatorlarga ajratish va qutiga moslash
                words = trans_text.split()
                lines = []
                cur_line = []
                for word in words:
                    test_line = " ".join(cur_line + [word])
                    bbox = draw.textbbox((0, 0), test_line, font=font)
                    if (bbox[2] - bbox[0]) <= (box_w - 4) or not cur_line:
                        cur_line.append(word)
                    else:
                        lines.append(" ".join(cur_line))
                        cur_line = [word]
                if cur_line:
                    lines.append(" ".join(cur_line))

                line_h = font_size + 2
                total_h = len(lines) * line_h
                if total_h > box_h and font_size > 7:
                    scale = box_h / total_h
                    font_size = max(7, int(font_size * scale * 0.95))
                    try:
                        font = ImageFont.truetype(font_path, font_size)
                    except Exception:
                        pass
                    line_h = font_size + 2
                    lines = []
                    cur_line = []
                    for word in words:
                        test_line = " ".join(cur_line + [word])
                        bbox = draw.textbbox((0, 0), test_line, font=font)
                        if (bbox[2] - bbox[0]) <= (box_w - 4) or not cur_line:
                            cur_line.append(word)
                        else:
                            lines.append(" ".join(cur_line))
                            cur_line = [word]
                    if cur_line:
                        lines.append(" ".join(cur_line))

                start_y = erase_top + max(0, (box_h - (len(lines) * line_h)) // 2)

                for l_idx, line_str in enumerate(lines):
                    line_bbox = draw.textbbox((0, 0), line_str, font=font)
                    line_w = line_bbox[2] - line_bbox[0]
                    if is_header or len(lines) == 1:
                        line_x = erase_left + max(0, (box_w - line_w) // 2)
                    else:
                        line_x = erase_left + 2
                    line_y = start_y + l_idx * line_h
                    draw.text((line_x, line_y), line_str, fill=text_rgb, font=font)

            out_buf = io.BytesIO()
            pil_img.convert("RGB").save(out_buf, format="PNG")
            return out_buf.getvalue()

        except Exception as e:
            logger.error(f"ImageTranslator xatolik: {e}")
            return image_bytes