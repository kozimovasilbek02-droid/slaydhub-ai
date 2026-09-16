# -*- coding: utf-8 -*-
"""
AI Presentation Content Generator
----------------------------------
Mavzu bo'yicha ilmiy reja, tuzilma va har bir slayd uchun chuqur kontent yaratuvchi
ko'p modulli (Gemini, Claude, OpenAI, OpenRouter) sun'iy intellekt dvigateli.
"""

import os
import re
import json
import logging
from typing import Dict, List, Any, Optional

from google import genai
from google.genai import types

logger = logging.getLogger(__name__)

DEFAULT_THEMES = ["uzbek_blue", "emerald_teal", "dark_slate", "modern_purple", "crimson_ruby"]

SYSTEM_PROMPT = """Siz professional taqdimotlar va ilmiy ma'ruzalar bo'yicha dunyodagi eng tajribali AI mutaxassisisiz.
Berilgan mavzu bo'yicha aniq, boy va professional tuzilgan taqdimot kontentini JSON formatida taqdim eting.

QOIDALAR:
1. Taqdimot tili: O'zbek tili (Lotin alifbosida). Matnlar tabiiy, ilmiy va jozibador bo'lishi shart.
2. Har bir slayd quyidagi turlardan biri bo'lishi kerak:
   - "hero": Kirish slaydi (sarlavha, izoh, kategoriya burchak yorlig'i)
   - "agenda": Mundarija va reja (5 tagacha bo'limlar raqam va qisqa izohi bilan)
   - "cards": 3 ta asosiy ustun / kartochka (har bir kartochka sarlavha va 2 ta mustaqil tahliliy banddan iborat)
   - "stats": 3 ta muhim raqamli ko'rsatkich yoki statistika (masalan, 85%, 2.5x, 1-o'rin)
   - "comparative": 2 ta qiyosiy ustun (Afzalliklar / Kamchiliklar, Kecha / Bugun, Nazariya / Amaliyot)
   - "conclusion": Xulosa va yakuniy tavsiyalar
3. Slaydlar ketma-ketligi:
   - 1-slayd doimo "hero"
   - 2-slayd doimo "agenda"
   - Oxirgi slayd doimo "conclusion"
   - Qolgan oraliq slaydlar mavzuning chuqur mohiyatini ochib beruvchi "cards", "stats", "comparative" turlaridan iborat bo'lsin.

JSON TUZILMASI NAMUNASI:
{
  "theme": "uzbek_blue",
  "topic": "Mavzu nomi",
  "slides": [
    {
      "slide_number": 1,
      "type": "hero",
      "badge": "KIRISH VA KONSEPSIYA",
      "title": "Taqdimotning Katta va Qiziqarli Sarlavhasi",
      "subtitle": "Mavzu bo'yicha 2-3 jumlali chuqur va tushunarli kirish izohi.",
      "category": "O'zbekistonning Yangi Tarixi"
    },
    {
      "slide_number": 2,
      "type": "agenda",
      "badge": "REJA VA MUNDARIJA",
      "title": "Taqdimot Rejasi va Asosiy Yo'nalishlar",
      "items": [
        {"num": "01", "title": "Birinchi asosiy bo'lim", "desc": "Bo'lim bo'yicha qisqacha mazmun"},
        {"num": "02", "title": "Ikkinchi asosiy bo'lim", "desc": "Bo'lim bo'yicha qisqacha mazmun"},
        {"num": "03", "title": "Uchinchi asosiy bo'lim", "desc": "Bo'lim bo'yicha qisqacha mazmun"}
      ]
    },
    {
      "slide_number": 3,
      "type": "cards",
      "badge": "1-BO'LIM TAHLILI",
      "title": "Mavzuning Asosiy Tushunchalari va Omili",
      "cards": [
        {"title": "1-Asosiy omil", "p1": "Omil bo'yicha birinchi tahliliy fikr va dalil.", "p2": "Ikkinchi amaliy jihat va xulosa."},
        {"title": "2-Asosiy omil", "p1": "Omil bo'yicha birinchi tahliliy fikr va dalil.", "p2": "Ikkinchi amaliy jihat va xulosa."},
        {"title": "3-Asosiy omil", "p1": "Omil bo'yicha birinchi tahliliy fikr va dalil.", "p2": "Ikkinchi amaliy jihat va xulosa."}
      ]
    }
  ]
}
Faqat sof JSON qaytaring, ortiqcha so'z va tushuntirish yozmang."""


class AIPresentationGenerator:
    """Mavzu bo'yicha to'liq taqdimot kontentini generatsiya qiluvchi universal AI xizmati."""

    def __init__(self, api_key: Optional[str] = None):
        self.gemini_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY", "")
        self.anthropic_key = os.environ.get("ANTHROPIC_API_KEY", "")
        self.openai_key = os.environ.get("OPENAI_API_KEY", "")
        self.openrouter_key = os.environ.get("OPENROUTER_API_KEY", "")

        self.gemini_client = genai.Client(api_key=self.gemini_key) if self.gemini_key else None

    async def generate_deck_spec(
        self,
        topic: str,
        slide_count: int = 10,
        theme: str = "uzbek_blue",
        provider: str = "auto"
    ) -> Dict[str, Any]:
        """
        Berilgan mavzu va slaydlar soni bo'yicha to'liq JSON spec yaratadi.
        """
        prompt = f"""Mavzu: "{topic}"
Slaydlar soni: Aniq {slide_count} ta slayd.
Tanlangan rang temasi: {theme}

Iltimos, ushbu mavzu yuzasidan aynan {slide_count} ta slayddan iborat mukammal, boy ma'lumotli va qiziqarli taqdimot kontentini JSON ko'rinishida generatsiya qiling."""

        # 1. Agar provider Claude bo'lsa
        if provider == "claude" and (self.anthropic_key or self.openrouter_key):
            try:
                res = self._call_claude(prompt)
                parsed = self._clean_and_parse_json(res)
                if parsed and "slides" in parsed:
                    return parsed
            except Exception as e:
                logger.warning(f"Claude chaqiruvida xatolik, Geminiga o'tiladi: {e}")

        # 2. Agar provider OpenAI bo'lsa
        if provider == "openai" and (self.openai_key or self.openrouter_key):
            try:
                res = self._call_openai(prompt)
                parsed = self._clean_and_parse_json(res)
                if parsed and "slides" in parsed:
                    return parsed
            except Exception as e:
                logger.warning(f"OpenAI chaqiruvida xatolik, Geminiga o'tiladi: {e}")

        # 3. Asosiy Gemini chaqiruvi (Gemini 3.1 Flash-Lite / 3.6 Flash)
        if self.gemini_client:
            models_to_try = ["gemini-3.1-flash-lite", "gemini-3.6-flash", "gemini-3.5-flash", "gemini-3.1-flash-lite-preview"]
            for m in models_to_try:
                try:
                    response = self.gemini_client.models.generate_content(
                        model=m,
                        contents=prompt,
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_PROMPT,
                            temperature=0.3,
                            response_mime_type="application/json"
                        )
                    )
                    parsed = self._clean_and_parse_json(response.text)
                    if parsed and "slides" in parsed and len(parsed["slides"]) > 0:
                        parsed["theme"] = theme
                        parsed["topic"] = topic
                        return parsed
                except Exception as ge:
                    logger.warning(f"Gemini ({m}) bilan generatsiyada xatolik: {ge}")

        # 4. Fallback: Agar barcha AI API lar vaqtincha mavjud bo'lmasa, strukturaviy shablon tuzamiz
        return self._generate_fallback_spec(topic, slide_count, theme)

    def _call_claude(self, prompt: str) -> str:
        import anthropic
        client = anthropic.Anthropic(api_key=self.anthropic_key or self.openrouter_key)
        resp = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )
        return resp.content[0].text

    def _call_openai(self, prompt: str) -> str:
        from openai import OpenAI
        client = OpenAI(api_key=self.openai_key or self.openrouter_key)
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"}
        )
        return resp.choices[0].message.content

    def _clean_and_parse_json(self, raw: str) -> Optional[Dict[str, Any]]:
        if not raw:
            return None
        text = raw.strip()
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
        try:
            return json.loads(text)
        except Exception:
            m = re.search(r'\{\s*"theme".*\}', text, re.DOTALL) or re.search(r'\{.*"slides".*\}', text, re.DOTALL)
            if m:
                try:
                    return json.loads(m.group(0))
                except Exception:
                    pass
        return None

    def _generate_fallback_spec(self, topic: str, slide_count: int, theme: str) -> Dict[str, Any]:
        """Internet yoki API bo'lmagan taqdirda professional strukturali fallback reja."""
        slides = [
            {
                "slide_number": 1,
                "type": "hero",
                "badge": "ILMIY-AMALIY TAQDIMOT",
                "title": topic,
                "subtitle": f"{topic} mavzusi bo'yicha asosiy tushunchalar, tahliliy ko'rsatkichlar va istiqbolli rejalar bayoni.",
                "category": "Zamonaviy Fan va Amaliyot"
            },
            {
                "slide_number": 2,
                "type": "agenda",
                "badge": "MUNDARIJA",
                "title": "Taqdimotning Asosiy Rejasi",
                "items": [
                    {"num": "01", "title": "Kirish va Asosiy Tushunchalar", "desc": "Mavzuning dolzarbligi va umumiy ta'rifi"},
                    {"num": "02", "title": "Asosiy Omillar va Tahlil", "desc": "Sohadagi mavjud holat va chuqur tahliliy ko'rsatkichlar"},
                    {"num": "03", "title": "Amaliyot va Qo'llanilish Sohalari", "desc": "Zamonaviy yondashuvlar va amaliy ahamiyat"},
                    {"num": "04", "title": "Xulosa va Tavsiyalar", "desc": "Kelgusi rivojlanish istiqbollari va xulosalar"}
                ]
            }
        ]

        # Oraliq slaydlar
        for i in range(3, slide_count):
            if i % 3 == 0:
                slides.append({
                    "slide_number": i,
                    "type": "stats",
                    "badge": f"{i-2}-BO'LIM: KO'RSATKICHLAR",
                    "title": f"{topic}: Asosiy Statistik Ko'rsatkichlar",
                    "metrics": [
                        {"value": "85%+", "label": "Samaradorlik Darajasi", "desc": "Sohada erishilgan ijobiy o'sish dinamikasi"},
                        {"value": "2.5x", "label": "Rivojlanish Sur'ati", "desc": "So'nggi yillardagi o'sish ko'rsatkichi"},
                        {"value": "TOP-10", "label": "Ustuvor Yo'nalish", "desc": "Strategik ahamiyatga ega tarmoqlar qatorida"}
                    ]
                })
            elif i % 2 == 0:
                slides.append({
                    "slide_number": i,
                    "type": "comparative",
                    "badge": f"{i-2}-BO'LIM: QIYOSIY TAHLIL",
                    "title": "Mavzu Yuzasidan Qiyosiy Tahlil va Afzalliklar",
                    "col1_title": "An'anaviy Yondashuv",
                    "col1_points": [
                        "Cheklangan resurslar va yuqori vaqt sarfi",
                        "Statik ma'lumotlar bilan ishlash cheklovlari"
                    ],
                    "col2_title": "Zamonaviy Innovatsion Yondashuv",
                    "col2_points": [
                        "Raqamli avtomatlashtirish va yuqori unumdorlik",
                        "Haqiqiy vaqt rejimida aniq natijalar olish imkoniyati"
                    ]
                })
            else:
                slides.append({
                    "slide_number": i,
                    "type": "cards",
                    "badge": f"{i-2}-BO'LIM: STRATEGIK OMILLAR",
                    "title": f"{topic}: Asosiy Aspektlar va Yo'nalishlar",
                    "cards": [
                        {
                            "title": "Nazariy Asoslar",
                            "p1": "Mavzuning fundamental ilmiy va uslubiy qoidalari.",
                            "p2": "Tushunchalarning tizimli tavsifi va qonuniyatlari."
                        },
                        {
                            "title": "Amaliy Ahamiyat",
                            "p1": "Zamonaviy tizimlarga integratsiya qilish mexanizmlari.",
                            "p2": "Real sharoitda samaradorlikni oshirish usullari."
                        },
                        {
                            "title": "Kutilayotgan Natija",
                            "p1": "Sohani takomillashtirish bo'yicha yangi bosqich.",
                            "p2": "Uzoq muddatli barqaror rivojlanish kafolati."
                        }
                    ]
                })

        # Oxirgi slayd
        slides.append({
            "slide_number": slide_count,
            "type": "conclusion",
            "badge": "YAKUNIY XULOSA",
            "title": f"{topic}: Yakuniy Xulosalar va Istiqbollar",
            "points": [
                "Taqdim etilgan mavzu bo'yicha barcha nazariy va amaliy jihatlar to'liq tahlil qilindi.",
                "Innovatsion va zamonaviy usullarni qo'llash orqali yuqori natijalarga erishish imkoniyati mavjud.",
                "Kelgusida sohani yanada rivojlantirish uchun tizimli chora-tadbirlar davom ettiriladi."
            ],
            "thank_you": "E'tiboringiz uchun rahmat!"
        })

        return {
            "theme": theme,
            "topic": topic,
            "slides": slides[:slide_count]
        }
