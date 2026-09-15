# -*- coding: utf-8 -*-
import os
import sys
import json
import re
import asyncio
import logging
from typing import Dict, Any, List, Optional
from google import genai
from google.genai import types

from bot.config import GEMINI_API_KEY

logger = logging.getLogger("NotebookLMService")

try:
    from notebooklm_mcp.config import ServerConfig, AuthConfig
    from notebooklm_mcp.client import NotebookLMClient
    HAS_NOTEBOOKLM_MCP = True
except Exception:
    HAS_NOTEBOOKLM_MCP = False


class NotebookLMService:
    """
    NotebookLM avtomatizatsiyasi va Gemini API asosidagi
    gibrid ilmiy hamda akademik slayd kontentini shakllantiruvchi xizmat.
    """
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or GEMINI_API_KEY
        self.client = genai.Client(api_key=self.api_key) if self.api_key else None

    async def query_notebooklm_automation(self, topic: str, timeout: int = 20) -> Optional[str]:
        """
        Tizimda mavjud bo'lgan notebooklm-mcp orqali brauzer sessiyasidan
        NotebookLM'dan mavzuga oid ilmiy tahlil va faktlarni olishga harakat qiladi.
        Agar brauzer ochiq bo'lmasa yoki sessiya mavjud bo'lmasa, None qaytaradi.
        """
        if not HAS_NOTEBOOKLM_MCP:
            return None

        # Foydalanuvchining persistent Chrome profili (avvaldan Google'ga login qilingan bo'lsa)
        profile_dir = os.environ.get("NOTEBOOKLM_PROFILE_DIR", r"c:\Users\user\chrome_profile")
        if not os.path.exists(profile_dir):
            profile_dir = r"c:\Users\user\.notebooklm_chrome_profile"

        cfg = ServerConfig(
            headless=True,
            timeout=timeout,
            auth=AuthConfig(
                profile_dir=profile_dir,
                use_persistent_session=True,
                auto_login=False
            )
        )
        client = None
        try:
            client = NotebookLMClient(cfg)
            # 1. Brauzerni ishga tushirish (maksimal 8 soniya kutish)
            await asyncio.wait_for(client.start(), timeout=8)
            # 2. Avtorizatsiya tekshirish (Google sessiyasi bormi?)
            is_auth = await asyncio.wait_for(client.authenticate(), timeout=8)
            if is_auth:
                logger.info(f"✅ NotebookLM ga to'g'ridan-to'g'ri ulandi! Mavzu: {topic}")
                query = f"Quyidagi mavzu bo'yicha ilmiy tushunchalar, muhim faktlar va mezonlarni xulosalab ber: {topic}"
                await client.send_message(query)
                resp = await client.get_response(wait_for_completion=True, max_wait=timeout)
                if resp and len(resp.strip()) > 30:
                    logger.info(f"✅ NotebookLM dan {len(resp)} belgili ilmiy faktlar qabul qilindi.")
                    return resp.strip()
            else:
                logger.info("ℹ️ NotebookLM sessiyasi tizimda faol emas, zaxira Gemini AI rejimiga o'tiladi.")
                return None
        except Exception as e:
            logger.info(f"ℹ️ NotebookLM avtomatizatsiyasi ishlamadi ({e}), zaxira Gemini AI rejimiga o'tiladi.")
            return None
        finally:
            if client and hasattr(client, 'close'):
                try:
                    await client.close()
                except Exception:
                    pass

    async def generate_slide_content(self, topic: str, slide_count: int = 10) -> Dict[str, Any]:
        """
        Mavzu bo'yicha belgilangan sondagi slaydlar uchun chuqur ilmiy kontent tuzadi.
        Gibrid arxitektura: Avval NotebookLM'dan faktlar olinadi, so'ngra Gemini bilan boyitiladi.
        """
        if not self.client:
            raise ValueError("Gemini API kaliti topilmadi!")

        # 1-qadam: NotebookLM dan faktlarni olishga urinish (agar sessiya mavjud bo'lsa)
        nlm_notes = None
        try:
            nlm_notes = await self.query_notebooklm_automation(topic, timeout=15)
        except Exception:
            nlm_notes = None

        nlm_context = ""
        if nlm_notes:
            nlm_context = f"\n\nDIQQAT: Quyidagi tahliliy materiallar foydalanuvchining NotebookLM manbalaridan olindi. Slaydlarni tuzishda aynan shu ilmiy faktlar va tushunchalarga tayaning:\n{nlm_notes}\n"

        prompt = f"""Siz Google NotebookLM kabi ishlaydigan yuqori intellektual akademik va ilmiy assistentsiz.
Mavzu: "{topic}"
Slaydlar umumiy soni: {slide_count} ta slayd.{nlm_context}

Vazifangiz: Mazkur mavzu bo'yicha O'zbekiston OTM talabalari va mutaxassislari uchun xalqaro darajadagi, chuqur tahliliy taqdimot kontentini shakllantiring.

Taqdimot struktura va maket (layout) qoidalari:
1. 1-slayd: layout_type = 'cover' (Title: To'liq akademik mavzu nomi, Subtitle: Ilmiy va amaliy tahliliy qo'llanma / 2026-yil)
2. 2-slayd: layout_type = 'agenda' (Mundarija: 4 ta asosiy reja bandi)
3. Mazmunga qarab quyidagi turlardan mosini tanlang (slaydlar bir xil bo'lib qolmasligi uchun xilma-xil maketlardan foydalaning):
   - Agar anatomik tuzilma, neyron, organ, mexanizm yoki tizim tahlil qilinsa:
     layout_type = 'diagram'
     "points": [
       "Komponent / Bo'lim: Strukturaviy xususiyati va lokalizatsiyasi",
       "Funksional mexanizm: Bajaradigan aniq fiziologik yoki amaliy roli",
       "Integratsiya: Boshqa tizimlar bilan uzviy bog'liqligi va natijasi"
     ]
   - Agar ikki xil tushuncha, ta'limot yoki holat taqqoslansa (masalan: Jonli vs Jonsiz, Berdyaev vs G'azzoliy, Simpatik vs Parasimpatik):
     layout_type = 'comparison'
     "comparison_data": {{
       "left_title": "Aniq chap tomon sarlavhasi",
       "left_points": ["Aspekt 1: Aniq tahlil", "Aspekt 2: Xususiyat va dalil", "Aspekt 3: Natija va xulosa"],
       "right_title": "Aniq o'ng tomon sarlavhasi",
       "right_points": ["Aspekt 1: Aniq tahlil", "Aspekt 2: Xususiyat va dalil", "Aspekt 3: Natija va xulosa"]
     }}
   - Agar ketma-ket bosqichlar, algoritmlar yoki jarayonlar bo'lsa (masalan: Bosqichlar, tartib, algoritm):
     layout_type = 'steps'
     "steps_data": [
       {{"step": "01", "title": "1-Bosqich nomi", "desc": "Bajariladigan aniq vazifa, mexanizm va mezon."}},
       {{"step": "02", "title": "2-Bosqich nomi", "desc": "Bajariladigan aniq vazifa, mexanizm va mezon."}},
       {{"step": "03", "title": "3-Bosqich nomi", "desc": "Bajariladigan aniq vazifa, mexanizm va mezon."}}
     ]
   - Agar miqdoriy ko'rsatkichlar, metrikalar, samaradorlik foizlari yoki indikatorlar tahlil qilinsa:
     layout_type = 'metrics'
     "metrics_data": [
       {{"label": "1-Ko'rsatkich nomi", "value": "88%", "desc": "Aniq ilmiy yoki statistik ko'rsatkich izohi"}},
       {{"label": "2-Ko'rsatkich nomi", "value": "94%", "desc": "Aniq ilmiy yoki statistik ko'rsatkich izohi"}},
       {{"label": "3-Ko'rsatkich nomi", "value": "82%", "desc": "Aniq ilmiy yoki statistik ko'rsatkich izohi"}}
     ]
   - Agar 3 ta asosiy element, toifa, ustun yoki omil tahlil qilinsa (masalan: 3 ta gormon, 3 ta tamoyil, 3 ta metod):
     layout_type = 'grid'
     "points": [
       "Kategoriya 1: Chuqur ilmiy tavsifi va amaliy roli",
       "Kategoriya 2: Chuqur ilmiy tavsifi va amaliy roli",
       "Kategoriya 3: Chuqur ilmiy tavsifi va amaliy roli"
     ]
   - Agar savol-javob yoki dolzarb muammolar bo'lsa:
     layout_type = 'qa'
     "qa_data": [
       {{"q": "Aniq ilmiy yoki amaliy savol?", "a": "To'liq asosli ilmiy javob va sharh."}},
       {{"q": "Ikkinchi dolzarb savol?", "a": "To'liq asosli ilmiy javob va sharh."}}
     ]
   - Standart tahliliy ro'yxat va xulosalar uchun:
     layout_type = 'cards'
     "points": [
       "Atama / Tushuncha: Chuqur ilmiy tavsifi va xususiyati",
       "Asosiy omil: Amaliy ko'rsatkichi va ahamiyati",
       "Natija / Tavsiya: Haqiqiy keys va daliliy misol"
     ]
4. Har bir punktda 'Sarlavha: Aniq fakt yoki tavsif' bo'lsin. Quruq, suvli umumiy gaplar bo'lmasin!
5. {slide_count}-slayd: layout_type = 'cards' (Xulosalar va amaliy tavsiyalar).

Javobni FAQAT toza JSON formatida quyidagi strukturada qaytaring:
{{
  "topic": "{topic}",
  "title": "Taqdimotning to'liq jozibali nomi",
  "subtitle": "Akademik va amaliy tahliliy qo'llanma / 2026-yil",
  "slides": [
    {{
      "slide_number": 1,
      "layout_type": "cover",
      "title": "Asosiy sarlavha",
      "subtitle": "Kichik sarlavha"
    }},
    {{
      "slide_number": 2,
      "layout_type": "agenda",
      "title": "Mundarija va Asosiy Rejalar",
      "points": [
        "1. Mavzuning nazariy asoslari",
        "2. Tahliliy modellar va qiyos",
        "3. Amaliy tajriba va xavfsizlik",
        "4. Xulosa va istiqbollar"
      ]
    }}
  ]
}}
"""
        models_to_try = ["gemini-3.6-flash", "gemini-3.1-pro-preview", "gemini-3.1-flash-lite"]
        last_error = None

        for model in models_to_try:
            try:
                response = self.client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=0.3,
                        response_mime_type="application/json"
                    )
                )
                text = response.text.strip()
                # Tozalash
                if text.startswith("```json"):
                    text = text[7:]
                elif text.startswith("```"):
                    text = text[3:]
                if text.endswith("```"):
                    text = text[:-3]
                text = text.strip()

                try:
                    parsed_json = json.loads(text)
                except Exception:
                    # Agar matn ichida qo'shimcha ma'lumotlar yoki izohlar bo'lsa
                    start_idx = text.find("{")
                    if start_idx != -1:
                        try:
                            decoder = json.JSONDecoder()
                            parsed_json, _ = decoder.raw_decode(text[start_idx:])
                        except Exception:
                            end_idx = text.rfind("}")
                            if end_idx != -1 and end_idx > start_idx:
                                parsed_json = json.loads(text[start_idx:end_idx+1])
                            else:
                                raise
                    else:
                        raise

                return parsed_json
            except Exception as e:
                last_error = e
                continue

        raise RuntimeError(f"Gemini orqali kontent generatsiya qilishda xatolik: {last_error}")
