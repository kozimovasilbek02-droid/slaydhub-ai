# -*- coding: utf-8 -*-
"""
backend.core.gemini_translator
------------------------------
Google Gemini orqali slayd matnlarini yuqori aniqlikda, sohaviy kontekst,
anti-overflow va parallel batch tarzda tarjima qiluvchi modul.
"""
from __future__ import annotations

import os
import json
import re
import time
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Any, Optional

from google import genai
from google.genai import types

from backend.core.transliteration import ensure_script, latin_to_cyrillic

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "gemini-3.1-flash-lite"
FALLBACK_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-3.1-flash-lite-preview",
    "gemini-3-flash-preview",
    "gemini-3.5-flash-lite",
    "gemini-3.6-flash",
]

def sanitize_text(text: str) -> str:
    """Boshqaruv belgilari va OpenXML artefaktlarini tozalaydi, probellarni saqlaydi."""
    if not text:
        return ""
    # _x000B_ va boshqaruv belgilarini probel bilan almashtiramiz (so'zlar yopishib qolmasligi uchun)
    t = re.sub(r'_x[0-9a-fA-F]{4}_', ' ', text)
    t = re.sub(r'[\u0001-\u0008\u000b\u000c\u000e-\u001f\u007f]', ' ', t)
    t = re.sub(r'[ \t]+', ' ', t)
    return t.strip()


DOMAIN_KEYWORDS: Dict[str, List[str]] = {
    "Tibbiyot, anatomiya, farmatsevtika va sog'liqni saqlash": [
        "tibbiyot", "meditsina", "vrach", "kasallik", "shifoxona", "infeksiya", "sanpin",
        "immunitet", "bemor", "terapiya", "jarrohlik", "doctor", "medicine", "hospital", "disease", "anatomy"
    ],
    "Tarix, arxeologiya va ijtimoiy fanlar": [
        "tarix", "history", "urush", "sulola", "asr", "davr", "neolit", "paleolit",
        "madaniyat", "arxeologiya", "podsholik", "imperiya", "inqilob", "revolution"
    ],
    "Iqtisodiyot, moliya, bank va investitsiyalar": [
        "iqtisod", "moliya", "bank", "investitsiya", "kredit", "soliq", "valyuta",
        "budget", "finance", "economy", "investment", "audit", "buxgalteriya", "accounting"
    ],
    "Axborot texnologiyalari, IT, AI va dasturlash": [
        "axborot", "texnologiya", "dasturlash", "sun'iy intellekt", "ai", "kompyuter",
        "it", "algoritm", "server", "software", "hardware", "data", "tarmoq", "kiberxavfsizlik"
    ],
    "Falsafa, dinshunoslik, axloq va ma'naviyat": [
        "falsafa", "din", "islom", "buddizm", "zardushtiylik", "mutafakkir",
        "falsafiy", "ahloq", "ma'naviyat", "hikmat", "philosophy", "religion", "ethics"
    ],
    "Biologiya, ekologiya, tabiat va geografiya": [
        "biologiya", "tabiat", "ekologiya", "hayvonot", "o'simlik", "geografiya",
        "iqlim", "yer", "sayyora", "nature", "biology", "geography", "environment"
    ],
    "Biznes, marketing, menejment va startaplar": [
        "biznes", "marketing", "menejment", "startap", "reklama", "strategiya",
        "swot", "porter", "hr", "kadrlar", "boshqaruv", "savdo", "business", "management"
    ],
    "Huquqshunoslik, qonunchilik va davlat boshqaruvi": [
        "huquq", "qonun", "sud", "konstitutsiya", "jinoyat", "fuqarolik",
        "shartnoma", "adliya", "law", "legal", "court", "justice"
    ],
    "Pedagogika, maktab ta'limi va metodika": [
        "ta'lim", "pedagogika", "maktab", "metodika", "o'qitish", "dars",
        "sinf", "education", "school", "teaching", "pedagogy"
    ]
}

def detect_domain(sample_text: str) -> str:
    """Taqdimot matni yoki sarlavhasidan tegishli ilmiy/biznes sohani avtomatik aniqlaydi."""
    if not sample_text:
        return "general"
    low = sample_text.lower()
    for domain_name, kw_list in DOMAIN_KEYWORDS.items():
        if any(kw in low for kw in kw_list):
            return domain_name
    return "Umumiy ta'limiy va ilmiy taqdimot"


class TranslationError(RuntimeError):
    """Tarjima xizmati xatolik qaytarganda."""


class GeminiTranslator:
    """
    Taqdimot matnlarini kontekst, terminologiya va slayd tuzilmasini saqlagan holda
    tarjima qiluvchi yuqori aniqlikdagi AI tarjimon dvigateli.
    """

    def __init__(self, api_key: Optional[str] = None, model_name: Optional[str] = None):
        self.api_key = (
            api_key
            or os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY", "")
        )
        self.client = genai.Client(api_key=self.api_key) if self.api_key else genai.Client()
        self.model_candidates = [model_name] if model_name else list(FALLBACK_MODELS)
        self.model_name = self.model_candidates[0]

    def translate_items_batch(
        self,
        items: List[Dict[str, Any]],
        target_script: str = "latin",
        domain: str = "general",
        glossary: Optional[Dict[str, str]] = None,
        stats: Optional[Dict[str, Any]] = None,
        batch_size: int = 60,
        max_workers: int = 1,
    ) -> List[Dict[str, Any]]:
        """
        Slayd elementlari ro'yxatini parallel batch tarzda tarjima qiladi.
        items: [{"id": "s1_sh0_p0", "text": "Strategic Vision", ...}, ...]
        """
        if not items:
            if stats is not None:
                stats.update(total=0, failed=0, failed_ids=[])
            return []

        # Agar domain ko'rsatilmagan bo'lsa, slayd matnlaridan avtomatik aniqlash
        if domain in ("general", "", None):
            combined_sample = " ".join([it.get("original_text") or it.get("text", "") for it in items[:15]])
            domain = detect_domain(combined_sample)

        batches = [items[i:i + batch_size] for i in range(0, len(items), batch_size)]

        if len(batches) == 1:
            results = self._translate_single_batch(batches[0], target_script, domain, glossary)
        else:
            results = []
            workers = min(max_workers, len(batches))
            with ThreadPoolExecutor(max_workers=workers) as executor:
                future_to_batch = {
                    executor.submit(self._translate_single_batch, b, target_script, domain, glossary): b
                    for b in batches
                }
                for future in as_completed(future_to_batch):
                    b = future_to_batch[future]
                    try:
                        res = future.result()
                        results.extend(res)
                    except Exception as exc:
                        logger.error("Batch tarjima xatosi: %s", exc)
                        for it in b:
                            orig = it.get("original_text") or it.get("text", "")
                            results.append({"id": it["id"], "translated_text": sanitize_text(orig)})

        order_map = {it["id"]: it for it in results}
        ordered_results = []
        failed_ids = []

        for it in items:
            orig = sanitize_text(it.get("original_text") or it.get("text", ""))
            matched = order_map.get(it["id"])
            if matched:
                tr_text = matched.get("translated_text", "")
                if not tr_text or tr_text.strip() == orig.strip():
                    # Matn tarjima qilinmagan bo'lishi mumkin (masalan qisqa nomlar bundan mustasno)
                    pass
                ordered_results.append(matched)
            else:
                ordered_results.append({"id": it["id"], "translated_text": orig})
                failed_ids.append(it["id"])

        if stats is not None:
            stats.update(total=len(items), failed=len(failed_ids), failed_ids=failed_ids)

        return ordered_results

    def _translate_single_batch(
        self,
        items: List[Dict[str, Any]],
        target_script: str = "latin",
        domain: str = "general",
        glossary: Optional[Dict[str, str]] = None,
        max_retries: int = 5,
    ) -> List[Dict[str, Any]]:
        is_cyrillic = target_script.lower() in ["cyrillic", "kirill", "uz-cyrl", "ўзбекча"]
        script_name = "O'zbek tili (Lotin yozuvi)" if not is_cyrillic else "Ўзбек тили (Кирилл ёзуви)"

        glossary_instructions = ""
        if glossary:
            terms = "; ".join([f"'{k}' => '{v}'" for k, v in glossary.items()])
            glossary_instructions = f"\nMUHIM QOIDA - Maxsus atamalar lug'atiga qat'iy amal qiling:\n{terms}\n"

        system_instruction = f"""Siz professional xalqaro PowerPoint taqdimotlari bo'yicha ekspert AI tarjimonsiz.
Vazifangiz berilgan slayd matnlarini (ruscha, inglizcha yoki boshqa tildagi) {script_name}ga professional, ravon, ma'noli va slayd ramkalariga sig'adigan darajada IXCHAM tarjima qilishdir.

Soha / Kontekst: {domain}
{glossary_instructions}

QAT'IY QOIDALAR:
1. HAR BIR RUSCHA VA INGLIZCHA MATNNI O'ZBEK TILIGA TO'LIQ O'GIRING:
   - "Появление неравенства и знати" => "Tengsizlik va zodagonlarning paydo bo'lishi"
   - "Улучшаются орудия труда" => "Mehnat qurollari takomillashmoqda"
   - "Улучшалась обработка почвы" => "Tuproqqa ishlov berish yaxshilandi"
   - "Легче стало готовить поле к посеву" => "Maydonni ekishga tayyorlash osonlashdi"
   - "Появились излишки" => "Ortiqcha mahsulot paydo bo'ldi"
   - "Неравное распределение" => "Tengsiz taqsimot"
   - "Неравенство" => "Tengsizlik"
   - "Охота" => "Ovchilik"
   - "Земледелие" => "Dehqonchilik"
   - "Собирательство" => "Terimchilik"
   - "Скотоводство" => "Chorvachilik"
   - "Dars rejasi" / "План урока" => "Dars rejasi"
   - "Повторение пройденного" => "O'tilganlarni takrorlash"
   - "Renaissance Genius" => "Uyg'onish davri dahosi"
   - "Executive Summary" => "Rahbarlik uchun xulosa"
   - "Contents" / "Table of contents" => "Mundarija"
   - "Work Report" => "Ish hisoboti"
2. IXCHAMLIK (ANTI-OVERFLOW):
   - Slayd bloklaridan toshib ketmasligi uchun cho'zilgan jumlalardan qoching.
   - Sarlavhalarni lo'nda va ixcham saqlang.
3. SONLAR, FOIZLAR, FORMULALAR:
   - Raqamlar, yillar (masalan: 1452-1519), foizlar va maxsus belgilarni o'zgartirmang.
4. TOZALASH:
   - Boshqaruv belgilarini (_x000B_, \\v, \\r) va ortiqcha probellarni tozalang.
5. JAVOB FORMATI:
   - Kiruvchi JSON massividagi har bir element uchun 'id' va 'translated' kalitlari bilan JSON massiv qaytaring.
6. SHABLON SARLAVHALARINI TO'G'RI TARJIMA QILISH:
   - "Agenda Style" => "Kun tartibi"
   - "Our Team Style" / "Team Style" => "Bizning jamoa"
   - "Infographic Style" => "Infografika"
   - "Portfolio Style" => "Portfolio"
   - "Real Estate" => "Ko'chmas mulk"
   - "Content Here" / "Contents Title" => "Mundarija"
   - Shablonlardagi "Style", "Layout" kabi sun'iy so'zlarni sarlavhaga qo'shmang, lo'nda va tabiiy nomlang.
7. LOREM IPSUM VA SHABLON MATNLARINI TO'LIQ O'ZBEKCHALASHTIRING:
   - "Lorem ipsum dolor sit amet..." kabi har qanday soxta matnlarni hech qachon shundayligicha qoldirmang!
   - Qisqa sarlavhalar uchun: "Mavzu bo'yicha qisqacha izoh" yoki "Taqdimotning qisqacha mazmuni";
   - Uzun matnlar uchun: "Ushbu bo'limda taqdimot mavzusi yuzasidan batafsil ma'lumotlar, asosiy ko'rsatkichlar va tahliliy xulosalar keltiriladi."
8. SHABLONNING O'ZINI MAQTASH/REKLAMA MATNLARINI:
   - O'rniga mavzuga mos professional tahliliy matn bering:
     * Qisqa sarlavhalar uchun: "Bo'lim mazmuni" yoki "Asosiy ko'rsatkichlar";
     * Uzun matnlar uchun: "Ushbu bo'limda taqdimot mavzusi yuzasidan batafsil ma'lumotlar, asosiy ko'rsatkichlar va amaliy tahlillar bayon etiladi."
"""

        prompt_payload = [
            {"id": it["id"], "text": sanitize_text(it.get("original_text") or it.get("text", ""))}
            for it in items
        ]
        user_content = json.dumps(prompt_payload, ensure_ascii=False, indent=2)

        result_map: Dict[str, str] = {}

        # 1. Agar GROQ_API_KEY mavjud bo'lsa, birinchi navbatda Groq orqali tarjima qilamiz
        groq_key = os.environ.get("GROQ_API_KEY")
        if groq_key:
            result_map = self._translate_via_groq(items, system_instruction, is_cyrillic)

        # 2. Agar Groq yetishmayotgan elementlar qoldirgan bo'lsa yoki ishlamasa, Gemini orqali to'ldiramiz
        missing_items = [it for it in items if not (result_map.get(it["id"]) or result_map.get(str(it["id"])))]
        if missing_items:
            gem_res = self._translate_via_gemini(missing_items, system_instruction, is_cyrillic)
            result_map.update(gem_res)

        # 3. Agar hali ham yetishmayotgan elementlar bo'lsa, favqulodda yagona tarjima
        still_missing = [it for it in items if not (result_map.get(it["id"]) or result_map.get(str(it["id"])))]
        for it in still_missing:
            orig = sanitize_text(it.get("original_text") or it.get("text", ""))
            em_tr = self._translate_single_item_emergency(orig, is_cyrillic)
            if em_tr and em_tr != orig:
                result_map[it["id"]] = em_tr
                result_map[str(it["id"])] = em_tr

        # 4. Yakuniy natijalarni yig'ish
        results = []
        for it in items:
            item_id = it["id"]
            orig = sanitize_text(it.get("original_text") or it.get("text", ""))
            tr_found = result_map.get(item_id) or result_map.get(str(item_id)) or ""
            
            if tr_found and tr_found.strip():
                results.append({"id": item_id, "translated_text": tr_found})
            else:
                em_tr = self._translate_single_item_emergency(orig, is_cyrillic)
                results.append({"id": item_id, "translated_text": em_tr if em_tr else orig})

        return results

    def _translate_via_gemini(self, items: List[Dict[str, Any]], system_instruction: str, is_cyrillic: bool) -> Dict[str, str]:
        """Gemini orqali batch tarjima."""
        if not items:
            return {}
        prompt_payload = [
            {"id": it["id"], "text": sanitize_text(it.get("original_text") or it.get("text", ""))}
            for it in items
        ]
        user_content = json.dumps(prompt_payload, ensure_ascii=False, indent=2)
        res_map = {}
        for cur_model in self.model_candidates:
            try:
                response = self.client.models.generate_content(
                    model=cur_model,
                    contents=user_content,
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction,
                        temperature=0.1,
                        response_mime_type="application/json"
                    )
                )
                raw_text = response.text.strip()
                raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                raw_text = re.sub(r"\s*```$", "", raw_text)
                m = re.search(r'\[\s*\{.*\}\s*\]', raw_text, re.DOTALL)
                if m:
                    raw_text = m.group(0)
                try:
                    import json_repair
                    parsed = json_repair.loads(raw_text)
                except Exception:
                    parsed = json.loads(raw_text)

                if isinstance(parsed, list):
                    for row in parsed:
                        if isinstance(row, dict) and "id" in row:
                            t_val = row.get("translated") or row.get("translated_text") or row.get("text") or ""
                            clean_t = sanitize_text(t_val)
                            if is_cyrillic:
                                clean_t = ensure_script(clean_t, "cyrillic")
                            res_map[row["id"]] = clean_t
                            res_map[str(row["id"])] = clean_t
                if res_map:
                    logger.info("Gemini orqali muvaffaqiyatli tarjima qilindi (%s, %d ta element)", cur_model, len(res_map) // 2)
                    return res_map
            except Exception as e:
                logger.warning("Gemini model %s xatosi: %s", cur_model, e)
                time.sleep(1.0)
        return res_map

    def _translate_single_item_emergency(self, text: str, is_cyrillic: bool) -> str:
        """Kichik favqulodda yagona matn tarjimasi (Groq yoki Gemini orqali)."""
        if not text or not text.strip():
            return ""
        clean_in = sanitize_text(text)
        script_name = "O'zbek tili (Lotin)" if not is_cyrillic else "Ўзбек тили (Кирилл)"

        # 1. Groq orqali urinish
        groq_key = os.environ.get("GROQ_API_KEY")
        if groq_key:
            import requests
            url = "https://api.groq.com/openai/v1/chat/completions"
            headers = {"Authorization": f"Bearer {groq_key}", "Content-Type": "application/json"}
            payload = {
                "model": "qwen/qwen3.8-27b",
                "messages": [
                    {"role": "system", "content": f"Siz professional tarjimonsiz. Ushbu jumlani {script_name}ga akademik va tabiiy tarjima qiling. Faqat tarjimani qaytaring, boshqa hech narsa yozmang."},
                    {"role": "user", "content": clean_in}
                ],
                "temperature": 0.1,
                "max_tokens": 150
            }
            try:
                r = requests.post(url, headers=headers, json=payload, timeout=20)
                if r.status_code == 200:
                    res_txt = r.json()["choices"][0]["message"]["content"].strip().strip('\'"`').strip()
                    if res_txt:
                        return sanitize_text(res_txt)
            except Exception:
                pass

        # 2. Gemini orqali urinish
        for m in self.model_candidates:
            try:
                res = self.client.models.generate_content(
                    model=m,
                    contents=f"Ushbu matnni {script_name}ga professional akademik tarjima qiling. Faqat tarjima matnini o'zini qaytaring: \"{clean_in}\""
                )
                out = res.text.strip().strip('\'"`').strip()
                if out:
                    return sanitize_text(out)
            except Exception:
                continue

        return clean_in

    def _translate_via_groq(self, items: List[Dict[str, Any]], system_instruction: str, is_cyrillic: bool) -> Dict[str, str]:
        groq_key = os.environ.get("GROQ_API_KEY")
        if not groq_key or not items:
            return {}
        import requests
        try:
            import json_repair
        except ImportError:
            json_repair = None

        url = "https://api.groq.com/openai/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {groq_key}",
            "Content-Type": "application/json"
        }
        prompt_payload = [
            {"id": it["id"], "text": sanitize_text(it.get("original_text") or it.get("text", ""))}
            for it in items
        ]
        user_content = json.dumps(prompt_payload, ensure_ascii=False, indent=2)
        models = ["qwen/qwen3.8-27b"]
        for m in models:
            try:
                payload = {
                    "model": m,
                    "messages": [
                        {"role": "system", "content": system_instruction + "\nMUHIM: Faqat to'g'ridan-to'g'ri JSON massiv qaytaring: [{\"id\": \"...\", \"translated\": \"...\"}]. Boshqa so'z qo'shmang."},
                        {"role": "user", "content": user_content}
                    ],
                    "temperature": 0.1,
                    "max_tokens": 1500
                }
                r = requests.post(url, headers=headers, json=payload, timeout=25)
                if r.status_code == 429:
                    logger.info("Groq 429 rate limit (%s). Zudlik bilan Gemini ga topshirilmoqda...", m)
                    return {}
                if r.status_code == 200:
                    raw_text = r.json()["choices"][0]["message"]["content"].strip()
                    raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                    raw_text = re.sub(r"\s*```$", "", raw_text)
                    m_match = re.search(r'\[\s*\{.*\}\s*\]', raw_text, re.DOTALL)
                    if m_match:
                        raw_text = m_match.group(0)
                    
                    parsed = None
                    try:
                        parsed = json.loads(raw_text)
                    except Exception:
                        if json_repair:
                            try:
                                parsed = json_repair.repair_json(raw_text, return_objects=True)
                            except Exception:
                                pass

                    res_map = {}
                    if isinstance(parsed, list):
                        for row in parsed:
                            if isinstance(row, dict) and "id" in row:
                                t_val = row.get("translated") or row.get("translated_text") or row.get("text") or ""
                                clean_t = sanitize_text(t_val)
                                if is_cyrillic:
                                    clean_t = ensure_script(clean_t, "cyrillic")
                                res_map[row["id"]] = clean_t
                                res_map[str(row["id"])] = clean_t
                    elif isinstance(parsed, dict):
                        for list_key in ["items", "translations", "result", "data"]:
                            if list_key in parsed and isinstance(parsed[list_key], list):
                                for row in parsed[list_key]:
                                    if isinstance(row, dict) and "id" in row:
                                        t_val = row.get("translated") or row.get("translated_text") or row.get("text") or ""
                                        clean_t = sanitize_text(t_val)
                                        if is_cyrillic:
                                            clean_t = ensure_script(clean_t, "cyrillic")
                                        res_map[row["id"]] = clean_t
                                        res_map[str(row["id"])] = clean_t
                                break

                    if res_map and len(res_map) > 0:
                        logger.info("Groq orqali muvaffaqiyatli tarjima qilindi (%s, %d ta element)", m, len(res_map) // 2)
                        return res_map
            except Exception as ge:
                logger.warning("Groq modeli %s xatosi: %s", m, ge)

        return {}





    def translate_single_text(self, text: str, target_script: str = "latin") -> str:
        """Yagona satrni (masalan fayl sarlavhasini) toza o'zbek tiliga o'giradi."""
        if not text or not text.strip():
            return "Taqdimot"
        is_cyrillic = target_script.lower() in ["cyrillic", "kirill", "uz-cyrl", "ўзбекча"]
        script_name = "O'zbek tili (Lotin yozuvi)" if not is_cyrillic else "Ўзбек тили (Кирилл ёзуви)"
        clean_in = sanitize_text(text)

        prompt = f"""Fayl sarlavhasini {script_name}ga qisqa, toza va professional tarjima qiling.
QOIDALAR:
1. FAQAT tarjima qilingan nomni qaytaring.
2. Oxiriga '_Tarjima', '_Ozbekcha' yoki qavslar qo'shmang.
3. '_' (pastki chiziq) ishlatmang, so'zlar orasida probel bo'lsin.

Matn: "{clean_in}\""""

        # 1. Birinchi navbatda Groq orqali to'g'ri va professional o'zbekcha nom olamiz
        groq_key = os.environ.get("GROQ_API_KEY")
        if groq_key:
            import requests
            url = "https://api.groq.com/openai/v1/chat/completions"
            headers = {"Authorization": f"Bearer {groq_key}", "Content-Type": "application/json"}
            payload = {
                "model": "qwen/qwen3.8-27b",
                "messages": [
                    {
                        "role": "user",
                        "content": (
                            f"Fayl mavzusi / sarlavhasi: '{clean_in}'.\n"
                            f"Ushbu taqdimot sarlavhasini {script_name}da ixcham, tabiiy, savodli va tushunarli qilib tarjima qiling.\n"
                            f"QOIDALAR:\n"
                            f"1. FAQAT tarjima qilingan sarlavha nomini qaytaring, boshqa hech narsa yozmang.\n"
                            f"2. Pastki chiziq (_) yoki keraksiz qavslar qo'ymang.\n"
                            f"3. Ruscha so'zlarni qoldirmang, to'liq o'zbekcha bo'lsin."
                        )
                    }
                ],
                "temperature": 0.1,
                "max_tokens": 70
            }
            try:
                r = requests.post(url, headers=headers, json=payload, timeout=20)
                if r.status_code == 200:
                    out = r.json()["choices"][0]["message"]["content"].strip().strip('\'"`').strip()
                    out = re.sub(r'(\s*[-_]?\s*(tarjima|ozbekcha|ўзбекча)[a-z]*)$', '', out, flags=re.IGNORECASE)
                    out = re.sub(r'[/\\:*?"<>|_]', ' ', out)
                    out = re.sub(r'\s+', ' ', out).strip()
                    if out and len(out) > 2:
                        return ensure_script(out, target_script) if is_cyrillic else out
            except Exception:
                pass

        # 2. Gemini orqali urinish
        for model in self.model_candidates:
            try:
                res = self.client.models.generate_content(
                    model=model,
                    contents=prompt
                )
                out = res.text.strip().strip('\'"').strip()
                out = re.sub(r'(\s*[-_]?\s*(tarjima|ozbekcha|ўзбекча)[a-z]*)$', '', out, flags=re.IGNORECASE)
                out = re.sub(r'[/\\:*?"<>|_]', ' ', out)
                out = re.sub(r'\s+', ' ', out).strip()
                if out:
                    return ensure_script(out, target_script) if is_cyrillic else out
            except Exception:
                continue

        # Zaxira (fallback)
        try:
            from deep_translator import GoogleTranslator
            gt = GoogleTranslator(source='auto', target='uz')
            tr = gt.translate(clean_in)
            out = re.sub(r'(\s*[-_]?\s*(tarjima|ozbekcha|ўзбекча)[a-z]*)$', '', tr, flags=re.IGNORECASE)
            out = re.sub(r'[/\\:*?"<>|_]', ' ', out)
            out = re.sub(r'\s+', ' ', out).strip()
            return ensure_script(out, target_script) if (out and is_cyrillic) else (out or clean_in)
        except Exception:
            return clean_in
