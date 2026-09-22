# -*- coding: utf-8 -*-
"""
core/visual_rhetoric.py
Visual Rhetoric Compiler for SlaydHub AI.
Transforms flat bullet-point lists into dynamic visual metaphors:
1. Causal Flow (Cause -> Effect -> Impact)
2. Comparative Balance (Thesis vs Antithesis)
3. Hierarchical Funnel (Filtration & Prioritization)
4. Cyclical Loop (Feedback & Continuous Iteration)
5. Multi-Metric Impact (Key Statistics & Leverage)
Also compiles rich Dual-Track Speaker Notes and Q&A briefings.
"""

import re
import math
from typing import Dict, Any, List, Optional, Tuple


class VisualRhetoric:
    """
    Analyzes presentation discourse and compiles rhetorical structures into visual metaphors.
    """

    # Rhetorical triggers for Uzbek and English
    CAUSAL_TRIGGERS = [
        r'\bnatijasida\b', r'\btufayli\b', r'\bolib\s+keladi\b', r'\bsababli\b',
        r'\bshuning\s+uchun\b', r'\bbosqichma-bosqich\b', r'\bta\'sirida\b',
        r'\bcreates\b', r'\bleads\s+to\b', r'\bcauses\b', r'\bresults\s+in\b',
        r'\btherefore\b', r'\bconsequently\b', r'\bdue\s+to\b'
    ]

    COMPARISON_TRIGGERS = [
        r'\bvs\b', r'\bqarama-qarshi\b', r'\bfarqli\s+ravishda\b', r'\bboshqa\s+tomondan\b',
        r'\bafzallik\b', r'\bkamchilik\b', r'\bsolishtirganda\b', r'\baks\s+holda\b',
        r'\bversus\b', r'\bcompared\s+to\b', r'\bon\s+the\s+other\s+hand\b',
        r'\bpros\s+and\s+cons\b', r'\bcontrast\b', r'\bunlike\b'
    ]

    CYCLE_TRIGGERS = [
        r'\bsikl\b', r'\bdoimiy\b', r'\bqayta\b', r'\btakrorlanuvchi\b',
        r'\baylanma\b', r'\biteratsiya\b', r'\bfeedback\b', r'\bloop\b',
        r'\bcycle\b', r'\bcontinuous\b', r'\brecurrent\b', r'\bprocess\b'
    ]

    FUNNEL_TRIGGERS = [
        r'\bboshlang\'ich\b', r'\bsaralash\b', r'\byakuniy\b', r'\bfunel\b',
        r'\bvoronka\b', r'\bfiltrlash\b', r'\bierarxiya\b', r'\bdaraja\b',
        r'\bfunnel\b', r'\bfilter\b', r'\bstage\b', r'\bpipeline\b'
    ]

    @classmethod
    def extract_semantic_items(cls, slide_data: Dict[str, Any]) -> List[Tuple[str, str]]:
        """Polymorphic semantic extraction of slide content"""
        items = []

        # 1. Cards / Steps / Metrics / Cycle
        for key in ["cards", "steps", "points", "theses", "metrics", "components", "levels", "cycle_steps", "key_takeaways", "recommendations"]:
            val = slide_data.get(key, [])
            if isinstance(val, list) and val:
                for item in val:
                    if isinstance(item, dict):
                        h = str(item.get("title", "") or item.get("name", "") or item.get("label", "") or item.get("step_number", "")).strip()
                        d = str(item.get("description", "") or item.get("desc", "") or item.get("value", "") or item.get("details", "")).strip()
                        if h or d:
                            items.append((h, d))
                    elif isinstance(item, str) and item.strip():
                        items.append((item.strip(), ""))
                break

        # 2. Side-by-Side (Comparison)
        if not items:
            l_pts = slide_data.get("left_points", [])
            r_pts = slide_data.get("right_points", [])
            if l_pts or r_pts:
                for lp in l_pts:
                    items.append((str(lp), "Chap tomondagi ko'rsatkich"))
                for rp in r_pts:
                    items.append((str(rp), "O'ng tomondagi ko'rsatkich"))

        # 3. Content block
        if not items:
            c = slide_data.get("content", "")
            if isinstance(c, list):
                items = [(str(x), "") for x in c if str(x).strip()]
            elif isinstance(c, str) and c.strip():
                lines = [l.strip() for l in c.split("\n") if l.strip()]
                for l in lines:
                    if ":" in l:
                        h, d = l.split(":", 1)
                        items.append((h.strip(), d.strip()))
                    elif "-" in l:
                        h, d = l.split("-", 1)
                        items.append((h.strip(), d.strip()))
                    else:
                        items.append((l, ""))

        return items

    @classmethod
    def detect_rhetoric_type(cls, text: str, bullets: Optional[List[str]] = None) -> str:
        """Determines the dominant rhetorical structure of a slide."""
        full_text = text.lower()
        if bullets:
            full_text += " " + " ".join(bullets).lower()

        for pat in cls.COMPARISON_TRIGGERS:
            if re.search(pat, full_text):
                return "comparative_balance"
        for pat in cls.CYCLE_TRIGGERS:
            if re.search(pat, full_text):
                return "cyclical_loop"
        for pat in cls.CAUSAL_TRIGGERS:
            if re.search(pat, full_text):
                return "causal_flow"
        for pat in cls.FUNNEL_TRIGGERS:
            if re.search(pat, full_text):
                return "hierarchical_funnel"

        return "matrix_grid"

    @classmethod
    def compile_dual_track_notes(
        cls,
        slide_title: str,
        slide_data: Dict[str, Any],
        topic: str,
        slide_index: int = 1,
        total_slides: int = 4
    ) -> str:
        """
        Generates rich, professional Dual-Track Speaker Notes for the presenter view:
        1. 🎙️ 60-90 Second Spiker Nutqi (Natural speech script)
        2. 📊 Chuqur Ilmiy va Amaliy Dalillar (Detailed evidence/proofs)
        3. ❓ Auditoriya Kutilayotgan Savollari va Javoblar (Q&A prep)
        """
        items = cls.extract_semantic_items(slide_data)
        flat_texts = [f"{h} {d}" for h, d in items]
        rhetoric_type = cls.detect_rhetoric_type(slide_title, flat_texts)

        # 1. Generate Spiker Nutqi
        script_paragraphs = []
        if slide_index == 1:
            script_paragraphs.append(
                f"Hurmatli tinglovchilar va mutaxassislar! Bugun biz '{topic}' mavzusi bo'yicha tadqiqotimizni boshlaymiz. "
                f"Ushbu taqdimotimizda mavzuning eng muhim poydevori, texnologik va ilmiy yechimlari hamda amaliy ahamiyatini batafsil tahlil qilamiz."
            )
        else:
            if slide_index == total_slides:
                script_paragraphs.append(f"Xulosa qilib aytganda, '{slide_title}' masalasi tadqiqotimizning hal qiluvchi nuqtasidir.")
            else:
                script_paragraphs.append(f"Navbatdagi e'tibor qaratishimiz lozim bo'lgan muhim masala — '{slide_title}'.")

        if items:
            bullet_summaries = []
            for idx, (h, d) in enumerate(items[:3], 1):
                clean_h = re.sub(r'[\*\#]', '', h).strip()
                if not clean_h and d:
                    clean_h = re.sub(r'[\*\#]', '', d).strip()
                if clean_h:
                    if idx == 1:
                        bullet_summaries.append(f"Eng avvalo, biz {clean_h} ga e'tibor qaratamiz.")
                    elif idx == 2:
                        bullet_summaries.append(f"Undan so'ng yana bir muhim omil — {clean_h} hisoblanadi.")
                    else:
                        bullet_summaries.append(f"Qolaversa, {clean_h} ham amaliyotda o'z tasdig'ini topgan.")

            if bullet_summaries:
                script_paragraphs.append(" ".join(bullet_summaries))

        # Dynamic Rhetorical Additions
        if rhetoric_type == "comparative_balance":
            script_paragraphs.append(
                "Ushbu qiyosiy tahlil shuni ko'rsatadiki, eski va an'anaviy usullar ma'lum bir nuqtada o'z interfeysini yo'qotadi, "
                "boshqa tomondan esa taklif etilayotgan yondashuv nafaqat resurslarni, balki vaqtni ham optimallashtiradi."
            )
        elif rhetoric_type == "causal_flow":
            script_paragraphs.append(
                "Bu yerdagi har bir jarayon bir-biriga uzviy bog'langan. Sabab va oqibat munosabatlari orqali biz "
                "kuchli mexanizmni yaratdik, bu bevosita muammoning o'zak yechimiga olib keladi."
            )
        elif rhetoric_type == "cyclical_loop":
            script_paragraphs.append(
                "Bu doimiy sikl bo'lib, har bir iteratsiyada tizim o'z-o'zini moslashtirib, oldingi xatolardan "
                "xulosa chiqargan holda keyingi bosqichda yuqoriroq samaradorlikka erishadi."
            )

        speaker_script = "\n\n".join(script_paragraphs)

        # Word count mapping to speech time (approx 120 WPM conversational pace)
        total_words = len(speaker_script.split())
        est_seconds = math.ceil((total_words / 120) * 60)
        # Pad script if it's too short (aiming for 50-90s)
        if est_seconds < 45:
            speaker_script += "\n\nShartli ravishda aytish mumkinki, ushbu jarayonlarni chuqur o'rganish amaliy ahamiyatga juda ham boydir."
            est_seconds += 10

        # 2. Chuqur Ilmiy Dalillar
        evidence_lines = []
        evidence_lines.append(f"• Mavzu Konteksti: {topic} (Ritorika: {rhetoric_type.upper()})")
        for idx, (h, d) in enumerate(items, 1):
            clean_h = re.sub(r'[\*\#]', '', h).strip()
            clean_d = re.sub(r'[\*\#]', '', d).strip()
            if clean_h:
                if clean_d:
                    evidence_lines.append(f"• {clean_h}: {clean_d}")
                else:
                    evidence_lines.append(f"• Asosiy tezis: {clean_h}")
        if not items:
            evidence_lines.append("• Tahlil: Ushbu ko'rsatkichlar empirik tajribalar va akademik adabiyotlar asosida tasdiqlangan.")

        evidence_text = "\n".join(evidence_lines)

        # 3. Kutilayotgan Auditoriya Savollari (Dynamic based on content)
        qa_text_parts = []
        if len(items) >= 1 and items[0][0]:
            clean_topic1 = re.sub(r'[\*\#]', '', items[0][0]).strip()
            qa_text_parts.append(
                f"1-Savol: Nima uchun taqdimotda aynan '{clean_topic1}' masalasiga bunchalik katta urg'u berilgan?\n"
                f"   Javob: Chunki amaliyot shuni ko'rsatadiki, aynan ushbu omil butun jarayonning xavfsizligi "
                f"va sifatini belgilaydigan eng muhim katalizator hisoblanadi."
            )
        else:
            qa_text_parts.append(
                f"1-Savol: Nima uchun aynan ushbu parametrlar tanlandi?\n"
                f"   Javob: Ular amaldagi benchmarklarda eng yuqori barqarorlik va ishonchlilik ko'rsatkichlarini namoyish etgan."
            )

        if len(items) >= 2 and items[1][0]:
            clean_topic2 = re.sub(r'[\*\#]', '', items[1][0]).strip()
            qa_text_parts.append(
                f"2-Savol: '{clean_topic2}' jarayoniga boshqacha yondashuv qo'llash mumkin emasmidi?\n"
                f"   Javob: Biz bir nechta alternativ usullarni tahlil qildik, biroq aynan tasdiqlangan va biz keltirgan "
                f"klassik model iqtisodiy va texnik jihatdan eng maqbul yo'l ekanligi isbotlandi."
            )
        else:
            qa_text_parts.append(
                f"2-Savol: Ushbu texnologiyaning eng katta to'sig'i nima deb bilasiz?\n"
                f"   Javob: Asosiy to'siq — infratuzilmani moslashtirish va odamlarni yangi jarayonga o'rgatish (Change Management)."
            )

        qa_text = "\n\n".join(qa_text_parts)

        notes_content = (
            f"===================================================\n"
            f"🎙️ SPIKER NUTQI ({est_seconds} - {est_seconds + 15} soniyalik jonli nutq ssenariysi):\n"
            f"---------------------------------------------------\n"
            f"{speaker_script}\n\n"
            f"===================================================\n"
            f"📊 CHUQUR ILMIY VA AMALIY DALILLAR:\n"
            f"---------------------------------------------------\n"
            f"{evidence_text}\n\n"
            f"===================================================\n"
            f"❓ KUTILAYOTGAN SAVOLLAR (Q&A):\n"
            f"---------------------------------------------------\n"
            f"{qa_text}\n"
            f"==================================================="
        )

        return notes_content
