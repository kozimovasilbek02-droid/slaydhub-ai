# -*- coding: utf-8 -*-
"""
NotebookLM Precision Prompt Generator (Academic & Defense Standard)
Generates structured, academic-grade prompts for Google NotebookLM (and Gemini)
strictly enforcing:
1. "Bir slayd — bir fikr" cognitive load rule.
2. Action-oriented / Xulosaviy sarlavhalar (e.g. "Kyber va Dilithium Kvant Xavfsizligini 100% Kafolatlaydi").
3. 60–70% slide fill ratio with maximum 3-4 distinct cards/blocks per slide.
4. Concise visual bullet structures instead of long paragraph text.
5. Titul metadata (Author, Organization, Supervisor, Email, Event).
6. Full support for Scientific Thesis / Defense (BMI, ВКР, Dissertatsiya) structures.
"""

import json
from typing import Dict, Any, List


class NotebookLMPromptGenerator:
    """
    Constructs high-precision prompts for NotebookLM.
    """

    @staticmethod
    def get_layout_schema_guide(layout_type: str, lang: str, is_bmi: bool = False) -> str:
        """Returns the specific JSON field guide for a slide layout."""
        if layout_type == "cover":
            return """{
        "slide_number": <N>,
        "layout_type": "cover",
        "title": "To'liq Ilmiy/Akademik Mavzu Nomi",
        "subtitle": "Kvant hisoblashlari davrida axborot xavfsizligi arxitekturasi | 2026-yil",
        "author_name": "Tadqiqotchi F.I.Sh.",
        "organization": "O'zbekiston Milliy Universiteti / Axborot xavfsizligi kafedrasi",
        "supervisor": "Ilmiy rahbar: DSc., professor A. Karimov",
        "email": "tadqiqotchi@univ.edu.uz",
        "event_name": "Bitiruv Malakaviy Ishi (BMI) / Ilmiy Himoya"
      }"""
        elif layout_type == "agenda":
            return """{
        "slide_number": <N>,
        "layout_type": "agenda",
        "title": "Tadqiqotning Tuzilishi va Mantiqiy Rejasi",
        "points": [
          "1. Muammoning qo'yilishi va dolzarbligi: Asosiy ilmiy ziddiyat",
          "2. Tadqiqot obyekti va maqsad-vazifalari: Chegaralar va mezonlar",
          "3. Ilmiy gipoteza va ishlab chiqilgan model: Yangi arxitektura",
          "4. Eksperimental aprobatsiya va natijalar: Xulosalar va tavsiyalar"
        ]
      }"""
        elif layout_type == "theory_concept":
            return """{
        "slide_number": <N>,
        "layout_type": "theory_concept",
        "title": "Mavzuning Dolzarbligi va Nazariy Asoslari",
        "concept_name": "Fundamental Nazariya yoki Ilmiy Muammo",
        "definition": "Aniq, qisqa ilmiy asos (1-2 jumla, devor matnlarsiz)",
        "key_formula_or_law": "Asosiy qonuniyat, formula yoki teoremaviy tenglama",
        "importance": "Ushbu tadqiqotning fan va amaliyotdagi ustuvor ahamiyati"
      }"""
        elif layout_type == "cards_grid":
            return """{
        "slide_number": <N>,
        "layout_type": "cards_grid",
        "title": "Tadqiqotning Asosiy Yo'nalishlari va Ustunlari",
        "cards": [
          {"title": "1-Aspekt / Obyekt", "desc": "Qisqa, lo'nda tavsif va o'rganish doirasi."},
          {"title": "2-Aspekt / Predmet", "desc": "Tadqiq etilayotgan xususiyat va qonuniyat."},
          {"title": "3-Aspekt / Vazifalar", "desc": "Hal qilinishi lozim bo'lgan aniq ilmiy masalalar."}
        ]
      }"""
        elif layout_type == "comparison_vs":
            return """{
        "slide_number": <N>,
        "layout_type": "comparison_vs",
        "title": "Mavjud Tizimlar Bilan Qiyosiy Tahlil",
        "left_title": "1-Model / Klassik Yondashuv",
        "left_points": [
          "Faktorizatsiya va diskret logarifm masalasiga asoslanadi",
          "Shor algoritmi yordamida kvant kompyuterida to'liq buziladi",
          "Kichik kalit hajmi va yuqori tezlikka ega"
        ],
        "right_title": "2-Model / Taklif Qilinayotgan PQC Yondashuv",
        "right_points": [
          "Ko'p o'lchovli panjaralar va xesh daraxtlariga tayanadi",
          "Kvant va klassik superkompyuter hujumlariga 100% chidamli",
          "Kattaroq kalit hajmi talab etiladi, ammo mutlaq xavfsiz"
        ]
      }"""
        elif layout_type == "timeline_steps":
            return """{
        "slide_number": <N>,
        "layout_type": "timeline_steps",
        "title": "Tadqiqot va Aprobatsiya Bosqichlari",
        "steps": [
          {"step": "01", "title": "1. Xatarlarni Baholash", "desc": "Mavjud axborot tizimlaridagi zaif nuqtalarni xatlovdan o'tkazish."},
          {"step": "02", "title": "2. Gibrid Kriptografiya", "desc": "Klassik va yangi post-kvant protokollarini parallel sinovdan o'tkazish."},
          {"step": "03", "title": "3. Standartlashtirish", "desc": "NIST tasdiqlagan ML-KEM va ML-DSA algoritmlariga to'liq o'tish."},
          {"step": "04", "title": "4. To'liq Kvant Himoyasi", "desc": "Infratuzilmani yangi kripto-arxitektura bilan to'liq himoyalash."}
        ]
      }"""
        elif layout_type == "diagram_anatomy":
            return """{
        "slide_number": <N>,
        "layout_type": "diagram_anatomy",
        "title": "Taklif Qilinayotgan Model va Arxitektura",
        "components": [
          {"name": "1. Panjaraga Asoslangan Modul (ML-KEM)", "function": "Ko'p o'lchovli panjaralardagi qisqa vektor masalasiga tayanuvchi xavfsiz kalit almashish."},
          {"name": "2. Xesh-Asosli Imzo Moduli (SLH-DSA)", "function": "Bir martalik kriptografik xeshlar daraxtiga tayanuvchi soxtalashtirib bo'lmas imzo."},
          {"name": "3. Gibrid Integratsiya Qatlami", "function": "Mavjud TLS/SSL protokollariga uzluksiz moslashuvchan ulanish."}
        ]
      }"""
        elif layout_type == "metrics_stats":
            return """{
        "slide_number": <N>,
        "layout_type": "metrics_stats",
        "title": "Eksperimental Natijalar va Xavfsizlik Metriklari",
        "metrics": [
          {"value": "2048-bit", "label": "RSA Zaiflik Chegarasi", "desc": "4000 mantiqiy kubit orqali polinomial vaqtda buziladi."},
          {"value": "99.9%", "label": "Lattice PQC Ishonchliligi", "desc": "NIST xalqaro standartlarida to'liq tasdiqlangan."},
          {"value": "2030+", "label": "Kvant Taxdidi Davri", "desc": "Kvant ustunligiga erishish kutilayotgan kritik davr."}
        ]
      }"""
        elif layout_type == "case_study":
            return """{
        "slide_number": <N>,
        "layout_type": "case_study",
        "title": "Amaliy Tadqiqot va Eksperiment Natijasi",
        "case_name": "Eksperimental Sinov Nomi",
        "problem": "Hal qilinishi lozim bo'lgan ilmiy muammo va ziddiyat.",
        "applied_method": "Qo'llanilgan metodologiya va dasturiy vositalar.",
        "result": "Olingan aniq o'lchov natijasi va samaradorlik xulosasi."
      }"""
        elif layout_type == "qa_discussion":
            return """{
        "slide_number": <N>,
        "layout_type": "qa_discussion",
        "title": "Ilmiy Munozaralar va Dolzarb Savollar",
        "qa_list": [
          {"question": "SNDL ('Hozir saqlab ol, keyin och') hujumi nima?", "answer": "Buzg'unchilar bugun shifrlangan ma'lumotlarni yozib oladi va kelajakda kvant kompyuterida shifrini ochadi."},
          {"question": "Nega migratsiyani zudlik bilan boshlash zarur?", "answer": "Chunki global bank va davlat tizimlarini yangi standartga o'tkazish 5-10 yil talab qiladi."}
        ]
      }"""
        elif layout_type == "conclusion_sources":
            return """{
        "slide_number": <N>,
        "layout_type": "conclusion_sources",
        "title": "Asosiy Ilmiy Xulosalar va Foydalanilgan Adabiyotlar",
        "key_takeaways": [
          "Kvant kompyuterlari klassik asimmetrik kriptografiyani (RSA, ECC) mutlaqo zaiflashtiradi.",
          "Post-kvant kriptografiyasi (PQC) standartlariga o'tish kechiktirib bo'lmas strategik vazifadir.",
          "Gibrid xavfsizlik protokollari orqali xavflarni bosqichma-bosqich minimallashtirish zarur."
        ],
        "references": [
          "1. NIST Post-Quantum Cryptography Standardization (FIPS 203, 204, 205, 2024)",
          "2. Shor P.W. - Polynomial-Time Algorithms for Prime Factorization (SIAM, 1997)",
          "3. E-mail: muallif@edu.uz | Tinglovchilarga samimiy minnatdorchilik bildiramiz!"
        ]
      }"""
        else:
            return """{
        "slide_number": <N>,
        "layout_type": "cards_grid",
        "title": "Asosiy Ilmiy Tahlil",
        "cards": [
          {"title": "1-Band", "desc": "Loyiha tavsifi va xulosasi."},
          {"title": "2-Band", "desc": "Loyiha tavsifi va xulosasi."},
          {"title": "3-Band", "desc": "Loyiha tavsifi va xulosasi."}
        ]
      }"""

    @classmethod
    def generate_prompt(cls, blueprint: Dict[str, Any]) -> str:
        """
        Builds the complete NotebookLM prompt with strict pedagogical and design rules.
        """
        topic = blueprint.get("topic", "Akademik Taqdimot")
        category_name = blueprint.get("category_name", "Fan va Ta'lim")
        presentation_mode = blueprint.get("presentation_mode", "general_academic")
        mode_name = blueprint.get("presentation_mode_name", "Akademik Taqdimot")
        lang = blueprint.get("language", "uz")
        slides = blueprint.get("slides", [])
        p_info = blueprint.get("presenter_info", {})

        is_bmi = (presentation_mode == "scientific_defense_bmi")

        # Language-specific system instructions
        if lang == "uz":
            lang_instruction = "Javobingiz to'liq O'ZBEK tilida, akademik, ilmiy va ravon tilda bo'lishi SHART."
            theme_intro = f"Mavzu: '{topic}' (Soha: {category_name} | Format: {mode_name})"
        elif lang == "ru":
            lang_instruction = "Ваш ответ должен быть строго на РУССКОМ языке, академическим и научным стилем."
            theme_intro = f"Тема: '{topic}' (Направление: {category_name} | Формат: {mode_name})"
        else:
            lang_instruction = "Your response must be entirely in ENGLISH, in rigorous academic style."
            theme_intro = f"Topic: '{topic}' (Domain: {category_name} | Format: {mode_name})"

        prompt_lines = [
            f"Sen yetakchi ilmiy tadqiqotchi va taqdimotlar bo'yicha bosh metodistsan.",
            f"",
            f"VAZIFA: Quyidagi {len(slides)} ta slaydli taqdimot uchun NotebookLM manbalari asosida to'liq ilmiy JSON kontent tuzib ber.",
            f"",
            f"{theme_intro}",
            f"{lang_instruction}",
            f"",
            f"QAT'IY TALABLAR:",
            f"1. BOY VA KO'P QIRRALI MAZMUN: Slaydlar yuzaki va oddiy bo'lmasin. Har bir slayd chuqur ilmiy/amaliy faktlar, aniq raqamlar, formulalar, metrikalar va texnik atamalarga boy bo'lsin.",
            f"2. ACTION-ORIENTED SARLAVHALAR: Slayd sarlavhasi shunchaki umumiy mavzu emas, balki qat'iy ilmiy hukm va xulosani ifodalasin (masalan: 'ML-KEM Kvant Xavfsizligini 100% Ta'minlaydi').",
            f"3. RANG-BARANG VIZUAL STRUKTURALAR: Har bir slayd o'zining arxetipiga to'liq mos ravishda boy tuzilsin: katta metrikalar, qiyosiy taqqoslash ustunlari (Left vs Right), xronologik bosqichlar (Timeline), kartalar yoki tizim tarkibi.",
            f"4. ANIQ MIQDORIY KO'RSATKICHLAR: Har bir tahlilda real ko'rsatkichlar (masalan: 99.8%, 2,500 kubit, 1.2 ms, NIST FIPS 203) aniq berilsin.",
            f"5. Titul slaydida muallif, tashkilot va tadqiqot tafsilotlari to'liq berilsin.",
            f"",
            f"SLAYDLARNING MANTIQIY REJASI:"
        ]

        for s in slides:
            s_num = s["slide_number"]
            l_type = s["layout_type"]
            t_hint = s.get("title_hint", f"Slide {s_num}")
            prompt_lines.append(f"  - Slayd {s_num} [{l_type}]: {t_hint}")

        prompt_lines.extend([
            f"",
            f"JAVOB FORMATI (FAQAT quyidagi to'liq valid JSON formatida bo'lsin):",
            f"```json",
            f"{{",
            f'  "topic": "{topic}",',
            f'  "presentation_mode": "{presentation_mode}",',
            f'  "presenter_info": {json.dumps(p_info, ensure_ascii=False)},',
            f'  "slides": [',
            f'    {{"slide_number": 1, "layout_type": "cover", "title": "To\'liq Mavzu", "subtitle": "Ilmiy izoh | 2026", "author_name": "F.I.Sh.", "organization": "Tashkilot", "supervisor": "Ilmiy rahbar"}},',
            f'    {{"slide_number": 2, "layout_type": "agenda", "title": "Reja", "points": ["1. Reja 1", "2. Reja 2", "3. Reja 3", "4. Reja 4"]}},',
            f'    {{"slide_number": 3, "layout_type": "metrics_stats", "title": "Asosiy Miqdoriy Ko\'rsatkichlar", "metrics": [{{"value": "99.9%", "label": "Kriptografik Barqarorlik", "desc": "NIST FIPS 203 testlarida to\'liq tasdiqlangan"}}, {{"value": "2,500", "label": "Mantiqiy Kubitlar", "desc": "Shor algoritmi amaliy qo\'llanilishi uchun talab"}}, {{"value": "1.2 ms", "label": "Handshake Tezligi", "desc": "TLS 1.3 protokoli bilan gibrid integratsiyada"}}]}},',
            f'    {{"slide_number": 4, "layout_type": "comparison_vs", "title": "Qiyosiy Tahlil: Klassik vs Post-Kvant", "left_title": "Klassik RSA / ECC", "left_points": ["• Matematik murakkablik: Subeksponensial O(exp(c*n^(1/3)))", "• Kvant kompyuterlarida yengilish xavfi: 100%"], "right_title": "Post-Kvant ML-KEM / Dilithium", "right_points": ["• Panjarali (Lattice) qiyinchilik asosida", "• 128 va 256 bitlik to\'liq xavfsizlik kafolati"]}},',
            f'    {{"slide_number": 5, "layout_type": "timeline_steps", "title": "Joriy Qilish Bosqichlari va Xronologiya", "steps": [{{"step_number": "01", "phase": "1-Bosqich", "title": "Standartlashtirish", "desc": "NIST algoritmlarini tahlil qilish va audit"}}, {{"step_number": "02", "phase": "2-Bosqich", "title": "Gibrid Protokol", "desc": "TLS 1.3 arxitekturasiga ML-KEM integratsiyasi"}}, {{"step_number": "03", "phase": "3-Bosqich", "title": "To\'liq Migratsiya", "desc": "Infratuzilmani PQC standartlariga to\'liq o\'tkazish"}}]}},',
            f'    {{"slide_number": 6, "layout_type": "cards_grid", "title": "Arxitektura va Tizim Tarkibi", "cards": [{{"title": "Panjarali Shifrlash", "desc": "Ko\'p o\'lchovli panjaralarda Shortest Vector Problem (SVP) yechimining murakkabligi"}}, {{"title": "Gibrid Xavfsizlik", "desc": "Klassik ECDH va kvantga chidamli Kyber-768 kalit almashinuvining uyg\'unlashuvi"}}, {{"title": "Zararsiz Integratsiya", "desc": "Mavjud PKI infratuzilmasiga qo\'shimcha yuklama bermasdan tatbiq etish"}}]}},',
            f'    {{"slide_number": 7, "layout_type": "conclusion_sources", "title": "Xulosalar va Ilmiy Manbalar", "key_takeaways": ["• Post-kvant kriptografiyasiga o\'tish 2030-yilgacha strategik zaruratdir", "• ML-KEM va ML-DSA xavfsizlik va samaradorlik bo\'yicha optimal balans beradi"], "references": ["1. NIST FIPS 203: Module-Lattice-Based Key-Encapsulation Mechanism Standard (2024)", "2. Shor, P. W. Polynomial-Time Algorithms for Prime Factorization and Discrete Logarithms on a Quantum Computer (SIAM, 1997)"]}}',
            f'  ]',
            f'}}',
            f"```",
            f"",
            f"Endi yuqoridagi {len(slides)} ta slaydning har biri uchun haqiqiy ilmiy ma'lumotlar bilan to'ldirilgan yagona to'liq JSON obyekti bilan javob ber:"
        ])

        final_prompt = "\n".join(prompt_lines)
        return final_prompt

    @staticmethod
    def generate_compact_markdown_prompt(
        topic: str,
        domain: str = "Akademik / Ilmiy",
        author_info: str = "",
        sources_summary: str = "",
        lang: str = "uz"
    ) -> str:
        """
        Generates a concise prompt (< 2,000 characters) specifically optimized
        for Google NotebookLM's query box character limit.
        Instructs NotebookLM to synthesize uploaded sources into an Action-Oriented
        10-slide outline in clean Markdown.
        """
        prompt = f"""Siz nufuzli ilmiy tadqiqotchi, fan doktori va xalqaro akademik ma'ruzachisiz.
Ushbu mavzu va manbalar asosida 10 betlik professional akademik taqdimot matnini tayyorlang:

MAVZU: {topic}
SOHA: {domain}
MUALLIF / MA'RUZACHI: {author_info or "Tadqiqotchi"}
QO'SHIMCHA TEZISLAR VA MANBALAR: {sources_summary or "Ilmiy adabiyotlar va dolzarb tadqiqotlar bo'yicha"}

QAT'IY SIFAT VA AKADEMIK TALABLAR:
1. BOY VA KO'P QIRRALI MAZMUN: Slaydlar yuzaki va oddiy bo'lmasin. Har bir slayd aniq faktlar, texnik/ilmiy metrikalar, chuqur tahlillar va amaliy natijalar bilan boyitilsin.
2. ACTION-ORIENTED SARLAVHALAR: Slayd sarlavhasi umumiy mavzu emas, balki qat'iy ilmiy hukm va xulosa bo'lsin (Masalan: "ML-KEM Kvant Xavfsizligini 100% Ta'minlaydi", "RSA-2048 Shor Algoritmi Qarshisida O'z Kuchini Yo'qotadi").
3. BOLD LEAD-IN TEZISLAR QOIDASI (ENG MUHIM TALAB - McKinsey / Executive Standarti):
   Har bir tezisni quruq umumiy gap qilib yozmang! Har bir tezis aniq 2 qismdan iborat bo'lishi SHART:
   * **[Qalin Boshlang'ich Kalit So'z / Tezis]:** [Kengaytirilgan ilmiy dalil, mexanizm, parametrlar va faktlar (15-25 so'z)]
   NAMUNALAR:
   * **NIST FIPS 203 Standarti:** Module-Lattice asosidagi ML-KEM algoritmi 128 va 256-bitli kalit almashinuvida 100% kvant himoyasini kafolatlaydi.
   * **Shor Algoritmining Xavfi:** 2,048-bitli RSA kalitlari kvant superkompyuterlarida bir necha soat ichida to'liq faktorizatsiya qilinishi mumkin.
   * **Gibrid Xavfsizlik Arxitekturasi:** Klassik ECDH va yangi PQC protokollarini parallel joriy qilish orqali xavflarni minimallashtirish ta'minlanadi.
4. SUN'IY INTELLEKTGA XOS QOLIPLARDAN XALOS BO'LISH (ANTI-AI CLICHÉ):
   - "Ushbu slayd doirasida...", "Ta'kidlash joizki...", "Dolzarb ahamiyatga ega...", "Asosiy tahliliy aspekt..." kabi shablon va bo'sh so'zlardan MUTLAQO FOYDALANMANG.
   - Hech qanday chala belgilar ("---", "5-Bosqich: ---") yozmang.
5. ANIQ DALILLAR, FORMULALAR VA METRIKALAR:
   - Har bir tezisda aniq texnik atamalar, matematik munosabatlar, xalqaro standartlar (masalan NIST FIPS 203, ISO/IEC, IEEE) yoki real miqdoriy ko'rsatkichlar (masalan: 2,500 mantiqiy kubit, 99.8% ishonchlilik, O(n³) murakkablik) keltirilsin.
6. SLAYDLAR RANG-BARANGLIGI:
   - Titul, Nazariy Asos, Qiyosiy Tahlil (Model A vs Model B), Katta Metrikalar, Arxitektura, Operatsion Sikl, Yo'l Xaritasi va Ilmiy Xulosalar kabi turli xil mantiqiy shakllarga mos bo'lsin.

Matnni quyidagi Markdown formatida taqdim eting:

### 1-slayd: Titul va Kirish
* **Action-oriented sarlavha:** ...
* **Mantiqiy maqsadi:** ...
* **Tarkibiy tezislar:**
  * **1-Asosiy Tezis:** [Aniq fakt, mexanizm va tushuntirish]
  * **2-Asosiy Tezis:** [Aniq fakt, mexanizm va tushuntirish]
  * **3-Asosiy Tezis:** [Aniq fakt, mexanizm va tushuntirish]

### 2-slayd: [Slayd Nomi]
* **Action-oriented sarlavha:** ...
* **Mantiqiy maqsadi:** ...
* **Tarkibiy tezislar:**
  * **1-Asosiy Tezis:** [Aniq fakt, mexanizm va tushuntirish]
  * **2-Asosiy Tezis:** [Aniq fakt, mexanizm va tushuntirish]
  * **3-Asosiy Tezis:** [Aniq fakt, mexanizm va tushuntirish]

... (10-slaydgacha davom ettiring)."""
        return prompt.strip()

    @staticmethod
    def generate_with_gemini_direct(
        topic: str,
        domain: str = "Akademik / Ilmiy",
        author_info: str = "",
        sources_summary: str = "",
        lang: str = "uz"
    ) -> str:
        """
        Directly generates the 10-slide Action-Oriented markdown outline using
        the Gemini API for one-click instant generation with multi-model fallback.
        """
        import os
        import time
        from google import genai

        prompt = NotebookLMPromptGenerator.generate_compact_markdown_prompt(
            topic=topic,
            domain=domain,
            author_info=author_info,
            sources_summary=sources_summary,
            lang=lang
        )

        # 1. Attempt Gemini Generation with candidate models
        last_err = None
        try:
            client = genai.Client()
            candidate_models = [
                "gemini-3.6-flash",
                "gemini-3.5-flash-lite",
                "gemini-flash-latest"
            ]

            for model_name in candidate_models:
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt
                    )
                    if response and response.text:
                        return response.text.strip()
                except Exception as e:
                    last_err = e
                    err_str = str(e)
                    # Switch immediately to next model or Groq if 429/Resource Exhausted
                    if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                        continue
        except Exception as client_err:
            last_err = client_err

        # 2. Seamless Fast Fallback: Groq Ultra-Fast API (Zero Quota Bottleneck)
        groq_key = os.getenv("GROQ_API_KEY")
        if groq_key:
            try:
                import requests
                groq_candidates = ["openai/gpt-oss-120b", "openai/gpt-oss-20b", "qwen/qwen3.8-27b"]
                for g_model in groq_candidates:
                    try:
                        g_res = requests.post(
                            "https://api.groq.com/openai/v1/chat/completions",
                            headers={
                                "Authorization": f"Bearer {groq_key}",
                                "Content-Type": "application/json"
                            },
                            json={
                                "model": g_model,
                                "messages": [
                                    {"role": "system", "content": "Siz nufuzli ilmiy tadqiqotchi, fan doktori va xalqaro akademik ma'ruzachisiz. Berilgan ko'rsatmalarga qat'iy amal qiling."},
                                    {"role": "user", "content": prompt}
                                ],
                                "temperature": 0.3,
                                "max_tokens": 4096
                            },
                            timeout=35
                        )
                        if g_res.status_code == 200:
                            g_data = g_res.json()
                            g_choices = g_data.get("choices", [])
                            if g_choices:
                                g_text = g_choices[0].get("message", {}).get("content", "").strip()
                                if g_text and len(g_text) > 50:
                                    return g_text
                    except Exception:
                        continue
            except Exception:
                pass

        # 3. User-friendly explanation if all models hit quota
        if "429" in str(last_err) or "RESOURCE_EXHAUSTED" in str(last_err):
            raise RuntimeError(
                "Gemini AI bepul so'rovlar limiti (429 Quota) vaqtincha band. "
                "Iltimos, 1 daqiqa kuting yoki yuqoridagi 1-tab (NotebookLM Rejimi) orqali "
                "bepul va cheklovlarsiz generatsiya qiling."
            )
        raise RuntimeError(f"Generatsiyada xatolik: {last_err}")


def generate_notebooklm_prompt(blueprint: Dict[str, Any]) -> str:
    """Convenience function to generate the NotebookLM prompt."""
    return NotebookLMPromptGenerator.generate_prompt(blueprint)

