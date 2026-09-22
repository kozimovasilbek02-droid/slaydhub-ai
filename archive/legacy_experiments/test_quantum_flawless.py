# -*- coding: utf-8 -*-
import os
import sys
from pptx import Presentation

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.academic_matcher import get_academic_matcher
from core.pptx_text_replacer import replace_presentation_content

matcher = get_academic_matcher()
bp = matcher.build_blueprint("Kvant kompyuterlari va asimmetrik kriptografiya", slide_count=10, language="uz")

template_path = bp["primary_pptx_path"]
print(f"Using template: {os.path.basename(template_path)}")

sample_content = {
    "topic": "Kvant kompyuterlari va asimmetrik kriptografiya",
    "slides": [
        {
            "slide_number": 1,
            "layout_type": "cover",
            "title": "Kvant kompyuterlari va asimmetrik kriptografiya",
            "subtitle": "Kvant hisoblashlari davrida axborot xavfsizligi arxitekturasi | 2026-yil"
        },
        {
            "slide_number": 2,
            "layout_type": "agenda",
            "points": [
                "1. Kvant hisoblashlarining matematik asoslari: Kubitlar va superpozitsiya tahlili",
                "2. Shor algoritmi va RSA zaifligi: Asimmetrik kalitlarning buzilish mexanizmi",
                "3. Post-kvant kriptografiyasi (PQC): Yangi xavfsiz shifrlash standartlari",
                "4. Xulosalar va amaliy tavsiyalar: Milliy xavfsizlik va migratsiya bosqichlari"
            ]
        },
        {
            "slide_number": 3,
            "layout_type": "theory_concept",
            "title": "Kvant Hisoblashlari Nazariy Asoslari",
            "concept_name": "Kvant Superpozitsiyasi va Kubit",
            "definition": "Klassik bitlardan farqli o'laroq, bir vaqtning o'zida |0> va |1> holatlarining chiziqli kombinatsiyasida bo'la oluvchi axborot birligi.",
            "key_formula_or_law": "|psi> = alpha|0> + beta|1>, |alpha|^2 + |beta|^2 = 1",
            "importance": "Eksponentsial parallel hisoblash imkoniyatini taqdim etadi."
        },
        {
            "slide_number": 4,
            "layout_type": "cards_grid",
            "title": "Zaiflikka Uchrovchi Asimmetrik Tizimlar",
            "cards": [
                {"title": "RSA Shifrlash", "desc": "Katta sonlarni tub ko'paytuvchilarga ajratish Shor algoritmi yordamida polinomial vaqtda yechiladi."},
                {"title": "ECC (Elliptik Egri Chiziqlar)", "desc": "Diskret logarifm masalasi kvant kompyuterlari tomonidan to'liq buziladi."},
                {"title": "Diffie-Hellman Kalit Almashinuvi", "desc": "Klassik kalit almashish protokoli kvant hujumlariga qarshi mutlaqo himoyasiz."}
            ]
        },
        {
            "slide_number": 5,
            "layout_type": "diagram_anatomy",
            "title": "Post-Kvant Kriptografiyasi (PQC) Arxitekturasi",
            "components": [
                {"name": "Lattice-Based Kriptografiya (Kyber/Dilithium)", "function": "Ko'p o'lchovli panjaralardagi eng qisqa vektor masalasiga asoslangan xavfsiz arxitektura."},
                {"name": "Hash-Based Imzolar (SPHINCS+)", "function": "Bir martalik kriptografik xesh funksiyalar daraxtiga tayanuvchi mustahkam himoya."},
                {"name": "Code-Based Tizimlar (Classic McEliece)", "function": "Xatoliklarni tuzatuvchi chiziqli kodlar orqali shifrlash."}
            ]
        },
        {
            "slide_number": 6,
            "layout_type": "comparison_vs",
            "title": "Klassik Kriptografiya vs Post-Kvant Kriptografiyasi",
            "left_title": "Klassik Asimmetrik Kriptografiya",
            "left_points": [
                "Faktorizatsiya va diskret logarifmga tayanadi",
                "Kvant kompyuterlari qarshisida mutlaqo zaif",
                "Kichik kalit hajmi va yuqori tezlik"
            ],
            "right_title": "Post-Kvant Kriptografiyasi (PQC)",
            "right_points": [
                "Panjaralar va xesh daraxtlariga asoslangan",
                "Kvant va klassik hujumlarga 100% chidamli",
                "Kattaroq kalit hajmi talab etiladi"
            ]
        },
        {
            "slide_number": 7,
            "layout_type": "timeline_steps",
            "title": "PQC Tizimlariga O'tish Bosqichlari",
            "steps": [
                {"title": "1. Xatarlarni Baholash", "desc": "Mavjud axborot tizimlaridagi asimmetrik shifrlash nuqtalarini xatlovdan o'tkazish."},
                {"title": "2. Gibrid Kriptografiya", "desc": "Klassik RSA/ECC va PQC algoritmlarini parallel qo'llash."},
                {"title": "3. Standartlashtirish", "desc": "NIST tomonidan tasdiqlangan Kyber va Dilithium standartlariga to'liq o'tish."},
                {"title": "4. To'liq Kvant Himoyasi", "desc": "Infratuzilmani kvantga chidamli arxitektura bilan to'liq yangilash."}
            ]
        },
        {
            "slide_number": 8,
            "layout_type": "metrics_stats",
            "title": "Kvant Xavfsizlik Metriklari va Ko'rsatkichlar",
            "metrics": [
                {"value": "2048-bit", "label": "RSA Zaiflik Chegarasi", "desc": "Taxminan 4000 mantiqiy kubit talab etiladi."},
                {"value": "99.9%", "label": "Lattice PQC Ishonchliligi", "desc": "NIST xalqaro testlarida tasdiqlangan xavfsizlik."},
                {"value": "2030+", "label": "Kvant Taxdidi Davri", "desc": "Kvant ustunligiga erishish kutilayotgan davr."}
            ]
        },
        {
            "slide_number": 9,
            "layout_type": "qa_discussion",
            "title": "Ilmiy Munozaralar va Dolzarb Savollar",
            "qa_list": [
                {"question": "'Hozir saqlab ol, keyin shifrni och' (SNDL) hujumi nima?", "answer": "Buzg'unchilar bugungi shifrlangan maxfiy ma'lumotlarni yozib oladi va kelajakda kvant kompyuteri paydo bo'lganda shifrini ochadi."},
                {"question": "Qachondan PQC ga o'tishni boshlash kerak?", "answer": "Zudlik bilan! Chunki global infratuzilmani yangi kripto-standartlarga o'tkazish 5-10 yil vaqt talab qiladi."}
            ]
        },
        {
            "slide_number": 10,
            "layout_type": "conclusion_sources",
            "title": "Ilmiy Xulosalar va Foydalanilgan Adabiyotlar",
            "key_takeaways": [
                "Kvant kompyuterlari asimmetrik kriptografiya xavfsizligiga tubdan xavf soladi.",
                "PQC (Post-Quantum Cryptography) standartlariga o'tish kechiktirib bo'lmas strategik vazifadir.",
                "Gibrid xavfsizlik protokollari orqali xavflarni minimallashtirish zarur."
            ],
            "references": [
                "1. NIST Post-Quantum Cryptography Standardization (2024)",
                "2. Shor P.W. - Polynomial-Time Algorithms for Prime Factorization (SIAM, 1997)"
            ]
        }
    ]
}

out_path = os.path.abspath(r"output\academic_studio\Kvant_kompyuterlari_FLAWLESS.pptx")
replace_presentation_content(template_path, sample_content, out_path, blueprint=bp)

# Also copy to Downloads for user inspection
user_download_path = r"C:\Users\user\Downloads\Kvant_kompyuterlari_va_asimmetrik_kripto_Customized.pptx"
import shutil
shutil.copyfile(out_path, user_download_path)

print(f"\n✅ FLAWLESS PPTX Created: {out_path} ({os.path.getsize(out_path):,} bytes)")

prs = Presentation(out_path)
for i, slide in enumerate(prs.slides):
    texts = [s.text_frame.text.strip() for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    print(f"\n--- SLIDE {i+1} ({len(texts)} non-empty shapes) ---")
    for t in texts:
        print(f"  [{len(t)} chars] {repr(t[:70])}")
