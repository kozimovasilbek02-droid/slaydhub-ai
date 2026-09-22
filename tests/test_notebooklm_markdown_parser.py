# -*- coding: utf-8 -*-
import unittest
from core.notebooklm_markdown_parser import NotebookLMMarkdownParser

QUANTUM_PRESENTATION_MARKDOWN = r"""
**'Kvant kompyuterlari va asimmetrik kriptografiya'** kabi chuqur ilmiy va strategik mavzudagi 10 betlik akademik taqdimot har bir slaydda **bitta aniq ilmiy g'oyani** (*Action-Oriented*) ochib beruvchi, 3-4 tadan oshmaydigan lo'nda tezislar va mantiqiy ketma-ketlikda tuzilishi lozim.

Quyida manbalarga tayanib tuzilgan **10 betlik slayd-taqdimotning mantiqiy strukturasi va tarkibiy mazmuni** keltirilgan:

---

### 1-slayd: Titul va Kirish
* **Action-oriented sarlavha:** Kvant kompyuterlari va asimmetrik kriptografiya: Tahdidlar va post-kvant o'tish strategiyasi.
* **Mantiqiy maqsadi:** Taqdimotning umumiy doirasini, dolzarbligini va tadqiqotning maqsad hamda vazifalarini belgilash.
* **Tarkibiy tezislar:**
  * Kvant hisoblash texnologiyalarining rivojlanishi va klassik kriptografik poydevorning zaiflashuvi.
  * Akademik tadqiqot maqsadi: Asimmetrik zaifliklarni tahlil qilish va PQC (*Post-Quantum Cryptography*) ga o'tish.
  * Tadqiqot ob'ektlari: RSA, ECC, AES, Shor va Grover algoritmlari hamda NIST PQC standartlari.

### 2-slayd: Klassik Asimmetrik Kriptografiyaning Matematik Asoslari
* **Action-oriented sarlavha:** Zamonaviy raqamli xavfsizlik va axborot almashinuvi RSA va ECC kabi asimmetrik shifrlarga tayanadi.
* **Mantiqiy maqsadi:** Hozirgi kunda ishlatilayotgan kriptotizimlarning qaysi qiyin matematik masalalarga asoslanganini ko'rsatish.
* **Tarkibiy tezislar:**
  * **RSA algoritmi:** Katta butun sonlarni tub ko'paytuvchilarga ajratish (faktorizatsiya) qiyinchiligiga asoslangan.
  * **ECC va Diffie-Hellman:** Elliptik egri chiziqlar va chekli maydonlardagi diskret logarifm masalasiga tayanadi.
  * **Klassik cheklov:** Ushbu matematik muammolarni klassik kompyuterlarda yechish subeksponensial vaqt talab etadi.

### 3-slayd: Shor Algoritmi va Asimmetrik Kriptografiyaning Inqirozi
* **Action-oriented sarlavha:** Shor algoritmi asimmetrik shifrlash va raqamli imzolarni polinom vaqtda to'liq buzadi.
* **Mantiqiy maqsadi:** Kvant kompyuterlarining asimmetrik shifrlarga bo'lgan strategik tahdidini ochib berish.
* **Tarkibiy tezislar:**
  * **Shor algoritmi (1994):** Kvant superpozitsiyasidan foydalanib faktorizatsiya va diskret logarifmni \\(O((\log N)^3)\\) murakkablikda yechadi.
  * **To'liq buziladigan standartlar:** RSA-2048, ECDSA, ECDH va Diffie-Hellman protokollari mutlaq zaiflashadi.
  * **Kubit ehtiyoji:** RSA-2048 va ECC-256 ni buzish uchun taxminan 1,200–2,500 ta mantiqiy kubit talab etiladi.

### 4-slayd: Grover Algoritmi va Simmetrik Shifrlash Barqarorligi
* **Action-oriented sarlavha:** Grover algoritmi simmetrik algoritmlarning effektiv kalit uzunligini teng yarimga qisqartiradi.
* **Mantiqiy maqsadi:** Simmetrik shifrlash (AES) hamda xesh-funsiyalarning kvant xavfsizligidagi o'rnini baholash.
* **Tarkibiy tezislar:**
  * **Grover algoritmi (1996):** Tartiblanmagan qidiruvni kvadratik tezlashtirib, operatsiyalar sonini \\(2^N\\) dan \\(2^{N/2}\\) ga tushiradi.
  * **AES-128 zaifligi:** Kalit uzunligining effektiv xavfsizligi 64-bit darajasiga tushib qoladi.
  * **Simmetrik yechim:** Kalit uzunligini AES-256 ga oshirish 128-bitli kvant barqarorligini saqlab qoladi.

### 5-slayd: "Harvest Now, Decrypt Later" (HNDL) Strategik Tahdidi
* **Action-oriented sarlavha:** HNDL taktikasi shifrlangan ma'lumotlarni bugundanoq to'plash va saqlash xavfini tug'diradi.
* **Mantiqiy maqsadi:** Nima uchun post-kvant xavfsizligiga o'tishni keyinga qoldirib bo'lmasligini tushuntirish.
* **Tarkibiy tezislar:**
  * **HNDL mexanizmi:** Hujumchilar shifrlangan trafikni bugun yozib oladi va kelajakda kvant kompyuterida shifrini ochadi.
  * **Kritik aktivlar:** Uzoq muddatli maxfiylikni talab qiluvchi davlat sirlari, tijorat va tibbiy ma'lumotlar xavf ostida.
  * **Sanoat xavotiri:** Tashkilotlarning 61% qismi HNDL ni eng birinchi navbatdagi kvant xavfi deb hisoblaydi.

### 6-slayd: NIST Post-Kvant Kriptografiya Standartlari (2024)
* **Action-oriented sarlavha:** NIST tomonidan tasdiqlangan PQC standartlari yangi kvant xavfsizligi poydevorini yaratdi.
* **Mantiqiy maqsadi:** Rasmiy tasdiqlangan va kvantga chidamli yangi algoritmlarni guruhlash.
* **Tarkibiy tezislar:**
  * **ML-KEM (FIPS 203):** Panjarali kriptografiyaga (*lattice-based*) asoslangan umumiy kalit almashish standarti.
  * **ML-DSA (FIPS 204):** Panjaradagi eng qisqa vektor masalasiga tayanuvchi raqamli imzo standarti.
  * **SLH-DSA (FIPS 205):** Xesh-funsiyalarga tayanuvchi, algebrik strukturasiz mustahkam imzo standarti.

### 7-slayd: Gibrid TLS 1.3 Kriptografik Arxitekturasi
* **Action-oriented sarlavha:** Gibrid TLS 1.3 modeli klassik va PQC algoritmlarini birlashtirib "chuqurlashtirilgan himoya" beradi.
* **Mantiqiy maqsadi:** O'tish davrida xavf-xatarlarni kamaytirishning amaliy texnik modelini ko'rsatish.
* **Tarkibiy tezislar:**
  * **Gibrid kalit almashinuvi:** ECDH va ML-KEM parallel ishlatilib, birgalikda seans kalitini shakllantiradi.
  * **Xavfsizlik kafolati:** Bitta algoritmda zaiflik topilsa ham ikkinchi algoritm maxfiylikni ta'minlaydi.
  * **TLS 1.3 roli:** Moslashuvchan Handshake strukturasi orqali tarmoq darajasida PQC inspeksiyasini amalga oshiradi.

### 8-slayd: Tizimli O'tish Metodologiyasi va Crypto-Agility
* **Action-oriented sarlavha:** Post-kvant migratsiyasi 4 bosqichli operatsion sikl va Crypto-Agility tamoyilini talab etadi.
* **Mantiqiy maqsadi:** Tashkilotlar uchun migratsiyaning amaliy harakatlar xaritasini berish.
* **Tarkibiy tezislar:**
  * **Avtomatlashtirilgan kashfiyot (Discovery):** CBOM hamda SAST/DAST va skanerlar orqali barcha kalitlar va algoritmlarni inventarizatsiya qilish.
  * **Prioritetlash va risk-analiz:** Kritik tizimlar hamda uzoq muddatli maxfiy ma'lumotlarni birinchi navbatda ajratish.
  * **Crypto-Agility:** Dasturiy kodda algoritmlarni osongina almashtirish imkonini yaratish.
  * **Doimiy monitoring:** Sertifikatlar va protokollarning mosligini avtomatik dashboardlar orqali tekshirib borish.

### 9-slayd: Global Migratsiya Tajribasi (AQSh va Ethereum Ekosistemasi)
* **Action-oriented sarlavha:** Dunyo davlatlari va blokcheyn ekosistemalari 2026–2030 yillarga mo'ljallangan qat'iy o'tish grafiklarini belgilamoqda.
* **Mantiqiy maqsadi:** Amaliy migratsiyaning xalqaro va davlat darajasidagi misollarini berish.
* **Tarkibiy tezislar:**
  * **AQSh Oq Uy M-26-15 Memorandumi:** Federal idoralardan PQC o'tish rejasini va avtomatlashtirilgan kashfiyotni talab qiladi.
  * **Ethereum Lean Roadmap:** Account Abstraction (EIP-8141) hamda 2029-yilgacha to'liq PQC imzo va STARK majburiyatlariga o'tishni rejalashtirgan.
  * **Tashkiliy xavfsizlik:** Zscaler va Thales kabi SSE provayderlari inline PQC inspeksiyasini va IPsec PPK larini amaliyotga joriy etmoqda.

### 10-slayd: Ilmiy Xulosalar va Strategik Tavsiyalar
* **Action-oriented sarlavha:** Kvant davrida axborot xavfsizligini ta'minlash bugundanoq kompleks va operatsion yondashuvni talab qiladi.
* **Mantiqiy maqsadi:** Taqdimotning yakuniy ilmiy va amaliy xulosalarini chiqarish.
* **Tarkibiy tezislar:**
  * **Eskirgan xavfsizlik:** Asimmetrik algoritmlardan (RSA/ECC) bosqichma-bosqich воздержаться va ularni PQC ga almashtirish shart.
  * **Infratuzilmaviy tayyorgarlik:** Gibrid TLS protokollari va agil kriptografik kutubxonalardan foydalanish lozim.
  * **Milliy va korporativ moslashuv:** Milliy kriptografik standartlarni kvant tahdidlariga moslashtirish hamda kadrlarni tayyorlash dolzarb vazifadir.
"""


class TestNotebookLMMarkdownParser(unittest.TestCase):

    def test_parse_quantum_presentation(self):
        result = NotebookLMMarkdownParser.parse(
            QUANTUM_PRESENTATION_MARKDOWN,
            topic="Kvant kompyuterlari va asimmetrik kriptografiya",
            default_metadata={
                "author_name": "Tadqiqotchi",
                "organization": "O'zbekiston Milliy Universiteti",
                "supervisor": "Prof. A. Karimov"
            }
        )

        slides = result.get("slides", [])
        self.assertEqual(len(slides), 10, f"Expected 10 slides, got {len(slides)}")

        # Verify Slide 1: Cover
        s1 = slides[0]
        self.assertEqual(s1["slide_number"], 1)
        self.assertEqual(s1["layout_type"], "cover")
        self.assertIn("Kvant kompyuterlari", s1["title"])
        self.assertEqual(s1["author_name"], "Tadqiqotchi")

        # Verify Slide 2: Theory Concept
        s2 = slides[1]
        self.assertEqual(s2["slide_number"], 2)
        self.assertIn(s2["layout_type"], ["theory_concept", "cards_grid"])
        self.assertGreater(len(s2["theses"]), 0)

        # Verify Slide 4: Comparison
        s4 = slides[3]
        self.assertEqual(s4["slide_number"], 4)
        self.assertIn(s4["layout_type"], ["comparison_vs", "cards_grid", "metrics_stats"])

        # Verify Slide 5: Metrics/Stats (61%)
        s5 = slides[4]
        self.assertEqual(s5["slide_number"], 5)
        self.assertIn(s5["layout_type"], ["metrics_stats", "cards_grid"])

        # Verify Slide 8: Process Cycle
        s8 = slides[7]
        self.assertEqual(s8["slide_number"], 8)
        self.assertIn(s8["layout_type"], ["process_cycle", "timeline_steps", "cards_grid"])

        # Verify Slide 9: Timeline Steps (2026-2030 Roadmap)
        s9 = slides[8]
        self.assertEqual(s9["slide_number"], 9)
        self.assertEqual(s9["layout_type"], "timeline_steps")

        # Verify Slide 10: Conclusion
        s10 = slides[9]
        self.assertEqual(s10["slide_number"], 10)
        self.assertEqual(s10["layout_type"], "conclusion_sources")
        self.assertIn("key_takeaways", s10)

        # Check visual diversity (at least 5 distinct archetypes among 10 slides)
        archetypes = set(s["layout_type"] for s in slides)
        self.assertGreaterEqual(len(archetypes), 5, f"Expected at least 5 distinct archetypes, got {len(archetypes)}: {archetypes}")

    def test_clean_formatting(self):
        cleaned = NotebookLMMarkdownParser._clean_formatting("**Bold Text** and `code`")
        self.assertEqual(cleaned, "Bold Text and code")


if __name__ == "__main__":
    unittest.main()
