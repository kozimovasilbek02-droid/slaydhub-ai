# -*- coding: utf-8 -*-
"""
tests/test_notebooklm_web_artifacts.py
Tests handling of real NotebookLM web copy-paste artifacts:
- Web citation buttons ("16more_horiz.", "21more_horiz")
- Concatenated inline headers ("TavsiyalarAction-oriented sarlavha:")
- Footnote badges ("[1]", "[2, 3]")
"""

import unittest
from core.notebooklm_markdown_parser import NotebookLMMarkdownParser


class TestNotebookLMWebArtifacts(unittest.TestCase):

    def test_strip_more_horiz_artifacts(self):
        sample = """
### 1-slayd: Titul va Kirish
* **Action-oriented sarlavha:** Kvant kompyuterlari va asimmetrik kriptografiya16more_horiz.
* **Mantiqiy maqsadi:** Umumiy maqsadni belgilash21more_horiz.
* **Tarkibiy tezislar:**
  * Kvant hisoblash texnologiyalari rivoji43more_horiz.
  * PQC standartlariga o'tish [1, 2]
"""
        parsed = NotebookLMMarkdownParser.parse(sample)
        self.assertEqual(len(parsed["slides"]), 1)
        s1 = parsed["slides"][0]
        self.assertNotIn("more_horiz", s1["title"])
        self.assertNotIn("16", s1["title"])
        self.assertNotIn("more_horiz", s1["purpose"])
        self.assertNotIn("more_horiz", s1["theses"][0])
        self.assertNotIn("[1, 2]", s1["theses"][1])

    def test_concatenated_headers_separation(self):
        sample = """
### 1-slayd: Kirish
* **Action-oriented sarlavha:** Kirish g'oyasi
* **Mantiqiy maqsadi:** Kirish maqsadi
* **Tarkibiy tezislar:**
  * Birinchi nuqta### 2-slayd: Taqqoslash
* **Action-oriented sarlavha:** Klassik vs Kvant
* **Mantiqiy maqsadi:** Taqqoslash maqsadi
* **Tarkibiy tezislar:**
  * Klassik kriptografiya zaiflashuvi.TavsiyalarAction-oriented sarlavha: 3-slayd Sarlavhasi
"""
        parsed = NotebookLMMarkdownParser.parse(sample)
        self.assertGreaterEqual(len(parsed["slides"]), 2)
        s1 = parsed["slides"][0]
        s2 = parsed["slides"][1]
        self.assertEqual(s1["slide_number"], 1)
        self.assertEqual(s2["slide_number"], 2)
        self.assertEqual(s2["layout_type"], "comparison_vs")

    def test_user_actual_screenshot_text(self):
        sample = """
### 1-slayd: Titul va Kirish
* **Action-oriented sarlavha:** Kvant kompyuterlari va asimmetrik kriptografiya: Tahdidlar va post-kvant o'tish strategiyasi.
* **Mantiqiy maqsadi:** Taqdimotning umumiy doirasini, dolzarbligini va tadqiqotning maqsad hamda vazifalarini belgilash.
* **Tarkibiy tezislar:**
  * Kvant hisoblash texnologiyalarining rivojlanishi va klassik kriptografik poydevorning zaiflashuvi.
  * Akademik tadqiqot maqsadi: Asimmetrik zaifliklarni tahlil qilish va PQC (*Post-Quantum Cryptography*) ga o'tish.
  * Tadqiqot ob'ektlari: RSA, ECC, AES, Shor va Grover algoritmlari hamda NIST standartlari.

---

### 2-slayd: Matematik Asoslar va Shor Algoritmining Asimmetrik Kriptotizimlarga Halokatli Ta'siri
* **Action-oriented sarlavha:** Shor algoritmi katta sonlarni faktorizatsiyalash va diskret logarifmni ko'phadli vaqtda yechish orqali RSA va ECC xavfsizligini nolga tenglashtiradi.
* **Mantiqiy maqsadi:** Kvant hisoblashning asimmetrik kriptografiyani buzish mexanizmini matematik va nazariy jihatdan isbotlash.
* **Tarkibiy tezislar:**
  * Shor algoritmining asosi: Kvant Furye almashtirishi (QFT) orqali davriylikni topish (period-finding).
  * Klassik murakkablik vs Kvant samaradorligi: Eksponensial murakkablikdan ko'phadli murakkablikka (O((log N)^3)) o'tish.
  * Zaiflik ko'lami: RSA (butun sonlarni ko'paytuvchilarga ajratish) va ECC (elliptik egri chiziqlarda diskret logarifm) to'liq yengiladi.

---

### 3-slayd: Simmetrik va Asimmetrik Kriptografiya Kvant Tahdidlari Taqqoslovi
* **Action-oriented sarlavha:** Kvant davrida asimmetrik tizimlar to'liq parchalanish xavfida, simmetrik tizimlar esa kalit uzunligini ikki barobar oshirish orqali saqlab qolinadi.
* **Mantiqiy maqsadi:** Kvant algoritmlarining ikki xil kriptografik arxitekturaga turlicha ta'sirini solishtirish.
* **Tarkibiy tezislar:**
  * Grover algoritmi va simmetrik kriptografiya: Brute-force qidiruvini kvadratik tezlashtiradi (O(N) dan O(sqrt(N)) ga).
  * Simmetrik himoya: AES-128 zaiflashadi, ammo AES-256 kvant hujumlariga qarshi yetarli xavfsizlik darajasini (128-bit) saqlaydi.
  * Asimmetrik tizimlarning muqarrar inqirozi: Kalit hajmini oshirish RSA/ECC xavfsizligini ta'minlay olmaydi, yangi arxitektura shart.
"""
        parsed = NotebookLMMarkdownParser.parse(sample)
        self.assertEqual(len(parsed["slides"]), 3)
        self.assertEqual(parsed["slides"][0]["layout_type"], "cover")
        self.assertEqual(parsed["slides"][1]["layout_type"], "theory_concept")
        self.assertEqual(parsed["slides"][2]["layout_type"], "comparison_vs")


if __name__ == "__main__":
    unittest.main()
