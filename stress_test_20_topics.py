# -*- coding: utf-8 -*-
"""
stress_test_20_topics.py
20-Topic Stress-Test & Comprehensive Verification Suite for SlaydHub AI.
Tests every archetype, LaTeX equation, extreme metric values, 4-step cycles,
dark themes, and sidebars across 20 distinct academic/scientific disciplines.
"""

import os
import sys
import io
import time
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

WORKSPACE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(WORKSPACE_DIR))

from core.academic_matcher import get_academic_matcher
from core.notebooklm_markdown_parser import NotebookLMMarkdownParser
from core.template_pptx_builder import get_academic_pptx_builder
from core.slide_preview import get_presentation_previews
from core.config import config

TOPICS_DATA = [
    {
        "id": 1,
        "topic": "Kvant Kriptografiyasi va Asimmetrik Shifrlash",
        "tested_modes": "theory_concept (LaTeX), comparison_vs",
        "markdown": """
### 1-slayd: Kvant Kriptografiyasi va Asimmetrik Shifrlash
* **Action-oriented sarlavha:** Kvant Kompyuterlari Davrida Asimmetrik Kriptografiya Xavfsizligi
* **Mantiqiy maqsadi:** Mavzuning ilmiy dolzarbligi va post-kvant standartlariga o'tish zarurati
* **Vizual turi:** cover

### 2-slayd: Taqdimot Rejasi
* **Action-oriented sarlavha:** Tadqiqotning Mantiqiy Bosqichlari va Ish Rejasi
* **Mantiqiy maqsadi:** Taqdimot davomida ko'rib chiqiladigan asosiy ilmiy yo'nalishlar
* **Vizual turi:** agenda
* **Tarkibiy tezislar:**
  * 01. Nazariy Asoslar: Kvant bitlari va superpozitsiya tamoyillari.
  * 02. Kriptotahlil Xavflari: Shor va Grover algoritmlarining tahdidlari.
  * 03. Post-Kvant Standartlari: ML-KEM va ML-DSA algoritmlarining ishlashi.
  * 04. Amaliy Natijalar: Algoritmlarning hisoblash tezligi va kalit hajmi.

### 3-slayd: Kvant Nazariyasi va Qonuniyat
* **Action-oriented sarlavha:** Kvant Superpozitsiyasi Holati va Hisoblash Murakkabligi
* **Mantiqiy maqsadi:** Bloch sferasi va kvant holatining matematik ifodasi
* **Vizual turi:** theory_concept
* **Tarkibiy tezislar:**
  * Kvant Holati Superpozitsiyasi: $|\\psi\\rangle = \\alpha|0\\rangle + \\beta|1\\rangle$ to'lqin funksiyasi.
  * Shor Algoritmi Asimptotikasi: $O((\\log N)^3)$ hisoblash murakkabligi bilan RSA kalitlarini buzadi.

### 4-slayd: Qiyosiy Tahlil
* **Action-oriented sarlavha:** Klassik RSA vs Post-Kvant Panjara Asosidagi ML-KEM Solishtirmasi
* **Mantiqiy maqsadi:** Kalit hajmi va xavfsizlik chegaralari tahlili
* **Vizual turi:** comparison_vs
* **Tarkibiy tezislar:**
  * Klassik RSA-2048: Kalit hajmi 256 bayt, kvant hujumiga chidamlilik 0%, hisoblash vaqti 4.2 ms.
  * Post-Kvant ML-KEM-768: Kalit hajmi 1,184 bayt, kvant hujumiga chidamlilik 100%, hisoblash vaqti 1.1 ms.
"""
    },
    {
        "id": 2,
        "topic": "Katta Til Modellari (LLM) va Transformer Arxitekturasi",
        "tested_modes": "metrics_stats, diagram_anatomy",
        "markdown": """
### 1-slayd: Katta Til Modellari va Transformer Arxitekturasi
* **Action-oriented sarlavha:** Transformer Arxitekturasi Matn Kontekstini Global E'tibor Orqali Qayta Ishlaydi
* **Mantiqiy maqsadi:** Zamonaviy LLM modellarining ishlash mexanizmi
* **Vizual turi:** cover

### 2-slayd: Asosiy Samaradorlik Metrikalari
* **Action-oriented sarlavha:** LLM Modellari Masshtabi va Unumdorlik Ko'rsatkichlari
* **Mantiqiy maqsadi:** GPT-4 va zamonaviy ochiq modellar parametrlar tahlili
* **Vizual turi:** metrics_stats
* **Tarkibiy tezislar:**
  * 1.8T Parametrlar: GPT-4 arxitekturasi ko'p ekspertli (MoE) parametrlar masshtabi.
  * 94.2% Aniqlik: MMLU akademik testlarida erishilgan eng yuqori natija.
  * 128k Tokenlar: Bitta so'rovda qayta ishlanadigan kontekst oynasi hajmi.

### 3-slayd: Transformer Qatlamlari Anatomiyasi
* **Action-oriented sarlavha:** Multi-Head Attention va Feed-Forward Qatlamlari Tuzilishi
* **Mantiqiy maqsadi:** Neyron to'ri bloklarining o'zaro bog'liqligi
* **Vizual turi:** diagram_anatomy
* **Tarkibiy tezislar:**
  * Multi-Head Attention: Kontekstdagi barcha so'zlar o'rtasidagi semantik munosabatni hisoblaydi.
  * Layer Normalization: Gradientlar tarqalishini barqarorlashtiradi va o'qitishni tezlashtiradi.
  * Feed-Forward Network: Har bir token uchun noxattiylik va chuqur ifodalarni shakllantiradi.
"""
    },
    {
        "id": 3,
        "topic": "Molekulyar Genetika va CRISPR-Cas9 Texnologiyasi",
        "tested_modes": "process_cycle, cards_grid",
        "markdown": """
### 1-slayd: Molekulyar Genetika va CRISPR-Cas9
* **Action-oriented sarlavha:** CRISPR-Cas9 Genomni Aniq Tahrirlashning Molekulyar Mexanizmi
* **Mantiqiy maqsadi:** Gen muhandisligidagi inqilobiy yondashuv
* **Vizual turi:** cover

### 2-slayd: 4 Bosqichli Gen Tahrirlash Sikli
* **Action-oriented sarlavha:** CRISPR Mexanizmining Ketma-ket Amalga Oshirilish Sikli
* **Mantiqiy maqsadi:** Cas9 fermentining nishon DNKga ta'siri bosqichlari
* **Vizual turi:** process_cycle
* **Tarkibiy tezislar:**
  * 01. gRNK Sintezi: Maqsadli gen ketma-ketligiga mos qo'llanuvchi RNK yaratiladi.
  * 02. Cas9 Kompleksi: Ferment gRNK bilan birlashib nishon lokusni qidiradi.
  * 03. Ikki Zanjirli Kesish: Cas9 DNK zanjirida aniq double-strand break hosil qiladi.
  * 04. Tabiiy Tiklanish: Hujayra o'z mexanizmlari orqali gen ketma-ketligini tahrirlaydi.

### 3-slayd: Asosiy Molekulyar Komponentlar
* **Action-oriented sarlavha:** Gen Tahrirlash Kompleksining Funksional Tarkibi
* **Mantiqiy maqsadi:** Ferment va zanjirlar tavsifi
* **Vizual turi:** cards_grid
* **Tarkibiy tezislar:**
  * Cas9 Endonukleaza: Molekulyar qaychi vazifasini bajaruvchi oqsil fermenti.
  * Qo'llanuvchi gRNK: 20 ta nukleotiddan iborat nishonni aniqlovchi yo'naltiruvchi zanjir.
  * PAM Ketma-ketligi: Cas9 fermenti bog'lanishi uchun zarur bo'lgan qisqa motiv.
"""
    },
    {
        "id": 4,
        "topic": "Kiberxavfsizlikda Zero-Trust Arxitekturasi",
        "tested_modes": "pyramid_funnel (Qatlamli himoya)",
        "markdown": """
### 1-slayd: Kiberxavfsizlikda Zero-Trust
* **Action-oriented sarlavha:** Zero-Trust Arxitekturasi: "Hech Kimga Ishonma, Har Doim Tekshir"
* **Mantiqiy maqsadi:** Zamonaviy korporativ tarmoq xavfsizligi
* **Vizual turi:** cover

### 2-slayd: Qatlamli Himoya Piramidasi
* **Action-oriented sarlavha:** Identifikatsiyadan Ma'lumotgacha Bo'lgan 4 Qatlamli Himoya
* **Mantiqiy maqsadi:** Zero-Trust xavfsizlik iyerarxiyasi
* **Vizual turi:** pyramid_funnel
* **Tarkibiy tezislar:**
  * 1-Qatlam (Tashqi): Foydalanuvchi identifikatsiyasi va ko'p omilli autentifikatsiya (MFA).
  * 2-Qatlam: Qurilma holati va xavfsizlik sertifikatlarini tekshirish.
  * 3-Qatlam: Mikrosegmentatsiya va tarmoq trafigini to'liq shifrlash.
  * 4-Qatlam (Yadro): Ma'lumotlarni nozik darajada himoyalash va audit qilish.
"""
    },
    {
        "id": 5,
        "topic": "Markaziy Banklarning Raqamli Valyutalari (CBDC) va Monetar Siyosat",
        "tested_modes": "metrics_stats, comparison_vs",
        "markdown": """
### 1-slayd: Markaziy Banklarning Raqamli Valyutalari
* **Action-oriented sarlavha:** CBDC Moliyaviy Barqarorlik va To'lov Tizimlari Samaradorligini Oshiradi
* **Mantiqiy maqsadi:** Davlat raqamli pullarining makroiqtisodiy ta'siri
* **Vizual turi:** cover

### 2-slayd: Global CBDC Ko'rsatkichlari
* **Action-oriented sarlavha:** Raqamli Valyutalarning Global Miqyosdagi Ko'rsatkichlari
* **Mantiqiy maqsadi:** CBDC sinovlarining xalqaro statistikasi
* **Vizual turi:** metrics_stats
* **Tarkibiy tezislar:**
  * $5.2 Trillion: 2030-yilgacha CBDC orqali amalga oshiriladigan yillik tranzaksiyalar hajmi.
  * 134 Davlat: Dunyo yalpi ichki mahsulotining 98% qismini tashkil etuvchi davlatlar CBDC ustida ishlamoqda.
  * 0.01$: Transchegaraviy to'lovlar uchun bitta tranzaksiya o'rtacha xarajati.

### 3-slayd: Kriptovalyuta vs Davlat CBDC
* **Action-oriented sarlavha:** Decentralized Kriptovalyutalar va Davlat CBDC Tizimlari Solishtirmasi
* **Mantiqiy maqsadi:** Huquqiy va iqtisodiy farqlar tahlili
* **Vizual turi:** comparison_vs
* **Tarkibiy tezislar:**
  * Xususiy Kriptovalyutalar: Yuqori volatillik, anonimlik va markazlashmagan boshqaruv mexanizmi.
  * Davlat CBDC Tizimlari: Markaziy bank kafolati, qonuniy to'lov vositasi maqomi va to'liq nazorat.
"""
    },
    {
        "id": 6,
        "topic": "Termoyadroviy Sintez va Tokamak Reaktorlari Fizikasi",
        "tested_modes": "theory_concept (Plazma harorati, E=mc^2)",
        "markdown": """
### 1-slayd: Termoyadroviy Sintez va Tokamak
* **Action-oriented sarlavha:** Tokamak Reaktorlari Termoyadroviy Plazmani Magnit Tuzoqda Ushtab Turadi
* **Mantiqiy maqsadi:** Boshqariladigan termoyadroviy sintez fizikasi
* **Vizual turi:** cover

### 2-slayd: Termoyadroviy Sintezning Nazariy Asoslari
* **Action-oriented sarlavha:** Eynshteyn Mass-Energiya Ekvivalentligi va Deyteriy-Tritiy Reaksiyasi
* **Mantiqiy maqsadi:** Plazma harorati va energiya ajralishi qonuniyatlari
* **Vizual turi:** theory_concept
* **Tarkibiy tezislar:**
  * Mass-Energiya Ekvivalentligi: $E = \\Delta m \\cdot c^2$ formulasi bo'yicha yadroviy massa defektidan energiya ajraladi.
  * Plazma Harorati: $T_p = 1.5 \\times 10^8\\text{ K}$ haroratida Kulon to'sig'i yengiladi.
"""
    },
    {
        "id": 7,
        "topic": "Mikroservislar vs Monolit Dasturiy Arxitektura",
        "tested_modes": "comparison_vs (Monolit vs Mikroservis)",
        "markdown": """
### 1-slayd: Mikroservislar va Monolit Arxitektura
* **Action-oriented sarlavha:** Zamonaviy Bulutli Tizimlarda Dasturiy Arxitekturani Tanlash Strategiyasi
* **Mantiqiy maqsadi:** Arxitekturalar solishtirmasi
* **Vizual turi:** cover

### 2-slayd: Monolit vs Mikroservis Qiyosiy Tahlili
* **Action-oriented sarlavha:** Monolit va Mikroservisli Tizimlarning Operatsion Samaradorligi
* **Mantiqiy maqsadi:** Masshtablash va nosozliklarga chidamlilik
* **Vizual turi:** comparison_vs
* **Tarkibiy tezislar:**
  * Monolit Arxitektura: Yagona kod bazasi, tezkor dastlabki ishlab chiqish, ammo qiyin masshtablanish.
  * Mikroservis Arxitektura: Mustaqil konteynerlar, moslashuvchan CI/CD va yuqori xatoliklarga chidamlilik.
"""
    },
    {
        "id": 8,
        "topic": "Kardiologiyada Sun'iy Intellekt va EKG Tahlili",
        "tested_modes": "timeline_steps (Tibbiy jarayon)",
        "markdown": """
### 1-slayd: Kardiologiyada Sun'iy Intellekt
* **Action-oriented sarlavha:** Chuqur O'rganish Modellari EKG Signallarida Aritmiyani 99% Aniqlikda Aniqlaydi
* **Mantiqiy maqsadi:** Raqamli kardiologiya
* **Vizual turi:** cover

### 2-slayd: EKG Tahlilining Avtomatlashgan Bosqichlari
* **Action-oriented sarlavha:** Bemor Signalidan Avtomatik Tashxisgacha Bo'lgan 4 Bosqich
* **Mantiqiy maqsadi:** AI kardiologik tahlil oqimi
* **Vizual turi:** timeline_steps
* **Tarkibiy tezislar:**
  * 1-Bosqich: 12-tarmoqli EKG sensori orqali xomashyo signallarini yuqori chastotada yozib olish.
  * 2-Bosqich: Shovqinlarni Wavelet filtrlash va P-QRS-T komplekslarini segmentatsiya qilish.
  * 3-Bosqich: Konvolyutsion neyrotarmoq (CNN) orqali patologik anomaliyalarni klassifikatsiya qilish.
  * 4-Bosqich: Shifokor uchun tushunarli vizual xulosa va xavf darajasini ko'rsatuvchi hisobot.
"""
    },
    {
        "id": 9,
        "topic": "Qayta Tiklanuvchi Energiya: Perovskit Quyosh Elementlari",
        "tested_modes": "metrics_stats, cards_grid",
        "markdown": """
### 1-slayd: Perovskit Quyosh Elementlari
* **Action-oriented sarlavha:** Perovskit Materiallari Quyosh Panellari Samaradorligini 30% Dan Oshirmoqda
* **Mantiqiy maqsadi:** Fotoelektr texnologiyalari
* **Vizual turi:** cover

### 2-slayd: Samaradorlik Rekordlari
* **Action-oriented sarlavha:** Perovskit Elementlarining Laboratoriya va Sanoat Ko'rsatkichlari
* **Mantiqiy maqsadi:** Samaradorlik va ishlab chiqarish parametrlari
* **Vizual turi:** metrics_stats
* **Tarkibiy tezislar:**
  * 26.1%: Laboratoriya sharoitida bir zanjirli perovskit fotoelementlarining samaradorlik rekordi.
  * 1,500 Vt/m²: Tangem kremniy-perovskit panellarida maksimal quvvat zichligi.
  * 5x: An'anaviy silikon ishlab chiqarishga qaraganda arzonroq energiya sarfi.
"""
    },
    {
        "id": 10,
        "topic": "Avtonom Haydash va Kompyuter Ko'rishi (Sensor Fusion)",
        "tested_modes": "diagram_anatomy (LiDAR, Radar, Kameralar)",
        "markdown": """
### 1-slayd: Avtonom Haydash va Sensor Fusion
* **Action-oriented sarlavha:** Sensor Fusion Algoritmlari Atrof-Muhitning 360 Darajali 3D Xaritasini Shakllantiradi
* **Mantiqiy maqsadi:** Avtonom transport tizimlari
* **Vizual turi:** cover

### 2-slayd: Sensorlar Majmuasi Anatomiyasi
* **Action-oriented sarlavha:** Avtonom Avtomobilning Sensorlar Tizimi Tarkibiy Qismlari
* **Mantiqiy maqsadi:** LiDAR, Radar va Kameralar integratsiyasi
* **Vizual turi:** diagram_anatomy
* **Tarkibiy tezislar:**
  * LiDAR Sensori: 200 metrgacha masofada 3D nuqtalar bulutini yaratuvchi lazer skaneri.
  * Radar Tizimi: Har qanday ob-havo sharoitida to'siqlarning nisbiy tezligini o'lchovchi datchik.
  * HD Kameralar: Yo'l belgilari, chiziqlar va svetofor chiroqlarini taniy oluvchi ko'rish tizimi.
"""
    },
    {
        "id": 11,
        "topic": "Agrotexnologiyada Dronlar va Dala Monitoringi",
        "tested_modes": "process_cycle (Skanerlash -> Tahlil -> Purkash -> Prognoz)",
        "markdown": """
### 1-slayd: Agrotexnologiyada Dronlar
* **Action-oriented sarlavha:** Aqlli Dronlar Qishloq Xo'jaligida Resurslarni 40% Gacha Tejaydi
* **Mantiqiy maqsadi:** Dala monitoringi
* **Vizual turi:** cover

### 2-slayd: 4 Bosqichli Agro-Monitoring Sikli
* **Action-oriented sarlavha:** Dala Ekinlarini Parvarishlashning Yopiq Sikli
* **Mantiqiy maqsadi:** Avtomatlashgan monitoring oqimi
* **Vizual turi:** process_cycle
* **Tarkibiy tezislar:**
  * 01. Multispektral Skanerlash: NDVI indeksi bo'yicha o'simliklarning xlorofill darajasi aniqlanadi.
  * 02. AI Tahlil va Xaritalash: Stress va kasallangan hududlarning aniq GPS koordinatalari olinadi.
  * 03. Differensial Purkash: Dronlar faqat zararlangan nuqtalarga pestitsid va o'g'it purkaydi.
  * 04. Hosildorlik Prognozi: Sensorlar va sun'iy yo'ldosh ma'lumotlari asosida hosil baholanadi.
"""
    },
    {
        "id": 12,
        "topic": "Kognitiv Psixologiya va Neyromarketing Qarorlari",
        "tested_modes": "three_pillars (3 ta Miya Tizimi)",
        "markdown": """
### 1-slayd: Kognitiv Psixologiya va Neyromarketing
* **Action-oriented sarlavha:** Neyromarketing Iste'molchilarning Qaror Qabul Qilish Mexanizmlarini Ochib Beradi
* **Mantiqiy maqsadi:** Kognitiv tahlil
* **Vizual turi:** cover

### 2-slayd: Neyro-Qarorlarning 3 Asosiy Ustuni
* **Action-oriented sarlavha:** Inson Miya Tizimlarining Xarid Jarayonidagi Rol va Funksiyalari
* **Mantiqiy maqsadi:** Uchlik model
* **Vizual turi:** three_pillars
* **Tarkibiy tezislar:**
  * Ratsional Neokorteks: Mantiqiy hisob-kitoblar, narx tahlili va mahsulot parametrlarini baholaydi.
  * Emotsional Limbik Tizim: Brend bilan hissiy bog'liqlik va vizual zavqlanish hissini shakllantiradi.
  * Instinktiv Reptil Miya: Asosiy xavfsizlik va zudlik bilan harakat qilish qarorlarini qabul qiladi.
"""
    },
    {
        "id": 13,
        "topic": "Koinot Astrofizikasi va Jeyms Uebb Teleskopi (JWST)",
        "tested_modes": "theory_concept (LaTeX, Qora fon)",
        "markdown": """
### 1-slayd: Koinot Astrofizikasi va JWST
* **Action-oriented sarlavha:** Jeyms Uebb Teleskopi Ilk Galaktikalar Shakllanishini Infratuzilma Orqali Kuzatadi
* **Mantiqiy maqsadi:** Kosmik astrofizika
* **Vizual turi:** cover

### 2-slayd: Spektral Siljish va Infratsil Qonuniyati
* **Action-oriented sarlavha:** Qizil Siljish Ko'rsatkichi va Koinotning Kengayish Qonuniyati
* **Mantiqiy maqsadi:** Spektral tahlil
* **Vizual turi:** theory_concept
* **Tarkibiy tezislar:**
  * Spektral Qizil Siljish: $z = \\frac{\\lambda_{\\text{obs}} - \\lambda_{\\text{emit}}}{\\lambda_{\\text{emit}}} \\ge 13.2$ qiymatiga erishildi.
  * Infratuzilma Diapazoni: $\\lambda = 0.6 - 28.3\\ \\mu\\text{m}$ to'lqin uzunligida koinot changlarini yorib o'tadi.
"""
    },
    {
        "id": 14,
        "topic": "Yashil Iqtisodiyot va Karbon Tashlamalarini Kamaytirish (ESG)",
        "tested_modes": "timeline_steps (2026 -> 2028 -> 2030)",
        "markdown": """
### 1-slayd: Yashil Iqtisodiyot va ESG
* **Action-oriented sarlavha:** Korxonalarda Karbon Neytralligiga O'tishning Xalqaro Standartlari
* **Mantiqiy maqsadi:** ESG boshqaruvi
* **Vizual turi:** cover

### 2-slayd: 2030-Yilgacha Karbon Kamaytirish Bosqichlari
* **Action-oriented sarlavha:** Parij Bitimi Doirasida Korporativ O'tish Rejasi
* **Mantiqiy maqsadi:** Vaqt chizig'i
* **Vizual turi:** timeline_steps
* **Tarkibiy tezislar:**
  * 2026-yil: Scope 1 va Scope 2 to'g'ridan-to'g'ri tashlamalarining to'liq raqamli auditi.
  * 2028-yil: Ishlab chiqarish energiyasining kamida 60% qismini qayta tiklanuvchi manbalarga o'tkazish.
  * 2030-yil: Net-Zero sertifikatiga erishish va barcha ta'minot zanjirlarida uglerod kompensatsiyasi.
"""
    },
    {
        "id": 15,
        "topic": "Nanotexnologiyalar: Saraton Hujayralariga Dori Yetkazish",
        "tested_modes": "diagram_anatomy (Liposoma, Nanozarrachalar)",
        "markdown": """
### 1-slayd: Nanotexnologiyalar va Onkologiya
* **Action-oriented sarlavha:** Maqsadli Nano-Kapsulalar Kimyoterapiya Toksikligini 80% Gacha Kamaytiradi
* **Mantiqiy maqsadi:** Dori yetkazish
* **Vizual turi:** cover

### 2-slayd: Nano-Yetkazish Tizimi Anatomiyasi
* **Action-oriented sarlavha:** Liposoma Kapsulasi va Maqsadli Ligandlar Tuzilishi
* **Mantiqiy maqsadi:** Nano-tuzilma
* **Vizual turi:** diagram_anatomy
* **Tarkibiy tezislar:**
  * Fosfolipid Liposoma: Dorini qon oqimida erib ketishdan asrovchi ikki qatlamli himoya qobig'i.
  * Monoklonal Antitanalar: Faqat o'sma hujayralari retseptorlariga birikuvchi aqlli yo'naltiruvchilar.
  * Magnit Nanozarrachalar: Tashqi magnit maydoni orqali o'sma o'chog'ida issiqlik ajratuvchi moddalar.
"""
    },
    {
        "id": 16,
        "topic": "Kiberurushlar va Xalqaro Gumanitar Huquq Yurisdiksiyasi",
        "tested_modes": "cards_grid (Tallin qo'llanmasi, Suverenitet)",
        "markdown": """
### 1-slayd: Kiberurushlar va Xalqaro Huquq
* **Action-oriented sarlavha:** Kiberhujumlar Davlat Suvereniteti va Jeneva Konvensiyalari Doirasida
* **Mantiqiy maqsadi:** Xalqaro huquq
* **Vizual turi:** cover

### 2-slayd: Xalqaro Huquqiy Mezonlar
* **Action-oriented sarlavha:** Kiberfazoda Qurolli Kuch Ishlatishni Regulyatsiya Qiluvchi Normalar
* **Mantiqiy maqsadi:** Huquqiy prinsiplar
* **Vizual turi:** cards_grid
* **Tarkibiy tezislar:**
  * Tallin Qo'llanmasi 2.0: Kiberoperatsiyalarda xalqaro gumanitar huquqni qo'llash bo'yicha xalqaro doktrina.
  * Davlat Suvereniteti: Tashqi serverlarga ruxsatsiz kirish suverenitetni buzish deb malakalanadi.
  * Hujum Atributsiyasi: Kiberhujumning haqiqiy muallifini qonuniy isbotlashning xalqaro standartlari.
"""
    },
    {
        "id": 17,
        "topic": "Katta Ma'lumotlar: Real Vaqt Rejimida Oqimli Qayta Ishlash",
        "tested_modes": "comparison_vs (Batch vs Stream Processing)",
        "markdown": """
### 1-slayd: Katta Ma'lumotlar va Oqimli Tizimlar
* **Action-oriented sarlavha:** Apache Flink va Kafka Yordamida Millisekundli Ma'lumotlar Oqimini Boshqarish
* **Mantiqiy maqsadi:** Oqimli qayta ishlash
* **Vizual turi:** cover

### 2-slayd: Batch vs Stream Processing Solishtirmasi
* **Action-oriented sarlavha:** Davriy Qayta Ishlash va Real Vaqt Rejimining Unumdorligi
* **Mantiqiy maqsadi:** Hadoop vs Flink
* **Vizual turi:** comparison_vs
* **Tarkibiy tezislar:**
  * Davriy (Batch) Processing: Katta hajmli arxiv ma'lumotlarini soatlar davomida qayta ishlaydi.
  * Oqimli (Stream) Processing: Tranzaksiyalarni ular sodir bo'lgan onda sub-sekund kechikishda tahlil qiladi.
"""
    },
    {
        "id": 18,
        "topic": "Sanoat Robototexnikasi va 6 Darajali Manipulyatorlar",
        "tested_modes": "timeline_steps (Kinematika -> Trayektoriya -> Nazorat)",
        "markdown": """
### 1-slayd: Sanoat Robototexnikasi
* **Action-oriented sarlavha:** 6 Erkinlik Darajasiga Ega Robotlar 0.02 mm Aniqlikda Operatsiyalarni Bajaradi
* **Mantiqiy maqsadi:** Robototexnika
* **Vizual turi:** cover

### 2-slayd: Robot Boshqaruvining 3 Bosqichi
* **Action-oriented sarlavha:** Matematik Modelidan Real Trayektoriyagacha Bo'lgan Boshqaruv
* **Mantiqiy maqsadi:** Kinematika zanjiri
* **Vizual turi:** timeline_steps
* **Tarkibiy tezislar:**
  * 1-Bosqich: To'g'ri va teskari kinematika tenglamalari orqali bo'g'in burchaklarini hisoblash.
  * 2-Bosqich: To'siqlardan qochuvchi silliq Spline trayektoriyalarini rejalashtirish.
  * 3-Bosqich: Yuqori tezlikda PID va adaptiv kuch nazorati orqali buyumni ushlash.
"""
    },
    {
        "id": 19,
        "topic": "Kvant Sensorlari va Gravitatsion To'lqinlar Deteksiyasi",
        "tested_modes": "theory_concept (LIGO, h ~ 10^-21)",
        "markdown": """
### 1-slayd: Kvant Sensorlari va Gravitatsion To'lqinlar
* **Action-oriented sarlavha:** Lazer Interferometriyasi Fazoviy Deformatsiyani Kvant Darajasida O'lchaydi
* **Mantiqiy maqsadi:** Gravitatsion to'lqinlar
* **Vizual turi:** cover

### 2-slayd: Gravitatsion To'lqinlar Deformatsiyasi
* **Action-oriented sarlavha:** LIGO Interferometrida Fazoviy Metrika O'zgarishi Formulalari
* **Mantiqiy maqsadi:** Deformatsiya o'lchami
* **Vizual turi:** theory_concept
* **Tarkibiy tezislar:**
  * Gravitatsion Deformatsiya: $h = \\frac{2 \\Delta L}{L} \\sim 10^{-21}$ proton diametridan 1000 barobar kichik.
  * Kvant Siqilgan Yorug'lik: Foton shovqinlarini kamaytirish orqali o'lchov aniqligini 50% ga oshiradi.
"""
    },
    {
        "id": 20,
        "topic": "Texnologik Startaplar Ekotizimi va Venchur Moliyalashtirish",
        "tested_modes": "pyramid_funnel (Pre-Seed -> Seed -> Series A -> B -> IPO)",
        "markdown": """
### 1-slayd: Texnologik Startaplar va Venchur
* **Action-oriented sarlavha:** Startaplar Rivojlanish Bosqichlari va Kapital Jalb Qilish Modeli
* **Mantiqiy maqsadi:** Venchur moliyasi
* **Vizual turi:** cover

### 2-slayd: Venchur Investitsiyalari Piramidasi
* **Action-oriented sarlavha:** G'oyadan Birja (IPO) Chiqishigacha Bo'lgan 4 Moliyaviy Bosqich
* **Mantiqiy maqsadi:** Moliyalashtirish qatlamlari
* **Vizual turi:** pyramid_funnel
* **Tarkibiy tezislar:**
  * 1-Qatlam (Yadro): Pre-Seed va Seed: Mahsulot prototipi (MVP) va dastlabki gipotezalarni tekshirish.
  * 2-Qatlam: Series A: Product-Market Fit ga erishish va sotuv kanallarini shakllantirish.
  * 3-Qatlam: Series B va C: Xalqaro bozorlarga chiqish va biznesni tezkor masshtablash.
  * 4-Qatlam (Cho'qqi): IPO va Strategik M&A: Ommaviy fond bozoriga chiqish va investorlar chiqishi.
"""
    }
]


def run_20_topics_stress_test(max_topics: Optional[int] = None):
    print("=" * 80)
    print("🔥 SLAYDHUB AI — 20 TA MAVZU VA BARCHA VIZUAL ARXETIPLAR STRESS-TESTI")
    print("=" * 80)

    matcher = get_academic_matcher()
    builder = get_academic_pptx_builder()

    results = []
    total_slides_count = 0
    t_start = time.time()

    preview_root = config.OUTPUT_ROOT / "preview_cache" / "stress_test_20"
    if preview_root.exists():
        shutil.rmtree(preview_root, ignore_errors=True)
    preview_root.mkdir(parents=True, exist_ok=True)

    topics_to_run = TOPICS_DATA[:max_topics] if max_topics else TOPICS_DATA

    for item in topics_to_run:
        t_id = item["id"]
        topic = item["topic"]
        tested_modes = item["tested_modes"]
        md_text = item["markdown"]

        print(f"\n[{t_id:02d}/20] 🧪 Sinov: '{topic}'")
        print(f"       🎯 Tekshirilayotgan rejimlar: {tested_modes}")

        t0 = time.time()
        try:
            # 1. Parse Markdown
            parsed = NotebookLMMarkdownParser.parse(md_text, topic=topic)
            assert len(parsed["slides"]) >= 2, f"Kamida 2 ta slayd kutilgan edi, topildi: {len(parsed['slides'])}"

            # 2. Build blueprint matching
            bp = matcher.build_blueprint(topic=topic, slide_count=len(parsed["slides"]), language="uz")

            # 3. Assemble Presentation
            out_pptx = builder.create_presentation(parsed, bp, engine_mode="harmonized")
            assert os.path.exists(out_pptx), "PPTX fayli diskda yaratilmadi!"
            size_kb = os.path.getsize(out_pptx) / 1024.0
            assert size_kb > 30.0, f"PPTX hajmi juda kichik: {size_kb:.1f} KB"

            # 4. Export Previews for visual audit
            previews = get_presentation_previews(out_pptx, max_slides=len(parsed["slides"]))
            elapsed = time.time() - t0
            total_slides_count += len(parsed["slides"])

            print(f"       ✅ PASS ({elapsed:.2f}s) — Slaydlar: {len(parsed['slides'])}, Hajmi: {size_kb:.1f}KB, Prevyular: {len(previews)} ta")
            results.append({
                "id": t_id,
                "topic": topic,
                "status": "PASS",
                "slides": len(parsed["slides"]),
                "size_kb": size_kb,
                "elapsed": elapsed,
                "out_pptx": out_pptx,
                "previews": previews
            })
        except Exception as e:
            elapsed = time.time() - t0
            print(f"       ❌ FAIL ({elapsed:.2f}s) — Xatolik: {e}")
            results.append({
                "id": t_id,
                "topic": topic,
                "status": "FAIL",
                "error": str(e),
                "elapsed": elapsed
            })

    total_time = time.time() - t_start
    passed = sum(1 for r in results if r["status"] == "PASS")
    failed = sum(1 for r in results if r["status"] == "FAIL")

    print("\n" + "=" * 80)
    print("📊 20 TA MAVZU STRESS-TEST NATIJALARI HISOBOTI")
    print("=" * 80)
    print(f"Jami Sinovdan O'tgan Mavzular:  {len(results)}")
    print(f"Muvaffaqiyatli (PASS):          {passed} / {len(results)} (100.0%)" if failed == 0 else f"Xatoliklar: {failed}")
    print(f"Jami Yaratilgan Slaydlar Soni:  {total_slides_count} ta")
    print(f"Umumiy Sarflangan Vaqt:        {total_time:.2f} soniya")
    print("=" * 80)

    return results


if __name__ == "__main__":
    run_20_topics_stress_test()
