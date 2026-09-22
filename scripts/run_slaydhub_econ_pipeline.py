# -*- coding: utf-8 -*-
"""
SlaydHub Bot - Iqtisodiyot va Moliya Slaydlari Konveyeri (ECON Pipeline)
=======================================================================
Ushbu skript SlaydHub bot (@SlaydHubUz_bot) dvigateli (TranslatorService, PPTXProcessor, GeminiTranslator)
orqali C:\\Users\\user\\Desktop\\Antigravity\\Soff.uz\\Slaydlar_ECON\\01_Sara_Slaydlar_ECON papkasidagi
barcha 225 ta saralangan taqdimotni O'zbek tili (Lotin)ga akademik standartlarda tarjima qiladi,
reklama va suvbelgilarni tozalaydi, matnlarni auto-fit qiladi va odamiylashtirilgan nomlar bilan
C:\\Users\\user\\Desktop\\Antigravity\\Soff.uz\\Tarjima_qilingan_slaydlar\\Slaydlar_ECON papkasiga saqlaydi.
"""
import os
import sys
import re
import json
import time
import shutil
import asyncio
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional
import xml.etree.ElementTree as ET
import zipfile

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

POWERPOINT_ROOT = Path(r"C:\Users\user\Desktop\Antigravity\Power Point")
SOFF_ROOT = Path(r"C:\Users\user\Desktop\Antigravity\Soff.uz")
for p in [str(POWERPOINT_ROOT), str(SOFF_ROOT)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from bot.services.translator_service import TranslatorService
from backend.core.gemini_translator import GeminiTranslator
from pptx import Presentation

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("SlaydHubEconPipeline")

SOURCE_DIR = SOFF_ROOT / "Slaydlar_ECON" / "01_Sara_Slaydlar_ECON"
TARGET_DIR = SOFF_ROOT / "Tarjima_qilingan_slaydlar" / "Slaydlar_ECON"
PROGRESS_FILE = TARGET_DIR / "batch_translation_progress.json"
TITLES_CACHE_FILE = TARGET_DIR / "titles_cache.json"

RUSSIAN_LATIN_PATTERNS = [
    r'\b(chto|chtob[yы]|eto|etot|etogo|dlya|vse|vsex|pri|ili|b[yы]t|b[yы]l|b[yы]la|b[yы]li|b[yы]lo)\b',
    r'\b(tsel|zadachi|issledovaniya|dann[yы]x|dann[yы]e|posle|mojno|nujno|yavlyaetsya|yavlyayutsya)\b',
    r'\b(cheloveka|sistem[yы]|urovne|uroven|ponyatie|klass|uroka|razvitiya|osnovn[yы]e)\b',
    r'\b(deyatelnost|predpriyatiya|organizatsii|pokazateli|stoimosti|proizvodstva)\b',
    r'\b(analiz|metod[yы]|printsip[yы]|formirovanie|kapital|resurs[yы]|upravlenie)\b'
]
UZBEK_MARKERS = [
    r"\b(va|bilan|uchun|ham|hamda|esa|lekin|chunki|bo['’`]?yicha|haqida|orqali)\b",
    r"\b(kerak|mumkin|bo['’`]?lgan|bo['’`]?lib|qilish|tizimi|asosiy|rivojlanish|maqsad|vazifalar)\b",
    r"\b(iqtisodiyot|moliya|korxona|boshqaruv|daromad|xarajat|tannarx|foyda|bank|kredit)\b"
]
ENGLISH_MARKERS = [
    r'\b(the|and|for|with|this|that|from|are|which|lecture|course|introduction|overview)\b',
    r'\b(system|systems|management|business|finance|economy|market|company)\b'
]

KNOWN_TOPIC_TRANSLATIONS = {
    "biznes-plan_predpriyatiya": "Korxona biznes rejasi",
    "ekonomika_organizatsii": "Tashkilot iqtisodiyoti",
    "ekonomika_organizatsii_0": "Tashkilot iqtisodiyoti va faoliyati",
    "finansy_i_denezhnye_fondy_predpriyatiya": "Korxonaning moliyasi va pul fondlari",
    "inflyatsiya": "Inflyatsiya va pul muomalasi",
    "kalkulyatsiya_sebestoimosti": "Mahsulot tannarxini kalkulyatsiya qilish",
    "kalkulyatsiya_sebestoimosti_produktsii": "Mahsulot tannarxini kalkulyatsiya qilish va xarajatlar",
    "menedzhment_swot_analiz1": "Menejmentda SWOT tahlili va rejalashtirish",
    "motivatsiya_personala": "Xodimlarni rag'batlantirish va motivatsiya tizimi",
    "oborotnye_fondy": "Korxonaning aylanma fondlari va mablag'lari",
    "personal_predpriyatiya": "Korxona xodimlari va kadrlar boshqaruvi",
    "piramida_potrebnostey": "Ehtiyojlar piramidasi va iste'mol xulqi",
    "pribyl_i_rentabelnost": "Korxona foydasi va rentabellik ko'rsatkichlari",
    "spetsifika_nauchnogo_issledovaniya_13_urok": "Iqtisodiy tadqiqotlar metodologiyasi",
    "vneshnyaya_i_vnutr.sreda_": "Tashkilotning tashqi va ichki muhiti",
    "rol_internet-bankinga_v_protsesse_kreditovaniya_fizicheskih_lits": "Jismoniy shaxslarni kreditlashda internet-bankingning o'rni",
    "tehnika_kross-prodazh": "Kross-sotuv (cross-selling) texnikasi va marketing",
    "01-vvedenie_v_menedzhment": "Menejment asoslariga kirish",
    "02-istoricheskie_tendentsii_shkoly_menedzhmenta": "Menejment maktablari va rivojlanish bosqichlari",
    "03_organizatsiya_kak_obekt_menedzh._vzaimodeystvie_chelov_i_organiz": "Tashkilot menejment obyekti sifatida",
    "04-tipy_struktur_upravleniya": "Boshqaruv tuzilmasi turlari",
    "05-vneshnyaya_i_vnutrennyaya_sreda_organizatsii": "Tashkilotning ichki va tashqi muhiti tahlili",
    "06_tsikl_menedzhmenta": "Menejment sikli va boshqaruv funksiyalari",
    "07planirovanie_v_menedzhm_1": "Menejmentda rejalashtirish jarayoni",
    "08_strategicheskoe_plnirovanie_2": "Strategik rejalashtirish va boshqaruv",
    "bankovskaya_sistema": "Bank tizimi va uning faoliyati",
    "denezhnoe_obrashchenie": "Pul muomalasi va uning qonuniyatlari",
    "izderzhki_proizvodstva": "Ishlab chiqarish xarajatlari va tannarx",
    "konkurentsiya": "Iqtisodiyotda raqobat va bozor mexanizmlari",
    "rynok_truda._bezrabotitsa": "Mehnat bozori va ishsizlik muammolari",
    "semeynyy_byudzhet": "Oila byudjeti va shaxsiy moliyaviy rejalashtirish",
    "svobodnaya_torgovlya._protekcionizm": "Erkin savdo va proteksionizm siyosati",
    "zarplata": "Mehnatga haq to'lash tizimlari va ish haqi"
}

def is_genuinely_uzbek(pptx_path: Path) -> bool:
    try:
        texts = []
        with zipfile.ZipFile(pptx_path, 'r') as z:
            for sname in sorted([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]):
                xml_data = z.read(sname)
                root = ET.fromstring(xml_data)
                for elem in root.iter():
                    if elem.tag.endswith('}t') and elem.text:
                        t = elem.text.strip()
                        if len(t) > 3:
                            texts.append(t)
        comb = ' '.join(texts).lower()
        if not comb.strip():
            return False
            
        uz_score = sum(len(re.findall(p, comb)) for p in UZBEK_MARKERS)
        ru_lat_score = sum(len(re.findall(p, comb)) for p in RUSSIAN_LATIN_PATTERNS)
        en_score = sum(len(re.findall(p, comb)) for p in ENGLISH_MARKERS)
        
        # Agar ruscha transliteratsiya 3 tadan ko'p bo'lsa yoki o'zbekchadan ko'p bo'lsa -> soxta
        if ru_lat_score >= 3 and ru_lat_score > uz_score:
            return False
        if ru_lat_score >= 5:
            return False
        # Agar inglizcha o'zbekchadan sezilarli ko'p bo'lsa -> tarjima qilinmagan
        if en_score >= 5 and en_score > uz_score:
            return False
        # Kamida 2 ta o'zbekcha grammatik marker bo'lishi kerak
        return uz_score >= 2
    except Exception:
        return False

class SlaydHubEconPipeline:
    def __init__(self):
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        self.progress = self._load_progress()
        self.titles_cache = self._load_titles_cache()
        self.translator_service = TranslatorService()
        self.translator_ai = GeminiTranslator()
        self.domain = "Iqtisodiyot, moliya, bank, menejment, marketing, biznes va buxgalteriya"

    def _load_progress(self) -> Dict[str, Any]:
        if PROGRESS_FILE.exists():
            try:
                with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_progress(self):
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.progress, f, ensure_ascii=False, indent=2)

    def _load_titles_cache(self) -> Dict[str, str]:
        if TITLES_CACHE_FILE.exists():
            try:
                with open(TITLES_CACHE_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _save_titles_cache(self):
        with open(TITLES_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(self.titles_cache, f, ensure_ascii=False, indent=2)

    def humanize_title(self, file_path: Path) -> str:
        stem = file_path.stem
        if stem in self.titles_cache:
            return self.titles_cache[stem]

        # Prefikslarni olib tashlash
        clean_stem = re.sub(r'^(Ekonomika_VUZ_|SPO_Econ2?_|VUZ_Econ2?_|VUZ_Econ_)', '', stem, flags=re.IGNORECASE)
        clean_stem = clean_stem.replace('__quot_', '').replace('_quot_', '').replace('quot_', '')
        clean_stem = clean_stem.replace('_', ' ').strip()
        clean_stem = re.sub(r'\s+', ' ', clean_stem)

        # Aniq lug'at bo'yicha tekshirish (uzunroq kalitlar birinchi tekshiriladi)
        sorted_known = sorted(KNOWN_TOPIC_TRANSLATIONS.items(), key=lambda x: len(x[0]), reverse=True)
        for k, v in sorted_known:
            if k.lower() in stem.lower() or k.lower() in clean_stem.lower():
                self.titles_cache[stem] = v
                self._save_titles_cache()
                return v

        # 1-slayddan sarlavhani o'qish
        s1_text = ""
        try:
            prs = Presentation(str(file_path))
            if len(prs.slides) > 0:
                p_texts = [p.text.strip() for sh in prs.slides[0].shapes if sh.has_text_frame for p in sh.text_frame.paragraphs if p.text.strip()]
                if p_texts:
                    s1_text = p_texts[0]
        except Exception:
            pass

        raw_query = f"{clean_stem} ({s1_text})" if s1_text and len(s1_text) < 60 and s1_text != clean_stem else clean_stem
        uz_title = self.translator_ai.translate_single_text(raw_query, target_script="latin")
        
        # Tozalash
        uz_title = re.sub(r'[/\\:*?"<>|_]', ' ', uz_title).strip()
        uz_title = re.sub(r'\s+', ' ', uz_title)
        if not uz_title or len(uz_title) < 3:
            uz_title = clean_stem.capitalize()

        self.titles_cache[stem] = uz_title
        self._save_titles_cache()
        return uz_title

    def is_already_done(self, file_path: Path) -> bool:
        fname = file_path.name
        if fname in self.progress and self.progress[fname].get("status") == "success":
            out_n = self.progress[fname].get("output_name")
            if out_n and (TARGET_DIR / out_n).exists() and (TARGET_DIR / out_n).stat().st_size > 1000:
                target_f = TARGET_DIR / out_n
                if is_genuinely_uzbek(target_f):
                    return True
                else:
                    logger.warning(f"⚠️ Qayta tarjima talab etiladi (soxta transliteratsiya): {target_f.name}")
        return False

    async def translate_one_file(self, file_path: Path) -> str:
        fname = file_path.name
        clean_title = self.humanize_title(file_path)
        out_filename = f"{clean_title}.pptx"
        output_path = TARGET_DIR / out_filename

        # Agar bu nomdagi fayl allaqachon boshqa manba fayli tomonidan yaratilgan bo'lsa
        if output_path.exists():
            orig_for_out = None
            for o_name, p_data in self.progress.items():
                if p_data.get("output_name") == out_filename and o_name != fname:
                    orig_for_out = o_name
                    break
            if orig_for_out:
                clean_title = f"{clean_title} 2"
                out_filename = f"{clean_title}.pptx"
                output_path = TARGET_DIR / out_filename

        logger.info(f"▶ Tarjima boshlanmoqda: {fname} -> {out_filename}")
        start_t = time.time()

        try:
            result = await self.translator_service.translate_presentation(
                input_pptx=str(file_path),
                output_pptx=str(output_path),
                target_script="latin",
                presentation_title=clean_title,
                domain=self.domain
            )

            if not result.get("success"):
                err = result.get("error", "Noma'lum xato")
                logger.error(f"  ❌ Tarjimada xato: {err}")
                self.progress[fname] = {"status": "failed", "error": err}
                self._save_progress()
                return "FAILED"

            # Qat'iy sifat va o'zbek tili tekshiruvi (barcha slaydlar bo'yicha)
            if not is_genuinely_uzbek(output_path):
                logger.error(f"  ❌ Sifat tekshiruvidan o'tmadi (transliteratsiya aniqlandi): {out_filename}")
                if output_path.exists():
                    output_path.unlink()
                self.progress[fname] = {"status": "failed", "error": "Transliteration detected, quality check failed"}
                self._save_progress()
                return "FAILED"

            elapsed = round(time.time() - start_t, 2)
            fin_s = result.get("final_slides", 0)
            init_s = result.get("initial_slides", 0)
            cleaned_ad = result.get("cleaned_ad_slides", 0)
            items_c = result.get("total_items", 0)

            logger.info(
                f"  ✅ MUVAFFAQITYATLI YAKUNLANDI ({elapsed} soniya)!\n"
                f"     📁 Fayl: {out_filename}\n"
                f"     📊 Slaydlar: {fin_s} ta (tozalangan reklama: {cleaned_ad} ta)\n"
                f"     📝 Matn bloklari: {items_c} ta"
            )

            self.progress[fname] = {
                "status": "success",
                "output_name": out_filename,
                "original_name": fname,
                "slides": fin_s,
                "items": items_c,
                "cleaned_ads": cleaned_ad,
                "time_sec": elapsed,
                "date": time.strftime("%Y-%m-%d %H:%M:%S")
            }
            self._save_progress()
            return "SUCCESS"

        except Exception as e:
            logger.error(f"  ❌ Kutilmagan xato ({fname}): {e}", exc_info=True)
            self.progress[fname] = {"status": "failed", "error": str(e)}
            self._save_progress()
            return "ERROR"

    def generate_catalog(self):
        logger.info("📚 Iqtisodiyot va moliya taqdimotlari katalogi yaratilmoqda...")
        catalog_path = TARGET_DIR / "KATALOG_TARJIMA_ECON.md"
        
        pptx_files = sorted([f for f in TARGET_DIR.glob("*.pptx") if not f.name.startswith("~$")])
        total_files = len(pptx_files)
        total_slides = 0
        total_size_bytes = 0

        # Invert progress map
        out_to_orig = {}
        for orig, data in self.progress.items():
            out_n = data.get("output_name")
            if out_n:
                out_to_orig[out_n] = {
                    "orig": orig,
                    "slides": data.get("slides", 0)
                }

        rows = []
        for idx, f in enumerate(pptx_files, 1):
            size_mb = f.stat().st_size / (1024 * 1024)
            total_size_bytes += f.stat().st_size
            
            info = out_to_orig.get(f.name, {})
            slides = info.get("slides", 0)
            if not slides:
                try:
                    with zipfile.ZipFile(f, 'r') as z:
                        slides = len([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])
                except Exception:
                    slides = 0
            total_slides += slides

            orig_name = info.get("orig", "-")
            clean_title = f.stem.replace('|', '/')
            orig_clean = orig_name.replace('|', '/')
            rows.append(f"| {idx} | **{clean_title}** | [{f.name}](./{f.name}) | {slides} ta | {size_mb:.2f} MB | `{orig_clean}` |")

        total_mb = total_size_bytes / (1024 * 1024)
        total_gb = total_mb / 1024

        md = f"""# SlaydHub.uz — Iqtisodiyot va Moliya Taqdimotlari (Slaydlar_ECON) Katalogi

Mazkur katalog `SlaydHubUz_bot` ning sun'iy intellekt tarjima va moslashtirish yadrosi orqali to'liq o'zbek tiliga (Lotin yozuvida) akademik standartlarda tarjima qilingan, barcha begona reklama va havolalari tozalangan, shrift va matnlari avtomatik moslashtirilgan iqtisodiy taqdimotlar to'plamini o'z ichiga oladi.

## Umumiy Statistika
- **Jami tarjima qilingan taqdimotlar:** {total_files} ta
- **Jami slaydlar soni:** {total_slides:,} ta
- **Umumiy fayl hajmi:** {total_mb:.2f} MB ({total_gb:.2f} GB)
- **Asosiy sohalar:** Buxgalteriya va audit, Moliya va bank ishi, Menejment va marketing, Mikro va makroiqtisodiyot, Biznes va tadbirkorlik, Xalqaro iqtisodiy munosabatlar, Ekonometrika va statistika.

---

## Taqdimotlar Jadvali

| № | Taqdimot Mavzusi (O'zbekcha) | Fayl Havolasi | Slaydlar | Hajmi | Asl Manba Fayli |
|---|---|---|---|---|---|
""" + "\n".join(rows) + f"""

---
*Katalog yaratilgan vaqt:* {time.strftime('%Y-%m-%d %H:%M:%S')}
*Tizim:* SlaydHub AI Bot Engine (TranslatorService & PPTXProcessor)
"""
        with open(catalog_path, "w", encoding="utf-8") as cf:
            cf.write(md)
        logger.info(f"✅ Katalog yaratildi: {catalog_path} ({total_files} ta taqdimot, {total_slides} ta slayd)")

    async def run(self, max_files: Optional[int] = None):
        logger.info("=" * 65)
        logger.info("🚀 SLAYDHUB BOT IQTISODIYOT SARA SLAYDLAR KONVEYERI")
        logger.info(f"📂 Manba papka: {SOURCE_DIR}")
        logger.info(f"📁 Qabul qiluvchi: {TARGET_DIR}")
        logger.info(f"🎓 Mutaxassislik: {self.domain}")
        logger.info("=" * 65)

        files = sorted([f for f in SOURCE_DIR.iterdir() if f.is_file() and f.suffix.lower() == ".pptx" and not f.name.startswith("~$")])

        all_pending: List[Path] = []
        already_done_count = 0

        for f in files:
            if self.is_already_done(f):
                already_done_count += 1
            else:
                all_pending.append(f)

        total_all = len(files)
        logger.info(f"📊 Jami sara slaydlar: {total_all} ta")
        logger.info(f"✅ Allaqachon tayyor: {already_done_count} ta")
        logger.info(f"⏳ Qayta ishlanishi kutilayotgan: {len(all_pending)} ta")

        if not all_pending:
            logger.info("🎉 Barcha sara slaydlar allaqachon muvaffaqiyatli tarjima qilingan!")
            self.generate_catalog()
            return

        limit = max_files if max_files is not None else len(all_pending)
        logger.info(f"🎯 Ushbu ishga tushirishda rejalashtirilgan: {min(limit, len(all_pending))} ta fayl\n")

        processed = 0
        success_count = 0
        failed_count = 0

        for idx, fpath in enumerate(all_pending, start=1):
            if processed >= limit:
                logger.info(f"Belgilangan hajm ({limit} ta fayl) bajarildi.")
                break

            logger.info(f"\n--- [{idx}/{len(all_pending)}] (Jami bajarilgan: {already_done_count + processed}/{total_all}) ---")
            res = await self.translate_one_file(fpath)
            processed += 1
            if res == "SUCCESS":
                success_count += 1
            else:
                failed_count += 1

            if processed < limit:
                time.sleep(2)

        self.generate_catalog()

        logger.info("\n" + "=" * 65)
        logger.info("🏁 KONVEYER ISHI YAKUNLANDI")
        logger.info(f"✨ Muvaffaqiyatli: {success_count} ta")
        logger.info(f"⚠️ Xatolik: {failed_count} ta")
        logger.info(f"📊 Jami hozirgi tayyor: {already_done_count + success_count}/{total_all}")
        logger.info("=" * 65)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="SlaydHub Bot Economics Slides Pipeline")
    parser.add_argument("--count", "-c", type=int, default=None, help="Maksimal ishlanadigan fayllar soni")
    args = parser.parse_args()

    pipeline = SlaydHubEconPipeline()
    asyncio.run(pipeline.run(max_files=args.count))
