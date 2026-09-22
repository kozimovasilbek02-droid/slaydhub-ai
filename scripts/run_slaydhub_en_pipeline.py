# -*- coding: utf-8 -*-
"""
SlaydHub Bot - Inglizcha Sara Slaydlar Konveyeri (Pipeline)
===========================================================
Ushbu skript SlaydHub bot (@SlaydHubUz_bot) dvigateli (TranslatorService, PPTXProcessor, GeminiTranslator)
orqali C:\\Users\\user\\Desktop\\Antigravity\\Soff.uz\\Slaydlar_EN\\01_Sara_Slaydlar_EN papkasidagi barcha
saralangan taqdimotlarni to'liq O'zbek tili (Lotin)ga tarjima qiladi, suvbelgi va reklamalarni tozalaydi,
auto-fit bilan ramkalarga moslaydi va C:\\Users\\user\\Desktop\\Antigravity\\Soff.uz\\Tarjima_qilingan_slaydlar\\Slaydlar_EN
papkasiga odamiylashtirilgan nomlar bilan joylaydi.
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
from typing import Dict, Any, List, Optional, Tuple

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

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("SlaydHubEnPipeline")

SOURCE_DIR = SOFF_ROOT / "Slaydlar_EN" / "01_Sara_Slaydlar_EN"
TARGET_DIR = SOFF_ROOT / "Tarjima_qilingan_slaydlar" / "Slaydlar_EN"
PROGRESS_FILE = TARGET_DIR / "batch_translation_progress.json"

PREFIX_TRANSLATIONS = {
    r"^AntennaComms_Lecture(\d+)": r"Simsiz aloqa va ko'p antennali tizimlar (\1-ma'ruza)",
    r"^QuantumComputing_Lecture0?(\d+)": r"Kvant hisoblash va kvant algoritmlari (\1-ma'ruza)",
    r"^CompilerTech_0?(\d+)[-_]?(.*)": r"Kompilyatorlar texnologiyasi va LLVM (\1-qism \2)",
    r"^CyberSecurity_Lecture_?(\d+)[-_]Slides.*": r"Kiberxavfsizlik asoslari (\1-ma'ruza)",
    r"^DataScience_Data_Science_and_Python_-_Lec0?(\d+)": r"Python orqali ma'lumotlar fani (\1-ma'ruza)",
    r"^CompNetworks_CN_Lecture_?(\d+)[-_]?(.*)": r"Kompyuter tarmoqlari va IP arxitekturasi (\1-ma'ruza \2)",
    r"^CompNetworks_Chapter_(\d+).*": r"Kompyuter tarmoqlari (\1-bob)",
    r"^CompNetworks_Network_Layer": r"Kompyuter tarmoqlari - Tarmoq pog'onasi",
    r"^DatabaseSystems_lec0?(\d+)": r"Ma'lumotlar bazasi tizimlari va arxitekturasi (\1-ma'ruza)",
    r"^EN_CS_0?(\d+)[-_]?(.*)": r"C++ dasturlash va dasturiy muhandislik (\1-ma'ruza \2)",
    r"^GaTech_CV_0?(\d+)": r"Kompyuter ko'rishi va tasvirlarni tahlil qilish (\1-ma'ruza)",
    r"^GaTech_7491_(.*)": r"Grafik hisoblash va animatsiya - \1",
    r"^CUDA_GPU_(.*)": r"CUDA va yuqori samarali GPU hisoblash - \1",
    r"^CMU_15781_(.*)": r"Sun'iy intellekt va mashinali o'rganish - \1",
    r"^EN_CMU_DFT_(.*)": r"Kvant materialshunoslik va DFT tahlili - \1",
    r"^SPO_Transport_(.*)": r"Avtomobil transporti va dvigatellar tizimi - \1",
    r"^SPO_Science_(.*)": r"Amaliy va tabiiy fanlar - \1",
    r"^SPO_Humanities_(.*)": r"Ijtimoiy-gumanitar fanlar - \1",
    r"^SPO_Arts_(.*)": r"San'at va arxitektura tarixi - \1",
    r"^EN_Cloud_(.*)": r"Bulutli hisoblash va mikroxizmatlar - \1",
    r"^EN_Research_(.*)": r"Ilmiy tadqiqot va akademik taqdimot - \1",
}

def humanize_en_filename(stem: str) -> str:
    for pat, repl in PREFIX_TRANSLATIONS.items():
        if re.search(pat, stem, re.IGNORECASE):
            res = re.sub(pat, repl, stem, flags=re.IGNORECASE)
            res = res.replace("_", " ").strip()
            res = re.sub(r"\s+", " ", res)
            return res
    t = stem.replace("_", " ").strip()
    t = re.sub(r"\s+", " ", t)
    return t

RUSSIAN_LATIN_PATTERNS = [
    r'\b(chto|chtob[yы]|eto|etot|etogo|dlya|vse|vsex|pri|ili|b[yы]t|b[yы]l|b[yы]la|b[yы]li|b[yы]lo)\b',
    r'\b(tsel|zadachi|issledovaniya|dann[yы]x|dann[yы]e|posle|mojno|nujno|yavlyaetsya|yavlyayutsya)\b',
    r'\b(cheloveka|sistem[yы]|urovne|uroven|ponyatie|klass|uroka|razvitiya|osnovn[yы]e)\b'
]
UZBEK_MARKERS = [
    r"\b(va|bilan|uchun|ham|hamda|esa|lekin|chunki|bo['’`]?yicha|haqida|orqali)\b",
    r"\b(kerak|mumkin|bo['’`]?lgan|bo['’`]?lib|qilish|tizimi|asosiy|rivojlanish|maqsad|vazifalar)\b"
]
ENGLISH_MARKERS = [
    r'\b(the|and|for|with|this|that|from|are|which|lecture|course|introduction|overview)\b',
    r'\b(system|systems|computer|algorithm|data|analysis|network|design)\b'
]

def is_genuinely_uzbek(pptx_path: Path) -> bool:
    try:
        import zipfile
        import xml.etree.ElementTree as ET
        texts = []
        with zipfile.ZipFile(pptx_path, 'r') as z:
            for sname in sorted([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')])[:6]:
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
        
        # Agar ruscha transliteratsiya o'zbekchadan ko'p bo'lsa -> soxta
        if ru_lat_score >= 3 and ru_lat_score > uz_score:
            return False
        # Agar inglizcha o'zbekchadan sezilarli ko'p bo'lsa -> tarjima qilinmagan
        if en_score >= 5 and en_score > uz_score:
            return False
        # Kamida 2 ta o'zbekcha grammatik marker bo'lishi kerak
        return uz_score >= 2
    except Exception:
        return False

class SlaydHubEnPipeline:
    def __init__(self):
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        self.progress = self._load_progress()
        self.translator_service = TranslatorService()
        self.domain = "Axborot texnologiyalari, IT, fan, muhandislik va oliy ta'lim"

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

    def is_already_done(self, file_path: Path) -> bool:
        fname = file_path.name
        fstem = file_path.stem
        clean_name = humanize_en_filename(fstem)

        target_f1 = TARGET_DIR / f"{clean_name}.pptx"
        target_f2 = TARGET_DIR / f"{fstem}.pptx"

        for cand in [target_f1, target_f2]:
            if cand.exists() and cand.stat().st_size > 1000:
                if is_genuinely_uzbek(cand):
                    return True
                else:
                    logger.warning(f"⚠️ Qayta tarjima talab etiladi (soxta transliteratsiya yoki tarjima qilinmagan): {cand.name}")

        return False

    async def translate_one_file(self, file_path: Path) -> str:
        fname = file_path.name
        fstem = file_path.stem
        clean_title = humanize_en_filename(fstem)
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

    async def run(self, max_files: Optional[int] = None):
        logger.info("=" * 65)
        logger.info("🚀 SLAYDHUB BOT INGLIZCHA SARA SLAYDLAR KONVEYERI")
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
                time.sleep(3)

        logger.info("\n" + "=" * 65)
        logger.info("🏁 KONVEYER ISHI YAKUNLANDI")
        logger.info(f"✨ Muvaffaqiyatli: {success_count} ta")
        logger.info(f"⚠️ Xatolik: {failed_count} ta")
        logger.info(f"📊 Jami hozirgi tayyor: {already_done_count + success_count}/{total_all}")
        logger.info("=" * 65)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="SlaydHub Bot English Slides Pipeline")
    parser.add_argument("--count", "-c", type=int, default=None, help="Maksimal ishlanadigan fayllar soni")
    args = parser.parse_args()

    pipeline = SlaydHubEnPipeline()
    asyncio.run(pipeline.run(max_files=args.count))
