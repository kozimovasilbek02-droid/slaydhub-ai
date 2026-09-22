# -*- coding: utf-8 -*-
"""
SlaydHub Bot - Tibbiyot Slaydlari Konveyeri (Pipeline)
=====================================================
Ushbu skript SlaydHub bot (@SlaydHubUz_bot) dvigateli (TranslatorService, PPTXProcessor, GeminiTranslator)
orqali C:\\Users\\user\\Desktop\\Antigravity\\Soff.uz\\Slaydlar_MED papkasidagi barcha tibbiy slaydlarni
kelgan joyidan davom ettirib, to'liq O'zbek tiliga (Lotin) tarjima qiladi, reklamalarni tozalaydi,
auto-fit qiladi va C:\\Users\\user\\Desktop\\Antigravity\\Soff.uz\\Tarjima_qilingan_slaydlar\\Slaydlar_MED
papkasiga kurslar iyerarxiyasini saqlagan holda joylaydi.
"""
import os
import sys
import json
import time
import shutil
import asyncio
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

# UTF-8 encoding
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Add Power Point workspace to path
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
logger = logging.getLogger("SlaydHubPipeline")

SOURCE_DIR = SOFF_ROOT / "Slaydlar_MED"
TARGET_DIR = SOFF_ROOT / "Tarjima_qilingan_slaydlar" / "Slaydlar_MED"
TEMP_CONV_DIR = TARGET_DIR / "_temp_conversions"
PROGRESS_FILE = TARGET_DIR / "batch_translation_progress.json"
MAP_FILE = Path(r"C:\Users\user\.gemini\antigravity\brain\29a95430-d6c6-4f1b-964d-b114848b3619\scratch\med_final_rename_map.json")

def load_rename_map() -> Dict[str, str]:
    if MAP_FILE.exists():
        try:
            with open(MAP_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def convert_ppt_to_pptx(ppt_path: Path, temp_dir: Path) -> Optional[Path]:
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        abs_ppt = str(ppt_path.resolve())
        out_pptx = temp_dir / f"{ppt_path.stem}_converted.pptx"
        abs_out = str(out_pptx.resolve())
        
        pres = ppt_app.Presentations.Open(abs_ppt, WithWindow=False)
        pres.SaveAs(abs_out, 24)  # 24 = ppSaveAsOpenXMLPresentation
        pres.Close()
        return out_pptx
    except Exception as e:
        logger.warning(f"PPT -> PPTX xatolik ({ppt_path.name}): {e}")
        return None

class SlaydHubMedPipeline:
    def __init__(self):
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
        TEMP_CONV_DIR.mkdir(parents=True, exist_ok=True)
        self.progress = self._load_progress()
        self.old_to_new = load_rename_map()
        self.new_to_old = {v: k for k, v in self.old_to_new.items()}
        self.translator_service = TranslatorService()
        self.domain = "Tibbiyot, anatomiya, farmatsevtika va sog'liqni saqlash"

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

    def is_already_done(self, course_name: str, file_path: Path) -> bool:
        target_course_dir = TARGET_DIR / course_name
        fname = file_path.name
        fstem = file_path.stem
        old_name = self.new_to_old.get(fname, fname)

        # 1. Diskda tayyor fayl bormi?
        if target_course_dir.exists():
            for existing in target_course_dir.iterdir():
                if not existing.is_file():
                    continue
                if existing.name.lower() == f"{fstem}.pptx".lower() or existing.name.lower() == fname.lower():
                    if existing.stat().st_size > 1000:
                        return True
                if old_name != fname:
                    old_stem = Path(old_name).stem
                    if existing.name.lower() == f"{old_stem}.pptx".lower() or existing.name.lower() == old_name.lower():
                        if existing.stat().st_size > 1000:
                            return True

        # 2. Progress faylda muvaffaqiyatli deb belgilanganmi?
        check_keys = [
            f"{course_name}/{old_name}".replace("\\", "/"),
            f"{course_name}/{fname}".replace("\\", "/"),
            old_name,
            fname
        ]
        for k in check_keys:
            if k in self.progress and self.progress[k].get("status") == "success":
                out_n = self.progress[k].get("output_name")
                if out_n and (target_course_dir / out_n).exists():
                    return True

        return False

    async def translate_one_file(self, course_name: str, file_path: Path) -> str:
        fname = file_path.name
        target_course_dir = TARGET_DIR / course_name
        target_course_dir.mkdir(parents=True, exist_ok=True)
        
        # Chiqish fayli nomi: Toza O'zbekcha nom
        out_filename = f"{file_path.stem}.pptx"
        output_path = target_course_dir / out_filename
        
        logger.info(f"▶ [{course_name}] Tarjima boshlanmoqda: {fname}")
        start_t = time.time()
        
        working_file = file_path
        is_temp = False
        
        if file_path.suffix.lower() == ".ppt":
            logger.info("  🔄 .ppt format PowerPoint orqali .pptx ga o'tkazilmoqda...")
            conv = convert_ppt_to_pptx(file_path, TEMP_CONV_DIR)
            if not conv or not conv.exists():
                logger.error(f"  ❌ .ppt konvertatsiya qilib bo'lmadi: {fname}")
                rel_key = f"{course_name}/{fname}".replace("\\", "/")
                self.progress[rel_key] = {"status": "failed", "error": "PPT conversion failed"}
                self._save_progress()
                return "FAILED_CONVERT"
            working_file = conv
            is_temp = True
            
        try:
            # SlaydHub bot TranslatorService orqali tarjima qilish
            result = await self.translator_service.translate_presentation(
                input_pptx=str(working_file),
                output_pptx=str(output_path),
                target_script="latin",
                presentation_title=file_path.stem,
                domain=self.domain
            )
            
            if not result.get("success"):
                err = result.get("error", "Noma'lum xato")
                logger.error(f"  ❌ Tarjimada xato: {err}")
                rel_key = f"{course_name}/{fname}".replace("\\", "/")
                self.progress[rel_key] = {"status": "failed", "error": err}
                self._save_progress()
                return "FAILED_TRANSLATION"
                
            elapsed = round(time.time() - start_t, 2)
            init_s = result.get("initial_slides", 0)
            fin_s = result.get("final_slides", 0)
            cleaned_ad = result.get("cleaned_ad_slides", 0)
            items_c = result.get("total_items", 0)
            
            logger.info(
                f"  ✅ MUVAFFAQITYATLI YAKUNLANDI ({elapsed} soniya)!\n"
                f"     📁 Fayl: {course_name}/{out_filename}\n"
                f"     📊 Slaydlar: {fin_s} ta (dastlabki {init_s} ta, tozalangan reklama: {cleaned_ad} ta)\n"
                f"     📝 Matn bloklari: {items_c} ta"
            )
            
            rel_key = f"{course_name}/{fname}".replace("\\", "/")
            self.progress[rel_key] = {
                "status": "success",
                "output_name": out_filename,
                "rel_subfolder": course_name,
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
            rel_key = f"{course_name}/{fname}".replace("\\", "/")
            self.progress[rel_key] = {"status": "failed", "error": str(e)}
            self._save_progress()
            return "ERROR"
        finally:
            if is_temp and working_file.exists():
                try:
                    working_file.unlink()
                except Exception:
                    pass

    async def run(self, max_files: Optional[int] = None):
        logger.info("=" * 65)
        logger.info("🚀 SLAYDHUB BOT TIBBIY SLAYDLAR KONVEYERI ISHGA TUSHDI")
        logger.info(f"📂 Manba papka: {SOURCE_DIR}")
        logger.info(f"📁 Qabul qiluvchi papka: {TARGET_DIR}")
        logger.info(f"🩺 Mutaxassislik sohasi: {self.domain}")
        logger.info("=" * 65)
        
        course_dirs = sorted([d for d in SOURCE_DIR.iterdir() if d.is_dir() and not d.name.startswith((".", "_"))])
        
        all_pending: List[Tuple[str, Path]] = []
        already_done_count = 0
        
        for cdir in course_dirs:
            cname = cdir.name
            files = sorted([f for f in cdir.iterdir() if f.suffix.lower() in (".pptx", ".ppt") and not f.name.startswith("~$")])
            for f in files:
                if self.is_already_done(cname, f):
                    already_done_count += 1
                else:
                    all_pending.append((cname, f))
                    
        total_all = already_done_count + len(all_pending)
        logger.info(f"📊 Jami slaydlar: {total_all} ta")
        logger.info(f"✅ Allaqachon tayyor bo'lganlar: {already_done_count} ta")
        logger.info(f"⏳ Qayta ishlanishi kutilayotganlar: {len(all_pending)} ta")
        
        if not all_pending:
            logger.info("🎉 Barcha slaydlar allaqachon muvaffaqiyatli tarjima qilingan!")
            return
            
        limit = max_files if max_files is not None else len(all_pending)
        logger.info(f"🎯 Ushbu ishga tushirishda rejalashtirilgan: {min(limit, len(all_pending))} ta fayl\n")
        
        processed = 0
        success_count = 0
        failed_count = 0
        
        for idx, (cname, fpath) in enumerate(all_pending, start=1):
            if processed >= limit:
                logger.info(f"Belgilangan hajm ({limit} ta fayl) bajarildi.")
                break
                
            logger.info(f"\n--- [{idx}/{len(all_pending)}] (Jami bajarilgan: {already_done_count + processed}/{total_all}) ---")
            res = await self.translate_one_file(cname, fpath)
            processed += 1
            if res == "SUCCESS":
                success_count += 1
            else:
                failed_count += 1
                
            # Rate limitga tushmaslik uchun qisqa tanaffus
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
    parser = argparse.ArgumentParser(description="SlaydHub Bot Medical Presentation Pipeline")
    parser.add_argument("--count", "-c", type=int, default=None, help="Maksimal ishlanadigan fayllar soni")
    args = parser.parse_args()
    
    pipeline = SlaydHubMedPipeline()
    asyncio.run(pipeline.run(max_files=args.count))
