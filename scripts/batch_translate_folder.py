# -*- coding: utf-8 -*-
"""
scripts.batch_translate_folder
------------------------------
Papka va barcha quyi papkalar ichidagi PPTX va PPT taqdimot fayllarini
SlideTranslate AI dvigateli orqali O'zbek tiliga (Lotin) to'liq tarjima qiluvchi,
suvbelgi va reklamalarni tozalovchi, rasmlarni adaptiv inpainting bilan o'giruvchi,
fayl nomlarini o'zbekchalashtiruvchi va papkalar iyerarxiyasini to'liq saqlovchi
avtomatlashtirilgan ommaviy tarjima tizimi.
"""
from __future__ import annotations

import os
import sys
import re
import json
import time
import shutil
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional

# UTF-8 stdout configuration for Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Add project root to sys.path
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from backend.core.gemini_translator import GeminiTranslator
from backend.core.pptx_processor import PPTXProcessor
from backend.core.transliteration import ensure_script

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("BatchTranslator")

PREFIX_MAP = {
    r"^Pravo_VUZ_?": "Huquq_OTM_",
    r"^Psikh_VUZ_?": "Psixologiya_OTM_",
    r"^Sotsiologiya_VUZ_?": "Sotsiologiya_OTM_",
    r"^Tekhnika_VUZ_?": "Texnika_OTM_",
    r"^Pedagogika_VUZ_?": "Pedagogika_OTM_",
    r"^Ekonomika_VUZ_?": "Iqtisodiyot_OTM_",
    r"^Meditsina_VUZ_?": "Tibbiyot_OTM_",
    r"^Biologiya_VUZ_?": "Biologiya_OTM_",
    r"^Fizika_VUZ_?": "Fizika_OTM_",
    r"^Kimyo_VUZ_?": "Kimyo_OTM_",
    r"^Khimiya_VUZ_?": "Kimyo_OTM_",
    r"^Istoriya_VUZ_?": "Tarix_OTM_",
    r"^Tarix_VUZ_?": "Tarix_OTM_",
    r"^Filosofiya_VUZ_?": "Falsafa_OTM_",
    r"^Ekologiya_VUZ_?": "Ekologiya_OTM_",
    r"^Geografiya_VUZ_?": "Geografiya_OTM_",
    r"^Informatika_VUZ_?": "Informatika_OTM_",
    r"^Matematika_VUZ_?": "Matematika_OTM_",
    r"^Dizayn_VUZ_?": "Dizayn_OTM_",
    r"^Arxitektura_VUZ_?": "Arxitektura_OTM_",
}

def translate_filename(raw_filename: str, translator: GeminiTranslator, cache: Dict[str, str]) -> str:
    """Fayl nomini o'zbek tiliga (Lotin) toza va tushunarli tarjima qiladi."""
    if raw_filename in cache:
        return cache[raw_filename]
        
    stem = Path(raw_filename).stem
    ext = ".pptx"
    
    prefix_uz = ""
    rest_stem = stem
    for pat, uz_p in PREFIX_MAP.items():
        if re.search(pat, stem, re.IGNORECASE):
            prefix_uz = uz_p
            rest_stem = re.sub(pat, "", stem, flags=re.IGNORECASE)
            break
            
    rest_stem = re.sub(r'_wecompress\.com_', '', rest_stem, flags=re.IGNORECASE)
    rest_stem = re.sub(r'[\(\[\{]\d+[\)\]\}]', '', rest_stem)
    rest_stem = re.sub(r'[-_]?(kopi[jy]a|copy|szhatyy|compressed)', '', rest_stem, flags=re.IGNORECASE)
    
    words_to_translate = rest_stem.replace('_', ' ').strip()
    
    if not words_to_translate:
        translated_stem = "Taqdimot"
    else:
        try:
            prompt = (
                "Ushbu ruscha yoki translit qilingan taqdimot mavzusini O'zbek tili (Lotin yozuvi)ga "
                "lo'nda va toza tarjima qiling (faqat tarjima qilingan nomni qaytaring, tushuntirishsiz):\n\n"
                f'"{words_to_translate}"'
            )
            res = translator.client.models.generate_content(
                model=translator.model_name,
                contents=prompt
            )
            trans = res.text.strip().strip('"\'').strip()
            trans = re.sub(r'[/\\:*?"<>|]', '', trans)
            trans = trans.replace("‘", "'").replace("’", "'").replace("`", "'")
            words = [w.strip() for w in trans.split() if w.strip()]
            translated_stem = "_".join(words)
        except Exception:
            translated_stem = rest_stem.strip('_')
            
    final_name = f"{prefix_uz}{translated_stem}{ext}"
    final_name = re.sub(r'_+', '_', final_name).strip('_')
    if not final_name.endswith('.pptx'):
        final_name += '.pptx'
        
    cache[raw_filename] = final_name
    return final_name


def convert_ppt_to_pptx(ppt_path: Path, temp_dir: Path) -> Optional[Path]:
    """Eski .ppt faylini PowerPoint COM orqali vaqtinchalik .pptx formatiga o'tkazadi."""
    try:
        import win32com.client
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        abs_ppt = str(ppt_path.resolve())
        out_pptx = temp_dir / f"{ppt_path.stem}_converted.pptx"
        abs_out = str(out_pptx.resolve())
        
        pres = ppt_app.Presentations.Open(abs_ppt, WithWindow=False)
        pres.SaveAs(abs_out, 24)  # 24 = ppSaveAsOpenXMLPresentation (.pptx)
        pres.Close()
        return out_pptx
    except Exception as e:
        logger.warning(f"PPT -> PPTX konvertatsiyasida xato ({ppt_path.name}): {e}")
        return None


def run_batch_translation(
    source_dir: str = r"C:\Users\user\Desktop\Antigravity\Soff.uz\Slaydlar_RU",
    target_parent_dir: str = r"C:\Users\user\Desktop\Antigravity\Soff.uz\Tarjima_qilingan_slaydlar",
    subfolder_name: str = "Slaydlar_RU",
    target_script: str = "latin"
):
    source_path = Path(source_dir)
    if not source_path.exists():
        logger.error(f"Manba katalogi topilmadi: {source_dir}")
        return

    target_root = Path(target_parent_dir) / subfolder_name
    target_root.mkdir(parents=True, exist_ok=True)
    temp_dir = target_root / "_temp_conversions"
    temp_dir.mkdir(parents=True, exist_ok=True)

    progress_file = target_root / "batch_translation_progress.json"
    cache_file = target_root / "filename_cache.json"

    # Load progress & cache
    progress_data: Dict[str, Any] = {}
    if progress_file.exists():
        try:
            progress_data = json.loads(progress_file.read_text(encoding="utf-8"))
        except Exception:
            progress_data = {}

    filename_cache: Dict[str, str] = {}
    if cache_file.exists():
        try:
            filename_cache = json.loads(cache_file.read_text(encoding="utf-8"))
        except Exception:
            filename_cache = {}

    translator = GeminiTranslator()

    # Barcha taqdimot fayllarini quyi papkalar bilan birga rekursiv yig'ish
    all_files: List[Path] = []
    for ext in [".pptx", ".ppt"]:
        all_files.extend(list(source_path.rglob(f"*{ext}")))

    all_files = sorted(list(set(all_files)), key=lambda x: str(x.relative_to(source_path)))
    total_files = len(all_files)

    logger.info(f"📂 Jami topilgan fayllar: {total_files} ta ({source_dir})")
    logger.info(f"📁 Natijalar saqlanadigan katalog: {target_root}")

    completed_count = sum(1 for v in progress_data.values() if isinstance(v, dict) and v.get("status") == "success")
    logger.info(f"🔄 Oldingi sessiyalardan yakunlangan: {completed_count}/{total_files}")

    start_total_time = time.time()

    for idx, file_path in enumerate(all_files, start=1):
        if not file_path.exists():
            continue

        rel_path = file_path.relative_to(source_path)
        rel_key = str(rel_path).replace("\\", "/")
        raw_name = file_path.name

        # Quyi papkani aniqlash
        sub_dir_parts = rel_path.parent
        target_sub_dir = target_root / sub_dir_parts
        target_sub_dir.mkdir(parents=True, exist_ok=True)

        # Allaqachon muvaffaqiyatli tarjima qilingan bo'lsa tekshirish
        if rel_key in progress_data and progress_data[rel_key].get("status") == "success":
            out_file_name = progress_data[rel_key].get("output_name")
            if out_file_name and (target_sub_dir / out_file_name).exists():
                logger.info(f"[{idx}/{total_files}] ⏩ O'tkazib yuborildi (avval tarjima qilingan): {rel_key} -> {out_file_name}")
                continue

        logger.info(f"\n[{idx}/{total_files}] ({((idx-1)/total_files)*100:.1f}%) 🚀 Ishlanmoqda: {rel_key}")
        file_start_time = time.time()

        # 1. Fayl nomini tarjima qilish
        translated_name = translate_filename(raw_name, translator, filename_cache)
        cache_file.write_text(json.dumps(filename_cache, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info(f"  📝 O'zbekcha nomi: {translated_name}")

        output_pptx_path = target_sub_dir / translated_name

        # 2. Faylni tayyorlash (.ppt bo'lsa .pptx ga o'tkazish)
        working_pptx_path = file_path
        is_temp_converted = False
        if file_path.suffix.lower() == ".ppt":
            logger.info("  ⚙️ .ppt fayli .pptx formatiga konvertatsiya qilinmoqda...")
            converted_path = convert_ppt_to_pptx(file_path, temp_dir)
            if not converted_path or not converted_path.exists():
                logger.error(f"  ❌ .ppt faylini ochib bo'lmadi: {rel_key}")
                progress_data[rel_key] = {"status": "failed", "error": "PPT conversion failed"}
                progress_file.write_text(json.dumps(progress_data, ensure_ascii=False, indent=2), encoding="utf-8")
                continue
            working_pptx_path = converted_path
            is_temp_converted = True

        try:
            # 3. Matnlarni ajratish
            extracted = PPTXProcessor.extract_presentation_data(str(working_pptx_path))
            slides = extracted.get("slides", [])
            all_items = []
            for s in slides:
                all_items.extend(s.get("items", []))

            total_items = len(all_items)
            total_slides = len(slides)
            logger.info(f"  📊 Slaydlar: {total_slides} ta, Matn bloklari: {total_items} ta")

            if total_items == 0:
                logger.warning("  ⚠️ Matn topilmadi, to'g'ridan-to'g'ri nusxa olinmoqda...")
                shutil.copy(str(working_pptx_path), str(output_pptx_path))
                progress_data[rel_key] = {
                    "status": "success",
                    "output_name": translated_name,
                    "rel_subfolder": str(sub_dir_parts),
                    "slides": total_slides,
                    "items": 0,
                    "time_sec": round(time.time() - file_start_time, 2)
                }
                progress_file.write_text(json.dumps(progress_data, ensure_ascii=False, indent=2), encoding="utf-8")
                continue

            # 4. Parallel batch tarjima
            stats: Dict[str, Any] = {}
            logger.info("  🌐 Matnlar Gemini AI orqali parallel tarjima qilinmoqda...")
            translated_items = translator.translate_items_batch(
                items=all_items,
                target_script=target_script,
                domain="Taqdimot, fan, ta'lim va ilmiy tahlil",
                stats=stats,
                batch_size=35,
                max_workers=4
            )

            translations_map = {
                it["id"]: it.get("translated_text") or it.get("text", "")
                for it in translated_items
            }

            # 5. Taqdimotga qo'llash, reklamalarni tozalash va rasmlarni tarjima qilish
            logger.info("  🎨 Matnlar joylanmoqda, rasmlar va sxemalar adaptiv inpainting bilan tarjima qilinmoqda...")
            PPTXProcessor.apply_translations_and_export(
                original_pptx_path=str(working_pptx_path),
                translations_map=translations_map,
                output_pptx_path=str(output_pptx_path),
                auto_fit=True,
                target_script=target_script,
                clean_watermarks=True,
                presentation_title=translated_name.replace(".pptx", "").replace("_", " ")
            )

            elapsed_file = round(time.time() - file_start_time, 2)
            logger.info(f"  ✅ Muvaffaqiyatli saqlandi! ({elapsed_file} soniya) -> {sub_dir_parts}/{translated_name}")

            progress_data[rel_key] = {
                "status": "success",
                "output_name": translated_name,
                "rel_subfolder": str(sub_dir_parts),
                "slides": total_slides,
                "items": total_items,
                "time_sec": elapsed_file
            }
            progress_file.write_text(json.dumps(progress_data, ensure_ascii=False, indent=2), encoding="utf-8")

        except Exception as proc_err:
            logger.error(f"  ❌ Xatolik yuz berdi ({rel_key}): {proc_err}", exc_info=True)
            progress_data[rel_key] = {
                "status": "failed",
                "error": str(proc_err)
            }
            progress_file.write_text(json.dumps(progress_data, ensure_ascii=False, indent=2), encoding="utf-8")

        finally:
            if is_temp_converted and working_pptx_path.exists():
                try:
                    working_pptx_path.unlink()
                except Exception:
                    pass

    # Cleanup temp directory
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass

    total_elapsed = round((time.time() - start_total_time) / 60, 2)
    successful = sum(1 for v in progress_data.values() if isinstance(v, dict) and v.get("status") == "success")
    failed = sum(1 for v in progress_data.values() if isinstance(v, dict) and v.get("status") == "failed")
    logger.info(f"\n{'='*60}")
    logger.info("🎉 BATCH TARJIMA YAKUNLANDI!")
    logger.info(f"📊 Jami: {total_files} | Muvaffaqiyatli: {successful} | Xatolik: {failed}")
    logger.info(f"⏱ Umumiy sarflangan vaqt: {total_elapsed} daqiqa")
    logger.info(f"📁 Barcha tayyor taqdimotlar: {target_root}")
    logger.info(f"{'='*60}\n")


if __name__ == '__main__':
    run_batch_translation()
