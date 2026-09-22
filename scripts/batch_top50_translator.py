# -*- coding: utf-8 -*-
import os
import sys
import json
import time
import shutil
import re
from typing import Dict, Any, List, Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from backend.core.pptx_processor import PPTXProcessor, safe_load_presentation
from backend.core.gemini_translator import GeminiTranslator
from bot.services.translator_service import TranslatorService

SOURCE_DIR = r"C:\Users\user\Desktop\slayd\Yuklab_Olingan_Top50"
TARGET_DIR = r"C:\Users\user\Desktop\slayd\Tarjima_Qilingan_Top50"
PROGRESS_FILE = os.path.join(TARGET_DIR, "progress.json")
TEMP_CONV_DIR = os.path.join(TARGET_DIR, "_temp_converted")

os.makedirs(TARGET_DIR, exist_ok=True)
os.makedirs(TEMP_CONV_DIR, exist_ok=True)

sys.stdout.reconfigure(encoding="utf-8")


def load_progress() -> Dict[str, Any]:
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "completed": [],
        "failed": {},
        "total_files": 0,
        "total_slides": 0,
        "total_items": 0
    }


def save_progress(progress: Dict[str, Any]):
    try:
        temp_p = PROGRESS_FILE + ".tmp"
        with open(temp_p, "w", encoding="utf-8") as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)
        if os.path.exists(PROGRESS_FILE):
            os.replace(temp_p, PROGRESS_FILE)
        else:
            os.rename(temp_p, PROGRESS_FILE)
    except Exception as e:
        print(f"[WARN] Progressni saqlashda xatolik: {e}")


def convert_ppt_to_pptx(ppt_path: str) -> Optional[str]:
    powerpoint = None
    pres = None
    try:
        import win32com.client
        base = os.path.splitext(os.path.basename(ppt_path))[0]
        out_pptx = os.path.join(TEMP_CONV_DIR, f"{base}_conv_{int(time.time()*1000)}.pptx")
        
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.DisplayAlerts = 0  # ppAlertsNone
        pres = powerpoint.Presentations.Open(os.path.abspath(ppt_path), WithWindow=False)
        pres.SaveAs(os.path.abspath(out_pptx), 24)
        
        if os.path.exists(out_pptx) and os.path.getsize(out_pptx) > 0:
            return out_pptx
    except Exception as e:
        print(f"  [ERROR] .ppt konvertatsiyasida xatolik ({os.path.basename(ppt_path)}): {e}")
    finally:
        try:
            if pres:
                pres.Close()
        except Exception:
            pass
        try:
            if powerpoint:
                powerpoint.Quit()
        except Exception:
            pass
    return None


def get_all_source_files() -> List[Dict[str, Any]]:
    tasks = []
    folders = sorted([d for d in os.listdir(SOURCE_DIR) if os.path.isdir(os.path.join(SOURCE_DIR, d))])
    
    for folder in folders:
        f_dir = os.path.join(SOURCE_DIR, folder)
        for root, _, files in os.walk(f_dir):
            for f in sorted(files):
                if f.startswith("~$") or "Термодинамика" in f or "Эволюция звёзд" in f:
                    continue
                ext = os.path.splitext(f)[1].lower()
                if ext in [".pptx", ".ppt"]:
                    full_path = os.path.join(root, f)
                    rel_path = os.path.relpath(full_path, SOURCE_DIR)
                    tasks.append({
                        "folder": folder,
                        "filename": f,
                        "ext": ext,
                        "full_path": full_path,
                        "rel_path": rel_path
                    })
    return tasks


def main():
    print("=" * 65)
    print("TOP-50 TAQDIMOTLARNI AVTONOM O'ZBEK TILIGA TARJIMA QILISH")
    print(f"Manba: {SOURCE_DIR}")
    print(f"Maqsad: {TARGET_DIR}")
    print("=" * 65)

    svc = TranslatorService()
    progress = load_progress()
    all_tasks = get_all_source_files()
    total_count = len(all_tasks)
    progress["total_files"] = total_count

    completed_set = set(progress["completed"])
    print(f"Jami topilgan fayllar: {total_count} ta")
    print(f"Oldin tarjima qilingan: {len(completed_set)} ta")
    print(f"Qolgan navbat: {total_count - len(completed_set)} ta fayl\n")

    for idx, task in enumerate(all_tasks, 1):
        rel_p = task["rel_path"]
        if rel_p in completed_set:
            continue

        folder_name = task["folder"]
        orig_filename = task["filename"]
        ext = task["ext"]
        source_file = task["full_path"]
        
        target_folder = os.path.join(TARGET_DIR, folder_name)
        os.makedirs(target_folder, exist_ok=True)

        print(f"\n[{idx}/{total_count}] Jild: {folder_name}")
        print(f"  Fayl: {orig_filename}")

        pptx_to_process = source_file
        is_temp_pptx = False

        # 1. Convert .ppt to .pptx if needed
        if ext == ".ppt":
            print("  .ppt formatidan .pptx ga o'tkazilmoqda...")
            conv_path = convert_ppt_to_pptx(source_file)
            if not conv_path:
                print(f"  Konvertatsiya qilib bo'lmadi, o'tkazib yuboriladi.")
                progress["failed"][rel_p] = "PPT conversion failed"
                save_progress(progress)
                continue
            pptx_to_process = conv_path
            is_temp_pptx = True

        # 2. Get Clean Uzbek Title
        try:
            print("  Taqdimot nomi o'zbekchaga tarjima qilinmoqda...")
            clean_title = svc.get_clean_presentation_title(orig_filename, target_script="latin")
            print(f"  Yangi nomi: {clean_title}")
        except Exception as e:
            clean_title = os.path.splitext(orig_filename)[0]
            clean_title = re.sub(r'[/\\:*?\"<>|_]', " ", clean_title).strip()
            print(f"  Nom tarjimasida fallback: {clean_title}")

        out_pptx_filename = f"{clean_title}.pptx"
        final_output_path = os.path.join(target_folder, out_pptx_filename)

        # 3. Translate Presentation
        start_t = time.time()
        try:
            pres_data = PPTXProcessor.extract_presentation_data(pptx_to_process)
            slides_count = len(pres_data.get("slides", []))
            items_to_translate = []
            for s in pres_data.get("slides", []):
                items_to_translate.extend(s.get("items", []))

            total_items = len(items_to_translate)
            print(f"  Slaydlar soni: {slides_count} ta, Matn bloklari: {total_items} ta")

            translations_map = {}
            if items_to_translate:
                batch_size = 40
                for b_idx in range(0, total_items, batch_size):
                    batch = items_to_translate[b_idx:b_idx + batch_size]
                    b_num = (b_idx // batch_size) + 1
                    total_batches = (total_items + batch_size - 1) // batch_size
                    print(f"    Partiya {b_num}/{total_batches} ({len(batch)} ta matn)...", end="\r")

                    translated_batch = None
                    for attempt in range(5):
                        try:
                            translated_batch = svc.translator.translate_items_batch(
                                items=batch,
                                target_script="latin",
                                domain="Ta'lim va fan"
                            )
                            break
                        except Exception as te:
                            wait_s = 2.0 * (attempt + 1)
                            print(f"\n    Gemini qayta urinish ({attempt+1}/5) {wait_s}s kutish: {te}")
                            time.sleep(wait_s)

                    if translated_batch:
                        for item in translated_batch:
                            translations_map[item["id"]] = item["translated_text"]
                    else:
                        for item in batch:
                            translations_map[item["id"]] = item["text"]
                    
                    time.sleep(1.0)

                print(f"    Matnlar tarjimasi yakunlandi! ({len(translations_map)}/{total_items})              ")

            # 4. Apply translations, layout optimization, title alignment and clean watermarks
            print("  Reklamalar tozalanmoqda va slaydlar shakllantirilmoqda...")
            PPTXProcessor.apply_translations_and_export(
                original_pptx_path=pptx_to_process,
                translations_map=translations_map,
                output_pptx_path=final_output_path,
                auto_fit=True,
                target_script="latin",
                clean_watermarks=True,
                presentation_title=clean_title,
                embed_thumbnail=False
            )

            dur = time.time() - start_t
            file_sz_mb = os.path.getsize(final_output_path) / (1024 * 1024)
            print(f"  MUVAFFAQIYAT: {out_pptx_filename} ({file_sz_mb:.2f} MB, {dur:.1f} soniya)")

            progress["completed"].append(rel_p)
            progress["total_slides"] += slides_count
            progress["total_items"] += total_items
            save_progress(progress)

        except Exception as pe:
            print(f"  Xatolik yuz berdi: {pe}")
            progress["failed"][rel_p] = str(pe)
            save_progress(progress)

        finally:
            if is_temp_pptx and os.path.exists(pptx_to_process):
                try:
                    os.remove(pptx_to_process)
                except Exception:
                    pass

    try:
        shutil.rmtree(TEMP_CONV_DIR, ignore_errors=True)
    except Exception:
        pass

    print("\n" + "=" * 65)
    print("BARCHA TAQDIMOTLARNI TARJIMA QILISH YAKUNLANDI!")
    print(f"Muvaffaqiyatli yakunlangan: {len(progress['completed'])} ta fayl")
    print(f"Jami slaydlar soni: {progress['total_slides']} ta")
    print(f"Jami tarjima qilingan matnlar: {progress['total_items']} ta")
    if progress["failed"]:
        print(f"Xatolik uchragan: {len(progress['failed'])} ta fayl")
    print(f"Natijalar papkasi: {TARGET_DIR}")
    print("=" * 65)


if __name__ == "__main__":
    main()
