import os
import sys
import zipfile
import asyncio
import shutil
import time
import json
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

POWERPOINT_ROOT = Path(r'C:\Users\user\Desktop\Antigravity\Power Point')
SOFF_ROOT = Path(r'C:\Users\user\Desktop\Antigravity\Soff.uz')
for p in [str(POWERPOINT_ROOT), str(SOFF_ROOT)]:
    if p not in sys.path:
        sys.path.insert(0, p)

from bot.services.translator_service import TranslatorService

orig_pptx = SOFF_ROOT / 'Slaydlar_EN' / '01_Sara_Slaydlar_EN' / 'www.cc.gatech.edu_01.pptx'
temp_dir = SOFF_ROOT / 'Temp_PPT_Convert'
temp_dir.mkdir(parents=True, exist_ok=True)
temp_stripped = temp_dir / 'stripped_01.pptx'
clean_name = "Kompyuter ko'rishi va tasvirlarni tahlil qilish (1-ma'ruza).pptx"
out_path = SOFF_ROOT / 'Tarjima_qilingan_slaydlar' / 'Slaydlar_EN' / clean_name

print(f"Original size: {orig_pptx.stat().st_size / (1024*1024):.2f} MB")

# 1. Create stripped version (replace files > 5MB with 1KB dummy)
large_items = {}
print("Analyzing zip infolist and stripping large media...")
with zipfile.ZipFile(orig_pptx, 'r') as zin, zipfile.ZipFile(temp_stripped, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.file_size > 5 * 1024 * 1024:
            print(f"  Stripping media: {item.filename} ({item.file_size / (1024*1024):.2f} MB)")
            large_items[item.filename] = zin.read(item.filename)
            zout.writestr(item.filename, b'\x00' * 1024)
        else:
            zout.writestr(item, zin.read(item.filename))

print(f"Stripped size: {temp_stripped.stat().st_size / (1024*1024):.2f} MB")
print("Starting translation...")
t0 = time.time()
ts = TranslatorService()
res = asyncio.run(ts.translate_presentation(
    input_pptx=str(temp_stripped),
    output_pptx=str(out_path),
    target_script='latin',
    presentation_title="Kompyuter ko'rishi va tasvirlarni tahlil qilish (1-ma'ruza)",
    domain="Axborot texnologiyalari, IT, fan, muhandislik va oliy ta'lim"
))

print("Translation completed. Result:", res.get("success"), f"in {time.time() - t0:.2f}s")

# 3. Restore large items into out_path
if res.get("success") and large_items:
    print("Restoring large media files into translated presentation...")
    temp_restored = temp_dir / 'restored_01.pptx'
    with zipfile.ZipFile(out_path, 'r') as zin, zipfile.ZipFile(temp_restored, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in large_items:
                zout.writestr(item.filename, large_items[item.filename])
            else:
                zout.writestr(item, zin.read(item.filename))
    shutil.move(str(temp_restored), str(out_path))
    print("Restored large media files successfully!")

print(f"Final file size: {out_path.stat().st_size / (1024*1024):.2f} MB")

# Record progress in batch_translation_progress.json
progress_file = SOFF_ROOT / 'Tarjima_qilingan_slaydlar' / 'Slaydlar_EN' / 'batch_translation_progress.json'
progress = {}
if progress_file.exists():
    try:
        with open(progress_file, 'r', encoding='utf-8') as pf:
            progress = json.load(pf)
    except Exception:
        pass

if res.get("success"):
    import datetime
    progress["www.cc.gatech.edu_01.pptx"] = {
        "status": "success",
        "output_name": clean_name,
        "original_name": "www.cc.gatech.edu_01.pptx",
        "slides": res.get("total_slides", 0),
        "items": res.get("total_items", 0),
        "cleaned_ads": res.get("watermarks_removed", 0),
        "time_sec": round(time.time() - t0, 2),
        "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    with open(progress_file, 'w', encoding='utf-8') as pf:
        json.dump(progress, pf, ensure_ascii=False, indent=2)
    print("Progress recorded for www.cc.gatech.edu_01.pptx")

# Clean temp stripped
try:
    temp_stripped.unlink(missing_ok=True)
except Exception:
    pass
