import os
import sys
import json
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

SOFF_ROOT = Path(r'C:\Users\user\Desktop\Antigravity\Soff.uz')
EN_DIR = SOFF_ROOT / 'Tarjima_qilingan_slaydlar' / 'Slaydlar_EN'
PROGRESS_FILE = EN_DIR / 'batch_translation_progress.json'
CATALOG_FILE = EN_DIR / 'KATALOG_TARJIMA_EN.md'

def get_slide_count_from_zip(pptx_path):
    try:
        with zipfile.ZipFile(pptx_path, 'r') as z:
            slides = [n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
            return len(slides)
    except Exception:
        return 0

def main():
    progress = {}
    if PROGRESS_FILE.exists():
        with open(PROGRESS_FILE, 'r', encoding='utf-8') as f:
            progress = json.load(f)

    # Invert progress map: output_name -> original_name & info
    out_to_orig = {}
    for orig, data in progress.items():
        out_name = data.get('output_name')
        if out_name:
            out_to_orig[out_name] = {
                'orig': orig,
                'slides': data.get('slides', 0),
                'items': data.get('items', 0)
            }

    pptx_files = sorted([f for f in EN_DIR.glob('*.pptx') if not f.name.startswith('~$')])

    total_files = len(pptx_files)
    total_slides = 0
    total_size_bytes = 0

    rows = []
    for idx, f in enumerate(pptx_files, 1):
        size_mb = f.stat().st_size / (1024 * 1024)
        total_size_bytes += f.stat().st_size
        
        info = out_to_orig.get(f.name, {})
        slides = info.get('slides')
        if not slides or slides == 0:
            slides = get_slide_count_from_zip(f)
        total_slides += slides

        orig_name = info.get('orig', '-')
        
        # Format table row
        # Escape pipe if any
        clean_title = f.stem.replace('|', '/')
        orig_clean = orig_name.replace('|', '/')
        
        rows.append(f"| {idx} | **{clean_title}** | [{f.name}](./{f.name}) | {slides} ta | {size_mb:.2f} MB | `{orig_clean}` |")

    total_size_mb = total_size_bytes / (1024 * 1024)
    total_size_gb = total_size_mb / 1024

    md_content = f"""# SlaydHub.uz — Inglizcha Slaydlar Kolleksiyasi (Slaydlar_EN) Katalogi

Mazkur katalog `SlaydHubUz_bot` ning sun'iy intellekt tarjima va moslashtirish yadrosi (Gemini / Qwen AI) orqali to'liq ingliz tilidan o'zbek tiliga (Lotin alifbosida) akademik standartlarda tarjima qilingan, barcha begona reklama va havolalari tozalangan, shrift va konteynerlari avtomatik sig'dirilgan (auto-fit) sifatli slaydlar to'plamini o'z ichiga oladi.

## Umumiy Statistika

- **Jami saralangan va tarjima qilingan taqdimotlar:** {total_files} ta
- **Jami slaydlar soni:** {total_slides:,} ta
- **Umumiy fayl hajmi:** {total_size_mb:.2f} MB ({total_size_gb:.2f} GB)
- **Asosiy yo'nalishlar:** Kompyuter ilmlari (Computer Science), Sun'iy intellekt (AI/ML), Kompyuter ko'rishi (Computer Vision), Dasturlash tillari va Kompilyatorlar (Compilers), Tarmoq va Kiberxavfsizlik (Networking & Cybersecurity), Simsiz aloqa (Wireless Communications & Antennas), Katta hajmdagi ma'lumotlar (Big Data & HPC).

---

## Taqdimotlar Jadvali

| № | Taqdimot Mavzusi (O'zbekcha) | Fayl Havolasi | Slaydlar | Hajmi | Asl Manba Fayli |
|---|---|---|---|---|---|
""" + "\n".join(rows) + f"""

---
*Katalog yaratilgan vaqt:* 2026-09-18
*Tizim:* SlaydHub AI Bot Engine (TranslatorService & PPTXProcessor)
"""

    with open(CATALOG_FILE, 'w', encoding='utf-8') as f:
        f.write(md_content)

    print(f"Katalog yaratildi: {CATALOG_FILE}")
    print(f"Jami fayllar: {total_files}, Jami slaydlar: {total_slides}, Jami hajm: {total_size_mb:.2f} MB")

if __name__ == '__main__':
    main()
