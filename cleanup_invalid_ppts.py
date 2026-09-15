import os
import zipfile
import shutil
import csv
import time

BASE_DIR = r"C:\Users\user\Desktop\Powepoint"
MASTER_CSV = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.csv")
MASTER_MD = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.md")

print("1. NOSOZ FAYLLARNI QIDIRISH VA O'CHIRISH...")
invalid_count = 0
invalid_size = 0

for root, _, files in os.walk(BASE_DIR):
    for f in files:
        path = os.path.join(root, f)
        is_invalid = False
        
        if f.lower().endswith(('.pptx', '.potx')):
            try:
                with zipfile.ZipFile(path, 'r') as z:
                    if '[Content_Types].xml' not in z.namelist():
                        is_invalid = True
            except Exception:
                is_invalid = True
                
        elif f.lower().endswith(('.ppt', '.pot')):
            try:
                with open(path, 'rb') as pf:
                    if pf.read(8) != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
                        is_invalid = True
            except Exception:
                is_invalid = True
                
        if is_invalid:
            try:
                size = os.path.getsize(path)
                os.remove(path)
                invalid_size += size
                invalid_count += 1
                # print(f"O'chirildi: {path}")
            except Exception as e:
                print(f"Xatolik (o'chirishda): {path} -> {e}")

print(f"-> O'chirildi: {invalid_count} ta fayl ({invalid_size / 1024 / 1024:.2f} MB)")

print("\n2. BO'SH PAPKALARNI TOZALASH...")
empty_folders = 0
# To delete empty folders, we traverse bottom-up
for root, dirs, files in os.walk(BASE_DIR, topdown=False):
    for d in dirs:
        dir_path = os.path.join(root, d)
        try:
            if not os.listdir(dir_path):
                os.rmdir(dir_path)
                empty_folders += 1
        except Exception:
            pass
print(f"-> O'chirildi: {empty_folders} ta bo'sh papka")

print("\n3. MASTER CSV NI YANGILASH (o'chirilganlarni olib tashlash)...")
if os.path.exists(MASTER_CSV):
    with open(MASTER_CSV, 'r', encoding='utf-8', errors='ignore') as f:
        reader = csv.reader(f)
        header = next(reader, None)
        valid_rows = []
        for row in reader:
            if len(row) >= 7:
                file_path = row[6]
                if os.path.exists(file_path):
                    valid_rows.append(row)
    
    with open(MASTER_CSV, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)
        if header:
            writer.writerow(header)
        writer.writerows(valid_rows)
    print(f"-> CSV yangilandi. Faol qatorlar: {len(valid_rows)}")

print("\n4. MASTER MD STATISTIKASINI YANGILASH...")
source_stats = {}
total_files = 0
total_bytes = 0
for item in os.listdir(BASE_DIR):
    item_path = os.path.join(BASE_DIR, item)
    if os.path.isdir(item_path) and not item.startswith('.'):
        count = size = 0
        for root, _, files in os.walk(item_path):
            for f in files:
                if f.lower().endswith(('.pptx', '.ppt', '.zip', '.potx', '.pot')):
                    count += 1
                    size += os.path.getsize(os.path.join(root, f))
        if count > 0:
            source_stats[item] = (count, size)
            total_files += count
            total_bytes += size
            
total_gb = total_bytes / (1024 ** 3)
now_str = time.strftime('%Y-%m-%d %H:%M:%S')

lines = [
    '# BARCHA TAQDIMOT VA SLAYDLAR MASTER KATALOGI\n',
    f'- **Jami Taqdimotlar Soni**: **{total_files:,} ta**',
    f'- **Umumiy Hajmi**: **{total_gb:.2f} GB** ({total_bytes / (1024*1024):.2f} MB)',
    f'- **Asosiy Papka**: `{BASE_DIR}`',
    f'- **Oxirgi Yangilanish**: `{now_str}`\n',
    '### Saytlar Kesimida Statistika:\n',
    '| # | Manba Sayt | Taqdimotlar Soni | Hajmi (MB) | Format | Holat |',
    '|---|------------|:----------------:|:----------:|:------:|:-----:|'
]

for idx, (source, (cnt, sz)) in enumerate(sorted(source_stats.items(), key=lambda x: x[1][0], reverse=True), 1):
    lines.append(f'| {idx} | **{source}** | **{cnt:,} ta** | {sz/(1024*1024):.2f} MB | `.pptx/.pot` | Faol |')

lines.append(f'\n**JAMI**: **{total_files:,} ta taqdimot ({total_gb:.2f} GB)**\n')

with open(MASTER_MD, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print(f"-> MD yangilandi: {total_files:,} ta ({total_gb:.2f} GB)")
print("\nBARCHA ISHLAR YAKUNLANDI!")
