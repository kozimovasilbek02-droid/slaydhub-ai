import os
import sys
import re
import csv
import time
import shutil
import hashlib
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"C:\Users\user\Desktop\Powepoint"
MASTER_CSV = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.csv")
MASTER_MD = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.md")
MIN_FREE_SPACE_GB = 1.5
NUM_THREADS = 10

SESSION_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': '*/*'
}

registry_titles = set()

def clean_name(s):
    if not s: return ""
    s = s.lower()
    s = re.sub(r'\.pptx|\.ppt|\.zip|\.potx|\.pot|\.rar', '', s)
    s = re.sub(r'[\W_]+', ' ', s).strip()
    return s

def clean_filename(name):
    clean = re.sub(r'[\\/*?:"<>|]', "", name)
    clean = re.sub(r'\s+', ' ', clean).strip()
    return clean[:120] if len(clean) > 120 else clean

def load_global_registry():
    global registry_titles
    print("GLOBAL DUPLIKATSIYA REGISTRI YUKLANMOQDA...")
    for root, _, files in os.walk(BASE_DIR):
        for f in files:
            if f.lower().endswith(('.pptx', '.ppt', '.zip', '.potx', '.pot')):
                registry_titles.add(clean_name(f))
    print(f"Registr tayyor: {len(registry_titles):,} nom")

def is_duplicate(title):
    return clean_name(title) in registry_titles

def check_disk_space():
    total, used, free = shutil.disk_usage("C:\\")
    free_gb = free / (1024 ** 3)
    return free_gb > MIN_FREE_SPACE_GB

def save_file(source_name, title, category, file_url, content, ext=".pptx"):
    if not check_disk_space():
        return False
    
    source_folder = os.path.join(BASE_DIR, source_name, "Prezentatsiyalar")
    safe_title = clean_filename(title)
    item_dir = os.path.join(source_folder, safe_title)
    os.makedirs(item_dir, exist_ok=True)
    target_file = os.path.join(item_dir, f"{safe_title}{ext}")
    
    with open(target_file, 'wb') as f:
        f.write(content)

    size_mb = len(content) / (1024 * 1024)
    registry_titles.add(clean_name(title))

    with open(MASTER_CSV, 'a', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow(["NEW", source_name, safe_title, category, f"{size_mb:.2f}", ext, target_file, file_url])
    print(f"[{source_name}] Yuklandi: {safe_title} ({size_mb:.2f} MB)")
    return True

def process_powerpointstyles():
    source_name = "PowerPointStyles"
    print(f"\n{source_name} YUKLASH BOSHLANDI...")
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)

    base_url = 'https://www.powerpointstyles.com/'
    template_links = set()

    for page in range(1, 50):
        url = f'https://www.powerpointstyles.com/all-templates-page-{page}.html' if page > 1 else 'https://www.powerpointstyles.com/all-templates.html'
        try:
            r = s.get(url, timeout=10)
            if r.status_code != 200:
                break
            soup = BeautifulSoup(r.text, 'html.parser')
            found = False
            for a in soup.find_all('a', href=True):
                href = a['href']
                if 'template-' in href and href.endswith('.html'):
                    template_links.add(urljoin(base_url, href))
                    found = True
            if not found:
                break
        except:
            break

    print(f"{source_name}: {len(template_links)} ta shablon topildi.")

    def download_template(t_url):
        try:
            tr = s.get(t_url, timeout=10)
            tsoup = BeautifulSoup(tr.text, 'html.parser')
            t_str = tsoup.title.string if tsoup.title else ""
            if not t_str:
                h2 = tsoup.find('h2')
                t_str = h2.text.strip() if h2 else ""
            title = t_str.replace("Powerpoint Template", "").replace("PowerPoint Template", "").strip()
            if not title or is_duplicate(title):
                return
            
            dl_url = None
            for a in tsoup.find_all('a', href=True):
                if 'download_id-' in a['href']:
                    dl_url = urljoin(base_url, a['href'])
                    break
            
            if not dl_url:
                return

            dl_res = s.get(dl_url, timeout=30)
            if dl_res.status_code == 200 and len(dl_res.content) > 10000:
                ext = '.pot' if b'Microsoft PowerPoint' in dl_res.content[:500] else '.pptx'
                save_file(source_name, title, "General", dl_url, dl_res.content, ext)
        except Exception as e:
            pass

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        for idx, _ in enumerate(executor.map(download_template, template_links), 1):
            if idx % 20 == 0 or idx == len(template_links):
                print(f"  {source_name}: {idx}/{len(template_links)} tekshirildi...")

def rebuild_master_stats():
    print("\nMASTER KATALOG YANGILANMOQDA...")
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

if __name__ == '__main__':
    load_global_registry()
    process_powerpointstyles()
    rebuild_master_stats()
    print("Done!")
