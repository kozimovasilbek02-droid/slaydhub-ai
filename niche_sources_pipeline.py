"""
YPPPT.com va GreatPPT.com manbalaridan bepul taqdimotlarni avtomatik yuklash.
Ikkala sayt ham Xitoy tilida bo'lib, to'g'ridan-to'g'ri .zip ichida .pptx fayllari bor.
"""
import os
import sys
import re
import csv
import time
import shutil
import hashlib
import zipfile
import io
import requests
import warnings
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"C:\Users\user\Desktop\Powepoint"
MASTER_CSV = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.csv")
MASTER_MD = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.md")

MIN_FREE_SPACE_GB = 1.5
NUM_THREADS = 10

SESSION_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
}

registry_titles = set()
registry_hashes = set()
downloaded_count = 0
downloaded_size_bytes = 0

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

def check_disk_space():
    total, used, free = shutil.disk_usage("C:\\")
    free_gb = free / (1024 ** 3)
    if free_gb < MIN_FREE_SPACE_GB:
        print(f"DISK TO'LDI: {free_gb:.2f} GB!")
        return False
    return True

def load_global_registry():
    global registry_titles, registry_hashes
    print("=" * 50)
    print("GLOBAL DUPLIKATSIYA REGISTRI YUKLANMOQDA...")
    print("=" * 50)
    
    csv_sources = [
        MASTER_CSV,
        r"C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar_Katalogi.csv",
        r"C:\Users\user\Desktop\Slaydlar\Slidenest\Slidenest_Katalogi.csv",
        r"C:\Users\user\Desktop\Slaydlar\Showeet\PPTX_Katalogi.csv",
        r"C:\Users\user\Desktop\Slaydlar\SlidesCarnival\SlidesCarnival_Katalogi.csv"
    ]
    for path in csv_sources:
        if os.path.exists(path):
            try:
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    for row in csv.reader(f):
                        for col in row[:4]:
                            n = clean_name(col)
                            if len(n) > 3 and not n.isdigit():
                                registry_titles.add(n)
                        for col in row:
                            if len(col) == 32 and re.match(r'^[a-fA-F0-9]{32}$', col):
                                registry_hashes.add(col.lower())
            except Exception:
                pass

    for base in [BASE_DIR, r"C:\Users\user\Desktop\SlideEgg_Downloads", r"C:\Users\user\Desktop\Slaydlar"]:
        if os.path.exists(base):
            for root, _, files in os.walk(base):
                for f in files:
                    if f.lower().endswith(('.pptx', '.ppt', '.zip')):
                        registry_titles.add(clean_name(f))

    print(f"Registr tayyor: {len(registry_titles):,} nom, {len(registry_hashes):,} hash\n")

def is_duplicate(title, file_hash=None):
    cn = clean_name(title)
    if cn in registry_titles:
        return True
    if file_hash and file_hash.lower() in registry_hashes:
        return True
    return False

def save_file(source_name, title, category, file_url, content, ext=".pptx"):
    global downloaded_count, downloaded_size_bytes
    if not check_disk_space():
        return False
    file_hash = hashlib.md5(content).hexdigest()
    if file_hash in registry_hashes:
        return False

    source_folder = os.path.join(BASE_DIR, source_name, "Prezentatsiyalar")
    safe_title = clean_filename(title)
    item_dir = os.path.join(source_folder, safe_title)
    os.makedirs(item_dir, exist_ok=True)
    target_file = os.path.join(item_dir, f"{safe_title}{ext}")
    with open(target_file, 'wb') as f:
        f.write(content)

    size_mb = len(content) / (1024 * 1024)
    downloaded_count += 1
    downloaded_size_bytes += len(content)
    registry_titles.add(clean_name(title))
    registry_hashes.add(file_hash)

    # Source CSV
    source_csv = os.path.join(BASE_DIR, source_name, f"{source_name}_Katalogi.csv")
    csv_new = not os.path.exists(source_csv)
    with open(source_csv, 'a', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if csv_new:
            w.writerow(["#","Manba","Nomi","Kategoriya","Hajmi(MB)","Turi","MD5","Manzil","Havola","Sana"])
        w.writerow([downloaded_count, source_name, safe_title, category, f"{size_mb:.2f}", ext, file_hash, target_file, file_url, time.strftime("%Y-%m-%d %H:%M:%S")])

    # Master CSV
    with open(MASTER_CSV, 'a', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow(["NEW", source_name, safe_title, category, f"{size_mb:.2f}", ext, target_file, file_url])

    print(f"  [{source_name}] Yuklandi: {safe_title} ({size_mb:.2f} MB)")
    return True

def extract_pptx_from_zip(content, title, source_name, category, file_url):
    """Zip fayldan .pptx/.ppt ni ajratib oladi va saqlaydi."""
    try:
        z = zipfile.ZipFile(io.BytesIO(content))
        ppt_names = [n for n in z.namelist() if any(n.lower().endswith(ext) for ext in ['.pptx', '.ppt', '.potx'])]
        if ppt_names:
            best = sorted(ppt_names, key=lambda x: (1 if x.lower().endswith('.pptx') else 0, -len(x)), reverse=True)[0]
            ppt_content = z.read(best)
            ext = os.path.splitext(best)[1].lower()
            return save_file(source_name, title, category, file_url, ppt_content, ext)
    except Exception:
        pass
    return False

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
                    if f.lower().endswith(('.pptx', '.ppt', '.zip', '.potx')):
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
        lines.append(f'| {idx} | **{source}** | **{cnt:,} ta** | {sz/(1024*1024):.2f} MB | `.pptx` | Faol |')
    lines.append(f'\n**JAMI**: **{total_files:,} ta taqdimot ({total_gb:.2f} GB)**\n')
    with open(MASTER_MD, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"Master Katalog yangilandi: {total_files:,} ta ({total_gb:.2f} GB)\n")

# ======================================================================
# PIPELINE: YPPPT.COM (Xitoy bepul taqdimotlar arxivi)
# ======================================================================
def process_ypppt():
    source_name = "YPPPT"
    print(f"\n{'='*50}")
    print(f"YPPPT.COM YUKLASH BOSHLANDI...")
    print(f"{'='*50}")
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)

    all_urls = set()
    # moban = shablonlar, jiaocheng = o'quv, tubiao = diagrammalar
    cats = ['moban', 'jiaocheng', 'tubiao']
    for cat in cats:
        for page in range(1, 50):
            url = f'https://www.ypppt.com/{cat}/' if page == 1 else f'https://www.ypppt.com/{cat}/list-{page}.html'
            try:
                r = s.get(url, timeout=10)
                r.encoding = 'utf-8'
                if r.status_code != 200:
                    break
                soup = BeautifulSoup(r.text, 'html.parser')
                found = False
                for a in soup.find_all('a', href=True):
                    href = a['href']
                    if '/article/' in href and '.html' in href:
                        all_urls.add(urljoin(url, href))
                        found = True
                if not found:
                    break
            except Exception:
                break

    print(f"YPPPT: {len(all_urls)} ta taqdimot sahifasi topildi.")

    def process_ypppt_page(page_url):
        try:
            res = s.get(page_url, timeout=15)
            res.encoding = 'utf-8'
            if res.status_code != 200:
                return
            soup = BeautifulSoup(res.text, 'html.parser')
            h1 = soup.find('h1')
            title = h1.text.strip() if h1 else ''
            if not title or is_duplicate(title):
                return

            # Find download page link
            dl_page_url = None
            for a in soup.find_all('a', href=True):
                href = a['href']
                if '/p/d.php' in href or 'down' in href.lower():
                    dl_page_url = urljoin(page_url, href)
                    break
            if not dl_page_url:
                return

            # Fetch the download page
            dr = s.get(dl_page_url, timeout=10)
            dr.encoding = 'utf-8'
            dsoup = BeautifulSoup(dr.text, 'html.parser')

            # Find zip link
            zip_url = None
            for a in dsoup.find_all('a', href=True):
                href = a['href']
                if any(ext in href.lower() for ext in ['.zip', '.rar', '.pptx']):
                    zip_url = urljoin(dl_page_url, href)
                    break
            if not zip_url:
                return

            # Download file
            file_res = s.get(zip_url, timeout=30)
            if file_res.status_code != 200 or len(file_res.content) < 15000:
                return

            content = file_res.content
            if content[:4] == b'PK\x03\x04':
                extract_pptx_from_zip(content, title, "YPPPT", "General", zip_url)
            elif content[:4] == b'PK\x05\x06' or len(content) > 30000:
                save_file("YPPPT", title, "General", zip_url, content, ".pptx")
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        futures = [executor.submit(process_ypppt_page, u) for u in all_urls]
        for idx, f in enumerate(as_completed(futures), 1):
            if idx % 50 == 0 or idx == len(all_urls):
                print(f"  YPPPT: {idx}/{len(all_urls)} sahifa tekshirildi...")

# ======================================================================
# PIPELINE: FreePPTBackgrounds.net (download categories)
# ======================================================================
def process_freepptbackgrounds():
    source_name = "FreePPTBackgrounds"
    print(f"\n{'='*50}")
    print(f"FreePPTBackgrounds.net YUKLASH BOSHLANDI...")
    print(f"{'='*50}")
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)

    # Get all posts from sitemap
    all_urls = set()
    for i in range(1, 10):
        sm_url = f'https://www.freepptbackgrounds.net/wp-sitemap-posts-post-{i}.xml'
        try:
            r = s.get(sm_url, timeout=15)
            if r.status_code != 200:
                break
            soup = BeautifulSoup(r.text, 'html.parser')
            locs = [l.text for l in soup.find_all('loc')]
            if not locs:
                break
            all_urls.update(locs)
        except Exception:
            break

    print(f"FreePPTBackgrounds: {len(all_urls)} ta sahifa topildi.")

    def process_fpb_page(page_url):
        try:
            res = s.get(page_url, timeout=15)
            if res.status_code != 200:
                return
            soup = BeautifulSoup(res.text, 'html.parser')
            h1 = soup.find('h1') or soup.find('h2', class_='entry-title')
            title = h1.text.strip() if h1 else ''
            if not title or is_duplicate(title):
                return

            # Find .potx or .pptx direct link in page
            dl_url = None
            for a in soup.find_all('a', href=True):
                href = a['href']
                if any(href.lower().endswith(ext) for ext in ['.pptx', '.potx', '.ppt', '.zip']):
                    dl_url = urljoin(page_url, href)
                    break

            if not dl_url:
                return

            file_res = s.get(dl_url, timeout=25)
            if file_res.status_code == 200 and len(file_res.content) > 10000:
                content = file_res.content
                ext = ".pptx"
                if dl_url.lower().endswith('.potx'):
                    ext = ".potx"
                elif dl_url.lower().endswith('.ppt'):
                    ext = ".ppt"
                elif dl_url.lower().endswith('.zip'):
                    extract_pptx_from_zip(content, title, source_name, "Backgrounds", dl_url)
                    return
                save_file(source_name, title, "Backgrounds", dl_url, content, ext)
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        futures = [executor.submit(process_fpb_page, u) for u in all_urls]
        for idx, f in enumerate(as_completed(futures), 1):
            if idx % 100 == 0 or idx == len(all_urls):
                print(f"  FreePPTBG: {idx}/{len(all_urls)} sahifa tekshirildi...")

# ======================================================================
def main():
    start_time = time.time()
    load_global_registry()
    process_ypppt()
    rebuild_master_stats()
    process_freepptbackgrounds()
    rebuild_master_stats()
    elapsed = time.time() - start_time
    print(f"\nYAKUNLANDI! Sarflangan vaqt: {elapsed / 60:.2f} min")
    print(f"Yangi yuklangan: {downloaded_count} ta ({downloaded_size_bytes / (1024*1024):.2f} MB)\n")

if __name__ == "__main__":
    main()
