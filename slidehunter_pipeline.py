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

# --- CONFIGURATION ---
BASE_DIR = r"C:\Users\user\Desktop\Powepoint"
SOURCE_NAME = "SlideHunter"
SOURCE_DIR = os.path.join(BASE_DIR, SOURCE_NAME)
PREZ_DIR = os.path.join(SOURCE_DIR, "Prezentatsiyalar")
SOURCE_CSV = os.path.join(SOURCE_DIR, f"{SOURCE_NAME}_Katalogi.csv")
SOURCE_MD = os.path.join(SOURCE_DIR, f"{SOURCE_NAME}_Katalogi.md")

MASTER_CSV = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.csv")
MASTER_MD = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.md")

MIN_FREE_SPACE_GB = 1.5
NUM_THREADS = 20

SESSION_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
}

registry_titles = set()
registry_hashes = set()
downloaded_count = 0
downloaded_size_bytes = 0

def clean_name(s):
    if not s:
        return ""
    s = s.lower()
    s = re.sub(r'\.pptx|\.ppt|\.zip|\.potx|\.pot', '', s)
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
        print(f"⚠️ [OGOHLANTIRISH] C: diskida bo'sh joy kam qoldi: {free_gb:.2f} GB!")
        return False
    return True

def load_global_registry():
    global registry_titles, registry_hashes
    print("==================================================")
    print("🔍 GLOBAL DUPLIKATSIYA REGISTRI YUKLANMOQDA...")
    print("==================================================")
    
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
                    reader = csv.reader(f)
                    for row in reader:
                        for col in row[:4]:
                            n = clean_name(col)
                            if len(n) > 3 and not n.isdigit() and 'katalog' not in n and 'taqdimot' not in n:
                                registry_titles.add(n)
                        for col in row:
                            if len(col) == 32 and re.match(r'^[a-fA-F0-9]{32}$', col):
                                registry_hashes.add(col.lower())
            except Exception as e:
                pass

    folder_sources = [
        BASE_DIR,
        r"C:\Users\user\Desktop\SlideEgg_Downloads",
        r"C:\Users\user\Desktop\Slaydlar"
    ]
    for base in folder_sources:
        if os.path.exists(base):
            for root, _, files in os.walk(base):
                for f in files:
                    if f.lower().endswith(('.pptx', '.ppt', '.zip')):
                        registry_titles.add(clean_name(f))

    print(f"✅ Global Registr tayyor: {len(registry_titles):,} ta unikal nom va {len(registry_hashes):,} ta MD5 hash yuklandi.\n")

def is_duplicate(title, file_hash=None):
    cn = clean_name(title)
    if cn in registry_titles:
        return True
    if file_hash and file_hash.lower() in registry_hashes:
        return True
    return False

def add_to_registry(title, file_hash=None):
    cn = clean_name(title)
    if cn:
        registry_titles.add(cn)
    if file_hash:
        registry_hashes.add(file_hash.lower())

def save_and_catalog(title, category, file_url, content, ext=".pptx"):
    global downloaded_count, downloaded_size_bytes
    
    if not check_disk_space():
        return False
        
    file_hash = hashlib.md5(content).hexdigest()
    if file_hash in registry_hashes:
        print(f"  ⏭️ [HASH DUBLIKAT] {title}")
        return False
        
    safe_title = clean_filename(title)
    item_dir = os.path.join(PREZ_DIR, safe_title)
    os.makedirs(item_dir, exist_ok=True)
    
    target_file = os.path.join(item_dir, f"{safe_title}{ext}")
    with open(target_file, 'wb') as f:
        f.write(content)
        
    size_mb = len(content) / (1024 * 1024)
    downloaded_count += 1
    downloaded_size_bytes += len(content)
    add_to_registry(title, file_hash)
    
    # 1. Update SlideHunter CSV
    csv_is_new = not os.path.exists(SOURCE_CSV)
    with open(SOURCE_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if csv_is_new:
            writer.writerow(["#", "Manba", "Taqdimot Nomi", "Kategoriya", "Hajmi (MB)", "Fayl Turi", "MD5", "Mahalliy Manzil", "Sayt Havolasi", "Sana"])
        writer.writerow([
            downloaded_count,
            SOURCE_NAME,
            safe_title,
            category,
            f"{size_mb:.2f}",
            ext,
            file_hash,
            target_file,
            file_url,
            time.strftime("%Y-%m-%d %H:%M:%S")
        ])
        
    # 2. Append to Master CSV
    master_is_new = not os.path.exists(MASTER_CSV)
    with open(MASTER_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if master_is_new:
            writer.writerow(["#", "Manba Sayt", "Taqdimot Nomi", "Kategoriya", "Fayl Hajmi (MB)", "Fayl Turi", "Mahalliy Manzil", "Sayt Havolasi"])
        writer.writerow([
            "NEW",
            SOURCE_NAME,
            safe_title,
            category,
            f"{size_mb:.2f}",
            ext,
            target_file,
            file_url
        ])

    print(f"  📥 [{SOURCE_NAME}] Yuklandi: {safe_title} ({size_mb:.2f} MB)")
    return True

def rebuild_master_stats():
    print("\n📊 MASTER KATALOG STATISTIKASI YANGILANMOQDA...")
    source_stats = {}
    total_files = 0
    total_bytes = 0
    
    for item in os.listdir(BASE_DIR):
        item_path = os.path.join(BASE_DIR, item)
        if os.path.isdir(item_path) and not item.startswith('.'):
            count = 0
            size = 0
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
    
    md_content = f"# 🌐 BARCHA TAQDIMOT VA SLAYDLAR MASTER KATALOGI\n\n"
    md_content += f"- **Jami Taqdimotlar Soni**: **{total_files:,} ta**\n"
    md_content += f"- **Umumiy Hajmi**: **{total_gb:.2f} GB** ({total_bytes / (1024 * 1024):.2f} MB)\n"
    md_content += f"- **Asosiy Papka**: `{BASE_DIR}`\n"
    md_content += f"- **Oxirgi Yangilanish**: `{time.strftime('%Y-%m-%d %H:%M:%S')}`\n\n"
    md_content += f"### 📊 Saytlar Kesimida To'liq Statistika:\n\n"
    md_content += f"| # | Manba Sayt | Taqdimotlar Soni | Hajmi (MB) | Format | Holat |\n"
    md_content += f"|---|------------|:----------------:|:----------:|:------:|:-----:|\n"
    
    for idx, (source, (cnt, sz)) in enumerate(sorted(source_stats.items(), key=lambda x: x[1][0], reverse=True), 1):
        sz_mb = sz / (1024 * 1024)
        md_content += f"| {idx} | **{source}** | **{cnt:,} ta** | {sz_mb:.2f} MB | `.pptx` | ✅ Faol |\n"
        
    md_content += f"\n**JAMI**: **{total_files:,} ta taqdimot ({total_gb:.2f} GB)**\n"
    
    with open(MASTER_MD, 'w', encoding='utf-8') as f:
        f.write(md_content)
        
    print(f"🏆 Master Katalog yangilandi: {total_files:,} ta taqdimot ({total_gb:.2f} GB) qayd etildi.\n")

def process_slidehunter():
    print(f"\n==================================================")
    print(f"🚀 {SOURCE_NAME} (NICHE PLATFORMA) YUKLASH BOSHLANDI...")
    print(f"==================================================")
    
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)
    
    clean_urls = set()
    sitemaps = [
        'https://slidehunter.com/slide_product-sitemap.xml',
        'https://slidehunter.com/post-sitemap.xml',
        'https://slidehunter.com/post-sitemap2.xml'
    ]
    for sm in sitemaps:
        try:
            r = s.get(sm, timeout=15)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, 'xml')
                for loc in soup.find_all('loc'):
                    u = loc.text.strip()
                    if not any(u.lower().endswith(ext) for ext in ['.jpg', '.png', '.jpeg', '.webp', '.svg']):
                        clean_urls.add(u)
        except Exception as e:
            print(f"  Sitemap xatosi ({sm}): {e}")

    urls = list(clean_urls)
    print(f"📋 {SOURCE_NAME}: {len(urls)} ta toza taqdimot sahifasi aniqlandi. Parallel yuklash boshlanmoqda...")
    
    def process_item(page_url):
        try:
            res = s.get(page_url, timeout=15)
            if res.status_code != 200:
                return
            soup = BeautifulSoup(res.text, 'html.parser')
            
            h1 = soup.find('h1')
            title = h1.text.strip() if h1 else os.path.basename(page_url.rstrip('/'))
            
            if not title or is_duplicate(title):
                return
                
            cat_elem = soup.find('a', rel='tag') or soup.find('span', class_='posted-in')
            category = cat_elem.text.strip() if cat_elem else "Business & Diagrams"
            
            # Step 1: Find did link
            did_url = None
            for a in soup.find_all('a', href=True):
                if '/download-template/?did=' in a['href']:
                    did_url = urljoin(page_url, a['href'])
                    break
                    
            if not did_url:
                return
                
            # Step 2: Fetch did page to get actual download link
            dr = s.get(did_url, timeout=15)
            if dr.status_code != 200:
                return
            dsoup = BeautifulSoup(dr.text, 'html.parser')
            
            dl_url = None
            for a in dsoup.find_all('a', class_='download-link', href=True):
                dl_url = urljoin(did_url, a['href'])
                break
            if not dl_url:
                for a in dsoup.find_all('a', href=True):
                    if '/download/' in a['href']:
                        dl_url = urljoin(did_url, a['href'])
                        break
                        
            if not dl_url:
                return
                
            # Step 3: Download file (Zip containing PPTX/PPT)
            file_res = s.get(dl_url, timeout=25)
            if file_res.status_code == 200 and len(file_res.content) > 10 * 1024:
                content = file_res.content
                
                # Check if it's a zip file
                if content[:4] == b'PK\x03\x04':
                    try:
                        z = zipfile.ZipFile(io.BytesIO(content))
                        ppt_names = [n for n in z.namelist() if any(n.lower().endswith(ext) for ext in ['.pptx', '.ppt', '.pot', '.potx'])]
                        if ppt_names:
                            # Extract primary presentation file
                            best_ppt = sorted(ppt_names, key=lambda x: (1 if x.lower().endswith('.pptx') else 0, len(x)), reverse=True)[0]
                            ppt_content = z.read(best_ppt)
                            ext = os.path.splitext(best_ppt)[1].lower()
                            save_and_catalog(title, category, dl_url, ppt_content, ext)
                            return
                    except Exception:
                        pass
                
                # Fallback if raw PPTX/PPT
                save_and_catalog(title, category, dl_url, content, ".pptx")
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        futures = [executor.submit(process_item, u) for u in urls]
        for idx, f in enumerate(as_completed(futures), 1):
            if idx % 100 == 0 or idx == len(urls):
                print(f"  ⏳ SlideHunter taraqqiyoti: {idx}/{len(urls)} sahifa ko'rib chiqildi...")

def main():
    start_time = time.time()
    load_global_registry()
    process_slidehunter()
    rebuild_master_stats()
    
    elapsed = time.time() - start_time
    print(f"\n🎉 [YAKUNLANDI] SlideHunter to'liq yuklandi!")
    print(f"⏱️ Sarflangan vaqt: {elapsed / 60:.2f} daqiqa")
    print(f"📥 Jami yangi yuklangan taqdimotlar: {downloaded_count} ta")
    print(f"💾 Jami yangi yuklangan hajm: {downloaded_size_bytes / (1024 * 1024):.2f} MB\n")

if __name__ == "__main__":
    main()
