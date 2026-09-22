import os
import sys
import re
import csv
import time
import shutil
import hashlib
import requests
import warnings
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, unquote

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
sys.stdout.reconfigure(encoding='utf-8')

# --- CONFIGURATION ---
BASE_DIR = r"C:\Users\user\Desktop\Powepoint"
MASTER_CSV = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.csv")
MASTER_MD = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.md")

MIN_FREE_SPACE_GB = 1.5
NUM_THREADS = 20

SESSION_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
}

# --- GLOBAL DUPLICATION REGISTRY ---
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
        print(f"⚠️ [OGOHLANTIRISH] C: diskida bo'sh joy kam qoldi: {free_gb:.2f} GB! Jarayon to'xtatilmoqda.")
        return False
    return True

def load_global_registry():
    global registry_titles, registry_hashes
    print("==================================================")
    print("🔍 GLOBAL DUPLIKATSIYA REGISTRI YUKLANMOQDA...")
    print("==================================================")
    
    csv_sources = [
        r"C:\Users\user\Desktop\Powepoint\BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.csv",
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
                print(f"  Registr CSV o'qishda ogohlantirish ({path}): {e}")

    folder_sources = [
        r"C:\Users\user\Desktop\Powepoint",
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

def save_and_catalog(source_name, title, category, file_url, content, ext=".pptx"):
    global downloaded_count, downloaded_size_bytes
    
    if not check_disk_space():
        return False
        
    file_hash = hashlib.md5(content).hexdigest()
    if file_hash in registry_hashes:
        print(f"  ⏭️ [HASH DUBLIKAT] {title}")
        return False
        
    source_folder = os.path.join(BASE_DIR, source_name)
    presentations_folder = os.path.join(source_folder, "Prezentatsiyalar")
    
    safe_title = clean_filename(title)
    item_dir = os.path.join(presentations_folder, safe_title)
    os.makedirs(item_dir, exist_ok=True)
    
    target_file = os.path.join(item_dir, f"{safe_title}{ext}")
    with open(target_file, 'wb') as f:
        f.write(content)
        
    size_mb = len(content) / (1024 * 1024)
    downloaded_count += 1
    downloaded_size_bytes += len(content)
    add_to_registry(title, file_hash)
    
    # 1. Update Source Specific CSV
    source_csv = os.path.join(source_folder, f"{source_name}_Katalogi.csv")
    csv_is_new = not os.path.exists(source_csv)
    with open(source_csv, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if csv_is_new:
            writer.writerow(["#", "Manba", "Taqdimot Nomi", "Kategoriya", "Hajmi (MB)", "Fayl Turi", "MD5", "Mahalliy Manzil", "Sayt Havolasi", "Sana"])
        writer.writerow([
            downloaded_count,
            source_name,
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
            source_name,
            safe_title,
            category,
            f"{size_mb:.2f}",
            ext,
            target_file,
            file_url
        ])

    print(f"  📥 [{source_name}] Yuklandi: {safe_title} ({size_mb:.2f} MB)")
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

# ==============================================================================
# PIPELINE 1: PRESENTATION MAGAZINE (1,918+ posts)
# ==============================================================================
def process_presentation_magazine():
    source_name = "PresentationMagazine"
    print(f"\n==================================================")
    print(f"🚀 [1/3] {source_name} YUKLASH BOSHLANDI...")
    print(f"==================================================")
    
    sitemap_url = "https://www.presentationmagazine.com/wp-sitemap-posts-post-1.xml"
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)
    
    try:
        r = s.get(sitemap_url, timeout=20)
        soup = BeautifulSoup(r.text, 'xml')
        urls = [loc.text.strip() for loc in soup.find_all('loc')]
    except Exception as e:
        print(f"❌ Sitemap yuklashda xatolik: {e}")
        return

    print(f"📋 {source_name}: {len(urls)} ta sahifa topildi. Parallel tahlil va yuklash boshlanmoqda...")
    
    def process_single_page(page_url):
        try:
            res = s.get(page_url, timeout=15)
            if res.status_code != 200:
                return
            page_soup = BeautifulSoup(res.text, 'html.parser')
            
            # Title
            h1 = page_soup.find('h1')
            title = h1.text.strip() if h1 else ""
            if not title:
                title_tag = page_soup.find('title')
                title = title_tag.text.split('-')[0].strip() if title_tag else ""
            if not title:
                title = os.path.basename(page_url).replace('.htm', '').replace('-', ' ')
                
            if is_duplicate(title):
                return
                
            # Category
            cat_elem = page_soup.find('span', class_='posted-in') or page_soup.find('a', rel='category tag')
            category = cat_elem.text.strip() if cat_elem else "General"
            
            # Download link
            dl_links = []
            for a in page_soup.find_all('a', href=True):
                href = a['href']
                if any(href.lower().endswith(ext) for ext in ['.pptx', '.ppt', '.pot', '.potx', '.zip']):
                    dl_links.append(urljoin(page_url, href))
                elif '/powerpoint-templates/' in href and any(ext in href.lower() for ext in ['.pptx', '.ppt', '.zip']):
                    dl_links.append(urljoin(page_url, href))

            if not dl_links:
                return
                
            target_url = dl_links[0]
            ext = ".pptx" if ".pptx" in target_url.lower() else (".ppt" if ".ppt" in target_url.lower() else ".zip")
            
            # Download file
            file_res = s.get(target_url, timeout=25)
            if file_res.status_code == 200 and len(file_res.content) > 10 * 1024:
                save_and_catalog(source_name, title, category, target_url, file_res.content, ext)
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        futures = [executor.submit(process_single_page, u) for u in urls]
        for idx, f in enumerate(as_completed(futures), 1):
            if idx % 100 == 0 or idx == len(urls):
                print(f"  ⏳ PresentationMagazine taraqqiyoti: {idx}/{len(urls)} sahifa ko'rib chiqildi...")

# ==============================================================================
# PIPELINE 2: SLIDESMANIA (Remaining 330+ Templates)
# ==============================================================================
def process_slidesmania():
    source_name = "SlidesMania"
    print(f"\n==================================================")
    print(f"🚀 [2/3] {source_name} QO'SHIMCHA YUKLASH BOSHLANDI...")
    print(f"==================================================")
    
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)
    
    all_urls = []
    for sitemap_url in ["https://slidesmania.com/post-sitemap.xml", "https://slidesmania.com/post-sitemap2.xml"]:
        try:
            r = s.get(sitemap_url, timeout=15)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, 'xml')
                all_urls.extend([loc.text.strip() for loc in soup.find_all('loc')])
        except Exception:
            pass
            
    print(f"📋 {source_name}: {len(all_urls)} ta sahifa topildi.")
    
    def process_single_page(page_url):
        try:
            res = s.get(page_url, timeout=15)
            if res.status_code != 200:
                return
            page_soup = BeautifulSoup(res.text, 'html.parser')
            
            h1 = page_soup.find('h1')
            title = h1.text.strip() if h1 else ""
            if not title or is_duplicate(title):
                return
                
            category = "Education & Creative"
            
            pptx_url = None
            for a in page_soup.find_all('a', href=True):
                href = a['href']
                if '.pptx' in href.lower() or 'download=powerpoint' in href.lower() or 'action=download_pptx' in href.lower():
                    pptx_url = urljoin(page_url, href)
                    break
                elif 'docs.google.com/presentation' in href and 'export/pptx' in href:
                    pptx_url = href
                    break
                    
            if not pptx_url:
                for a in page_soup.find_all('a', href=True):
                    href = a['href']
                    if 'docs.google.com/presentation/d/' in href:
                        match = re.search(r'/presentation/d/([a-zA-Z0-9_-]+)', href)
                        if match:
                            doc_id = match.group(1)
                            pptx_url = f"https://docs.google.com/presentation/d/{doc_id}/export/pptx"
                            break

            if not pptx_url:
                return
                
            file_res = s.get(pptx_url, timeout=25)
            if file_res.status_code == 200 and len(file_res.content) > 15 * 1024:
                save_and_catalog(source_name, title, category, pptx_url, file_res.content, ".pptx")
        except Exception:
            pass

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        futures = [executor.submit(process_single_page, u) for u in all_urls]
        for idx, f in enumerate(as_completed(futures), 1):
            if idx % 100 == 0 or idx == len(all_urls):
                print(f"  ⏳ SlidesMania taraqqiyoti: {idx}/{len(all_urls)} sahifa ko'rib chiqildi...")

# ==============================================================================
# MAIN SUPERVISOR RUNNER
# ==============================================================================
def main():
    start_time = time.time()
    load_global_registry()
    
    # 1. Presentation Magazine
    process_presentation_magazine()
    rebuild_master_stats()
    
    # 2. SlidesMania
    process_slidesmania()
    rebuild_master_stats()
    
    elapsed = time.time() - start_time
    print(f"\n🎉 [YAKUNLANDI] Jarayon muvaffaqiyatli yakunlandi!")
    print(f"⏱️ Sarflangan vaqt: {elapsed / 60:.2f} daqiqa")
    print(f"📥 Jami yangi yuklangan taqdimotlar: {downloaded_count} ta")
    print(f"💾 Jami yangi yuklangan hajm: {downloaded_size_bytes / (1024 * 1024):.2f} MB\n")

if __name__ == "__main__":
    main()
