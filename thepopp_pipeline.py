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
SOURCE_NAME = "ThePOPP"
SOURCE_DIR = os.path.join(BASE_DIR, SOURCE_NAME)
PREZ_DIR = os.path.join(SOURCE_DIR, "Prezentatsiyalar")
SOURCE_CSV = os.path.join(SOURCE_DIR, f"{SOURCE_NAME}_Katalogi.csv")
SOURCE_MD = os.path.join(SOURCE_DIR, f"{SOURCE_NAME}_Katalogi.md")

MASTER_CSV = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.csv")
MASTER_MD = os.path.join(BASE_DIR, "BARCHA_TAQDIMOTLAR_MASTER_KATALOGI.md")

MIN_FREE_SPACE_GB = 1.5
NUM_THREADS = 5

SESSION_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
}

registry_titles = set()
registry_hashes = set()

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

def load_global_registry():
    global registry_titles, registry_hashes
    if os.path.exists(MASTER_CSV):
        try:
            with open(MASTER_CSV, 'r', encoding='utf-8', errors='ignore') as f:
                reader = csv.reader(f)
                for row in reader:
                    for col in row[:4]:
                        n = clean_name(col)
                        if len(n) > 3 and not n.isdigit():
                            registry_titles.add(n)
        except Exception:
            pass

def is_duplicate(title, file_hash=None):
    cn = clean_name(title)
    if cn in registry_titles:
        return True
    if file_hash and file_hash.lower() in registry_hashes:
        return True
    return False

def save_and_catalog(title, category, file_url, content, ext=".pptx"):
    file_hash = hashlib.md5(content).hexdigest()
    if file_hash in registry_hashes:
        return False
        
    safe_title = clean_filename(title)
    item_dir = os.path.join(PREZ_DIR, safe_title)
    os.makedirs(item_dir, exist_ok=True)
    
    target_file = os.path.join(item_dir, f"{safe_title}{ext}")
    with open(target_file, 'wb') as f:
        f.write(content)
        
    size_mb = len(content) / (1024 * 1024)
    registry_titles.add(clean_name(title))
    registry_hashes.add(file_hash)
    
    csv_is_new = not os.path.exists(SOURCE_CSV)
    with open(SOURCE_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if csv_is_new:
            writer.writerow(["#", "Manba", "Taqdimot Nomi", "Kategoriya", "Hajmi (MB)", "Fayl Turi", "MD5", "Mahalliy Manzil", "Sayt Havolasi", "Sana"])
        writer.writerow([1, SOURCE_NAME, safe_title, category, f"{size_mb:.2f}", ext, file_hash, target_file, file_url, time.strftime("%Y-%m-%d %H:%M:%S")])
        
    with open(MASTER_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["NEW", SOURCE_NAME, safe_title, category, f"{size_mb:.2f}", ext, target_file, file_url])

    print(f"  📥 [{SOURCE_NAME}] Yuklandi: {safe_title} ({size_mb:.2f} MB)")
    return True

def process_thepopp():
    print(f"\n==================================================")
    print(f"🚀 {SOURCE_NAME} (PREMIUM DIZAYN) YUKLASH BOSHLANDI...")
    print(f"==================================================")
    s = requests.Session()
    s.headers.update(SESSION_HEADERS)
    
    all_templates = set()
    for page in range(1, 5):
        url = f'https://thepopp.com/templates/page/{page}/' if page > 1 else 'https://thepopp.com/templates/'
        try:
            r = s.get(url, timeout=10)
            if r.status_code != 200:
                break
            soup = BeautifulSoup(r.text, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                if '/templates/' in href and href != 'https://thepopp.com/templates/' and '/page/' not in href:
                    all_templates.add(href)
        except Exception:
            break

    print(f"📋 {SOURCE_NAME}: {len(all_templates)} ta shablon topildi.")
    
    for t_url in all_templates:
        try:
            res = s.get(t_url, timeout=15)
            if res.status_code != 200:
                continue
            soup = BeautifulSoup(res.text, 'html.parser')
            h1 = soup.find('h1')
            title = h1.text.strip() if h1 else os.path.basename(t_url.rstrip('/'))
            
            if is_duplicate(title):
                continue
                
            dl_btn = None
            for a in soup.find_all('a', href=True):
                if '/download/' in a['href']:
                    dl_btn = urljoin(t_url, a['href'])
                    break
                    
            if not dl_btn:
                continue
                
            file_res = s.get(dl_btn, timeout=25)
            if file_res.status_code == 200 and len(file_res.content) > 10 * 1024:
                content = file_res.content
                if content[:4] == b'PK\x03\x04':
                    try:
                        z = zipfile.ZipFile(io.BytesIO(content))
                        ppt_names = [n for n in z.namelist() if any(n.lower().endswith(ext) for ext in ['.pptx', '.ppt'])]
                        if ppt_names:
                            best_ppt = sorted(ppt_names, key=lambda x: (1 if 'noanimation' not in x.lower() else 0, len(x)), reverse=True)[0]
                            ppt_content = z.read(best_ppt)
                            ext = os.path.splitext(best_ppt)[1].lower()
                            save_and_catalog(title, "Creative & Design", dl_btn, ppt_content, ext)
                            continue
                    except Exception:
                        pass
                save_and_catalog(title, "Creative & Design", dl_btn, content, ".pptx")
        except Exception as e:
            print(f"  Xato ({t_url}): {e}")

def main():
    load_global_registry()
    process_thepopp()
    print(f"✅ ThePOPP yakunlandi.")

if __name__ == "__main__":
    main()
