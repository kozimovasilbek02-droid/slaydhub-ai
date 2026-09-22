import os
import re
import sys
import json
import time
import shutil
import urllib.parse
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import requests

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Base configuration
BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "downloaded_templates"
CATALOG_FILE = OUTPUT_DIR / "catalog.json"
MANIFEST_FILE = OUTPUT_DIR / "manifest.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
}

CATEGORY_KEYWORDS = {
    "Business_Corporate": [
        "business", "corporate", "startup", "pitch", "deck", "company", "strategy",
        "roadmap", "meeting", "finance", "agenda", "proposal", "consulting", "office",
        "career", "management", "work", "project", "resume", "cv", "invoice", "plan",
        "report", "investment", "executive", "kpi", "operations", "commercial", "enterprise"
    ],
    "Education_Academic": [
        "education", "school", "lesson", "teacher", "student", "history", "math",
        "geography", "biology", "science", "thesis", "classroom", "university", "college",
        "academic", "curriculum", "grade", "preschool", "english", "physics", "chemistry",
        "exam", "homework", "lecture", "grammar", "literature", "kindergarten", "study", "slidesclass"
    ],
    "Marketing_SocialMedia": [
        "marketing", "sales", "social-media", "instagram", "tiktok", "youtube", "brand",
        "product", "advertising", "campaign", "lead", "seo", "customer", "funnel",
        "market", "ecommerce", "influencer", "promo", "newsletter", "public-relations", "content"
    ],
    "Infographics_Charts": [
        "infographic", "chart", "diagram", "timeline", "matrix", "comparison", "dashboard",
        "graph", "pyramid", "cycle", "steps", "funnel", "workflow", "stats", "data", "process",
        "hierarchy", "tables", "venn"
    ],
    "Medical_Healthcare": [
        "medical", "health", "doctor", "hospital", "clinic", "pharmacy", "care",
        "disease", "anatomy", "mental", "dental", "nurse", "wellness", "fitness",
        "nutrition", "therapy", "medicine", "virus", "dna", "body", "healthcare"
    ],
    "Creative_Design": [
        "creative", "aesthetic", "art", "photography", "portfolio", "modern", "vintage",
        "geometric", "pastel", "dark", "cute", "colorful", "minimal", "minimalist",
        "elegant", "fashion", "typography", "drawing", "retro", "neon", "abstract", "illustration"
    ],
    "Culture_Events_Holidays": [
        "holiday", "christmas", "halloween", "party", "event", "wedding", "festival",
        "birthday", "culture", "music", "travel", "food", "sports", "game", "celebration",
        "ramadan", "easter", "thanksgiving", "new-year", "valentine"
    ]
}

def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name[:120] if len(name) > 120 else name

def categorize_template(title: str, url: str) -> str:
    slug = url.rstrip('/').split('/')[-1].replace('-', ' ')
    combined = f"{title} {slug}".lower()
    for cat, keywords in CATEGORY_KEYWORDS.items():
        if any(re.search(r'\b' + re.escape(kw) + r'\b', combined) for kw in keywords):
            return cat
    return "General_Templates"

def get_sitemap_urls() -> list:
    print("🔍 Fetching sitemaps from SlidesCarnival...", flush=True)
    post_sitemaps = [
        f"https://www.slidescarnival.com/post-sitemap{i if i > 1 else ''}.xml"
        for i in range(1, 9)
    ]
    all_urls = []
    session = requests.Session()
    session.headers.update(HEADERS)
    
    for sm in post_sitemaps:
        try:
            resp = session.get(sm, timeout=15)
            if resp.status_code == 200:
                root = ET.fromstring(resp.content)
                urls = [elem.text for elem in root.findall('.//{http://www.sitemaps.org/schemas/sitemap/0.9}loc')]
                all_urls.extend(urls)
                print(f"  ✓ {sm.split('/')[-1]}: {len(urls)} URLs", flush=True)
        except Exception as e:
            print(f"  ⚠️ Error fetching {sm}: {e}", flush=True)
            
    print(f"Total candidate URLs from sitemaps: {len(all_urls)}", flush=True)
    return list(set(all_urls))

def extract_template_info(url: str, session: requests.Session) -> dict:
    try:
        resp = session.get(url, timeout=12)
        if resp.status_code != 200:
            return None
        html = resp.text
        
        # Check for direct download pptx link
        pptx_links = re.findall(r'href=[\'"]([^\'"]+?/pptx)[\'"]', html)
        if not pptx_links:
            # Check GCS direct links
            pptx_links = re.findall(r'(https://storage\.googleapis\.com/sc_templates/[^\'"]+\.pptx)', html)
            
        if not pptx_links:
            sc_match = re.search(r'https://www\.slidescarnival\.com/download/([a-zA-Z0-9_-]+)/([a-zA-Z0-9_-]+)/pptx', html)
            if sc_match:
                pptx_links = [sc_match.group(0)]
                
        if not pptx_links:
            return None
            
        download_url = pptx_links[0]
        
        # Extract clean title
        title_match = re.search(r'<title>(.*?)</title>', html, re.IGNORECASE)
        title = title_match.group(1).split('.')[0].split('–')[0].split('-')[0].strip() if title_match else ""
        if not title or "SlidesCarnival" in title:
            slug = url.rstrip('/').split('/')[-1]
            title = slug.replace('-', ' ').title()
            
        category = categorize_template(title, url)
        
        return {
            "page_url": url,
            "title": title,
            "category": category,
            "download_url": download_url,
        }
    except Exception:
        return None

def build_catalog(candidate_urls: list, max_workers: int = 15) -> list:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    existing_catalog = {}
    if CATALOG_FILE.exists():
        try:
            with open(CATALOG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                existing_catalog = {item['page_url']: item for item in data}
                print(f"Loaded {len(existing_catalog)} existing templates from catalog.", flush=True)
        except Exception:
            pass

    urls_to_fetch = [u for u in candidate_urls if u not in existing_catalog]
    print(f"🔎 Scanning {len(urls_to_fetch)} new candidate pages for PPTX downloads...", flush=True)
    
    catalog = list(existing_catalog.values())
    processed = 0
    found = len(catalog)

    with requests.Session() as session:
        session.headers.update(HEADERS)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_url = {executor.submit(extract_template_info, u, session): u for u in urls_to_fetch}
            for future in as_completed(future_to_url):
                processed += 1
                res = future.result()
                if res:
                    catalog.append(res)
                    found += 1
                if processed % 50 == 0 or processed == len(urls_to_fetch):
                    print(f"  Progress: {processed}/{len(urls_to_fetch)} scanned | Found PPTX templates: {found}", flush=True)
                    with open(CATALOG_FILE, 'w', encoding='utf-8') as f:
                        json.dump(catalog, f, indent=2, ensure_ascii=False)

    # Re-categorize all items with the latest categories
    for item in catalog:
        item['category'] = categorize_template(item.get('title', ''), item.get('page_url', ''))

    print(f"✅ Catalog build complete! Total PPTX templates found: {len(catalog)}", flush=True)
    with open(CATALOG_FILE, 'w', encoding='utf-8') as f:
        json.dump(catalog, f, indent=2, ensure_ascii=False)
    return catalog

def download_single_template(item: dict, session: requests.Session) -> dict:
    url = item['download_url']
    category = item.get('category', 'General_Templates')
    title = sanitize_filename(item.get('title', 'Template'))
    cat_dir = OUTPUT_DIR / category
    cat_dir.mkdir(parents=True, exist_ok=True)

    dest_file = cat_dir / f"{title}.pptx"
    
    # If file already exists and is valid PPTX, skip
    if dest_file.exists() and dest_file.stat().st_size > 50000:
        return {
            "status": "skipped",
            "title": title,
            "category": category,
            "path": str(dest_file),
            "size": dest_file.stat().st_size
        }

    max_retries = 3
    for attempt in range(max_retries):
        try:
            resp = session.get(url, timeout=30, stream=True, allow_redirects=True)
            if resp.status_code == 200:
                temp_file = dest_file.with_suffix('.tmp')
                with open(temp_file, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)
                
                file_size = temp_file.stat().st_size
                if file_size > 10000:
                    with open(temp_file, 'rb') as f:
                        header = f.read(4)
                    if header == b'PK\x03\x04':
                        if dest_file.exists():
                            dest_file.unlink()
                        temp_file.rename(dest_file)
                        return {
                            "status": "downloaded",
                            "title": title,
                            "category": category,
                            "path": str(dest_file),
                            "size": file_size
                        }
                if temp_file.exists():
                    temp_file.unlink()
        except Exception:
            time.sleep(1 + attempt)

    return {
        "status": "failed",
        "title": title,
        "category": category,
        "url": url,
        "error": "Download failed or invalid PPTX file"
    }

def print_status_summary():
    if not OUTPUT_DIR.exists():
        print("No templates directory found.")
        return

    print("=" * 60)
    print("📊 SLIDESCARNIVAL DOWNLOAD STATUS SUMMARY")
    print(f"Output Directory: {OUTPUT_DIR}")
    print("=" * 60)

    total_files = 0
    total_size_bytes = 0
    cat_counts = {}

    for cat_dir in OUTPUT_DIR.iterdir():
        if cat_dir.is_dir():
            files = list(cat_dir.glob("*.pptx"))
            cat_size = sum(f.stat().st_size for f in files)
            cat_counts[cat_dir.name] = (len(files), cat_size)
            total_files += len(files)
            total_size_bytes += cat_size

    for cat, (count, size_bytes) in sorted(cat_counts.items(), key=lambda x: -x[1][0]):
        size_mb = size_bytes / (1024 * 1024)
        print(f"  📂 {cat.ljust(25)}: {count:4d} files ({size_mb:7.1f} MB)")

    total_mb = total_size_bytes / (1024 * 1024)
    total_gb = total_mb / 1024
    print("-" * 60)
    print(f"  🏆 TOTAL DOWNLOADED: {total_files} PPTX files ({total_mb:.1f} MB / {total_gb:.2f} GB)")
    print("=" * 60)

def run_downloader(target_category: str = "all", limit: int = 0, workers: int = 8):
    print("=" * 60, flush=True)
    print("🚀 SLIDESCARNIVAL BULK PPTX DOWNLOADER", flush=True)
    print(f"Output Directory: {OUTPUT_DIR}", flush=True)
    print(f"Workers: {workers} | Category filter: {target_category} | Limit: {'No limit' if limit <= 0 else limit}", flush=True)
    print("=" * 60, flush=True)

    # Step 1: Load or create catalog
    catalog = []
    if CATALOG_FILE.exists():
        try:
            with open(CATALOG_FILE, 'r', encoding='utf-8') as f:
                catalog = json.load(f)
                print(f"Loaded existing catalog with {len(catalog)} templates.", flush=True)
        except Exception:
            pass

    if not catalog:
        candidate_urls = get_sitemap_urls()
        catalog = build_catalog(candidate_urls, max_workers=15)

    # Ensure all items have fresh categories
    for item in catalog:
        item['category'] = categorize_template(item.get('title', ''), item.get('page_url', ''))

    # Filter catalog
    if target_category.lower() != "all":
        catalog = [item for item in catalog if item.get('category', '').lower() == target_category.lower()]
        print(f"Filtered by category '{target_category}': {len(catalog)} templates", flush=True)

    if limit > 0:
        catalog = catalog[:limit]
        print(f"Applied limit: downloading top {len(catalog)} templates", flush=True)

    # Step 2: Load manifest
    manifest = {}
    if MANIFEST_FILE.exists():
        try:
            with open(MANIFEST_FILE, 'r', encoding='utf-8') as f:
                manifest = json.load(f)
        except Exception:
            manifest = {}

    print(f"\n📥 Starting download for {len(catalog)} templates...", flush=True)
    downloaded_count = 0
    skipped_count = 0
    failed_count = 0
    total_bytes = 0
    start_time = time.time()

    with requests.Session() as session:
        session.headers.update(HEADERS)
        with ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_item = {executor.submit(download_single_template, item, session): item for item in catalog}
            completed = 0
            total = len(future_to_item)

            for future in as_completed(future_to_item):
                completed += 1
                res = future.result()
                status = res.get('status')
                item_url = future_to_item[future]['download_url']
                manifest[item_url] = res

                if status == "downloaded":
                    downloaded_count += 1
                    total_bytes += res.get('size', 0)
                    size_mb = res.get('size', 0) / (1024 * 1024)
                    print(f"[{completed}/{total}] ✓ [{res['category']}] {res['title']} ({size_mb:.1f} MB)", flush=True)
                elif status == "skipped":
                    skipped_count += 1
                else:
                    failed_count += 1
                    print(f"[{completed}/{total}] ✗ Failed: {res.get('title')}", flush=True)

                if completed % 10 == 0 or completed == total:
                    with open(MANIFEST_FILE, 'w', encoding='utf-8') as f:
                        json.dump(manifest, f, indent=2, ensure_ascii=False)

    elapsed = time.time() - start_time
    total_mb = total_bytes / (1024 * 1024)
    print("\n" + "=" * 60, flush=True)
    print("🎉 DOWNLOAD BATCH COMPLETE!", flush=True)
    print(f"✓ Newly Downloaded: {downloaded_count} files ({total_mb:.1f} MB)", flush=True)
    print(f"⏭ Skipped (Already existed): {skipped_count}", flush=True)
    print(f"✗ Failed: {failed_count}", flush=True)
    print(f"⏱ Time elapsed: {elapsed:.1f} seconds", flush=True)
    print("=" * 60, flush=True)
    print_status_summary()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Bulk download PPTX templates from SlidesCarnival")
    parser.add_argument("--discover", action="store_true", help="Only discover and build catalog")
    parser.add_argument("--status", action="store_true", help="Show download status summary")
    parser.add_argument("--category", default="all", help="Category to download (default: all)")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of templates (0 for all)")
    parser.add_argument("--workers", type=int, default=8, help="Number of download threads (default: 8)")
    args = parser.parse_args()

    if args.status:
        print_status_summary()
    elif args.discover:
        urls = get_sitemap_urls()
        build_catalog(urls, max_workers=15)
    else:
        run_downloader(target_category=args.category, limit=args.limit, workers=args.workers)
