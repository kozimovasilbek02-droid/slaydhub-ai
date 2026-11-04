# -*- coding: utf-8 -*-
"""
core/image_fetcher.py
High-resolution, topic-relevant scientific & academic image fetcher for SlaydHub AI.
Searches Wikimedia Commons and open educational media sources to fetch authentic
diagrams, illustrations, and photos matching the presentation's scientific topic.
Caches images locally in output/image_cache/ for instant reuse.
Enforces strict 512MB RAM & DecompressionBomb safeguards (max 1280x720, < 2.5MP).
"""

import os
import re
import io
import json
import hashlib
import logging
import urllib.request
import urllib.parse
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from PIL import Image

# 512MB RAM Himoyasi: DecompressionBomb hujumlari va gigant rasmlardan saqlanish
Image.MAX_IMAGE_PIXELS = 10_000_000

from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

from core.config import config
from core.media_harmonizer import MediaHarmonizer

logger = logging.getLogger("ImageFetcher")

CACHE_DIR = config.OUTPUT_ROOT / "image_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

# Tezkor in-memory kesh (bir xil qidiruvlarni qayta-qayta bajarmaslik uchun)
_SEARCH_CACHE: Dict[str, Optional[str]] = {}

# Curated English search term mappings for popular academic & technical domains
TOPIC_KEYWORD_MAP: Dict[str, List[str]] = {
    "kvant": ["Quantum computing", "Quantum processor", "Bloch sphere", "Superconducting quantum computing"],
    "kripto": ["Cryptography", "Cryptographic key", "Lattice-based cryptography", "Cipher"],
    "xavfsizlik": ["Computer security", "Cybersecurity", "Network security topology", "Data encryption"],
    "tibbiyot": ["Medicine", "Medical research", "Microscope laboratory", "Biomedical engineering"],
    "biologiya": ["DNA structure", "Molecular biology", "Cell biology diagram", "Genetics"],
    "sun'iy intellekt": ["Artificial intelligence", "Neural network diagram", "Deep learning graph", "Machine learning"],
    "ai": ["Artificial neural network", "Machine learning", "Data visualization", "Algorithm flowchart"],
    "iqtisod": ["Economics chart", "Global finance", "Stock market analysis", "Macroeconomics"],
    "energetika": ["Solar energy panel", "Wind turbine renewable energy", "Clean energy transition"],
    "fizika": ["Particle physics", "Electromagnetic wave", "Atomic orbital", "Optics laser"],
    "matematika": ["Mathematical graph", "Geometry fractal", "Calculus surface", "Topology"],
    "tarix": ["Historical manuscript", "Ancient architecture", "Archaeological artifact"]
}


class AcademicImageFetcher:
    """
    Fetches, resizes, and caches topic-relevant academic diagrams and photos.
    """

    @classmethod
    def extract_search_terms(cls, topic: str, slide_title: str = "") -> List[str]:
        """Derives the best English search terms from Uzbek/Russian topic titles."""
        combined = f"{topic} {slide_title}".lower()
        terms = []

        for key, query_list in TOPIC_KEYWORD_MAP.items():
            if key in combined:
                terms.extend(query_list)

        if not terms:
            clean = re.sub(r'[^a-zA-Z0-9\s]', ' ', topic).strip()
            terms.append(clean if len(clean) > 3 else "Science technology research")

        return terms[:3]

    @classmethod
    def search_wikimedia_image(cls, query: str) -> Optional[str]:
        """
        Searches Wikimedia Commons for a safe-sized, public domain or CC-licensed image.
        Returns the image download URL.
        """
        if query in _SEARCH_CACHE:
            return _SEARCH_CACHE[query]

        try:
            encoded_q = urllib.parse.quote(query)
            url = (
                f"https://commons.wikimedia.org/w/api.php?action=query&generator=search"
                f"&gsrsearch={encoded_q}&gsrnamespace=6&gsrlimit=5"
                f"&prop=imageinfo&iiprop=url|mime|size|dimensions&format=json"
            )
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "SlaydHubAI/5.2 (academic-presentation-studio; contact: dev@slaydhub.ai)"}
            )

            with urllib.request.urlopen(req, timeout=4) as response:
                data = json.loads(response.read().decode("utf-8"))

            pages = data.get("query", {}).get("pages", {})
            for pid, pinfo in pages.items():
                img_info = pinfo.get("imageinfo", [{}])[0]
                img_url = img_info.get("url", "")
                mime = img_info.get("mime", "").lower()
                size = img_info.get("size", 0)
                width = img_info.get("width", 0)
                height = img_info.get("height", 0)

                # 512MB RAM himoyasi: Decompression Bomb rasmlarni (150MP) chetlab o'tish
                if width * height > 2_500_000 or width > 2560 or height > 1600:
                    logger.debug(f"Hajmi juda katta rasm o'tkazib yuborildi ({width}x{height}): {img_url}")
                    continue

                if img_url and any(mime.endswith(ext) for ext in ["jpeg", "jpg", "png"]):
                    if 10_000 < size < 6_000_000:
                        _SEARCH_CACHE[query] = img_url
                        return img_url

        except Exception as e:
            logger.debug(f"Wikimedia search error for query '{query}': {e}")

        _SEARCH_CACHE[query] = None
        return None

    @classmethod
    def download_and_cache_image(cls, img_url: str) -> Optional[str]:
        """Downloads image, downscales to max 1280x720, and caches as clean JPEG (<200KB)."""
        try:
            url_hash = hashlib.md5(img_url.encode("utf-8")).hexdigest()[:12]
            cache_path = CACHE_DIR / f"academic_{url_hash}.jpg"

            if cache_path.exists() and cache_path.stat().st_size > 1024:
                return str(cache_path)

            req = urllib.request.Request(
                img_url,
                headers={"User-Agent": "SlaydHubAI/5.2 (academic-presentation-studio)"}
            )
            with urllib.request.urlopen(req, timeout=6) as resp:
                img_data = resp.read()

            if len(img_data) < 1000 or len(img_data) > 8_000_000:
                return None

            # PIL orqali darhol 1280x720 gacha kichraytirish va RGB JPEG sifatida saqlash
            with Image.open(io.BytesIO(img_data)) as im:
                if im.mode in ("RGBA", "P"):
                    im = im.convert("RGB")
                im.thumbnail((1280, 720), Image.Resampling.LANCZOS)
                im.save(cache_path, format="JPEG", quality=82, optimize=True)

            return str(cache_path)

        except Exception as e:
            logger.debug(f"Failed to download/cache image from '{img_url}': {e}")
            return None

    @classmethod
    def get_image_for_slide(cls, topic: str, slide_title: str = "", archetype: str = "") -> Optional[str]:
        """
        Retrieves a local high-quality image path matching the slide topic.
        Render 512MB RAM & Zero-Hang: Does not perform blocking live HTTP requests
        during real-time slide generation unless ENABLE_LIVE_IMAGE_SCRAPING=1.
        """
        queries = cls.extract_search_terms(topic, slide_title)

        # 1. Avval mavjud keshdan tekshiramiz (bir lahzalik <1ms)
        for q in queries:
            url_hash = hashlib.md5(q.encode("utf-8")).hexdigest()[:12]
            cached_cand = CACHE_DIR / f"academic_{url_hash}.jpg"
            if cached_cand.exists() and cached_cand.stat().st_size > 1024:
                return str(cached_cand)

        # 2. Jonli internet qidiruvi yoqilmagan bo'lsa (default), server qotib qolmasligi uchun None qaytaramiz
        if os.getenv("ENABLE_LIVE_IMAGE_SCRAPING", "0").lower() not in ("1", "true", "yes"):
            return None

        # 3. Jonli qidiruv (faqat ruxsat berilganda va qisqa timeout bilan)
        for q in queries:
            img_url = cls.search_wikimedia_image(q)
            if img_url:
                local_path = cls.download_and_cache_image(img_url)
                if local_path and os.path.exists(local_path):
                    return local_path

        return None

    @classmethod
    def inject_image_into_slide(
        cls,
        slide,
        image_path: str,
        sw_pt: float = 960.0,
        sh_pt: float = 540.0,
        scale_factor: float = 1.0,
        layout_type: str = ""
    ) -> bool:
        """
        Injects an academic image into an unfilled picture placeholder or dedicated graphic zone.
        NEVER overwrites authentic template photos or graphics to preserve design fidelity.
        """
        if not image_path or not os.path.exists(image_path):
            return False

        try:
            # 1. Faqat bo'sh qolgan rasmiy rasm joylagichlari (picture placeholder) mavjud bo'lsa to'ldiramiz
            for ph in slide.placeholders:
                try:
                    if ph.placeholder_format.type in (6, 18):  # PICTURE or MEDIA placeholder
                        slide.shapes.add_picture(
                            image_path, ph.left, ph.top, ph.width, ph.height
                        )
                        slide.shapes._spTree.remove(ph._element)
                        logger.info(f"Populated picture placeholder with '{image_path}'")
                        return True
                except Exception:
                    pass

            # 2. Agar slayd dizaynida fotosuratlar bo'lmasa va maxsus diagramma maketi bo'lsa
            existing_images = MediaHarmonizer.find_image_shapes(slide)
            if existing_images:
                # Shablonning o'z fotosurati bor, uni buzmasdan saqlaymiz!
                return False

            image_friendly_layouts = {
                "theory_concept", "case_study", "diagram_anatomy", 
                "problem_solution", "qa_discussion"
            }
            if layout_type and layout_type not in image_friendly_layouts:
                return False

            occupied_boxes = []
            for s in slide.shapes:
                try:
                    l = s.left.pt if hasattr(s.left, "pt") else float(s.left) / 12700.0
                    t = s.top.pt if hasattr(s.top, "pt") else float(s.top) / 12700.0
                    w = s.width.pt if hasattr(s.width, "pt") else float(s.width) / 12700.0
                    h = s.height.pt if hasattr(s.height, "pt") else float(s.height) / 12700.0
                    occupied_boxes.append((l, t, l + w, t + h))
                except Exception:
                    pass

            req_w = 4.0 * 72.0 * scale_factor
            req_h = 3.2 * 72.0 * scale_factor
            safe_l, safe_t = MediaHarmonizer.find_safe_image_placement(sw_pt, sh_pt, occupied_boxes, req_w, req_h)

            pic = slide.shapes.add_picture(
                image_path,
                Inches(safe_l / 72.0),
                Inches(safe_t / 72.0),
                Inches(req_w / 72.0),
                Inches(req_h / 72.0)
            )
            MediaHarmonizer.replace_image_with_file(pic, image_path)
            logger.info(f"Added collision-free illustration '{image_path}' at ({safe_l:.1f}, {safe_t:.1f})")
            return True

        except Exception as e:
            logger.debug(f"Could not inject image into slide: {e}")
            return False
