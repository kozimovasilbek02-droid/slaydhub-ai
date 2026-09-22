# -*- coding: utf-8 -*-
"""
core/image_fetcher.py
High-resolution, topic-relevant scientific & academic image fetcher for SlaydHub AI.
Searches Wikimedia Commons and open educational media sources to fetch authentic
diagrams, illustrations, and photos matching the presentation's scientific topic.
Caches images locally in output/image_cache/ for instant reuse.
"""

import os
import re
import json
import hashlib
import logging
import urllib.request
import urllib.parse
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
from PIL import Image

from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

from core.config import config
from core.media_harmonizer import MediaHarmonizer

logger = logging.getLogger("ImageFetcher")

CACHE_DIR = config.OUTPUT_ROOT / "image_cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

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
            # Clean non-alphanumeric words and use topic directly
            clean = re.sub(r'[^a-zA-Z0-9\s]', ' ', topic).strip()
            terms.append(clean if len(clean) > 3 else "Science technology research")

        return terms[:3]

    @classmethod
    def search_wikimedia_image(cls, query: str) -> Optional[str]:
        """
        Searches Wikimedia Commons for a high-resolution, public domain or CC-licensed image.
        Returns the image download URL.
        """
        try:
            encoded_q = urllib.parse.quote(query)
            url = (
                f"https://commons.wikimedia.org/w/api.php?action=query&generator=search"
                f"&gsrsearch={encoded_q}&gsrnamespace=6&gsrlimit=5"
                f"&prop=imageinfo&iiprop=url|mime|size&format=json"
            )
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "SlaydHubAI/3.5 (academic-presentation-studio; contact: dev@slaydhub.ai)"}
            )

            with urllib.request.urlopen(req, timeout=6) as response:
                data = json.loads(response.read().decode("utf-8"))

            pages = data.get("query", {}).get("pages", {})
            for pid, pinfo in pages.items():
                img_info = pinfo.get("imageinfo", [{}])[0]
                img_url = img_info.get("url", "")
                mime = img_info.get("mime", "").lower()
                size = img_info.get("size", 0)

                # Prioritize valid JPG and PNG raster images (avoid WebM or giant files > 15MB)
                if img_url and any(mime.endswith(ext) for ext in ["jpeg", "jpg", "png"]):
                    if 10_000 < size < 15_000_000:
                        return img_url

        except Exception as e:
            logger.debug(f"Wikimedia search error for query '{query}': {e}")

        return None

    @classmethod
    def download_and_cache_image(cls, img_url: str) -> Optional[str]:
        """Downloads the image from URL and caches it locally as a clean PNG/JPG."""
        try:
            url_hash = hashlib.md5(img_url.encode("utf-8")).hexdigest()[:12]
            ext = ".jpg" if ".png" not in img_url.lower() else ".png"
            cache_path = CACHE_DIR / f"academic_{url_hash}{ext}"

            if cache_path.exists() and cache_path.stat().st_size > 1024:
                return str(cache_path)

            req = urllib.request.Request(
                img_url,
                headers={"User-Agent": "SlaydHubAI/3.5 (academic-presentation-studio)"}
            )
            with urllib.request.urlopen(req, timeout=10) as resp:
                img_data = resp.read()

            with open(cache_path, "wb") as f:
                f.write(img_data)

            # Verify that it's a valid openable image
            with Image.open(cache_path) as im:
                im.verify()

            return str(cache_path)

        except Exception as e:
            logger.debug(f"Failed to download/cache image from '{img_url}': {e}")
            return None

    @classmethod
    def get_image_for_slide(cls, topic: str, slide_title: str = "", archetype: str = "") -> Optional[str]:
        """
        Retrieves a local high-quality image path matching the slide topic.
        """
        queries = cls.extract_search_terms(topic, slide_title)
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
        Replaces an existing picture shape or picture placeholder, or places
        an elegant image card in an open graphic zone on the right side of the slide.
        """
        if not image_path or not os.path.exists(image_path):
            return False

        try:
            # 1. Look for existing picture shapes or AutoShapes with blipFill (<p:spPr//a:blipFill>)
            # and replace with Smart Center-Cropping to eliminate aspect-ratio distortion
            candidate_shapes = MediaHarmonizer.find_image_shapes(slide)
            for shp in candidate_shapes:
                if MediaHarmonizer.replace_image_with_file(shp, image_path):
                    logger.info(f"MediaHarmonizer: Replaced image in '{shp.name}' with Smart Center-Crop from '{image_path}'")
                    return True

            # 2. Look for an unfilled picture placeholder
            for ph in slide.placeholders:
                try:
                    if ph.placeholder_format.type in (6, 18):  # PICTURE or MEDIA placeholder
                        slide.shapes.add_picture(
                            image_path, ph.left, ph.top, ph.width, ph.height
                        )
                        # Remove the old placeholder element
                        slide.shapes._spTree.remove(ph._element)
                        logger.info(f"Populated picture placeholder with '{image_path}'")
                        return True
                except Exception:
                    pass

            # 3. If no native picture slot exists, place an illustration card ONLY for layouts
            # with visual space (avoiding full-width card grids, timelines, or title slides)
            image_friendly_layouts = {
                "theory_concept", "case_study", "diagram_anatomy", 
                "problem_solution", "qa_discussion"
            }
            if layout_type and layout_type not in image_friendly_layouts:
                logger.debug(f"Skipping floating picture overlay for layout '{layout_type}'")
                return False

            # Collect occupied boxes to ensure zero collision
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

            req_w = 4.2 * 72.0 * scale_factor
            req_h = 3.6 * 72.0 * scale_factor
            safe_l, safe_t = MediaHarmonizer.find_safe_image_placement(sw_pt, sh_pt, occupied_boxes, req_w, req_h)

            pic = slide.shapes.add_picture(
                image_path,
                Inches(safe_l / 72.0),
                Inches(safe_t / 72.0),
                Inches(req_w / 72.0),
                Inches(req_h / 72.0)
            )
            # Apply center crop to the new picture shape
            MediaHarmonizer.replace_image_with_file(pic, image_path)
            logger.info(f"Added new collision-free scientific illustration '{image_path}' to slide at ({safe_l:.1f}, {safe_t:.1f})")
            return True

        except Exception as e:
            logger.debug(f"Could not inject image into slide: {e}")
            return False
