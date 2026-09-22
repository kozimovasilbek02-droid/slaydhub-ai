# -*- coding: utf-8 -*-
"""
Slide & Template Preview Engine (Ultra-Fast Cached Visual Pipeline)
Provides:
1. Instant zero-dependency template thumbnail extraction from PPTX zip package (docProps/thumbnail.jpeg).
2. Complete presentation slide deck preview rendering (Slide 1..N) via PowerPoint COM with persistent disk caching.
3. High-res preview support for Streamlit interactive slide viewers.
"""

import os
import io
import time
import zipfile
import hashlib
import logging
from typing import List, Optional
from pathlib import Path

logger = logging.getLogger("SlidePreviewEngine")

# Cache directory for extracted and rendered slide thumbnails
PREVIEW_CACHE_DIR = Path("output/preview_cache").resolve()
PREVIEW_CACHE_DIR.mkdir(parents=True, exist_ok=True)


def _get_file_hash(file_path: str) -> str:
    """Computes a lightweight unique hash from file path, size and modification time."""
    try:
        p = os.path.abspath(file_path)
        stat = os.stat(p)
        token = f"{p.lower()}_{stat.st_size}_{stat.st_mtime}"
        return hashlib.md5(token.encode("utf-8")).hexdigest()[:16]
    except Exception:
        return hashlib.md5(file_path.encode("utf-8")).hexdigest()[:16]


def get_template_thumbnail(pptx_path: str) -> Optional[str]:
    """
    Retrieves or generates a thumbnail image for a PPTX template.
    Fast path: extracts embedded `docProps/thumbnail.jpeg` from zip in < 5ms.
    Fallback: uses PowerPoint COM to export Slide 1.
    Returns the absolute path to the cached image, or None if unavailable.
    """
    if not pptx_path or not os.path.exists(pptx_path):
        return None

    file_hash = _get_file_hash(pptx_path)
    cached_thumb = PREVIEW_CACHE_DIR / f"{file_hash}_thumb.jpg"

    if cached_thumb.exists() and cached_thumb.stat().st_size > 500:
        return str(cached_thumb)

    # 1. Fast Path: Extract embedded docProps/thumbnail.jpeg from PPTX zip
    try:
        with zipfile.ZipFile(pptx_path, "r") as z:
            thumb_names = [n for n in z.namelist() if n.lower().startswith("docprops/thumbnail.")]
            if thumb_names:
                thumb_bytes = z.read(thumb_names[0])
                if len(thumb_bytes) > 200:
                    with open(cached_thumb, "wb") as f_out:
                        f_out.write(thumb_bytes)
                    logger.debug(f"Extracted zip thumbnail for {os.path.basename(pptx_path)}")
                    return str(cached_thumb)
    except Exception as e:
        logger.debug(f"Zip thumbnail extraction failed for {pptx_path}: {e}")

    # 2. Fallback Path: Render Slide 1 via PowerPoint COM
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        abs_pptx = os.path.abspath(pptx_path)
        pres = ppt.Presentations.Open(abs_pptx, WithWindow=False)
        if pres.Slides.Count >= 1:
            pres.Slides(1).Export(str(cached_thumb), "JPG", 640, 360)
        pres.Close()
        ppt.Quit()
        if cached_thumb.exists():
            return str(cached_thumb)
    except Exception as e:
        logger.debug(f"PowerPoint COM thumbnail export failed for {pptx_path}: {e}")

    return None


def get_presentation_previews(pptx_path: str, max_slides: int = 30) -> List[str]:
    """
    Generates and returns high-resolution slide preview image paths for all slides in a presentation.
    Caches slide images inside `output/preview_cache/<file_hash>/slide_<idx>.jpg`.
    """
    if not pptx_path or not os.path.exists(pptx_path):
        return []

    file_hash = _get_file_hash(pptx_path)
    deck_cache_dir = PREVIEW_CACHE_DIR / file_hash
    deck_cache_dir.mkdir(parents=True, exist_ok=True)

    # Check if previews were already generated
    cached_slides = sorted(
        deck_cache_dir.glob("slide_*.jpg"),
        key=lambda p: int(p.stem.split("_")[1]) if p.stem.split("_")[1].isdigit() else 999
    )
    if cached_slides:
        return [str(p) for p in cached_slides]

    # Render all slides via PowerPoint COM
    slide_paths = []
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        abs_pptx = os.path.abspath(pptx_path)
        pres = ppt.Presentations.Open(abs_pptx, WithWindow=False)
        count = min(pres.Slides.Count, max_slides)

        for s_idx in range(1, count + 1):
            out_file = str(deck_cache_dir / f"slide_{s_idx}.jpg")
            pres.Slides(s_idx).Export(out_file, "JPG", 1280, 720)
            if os.path.exists(out_file):
                slide_paths.append(out_file)

        pres.Close()
        ppt.Quit()
        logger.info(f"Rendered {len(slide_paths)} slide previews for {os.path.basename(pptx_path)}")
        return slide_paths
    except Exception as e:
        logger.warning(f"PowerPoint COM batch preview export failed for {pptx_path}: {e}")

    # Fallback if COM export failed: extract cover thumbnail if available
    cover_thumb = get_template_thumbnail(pptx_path)
    if cover_thumb and os.path.exists(cover_thumb):
        return [cover_thumb]

    return []


def export_pptx_to_pdf(pptx_path: str, output_pdf_path: Optional[str] = None) -> Optional[str]:
    """
    Converts an authentic PPTX presentation into a high-quality PDF document
    via PowerPoint COM automation (ppSaveAsPDF = 32).
    """
    if not pptx_path or not os.path.exists(pptx_path):
        return None

    target_pdf = output_pdf_path or (os.path.splitext(pptx_path)[0] + ".pdf")

    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        abs_pptx = os.path.abspath(pptx_path)
        abs_pdf = os.path.abspath(target_pdf)
        pres = ppt.Presentations.Open(abs_pptx, WithWindow=False)
        # 32 = ppSaveAsPDF
        pres.SaveAs(abs_pdf, 32)
        pres.Close()
        ppt.Quit()
        if os.path.exists(abs_pdf):
            logger.info(f"Successfully exported PPTX to PDF: {abs_pdf}")
            return abs_pdf
    except Exception as e:
        logger.warning(f"PowerPoint COM PDF export failed for {pptx_path}: {e}")

    return None
