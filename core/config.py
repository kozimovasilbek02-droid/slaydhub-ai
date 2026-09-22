# -*- coding: utf-8 -*-
r"""
core/config.py
Centralized configuration manager for SlaydHub AI.
Eliminates hardcoded paths (C:\Users\user...) across the project.
Provides dynamic path resolution based on Path.home(), project root, and .env overrides.
"""

import os
import sys
from pathlib import Path
from typing import List, Optional


def _load_dotenv(dotenv_path: Path) -> None:
    """Lightweight .env loader without requiring third-party dependencies."""
    if not dotenv_path.exists() or not dotenv_path.is_file():
        return
    try:
        with open(dotenv_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, val = line.split("=", 1)
                key = key.strip()
                val = val.strip().strip("'\"")
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception:
        pass


class AppConfig:
    """Dynamic Application Configuration."""

    def __init__(self):
        # Base project root (two levels up from this file: core/config.py -> Power Point)
        self.BASE_DIR: Path = Path(__file__).resolve().parent.parent

        # Try to load .env from project root
        _load_dotenv(self.BASE_DIR / ".env")

        # Output Directories
        self.OUTPUT_ROOT: Path = Path(os.getenv("SLAYDHUB_OUTPUT_DIR", str(self.BASE_DIR / "output")))
        self.OUTPUT_STUDIO_DIR: Path = self.OUTPUT_ROOT / "academic_studio"
        self.OUTPUT_STUDIO_DIR.mkdir(parents=True, exist_ok=True)

        # Index File Path
        self.INDEX_CACHE_PATH: Path = Path(
            os.getenv(
                "SLAYDHUB_INDEX_PATH",
                str(self.BASE_DIR / "core" / "academic_template_index_native.json")
            )
        )
        self.FALLBACK_INDEX_PATH: Path = self.BASE_DIR / "core" / "academic_template_index.json"

        # Blacklisted low-quality template folders (requested by user)
        self.BLACKLISTED_FOLDERS: List[str] = [
            "tarjima_qilingan_top50",
            "yuklab_olingan_top50"
        ]

        # Template Library Search Roots (Cross-platform and multi-drive compatible)
        self.TEMPLATE_DIRS: List[str] = self._resolve_template_dirs()

    def _resolve_template_dirs(self) -> List[str]:
        """Dynamically identifies available authentic template directories."""
        home = Path.home()
        env_dirs = os.getenv("SLAYDHUB_TEMPLATE_DIRS", "")

        candidate_dirs: List[Path] = []

        if env_dirs:
            for p_str in env_dirs.split(";"):
                p = Path(p_str.strip())
                if p.exists():
                    candidate_dirs.append(p)

        # Default standard search paths
        standard_candidates = [
            home / "Desktop" / "slayd",
            home / "Downloads",
            Path("G:/Мой диск/Slaydlar/Slidenest"),
            Path("G:/Мой диск/Powepoint/PresentationGO/Prezentatsiyalar"),
            Path("G:/Мой диск/Powepoint/PPTMON/Prezentatsiyalar"),
            Path("G:/Мой диск/Slaydlar/SlidesCarnival/Prezentatsiyalar"),
            Path("G:/Мой диск/Powepoint/AllPPT/Prezentatsiyalar"),
            Path("G:/Мой диск/Powepoint/Showeet/Prezentatsiyalar"),
            Path("G:/Мой диск/Powepoint/SlideHunter/Prezentatsiyalar"),
            Path("G:/Мой диск/Powepoint/PresentationMagazine/Prezentatsiyalar"),
            Path("G:/Мой диск/Powepoint/SlidesMania/Prezentatsiyalar"),
            Path("G:/Мой диск/SlideEgg_Downloads/Prezentatsiyalar")
        ]

        for cand in standard_candidates:
            if cand.exists() and cand not in candidate_dirs:
                candidate_dirs.append(cand)

        return [str(p) for p in candidate_dirs]

    @property
    def gemini_api_key(self) -> Optional[str]:
        """Safely returns the Gemini / Google API Key."""
        return os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")

    def is_folder_blacklisted(self, path: str) -> bool:
        """Checks if a given path contains any blacklisted folder keywords."""
        norm = str(path).lower().replace("\\", "/")
        return any(b in norm for b in self.BLACKLISTED_FOLDERS)


# Global Singleton
config = AppConfig()

# Module-level convenience aliases
BASE_DIR = config.BASE_DIR
PROJECT_ROOT = config.BASE_DIR
OUTPUT_ROOT = config.OUTPUT_ROOT
OUTPUT_STUDIO_DIR = config.OUTPUT_STUDIO_DIR
INDEX_CACHE_PATH = config.INDEX_CACHE_PATH
FALLBACK_INDEX_PATH = config.FALLBACK_INDEX_PATH
TEMPLATE_DIRS = config.TEMPLATE_DIRS
is_folder_blacklisted = config.is_folder_blacklisted
