# -*- coding: utf-8 -*-
"""
core/penpot_themes.py
Penpot-Inspired Professional Design System & Theme Presets for SlaydHub AI.

Provides carefully balanced color palettes, typography scales,
and unboxed card geometries modeled after high-end Penpot design templates.
"""

from typing import Dict, Any, List


PENPOT_THEMES: Dict[str, Dict[str, Any]] = {
    "tech_minimal": {
        "name": "Tech Minimal (Penpot Blueprint)",
        "bg_slide": "#FAFAFC",
        "primary": "#2563EB",       # Royal Blue
        "secondary": "#7C3AED",     # Violet
        "accent": "#06B6D4",        # Cyan
        "text_title": "#0F172A",    # Slate 900
        "text_body": "#334155",     # Slate 700
        "text_muted": "#64748B",    # Slate 500
        "card_bg": "#FFFFFF",
        "card_border": "#E2E8F0",
        "accent_bar": "#2563EB",
        "font_title": "Montserrat",
        "font_body": "Segoe UI",
        "card_radius": 12,          # pt
        "shadow_color": "#CBD5E1"
    },
    "academic_navy": {
        "name": "Academic Navy (Penpot Corporate)",
        "bg_slide": "#FFFFFF",
        "primary": "#0A3B75",       # Deep Navy
        "secondary": "#D97706",     # Warm Amber
        "accent": "#2563EB",        # Classic Blue
        "text_title": "#1E293B",    # Navy Slate
        "text_body": "#475569",     # Body Slate
        "text_muted": "#94A3B8",    # Muted
        "card_bg": "#F8FAFC",
        "card_border": "#E2E8F0",
        "accent_bar": "#0A3B75",
        "font_title": "Montserrat",
        "font_body": "Segoe UI",
        "card_radius": 8,
        "shadow_color": "#E2E8F0"
    },
    "emerald_science": {
        "name": "Emerald Bio-Science",
        "bg_slide": "#F0FDF4",
        "primary": "#059669",       # Emerald Green
        "secondary": "#0D9488",     # Teal
        "accent": "#F59E0B",        # Gold
        "text_title": "#064E3B",    # Deep Forest
        "text_body": "#1F2937",
        "text_muted": "#4B5563",
        "card_bg": "#FFFFFF",
        "card_border": "#D1FAE5",
        "accent_bar": "#059669",
        "font_title": "Montserrat",
        "font_body": "Segoe UI",
        "card_radius": 10,
        "shadow_color": "#E5E7EB"
    },
    "cyber_dark": {
        "name": "Cyber Dark Modern",
        "bg_slide": "#0F172A",      # Slate 900
        "primary": "#38BDF8",       # Sky Blue
        "secondary": "#A855F7",     # Purple
        "accent": "#F43F5E",        # Rose
        "text_title": "#F8FAFC",    # White
        "text_body": "#E2E8F0",     # Light Slate
        "text_muted": "#94A3B8",
        "card_bg": "#1E293B",       # Slate 800
        "card_border": "#334155",   # Slate 700
        "accent_bar": "#38BDF8",
        "font_title": "Montserrat",
        "font_body": "Segoe UI",
        "card_radius": 14,
        "shadow_color": "#020617"
    }
}


class PenpotThemeManager:
    """Helper for selecting and applying Penpot theme presets to python-pptx presentations."""

    @classmethod
    def get_theme(cls, theme_key: str = "tech_minimal") -> Dict[str, Any]:
        return PENPOT_THEMES.get(theme_key, PENPOT_THEMES["tech_minimal"])

    @classmethod
    def list_themes(cls) -> List[str]:
        return list(PENPOT_THEMES.keys())
