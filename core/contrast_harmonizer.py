# -*- coding: utf-8 -*-
"""
core/contrast_harmonizer.py
WCAG 2.1 Contrast Engine & Dynamic Theme Typography Colorizer.
Provides precise sRGB luminance calculations and automatic high-contrast text adaptation
to eliminate legibility issues on recolored theme shapes.
"""

from typing import Tuple, Union, Dict, Any
from pptx.dml.color import RGBColor

from core.penpot_themes import PenpotThemeManager


class ContrastHarmonizer:
    @staticmethod
    def _srgb_to_lin(c: float) -> float:
        """Expands sRGB channel to linear values for proper luminance calculation."""
        v = c / 255.0
        if v <= 0.04045:
            return v / 12.92
        return ((v + 0.055) / 1.055) ** 2.4

    @classmethod
    def calculate_luminance(cls, hex_color: str) -> float:
        """Calculates WCAG 2.1 Relative Luminance of a hex color."""
        hex_str = str(hex_color).lstrip('#')
        if len(hex_str) != 6:
            return 1.0 # fallback

        r = int(hex_str[0:2], 16)
        g = int(hex_str[2:4], 16)
        b = int(hex_str[4:6], 16)

        lin_r = cls._srgb_to_lin(r)
        lin_g = cls._srgb_to_lin(g)
        lin_b = cls._srgb_to_lin(b)

        return 0.2126 * lin_r + 0.7152 * lin_g + 0.0722 * lin_b

    @classmethod
    def calculate_contrast_ratio(cls, hex1: str, hex2: str) -> float:
        """Calculates WCAG contrast ratio between two hex colors."""
        l1 = cls.calculate_luminance(hex1)
        l2 = cls.calculate_luminance(hex2)

        l_lighter = max(l1, l2)
        l_darker = min(l1, l2)

        return (l_lighter + 0.05) / (l_darker + 0.05)

    @classmethod
    def get_high_contrast_text_color(cls, bg_hex: str, theme_key: str = "tech_minimal") -> str:
        """
        Calculates whether white or the theme's dark text color provides a better
        WCAG contrast ratio against the given background.
        Returns the hex code without '#'.
        """
        theme = PenpotThemeManager.get_theme(theme_key)
        dark_text_hex = str(theme.get("text_title", "0F172A")).lstrip('#')
        light_text_hex = "FFFFFF"

        cr_with_dark = cls.calculate_contrast_ratio(bg_hex, dark_text_hex)
        cr_with_light = cls.calculate_contrast_ratio(bg_hex, light_text_hex)

        # Strict WCAG AAA Contrast Optimization: Pick the text color that yields the maximum contrast ratio
        if cr_with_light >= cr_with_dark:
            return light_text_hex
        return dark_text_hex

    @classmethod
    def rgb_to_hex(cls, rgb: RGBColor) -> str:
        """Converts RGBColor tuple to hex string."""
        return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"

    @classmethod
    def ensure_text_contrast(cls, shape, bg_hex: str, theme_key: str = "tech_minimal") -> None:
        """
        Examines text inside an OpenXML text frame. Automatically forces the runs to optimal
        high-contrast text colors corresponding to the active penpot theme based on the shape's background.
        """
        if not hasattr(shape, "has_text_frame") or not shape.has_text_frame:
            return

        bg_hex = str(bg_hex).lstrip('#')
        optimal_hex = cls.get_high_contrast_text_color(bg_hex, theme_key)

        optimal_rgb = RGBColor(
            int(optimal_hex[0:2], 16),
            int(optimal_hex[2:4], 16),
            int(optimal_hex[4:6], 16)
        )

        for paragraph in shape.text_frame.paragraphs:
            if not paragraph.runs and paragraph.text.strip():
                try:
                    paragraph.font.color.rgb = optimal_rgb
                except Exception:
                    pass
            for run in paragraph.runs:
                # We overwrite the text color to the optimal contrasting color
                run.font.color.rgb = optimal_rgb

    @classmethod
    def get_theme_text_colors(cls, theme_key: str = "tech_minimal", is_dark_bg: bool = False) -> Dict[str, RGBColor]:
        """
        Dynamically extracts and prepares RGBColor objects for Typography classes
        (Title, Lead-in, Body, Muted) linked to the active Penpot theme.
        Eliminates legacy hardcoded blue colors.
        """
        theme = PenpotThemeManager.get_theme(theme_key)

        def hex_to_rgb(hex_str: str) -> RGBColor:
            hex_str = str(hex_str).lstrip('#')
            return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))

        if is_dark_bg:
            # On dark backgrounds we use light, highly readable colors.
            accent_light = "38BDF8" # Default sky blue fallback for dark
            if "primary" in theme:
                # Calculate if the theme's primary color is visible enough on dark bg
                cr = cls.calculate_contrast_ratio("000000", theme["primary"])
                accent_light = theme["primary"] if cr >= 4.5 else "38BDF8"

            return {
                "title": hex_to_rgb("#FFFFFF"),
                "lead_in": hex_to_rgb(accent_light),
                "body": hex_to_rgb("#F1F5F9"),
                "muted": hex_to_rgb("#94A3B8")
            }
        else:
            return {
                "title": hex_to_rgb(theme.get("text_title", "#0F172A")),
                "lead_in": hex_to_rgb(theme.get("primary", "#2563EB")),
                "body": hex_to_rgb(theme.get("text_body", "#334155")),
                "muted": hex_to_rgb(theme.get("text_muted", "#64748B"))
            }

