# -*- coding: utf-8 -*-
"""
core/utils/colors.py
Centralized color utility module for SlaydHub AI.
Eliminates duplicated hex_to_rgb implementations across the project.
Provides robust color conversion, luminance calculation, contrast text selection,
and standard academic color palettes.
"""

import re
from typing import Union, Tuple
from pptx.dml.color import RGBColor


def hex_to_rgb(hex_code: Union[str, RGBColor, Tuple[int, int, int]], default: Tuple[int, int, int] = (30, 41, 59)) -> RGBColor:
    """
    Safely converts a hex string (e.g. '#1E293B', '1E293B', '#FFF', 'FFF') or tuple to a pptx RGBColor.
    Returns default (dark slate) if hex_code is invalid.
    """
    if isinstance(hex_code, RGBColor):
        return hex_code

    if isinstance(hex_code, (tuple, list)) and len(hex_code) >= 3:
        try:
            return RGBColor(int(hex_code[0]), int(hex_code[1]), int(hex_code[2]))
        except (ValueError, TypeError):
            return RGBColor(*default)

    if not isinstance(hex_code, str):
        return RGBColor(*default)

    clean_hex = hex_code.strip().lstrip('#')

    # Expand short format #RGB to #RRGGBB
    if len(clean_hex) == 3:
        clean_hex = ''.join(c * 2 for c in clean_hex)

    if not re.match(r'^[0-9a-fA-F]{6}$', clean_hex):
        return RGBColor(*default)

    try:
        r = int(clean_hex[0:2], 16)
        g = int(clean_hex[2:4], 16)
        b = int(clean_hex[4:6], 16)
        return RGBColor(r, g, b)
    except (ValueError, TypeError):
        return RGBColor(*default)


def rgb_to_hex(r_or_color: Union[RGBColor, Tuple[int, int, int], int], g: int = 0, b: int = 0) -> str:
    """
    Converts RGB values or an RGBColor object to uppercase hex string (e.g. '#1E293B').
    """
    if isinstance(r_or_color, (RGBColor, tuple, list)):
        return f"#{int(r_or_color[0]):02X}{int(r_or_color[1]):02X}{int(r_or_color[2]):02X}"
    return f"#{int(r_or_color):02X}{int(g):02X}{int(b):02X}"


def get_luminance(color: Union[RGBColor, str, Tuple[int, int, int]]) -> float:
    """
    Calculates perceived luminance (0 to 255) using standard ITU-R BT.709 coefficients.
    """
    rgb = hex_to_rgb(color) if not isinstance(color, RGBColor) else color
    return 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]


def get_contrast_text_color(bg_color: Union[RGBColor, str, Tuple[int, int, int]]) -> RGBColor:
    """
    Returns high-contrast text color (dark navy/charcoal for light backgrounds, white for dark backgrounds).
    """
    lum = get_luminance(bg_color)
    if lum > 140:
        return RGBColor(15, 23, 42)  # Dark Navy Slate #0F172A
    return RGBColor(255, 255, 255)   # Pure White


# Standard Academic Color Palettes
ACADEMIC_PALETTES = {
    "deep_navy": {
        "primary": "#0F172A",
        "secondary": "#1E293B",
        "accent": "#38BDF8",
        "background": "#F8FAFC",
        "text_dark": "#0F172A",
        "text_light": "#FFFFFF"
    },
    "emerald_academic": {
        "primary": "#064E3B",
        "secondary": "#047857",
        "accent": "#10B981",
        "background": "#F0FDF4",
        "text_dark": "#064E3B",
        "text_light": "#FFFFFF"
    },
    "royal_indigo": {
        "primary": "#1E1B4B",
        "secondary": "#3730A3",
        "accent": "#6366F1",
        "background": "#EEF2FF",
        "text_dark": "#1E1B4B",
        "text_light": "#FFFFFF"
    },
    "ruby_crimson": {
        "primary": "#881337",
        "secondary": "#BE123C",
        "accent": "#F43F5E",
        "background": "#FFF1F2",
        "text_dark": "#4C0519",
        "text_light": "#FFFFFF"
    },
    "corporate_slate": {
        "primary": "#1E293B",
        "secondary": "#475569",
        "accent": "#0284C7",
        "background": "#F1F5F9",
        "text_dark": "#0F172A",
        "text_light": "#FFFFFF"
    }
}
