# -*- coding: utf-8 -*-
"""
core/theme_harmonizer.py
Vector Accent Shape Color Harmonizer for SlaydHub AI.

Harmonizes small-to-medium vector accent shapes (arrows, step badges, icon backgrounds,
colored indicator ribbons) with the active Penpot theme palette (primary, secondary, accent)
while strictly preserving background cards, canvases, photos, and tables (Zero-Defect).
"""

import logging
from typing import Dict, Any, Optional, List
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

from core.penpot_themes import PenpotThemeManager
from core.contrast_harmonizer import ContrastHarmonizer
from core.smart_art_harmonizer import SmartArtGuardian, OutlineHarmonizer

logger = logging.getLogger("ThemeHarmonizer")


class ThemeHarmonizer:
    """
    Intelligent Vector Accent Shape Harmonizer.
    Analyzes slide geometry and shape semantics to adapt visual accents to Penpot palettes.
    """

    # Maximum area fraction for a shape to be classified as an accent (e.g., max 25% of slide)
    MAX_ACCENT_AREA_FRACTION = 0.25

    # Minimum dimension to avoid zero-division or invisible dust shapes
    MIN_SHAPE_DIMENSION_PT = 4.0

    @classmethod
    def harmonize_slide_accents(
        cls,
        slide,
        theme_key: str = "tech_minimal",
        slide_width_pt: float = 960.0,
        slide_height_pt: float = 540.0
    ) -> int:
        """
        Scans all shapes on the slide and harmonizes accent shapes to the theme palette.
        Returns the count of harmonized accent shapes.
        """
        theme = PenpotThemeManager.get_theme(theme_key)
        primary_hex = theme.get("primary", "#2563EB").replace("#", "")
        secondary_hex = theme.get("secondary", "#7C3AED").replace("#", "")
        accent_hex = theme.get("accent", "#06B6D4").replace("#", "")

        palette_cycle = [primary_hex, secondary_hex, accent_hex]
        slide_area = slide_width_pt * slide_height_pt

        harmonized_count = 0
        accent_idx = 0

        for shape in slide.shapes:
            if not cls._is_accent_candidate(shape, slide_width_pt, slide_height_pt, slide_area):
                continue

            target_hex = palette_cycle[accent_idx % len(palette_cycle)]
            success = cls._apply_shape_color(shape, target_hex)
            if success:
                # Inoculate text inside recolored badge/shape with WCAG AAA optimal contrast
                ContrastHarmonizer.ensure_text_contrast(shape, target_hex, theme_key)
                harmonized_count += 1
                accent_idx += 1

        logger.info(f"ThemeHarmonizer: Harmonized {harmonized_count} accent shapes to theme '{theme_key}'")
        # Harmonize vector outlines across slide shapes to eliminate ghost borders
        OutlineHarmonizer.harmonize_slide_outlines(slide, theme_key=theme_key)
        return harmonized_count

    @classmethod
    def _is_accent_candidate(
        cls,
        shape,
        slide_width_pt: float,
        slide_height_pt: float,
        slide_area: float
    ) -> bool:
        """
        Determines if a shape is a true visual accent (arrow, badge, decorative pill)
        and NOT a background container, full-bleed image, or text card.
        """
        # 0. Shield SmartArt diagrams from accent recoloring
        if SmartArtGuardian.is_smart_art(shape):
            return False

        # 1. Reject non-vector or complex structural types
        if shape.shape_type in [
            MSO_SHAPE_TYPE.PICTURE,
            MSO_SHAPE_TYPE.TABLE,
            MSO_SHAPE_TYPE.MEDIA,
            MSO_SHAPE_TYPE.CHART,
            MSO_SHAPE_TYPE.GROUP,
            MSO_SHAPE_TYPE.LINE
        ]:
            return False

        # 2. Extract dimensions
        try:
            w_pt = shape.width.pt
            h_pt = shape.height.pt
        except Exception:
            return False

        if w_pt < cls.MIN_SHAPE_DIMENSION_PT or h_pt < cls.MIN_SHAPE_DIMENSION_PT:
            return False

        # 3. Reject full-slide or large background containers (> 30% area or > 80% span)
        shape_area = w_pt * h_pt
        if (shape_area / slide_area) > cls.MAX_ACCENT_AREA_FRACTION:
            return False

        if (w_pt / slide_width_pt) > 0.75 and (h_pt / slide_height_pt) > 0.40:
            return False

        # 4. Reject pure text containers with transparent or white background
        # If shape has a large body of text (> 50 chars), it is a content block, not an accent icon/badge
        if shape.has_text_frame:
            text_content = shape.text_frame.text.strip()
            if len(text_content) > 35:
                return False

        # 5. Check if shape actually has a fill (solidFill)
        spPr = shape._element.xpath('./p:spPr')
        if not spPr:
            return False

        # Check existing fill: must have a solid fill or gradient fill to recolor
        fills = spPr[0].xpath('./a:solidFill | ./a:gradFill')
        if not fills:
            return False

        return True

    @classmethod
    def _apply_shape_color(cls, shape, hex_color: str) -> bool:
        """
        Applies solidFill with the given hex color to the shape using clean OpenXML XPath.
        """
        try:
            spPr_lst = shape._element.xpath('./p:spPr')
            if not spPr_lst:
                return False
            spPr = spPr_lst[0]

            # Remove existing fill tags safely
            for child in list(spPr):
                if child.tag.endswith('Fill'):
                    spPr.remove(child)

            # Insert clean solidFill with srgbClr
            solidFill = OxmlElement('a:solidFill')
            srgbClr = OxmlElement('a:srgbClr')
            srgbClr.set('val', hex_color.upper())
            solidFill.append(srgbClr)
            spPr.append(solidFill)

            # Harmonize outline stroke if shape possesses an existing border line
            OutlineHarmonizer.harmonize_shape_outline(shape, hex_color)

            return True
        except Exception as e:
            logger.debug(f"ThemeHarmonizer: Failed to color shape: {e}")
            return False
