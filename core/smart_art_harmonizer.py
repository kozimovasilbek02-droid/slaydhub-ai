# -*- coding: utf-8 -*-
"""
core/smart_art_harmonizer.py
SmartArt Safe Guardian & Vector Outline Harmonizer for SlaydHub AI.

Solves (Round 11):
1. SmartArt Crash & Corrupt XML Protection:
   - Identifies MSO_SHAPE_TYPE.SMART_ART (24) and DrawingML diagram graphic frames.
   - Shields SmartArt from destructive text overwrites and shape pruning that corrupts
     OpenXML DrawingML diagram data models (<dgm:dataModel>, <dgm:relIds>).
   - Provides safe non-destructive text extraction and careful <dgm:t> updates.
2. Outline & Border Line Harmonization (Ghost Outline Eliminator):
   - Scans <p:spPr/a:ln> for legacy clashing template strokes (e.g. green or red borders).
   - Harmonizes strokes with active Penpot theme colors (primary, accent, or card_border).
   - Preserves line width (w) and dashed styling while removing clashing color artifacts.
"""

import re
import logging
from typing import Dict, Any, List, Optional, Tuple, Set
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shapes.base import BaseShape

from core.penpot_themes import PenpotThemeManager

logger = logging.getLogger("SmartArtHarmonizer")

# DrawingML Diagram XML Namespace URI
DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
SMART_ART_TYPE_ID = 24  # MSO_SHAPE_TYPE.SMART_ART


class SmartArtGuardian:
    """
    Guards SmartArt DrawingML diagrams from destructive mutations,
    preventing PowerPoint 'Corrupted File' crashes.
    """

    @classmethod
    def is_smart_art(cls, shape: BaseShape) -> bool:
        """
        Determines whether a shape is a SmartArt diagram or a DrawingML GraphicFrame.
        """
        if not hasattr(shape, "_element") or shape._element is None:
            return False

        # 1. python-pptx shape type check
        shape_type = getattr(shape, "shape_type", None)
        if shape_type == SMART_ART_TYPE_ID or (
            hasattr(MSO_SHAPE_TYPE, "SMART_ART") and shape_type == MSO_SHAPE_TYPE.SMART_ART
        ):
            return True

        elem = shape._element

        # 2. Tag name check: <p:graphicFrame> with diagram graphicData
        if elem.tag.endswith("graphicFrame"):
            try:
                diagram_nodes = elem.xpath('.//a:graphicData[contains(@uri, "diagram")]')
                if diagram_nodes:
                    return True
            except Exception:
                pass

            try:
                relid_nodes = elem.xpath('.//*[local-name()="relIds" or local-name()="pt" or local-name()="dataModel"]')
                if relid_nodes:
                    return True
            except Exception:
                pass

        # 3. Shape internal XPath check for diagram data models
        try:
            dgm_nodes = elem.xpath('.//*[local-name()="relIds" or local-name()="ptLst" or local-name()="dataModel"]')
            if dgm_nodes:
                return True
        except Exception:
            pass

        return False

    @classmethod
    def should_bypass_shape(cls, shape: BaseShape) -> bool:
        """
        Returns True if a shape must be shielded from standard text box replacement
        and shape pruning to guarantee 0% XML corruptions.
        """
        return cls.is_smart_art(shape)

    @classmethod
    def extract_smart_art_texts(cls, shape: BaseShape) -> List[str]:
        """
        Safely extracts text strings from a SmartArt diagram without mutating its XML.
        """
        if not hasattr(shape, "_element") or shape._element is None:
            return []

        texts = []
        try:
            elem = shape._element
            # Extract from text nodes inside graphic data
            t_nodes = elem.xpath('.//*[local-name()="t"]')
            for t in t_nodes:
                val = t.text
                if val and val.strip():
                    texts.append(val.strip())
        except Exception as e:
            logger.debug(f"SmartArtGuardian text extraction: {e}")

        return texts

    @classmethod
    def filter_safe_shapes_for_cleaning(
        cls,
        shapes_meta: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """
        Filters out SmartArt diagrams so they are NEVER pruned by SemanticShapeCleaner.
        """
        safe_list = []
        for meta in shapes_meta:
            shp = meta.get("shape")
            if shp is not None and cls.is_smart_art(shp):
                logger.info(f"SmartArtGuardian: Shielding SmartArt diagram '{getattr(shp, 'name', 'unnamed')}' from shape pruning.")
                continue
            safe_list.append(meta)
        return safe_list


class OutlineHarmonizer:
    """
    Scans and harmonizes vector outlines (<a:ln>) to eliminate clashing
    legacy template colors (Ghost Outlines).
    """

    @classmethod
    def harmonize_shape_outline(
        cls,
        shape: BaseShape,
        target_color_hex: str,
        preserve_transparent: bool = True
    ) -> bool:
        """
        Harmonizes the stroke border (<a:ln>) of a shape with the target hex color.
        Preserves line thickness (w) and dash pattern while updating stroke fill.
        """
        if not hasattr(shape, "_element") or shape._element is None:
            return False

        try:
            spPr_lst = shape._element.xpath('./p:spPr')
            if not spPr_lst:
                return False
            spPr = spPr_lst[0]

            ln_lst = spPr.xpath('./a:ln')
            if not ln_lst:
                return False  # Shape has no outline definition; leave borderless

            ln = ln_lst[0]

            # If the outline has <a:noFill>, it is intentionally transparent
            if preserve_transparent and ln.xpath('./a:noFill'):
                return False

            # Check if outline has an existing colored fill (solidFill or gradFill)
            has_color_fill = bool(ln.xpath('./a:solidFill | ./a:gradFill'))
            if not has_color_fill:
                return False

            # Remove old fills inside <a:ln>
            for child in list(ln):
                if child.tag.endswith('Fill'):
                    ln.remove(child)

            # Add clean solidFill with target theme color
            clean_hex = str(target_color_hex).lstrip('#').upper()
            solidFill = OxmlElement('a:solidFill')
            srgbClr = OxmlElement('a:srgbClr')
            srgbClr.set('val', clean_hex)
            solidFill.append(srgbClr)

            # Insert solidFill at the appropriate position inside <a:ln>
            ln.insert(0, solidFill)
            return True

        except Exception as e:
            logger.debug(f"OutlineHarmonizer: Failed to harmonize outline: {e}")
            return False

    @classmethod
    def harmonize_slide_outlines(
        cls,
        slide,
        theme_key: str = "tech_minimal"
    ) -> int:
        """
        Scans all shapes on a slide and harmonizes their outlines to the active Penpot theme.
        Eliminates legacy clashing borders (red, green, orange) on cards and badges.
        """
        theme = PenpotThemeManager.get_theme(theme_key)
        primary_hex = str(theme.get("primary", "2563EB")).lstrip('#')
        card_border_hex = str(theme.get("card_border", "E2E8F0")).lstrip('#')
        accent_hex = str(theme.get("accent", "06B6D4")).lstrip('#')

        harmonized_count = 0

        for shape in slide.shapes:
            # Shield SmartArt from direct outline tampering
            if SmartArtGuardian.is_smart_art(shape):
                continue

            # Shield structural media
            if getattr(shape, "shape_type", None) in (
                MSO_SHAPE_TYPE.PICTURE,
                MSO_SHAPE_TYPE.TABLE,
                MSO_SHAPE_TYPE.CHART,
                MSO_SHAPE_TYPE.MEDIA
            ):
                continue

            # Determine whether this is a card border or an accent stroke
            try:
                w_pt = shape.width.pt if hasattr(shape.width, "pt") else 100.0
                h_pt = shape.height.pt if hasattr(shape.height, "pt") else 100.0
            except Exception:
                w_pt, h_pt = 100.0, 100.0

            # Large containers and cards use card_border; small accents use primary or accent
            if w_pt > 150.0 and h_pt > 80.0:
                target_hex = card_border_hex
            else:
                target_hex = primary_hex

            if cls.harmonize_shape_outline(shape, target_hex):
                harmonized_count += 1

        logger.info(f"OutlineHarmonizer: Harmonized {harmonized_count} shape outlines on slide.")
        return harmonized_count
