# -*- coding: utf-8 -*-
"""
core/font_harmonizer.py
Smart Dual-Font Pairing & OpenXML Typography Harmonizer.
Applies Penpot-themed typography hierarchy globally across a presentation.
Forces both <a:latin> and <a:cs> (Complex Script) run properties to guarantee
flawless rendering of Uzbek (O', G') and Cyrillic characters.
"""

from typing import Dict, Any, Optional, Union
from pptx.oxml.xmlchemy import OxmlElement
from pptx.shapes.base import BaseShape
from pptx.enum.shapes import MSO_SHAPE_TYPE

from core.penpot_themes import PenpotThemeManager


class FontHarmonizer:

    @classmethod
    def apply_font_to_rPr(cls, rPr, font_name: str) -> None:
        """Injects both latin and cs (Complex Script) typefaces into an OpenXML <a:rPr> adhering to schema order."""
        if rPr is None or not font_name:
            return

        # 1. Latin typeface
        latins = rPr.xpath('./a:latin')
        if latins:
            latins[0].set('typeface', font_name)
        else:
            latin = OxmlElement('a:latin')
            latin.set('typeface', font_name)
            # Insert before hyperlinks/extensions or append
            hlinks = rPr.xpath('./a:hlinkClick | ./a:hlinkHover | ./a:extLst')
            if hlinks:
                hlinks[0].addprevious(latin)
            else:
                rPr.append(latin)

        # 2. Complex Script typeface (vital for Uzbek O', G' and Cyrillic)
        css = rPr.xpath('./a:cs')
        if css:
            css[0].set('typeface', font_name)
        else:
            cs = OxmlElement('a:cs')
            cs.set('typeface', font_name)
            latins = rPr.xpath('./a:latin')
            if latins:
                latins[0].addnext(cs)
            else:
                hlinks = rPr.xpath('./a:hlinkClick | ./a:hlinkHover | ./a:extLst')
                if hlinks:
                    hlinks[0].addprevious(cs)
                else:
                    rPr.append(cs)

    @classmethod
    def harmonize_text_container(cls, txBody, target_font: str) -> None:
        """Ensures all runs (<a:r>) and paragraph defaults have proper rPr and font tags."""
        if txBody is None:
            return

        # 1. Ensure all runs have rPr and set font
        for r_elem in txBody.xpath('.//a:r'):
            rPr_nodes = r_elem.xpath('./a:rPr')
            if not rPr_nodes:
                rPr = OxmlElement('a:rPr')
                r_elem.insert(0, rPr)
            else:
                rPr = rPr_nodes[0]
            cls.apply_font_to_rPr(rPr, target_font)

        # 2. Paragraph default properties
        for defRPr in txBody.xpath('.//a:defRPr'):
            cls.apply_font_to_rPr(defRPr, target_font)

        # 3. End-of-paragraph marks
        for endParaRPr in txBody.xpath('.//a:endParaRPr'):
            cls.apply_font_to_rPr(endParaRPr, target_font)

    @classmethod
    def harmonize_shape_fonts(
        cls,
        shape: BaseShape,
        font_title: str,
        font_body: str,
        is_title: bool = False
    ) -> None:
        """
        Recursively harmonizes fonts in a shape based on typography hierarchy.
        Titles get font_title (e.g. Montserrat); body/tables get font_body (e.g. Segoe UI).
        """
        if not hasattr(shape, "shape_type"):
            return

        # Determine if this specific shape is a title
        is_shape_title = is_title
        if getattr(shape, "is_placeholder", False):
            try:
                ph_type = shape.placeholder_format.type
                # 1 = Title, 3 = Center Title, 4 = Subtitle
                if ph_type in (1, 3, 4):
                    is_shape_title = True
            except Exception:
                pass

        # Name-based title heuristic
        s_name = getattr(shape, "name", "").lower()
        if any(k in s_name for k in ["title", "titre", "header", "encabezado", "sarlavha", "headline"]):
            is_shape_title = True

        target_font = font_title if is_shape_title else font_body

        if shape.has_text_frame:
            try:
                txBody = shape.text_frame._txBody
                cls.harmonize_text_container(txBody, target_font)
            except Exception:
                pass

        elif shape.has_table:
            try:
                tbl = shape.table._tbl
                for tc in tbl.xpath('.//a:tc'):
                    for txBody in tc.xpath('./a:txBody'):
                        cls.harmonize_text_container(txBody, font_body)
            except Exception:
                pass

        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                for child in shape.shapes:
                    cls.harmonize_shape_fonts(child, font_title, font_body, is_shape_title)
            except Exception:
                pass

    @classmethod
    def harmonize_presentation_fonts(cls, prs, theme_key: Union[str, Dict[str, Any]] = "tech_minimal") -> None:
        """
        Global Pass: Applies Penpot typography hierarchy across all slides in the deck.
        Scans master slides, layouts, normal slides, and speaker notes.
        """
        if isinstance(theme_key, dict):
            theme = theme_key
        else:
            theme = PenpotThemeManager.get_theme(theme_key)

        font_title = theme.get("font_title", "Montserrat")
        font_body = theme.get("font_body", "Segoe UI")

        containers = []
        try:
            containers.extend(prs.slide_masters)
            for master in prs.slide_masters:
                containers.extend(master.slide_layouts)
            containers.extend(prs.slides)
        except Exception:
            if hasattr(prs, "slides"):
                containers.extend(prs.slides)

        for container in containers:
            for shape in container.shapes:
                cls.harmonize_shape_fonts(shape, font_title, font_body)

            # Harmonize Speaker Notes (Always body font)
            if hasattr(container, "has_notes_slide") and container.has_notes_slide:
                try:
                    for sp in container.notes_slide.shapes:
                        cls.harmonize_shape_fonts(sp, font_title, font_body, is_title=False)
                except Exception:
                    pass
