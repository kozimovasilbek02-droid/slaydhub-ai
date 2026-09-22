# -*- coding: utf-8 -*-
"""
core/bullet_harmonizer.py
Native Bullet & Master-Indent Harmonizer (Dual-Brain Round 13)
Preserves authentic template bullet characters, Wingdings fonts, custom colors,
and hanging indents while reconstructing multi-level hierarchical lists (lvl="1", lvl="2")
with 100% DrawingML XML schema compliance.
"""

import re
import copy
import logging
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any
import lxml.etree as etree
from pptx.shapes.base import BaseShape

logger = logging.getLogger("BulletHarmonizer")

# DrawingML Namespace
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Canonical child element order in DrawingML <a:pPr> (§21.1.2.2.7)
PPR_CHILD_ORDER = {
    "lnSpc": 10,
    "spcBfr": 20,
    "spcAft": 30,
    "buClrTx": 40,
    "buClr": 40,
    "buSzTx": 50,
    "buSzPct": 50,
    "buSzPts": 50,
    "buFontTx": 60,
    "buFont": 60,
    "buNone": 70,
    "buAutoNum": 70,
    "buChar": 70,
    "buBlip": 70,
    "tabLst": 80,
    "defRPr": 90,
    "extLst": 100,
}

# Standard EMUs (1 inch = 914,400 EMUs, 1 pt = 12,700 EMUs)
DEFAULT_MAR_L_LVL0 = 345600   # ~27.2 pt (0.378 in)
DEFAULT_INDENT_LVL0 = -180000 # ~ -14.2 pt (-0.197 in hanging indent)
INDENT_STEP_PER_LVL = 345600  # ~27.2 pt step per hierarchy level


@dataclass
class BulletStyle:
    """Encapsulates captured OpenXML bullet attributes from a template paragraph."""
    level: int = 0
    bullet_type: str = "none"  # 'char', 'autoNum', 'blip', 'none', 'inherit'
    char: Optional[str] = None
    font_typeface: Optional[str] = None
    color_hex: Optional[str] = None
    scheme_color: Optional[str] = None
    size_pct: Optional[int] = None
    size_pts: Optional[int] = None
    mar_l: Optional[int] = None
    indent: Optional[int] = None
    has_explicit_bullet: bool = False
    raw_bu_nodes: List[Any] = field(default_factory=list)


@dataclass
class HierarchyItem:
    """Represents a structured hierarchical bullet or text item."""
    level: int = 0
    is_bullet: bool = False
    lead_in: Optional[str] = None
    body: str = ""
    clean_text: str = ""


class BulletHarmonizer:
    """
    Central engine for capturing, preserving, and applying native PowerPoint bullets
    and multi-level indentation structures with zero XML corruption.
    """

    @staticmethod
    def order_ppr_children(ppr_element: etree._Element) -> None:
        """
        Sorts the child elements of <a:pPr> into the strict canonical order
        required by the DrawingML OpenXML specification (ECMA-376).
        """
        if ppr_element is None or len(ppr_element) <= 1:
            return

        children = list(ppr_element)
        def sort_key(elem):
            tag_name = elem.tag.split("}")[-1]
            return PPR_CHILD_ORDER.get(tag_name, 999)

        sorted_children = sorted(children, key=sort_key)
        for child in sorted_children:
            ppr_element.append(child)

    @classmethod
    def extract_paragraph_bullet_style(cls, paragraph) -> BulletStyle:
        """
        Extracts bullet properties, typeface, color, and margins from an existing
        paragraph's <a:pPr> element.
        """
        style = BulletStyle(level=getattr(paragraph, "level", 0))
        pPr = paragraph._p.find(f"{{{A_NS}}}pPr")
        if pPr is None:
            return style

        # Extract margins
        if "marL" in pPr.attrib:
            try:
                style.mar_l = int(pPr.attrib["marL"])
            except ValueError:
                pass
        if "indent" in pPr.attrib:
            try:
                style.indent = int(pPr.attrib["indent"])
            except ValueError:
                pass
        if "lvl" in pPr.attrib:
            try:
                style.level = int(pPr.attrib["lvl"])
            except ValueError:
                pass

        # Search for bullet-specific nodes
        bu_nodes = []
        for child in pPr:
            tag = child.tag.split("}")[-1]
            if tag.startswith("bu"):
                bu_nodes.append(copy.deepcopy(child))
                if tag == "buChar":
                    style.bullet_type = "char"
                    style.char = child.attrib.get("char")
                    style.has_explicit_bullet = True
                elif tag == "buFont":
                    style.font_typeface = child.attrib.get("typeface")
                elif tag == "buClr":
                    srgb = child.find(f"{{{A_NS}}}srgbClr")
                    if srgb is not None:
                        style.color_hex = srgb.attrib.get("val")
                    scheme = child.find(f"{{{A_NS}}}schemeClr")
                    if scheme is not None:
                        style.scheme_color = scheme.attrib.get("val")
                elif tag == "buSzPct":
                    try:
                        style.size_pct = int(child.attrib.get("val", "100000"))
                    except ValueError:
                        pass
                elif tag == "buSzPts":
                    try:
                        style.size_pts = int(child.attrib.get("val", "0"))
                    except ValueError:
                        pass
                elif tag == "buAutoNum":
                    style.bullet_type = "autoNum"
                    style.has_explicit_bullet = True
                elif tag == "buBlip":
                    style.bullet_type = "blip"
                    style.has_explicit_bullet = True
                elif tag == "buNone":
                    style.bullet_type = "none"
                    style.has_explicit_bullet = False

        style.raw_bu_nodes = bu_nodes
        return style

    @classmethod
    def extract_shape_bullet_styles(cls, shape: BaseShape) -> Dict[int, BulletStyle]:
        """
        Inspects all paragraphs in a template shape (and its layout/master if placeholder)
        to build a dictionary of BulletStyle by hierarchy level {0: style0, 1: style1, ...}.
        """
        styles: Dict[int, BulletStyle] = {}
        if not shape or not getattr(shape, "has_text_frame", False):
            return styles

        tf = shape.text_frame
        for p in tf.paragraphs:
            lvl = getattr(p, "level", 0)
            p_style = cls.extract_paragraph_bullet_style(p)
            if p_style.has_explicit_bullet or p_style.mar_l is not None:
                if lvl not in styles:
                    styles[lvl] = p_style

        # If shape is a placeholder without explicit bullet, inspect slide layout/master
        if 0 not in styles and getattr(shape, "is_placeholder", False):
            try:
                slide = shape.part
                # Attempt to extract from layout shape
                layout = getattr(slide, "slide_layout", None)
                if layout:
                    ph_idx = shape.placeholder_format.idx
                    for l_shape in layout.shapes:
                        if getattr(l_shape, "is_placeholder", False) and l_shape.placeholder_format.idx == ph_idx:
                            if l_shape.has_text_frame:
                                for lp in l_shape.text_frame.paragraphs:
                                    lp_style = cls.extract_paragraph_bullet_style(lp)
                                    if lp_style.has_explicit_bullet:
                                        styles[lp_style.level] = lp_style
                                        break
            except Exception as e:
                logger.debug(f"Placeholder bullet lookup skipped: {e}")

        return styles

    @classmethod
    def parse_hierarchical_lines(cls, text: str) -> List[HierarchyItem]:
        """
        Parses raw text into structured HierarchyItem objects, detecting:
        - Hierarchy Level: 0 (main), 1 (sub-bullet), 2 (sub-sub-bullet) via indents and bullet cues.
        - Bullet Presence: distinguishes bullet items from prose paragraphs.
        - Bold Lead-In: separates '**Lead-In:** body' or 'Lead-In: body' without manual bullet strings.
        """
        if not text:
            return []

        raw_lines = text.split("\n")
        items: List[HierarchyItem] = []

        bullet_pattern = re.compile(r'^\s*([•\-\*–—▪▫◦►✔➔\u2022\u25cf\u25b6\u2714\u2713]|\d+[\.\)]|[a-zA-Z][\.\)])\s*(.*)$')
        lead_in_pattern = re.compile(r'^(?:\*\*(.*?)\*\*[:\s]*|([A-Za-z0-9\s\-_]{2,32}):\s+)(.*)$')

        prev_level = 0

        for line in raw_lines:
            stripped = line.strip()
            if not stripped:
                continue

            # 1. Determine indentation level from leading spaces / tabs
            leading_whitespace = len(line) - len(line.lstrip())
            # Check tabs
            tabs_count = line[:leading_whitespace].count("\t")
            spaces_count = leading_whitespace - tabs_count
            effective_indent = spaces_count + (tabs_count * 4)

            level = 0
            if effective_indent >= 4:
                level = 2
            elif effective_indent >= 2:
                level = 1

            # 2. Check for bullet marker
            m_bullet = bullet_pattern.match(line)
            is_bullet = False
            clean_content = stripped

            if m_bullet:
                is_bullet = True
                marker = m_bullet.group(1)
                clean_content = m_bullet.group(2).strip()

                # If unindented, but starts with a sub-marker (e.g. '-' or '◦') while previous was level 0 with major bullet
                if level == 0 and marker in ("-", "–", "—", "◦", "▫", "*") and prev_level >= 0 and len(items) > 0 and items[-1].is_bullet:
                    # Check if previous item had a numbered or round bullet
                    level = 1

            # 3. Detect Bold Lead-In
            lead_in = None
            body = clean_content

            m_lead = lead_in_pattern.match(clean_content)
            if m_lead:
                raw_lead = m_lead.group(1) or m_lead.group(2)
                raw_body = m_lead.group(3)
                if raw_lead and raw_body:
                    words = raw_lead.strip().split()
                    if 1 <= len(words) <= 7 and len(raw_lead.strip()) <= 45:
                        lead_in = raw_lead.strip()
                        body = raw_body.strip()

            clean_text = f"{lead_in}: {body}" if lead_in else body

            items.append(HierarchyItem(
                level=level,
                is_bullet=is_bullet,
                lead_in=lead_in,
                body=body,
                clean_text=clean_text
            ))
            prev_level = level

        return items

    @classmethod
    def apply_native_bullet(
        cls,
        paragraph,
        bullet_style: Optional[BulletStyle] = None,
        level: int = 0,
        is_bullet: bool = True,
        theme_accent_hex: Optional[str] = None,
        scale_factor: float = 1.0
    ) -> None:
        """
        Applies native OpenXML bullet and hanging indentation to a paragraph with
        strict DrawingML schema ordering (§21.1.2.2.7).
        Guarantees zero double-bullets, correct line-wrap indentation, and style fidelity.
        """
        pPr = paragraph._p.get_or_add_pPr()
        paragraph.level = level
        pPr.set("lvl", str(level))

        # 1. Clean existing bullet nodes to prevent duplicate or conflicting tags
        for child in list(pPr):
            tag = child.tag.split("}")[-1]
            if tag.startswith("bu"):
                pPr.remove(child)

        # 2. Compute Margins and Indent
        if is_bullet:
            base_mar_l = DEFAULT_MAR_L_LVL0
            base_indent = DEFAULT_INDENT_LVL0

            # Inherit from template bullet style if available
            if bullet_style and bullet_style.mar_l is not None:
                base_mar_l = bullet_style.mar_l
            if bullet_style and bullet_style.indent is not None:
                base_indent = bullet_style.indent

            # Scale margins for sub-levels
            curr_mar_l = int(base_mar_l + (level * INDENT_STEP_PER_LVL * scale_factor))
            curr_indent = int(base_indent * scale_factor)

            pPr.set("marL", str(curr_mar_l))
            pPr.set("indent", str(curr_indent))

            # 3. Apply Bullet Elements
            # Option A: Clone authentic template bullet if present and level 0
            if bullet_style and bullet_style.has_explicit_bullet and bullet_style.raw_bu_nodes and level == 0:
                for bu_elem in bullet_style.raw_bu_nodes:
                    tag = bu_elem.tag.split("}")[-1]
                    if tag != "buNone":
                        pPr.append(copy.deepcopy(bu_elem))
            else:
                # Option B: Synthesize crisp native bullet
                # Level 0: Solid round bullet '•'
                # Level 1: Clean hyphen-en '–' or '◦'
                # Level 2: Square bullet '▪'
                bullet_char = "•"
                if level == 1:
                    bullet_char = "–"
                elif level >= 2:
                    bullet_char = "▪"

                # If template had a custom char, inherit or adapt it
                if bullet_style and bullet_style.char and level == 0:
                    bullet_char = bullet_style.char

                # Bullet Color: scheme/accent or inherit text color
                if theme_accent_hex:
                    clean_hex = theme_accent_hex.lstrip("#").upper()
                    bu_clr = etree.Element(f"{{{A_NS}}}buClr")
                    etree.SubElement(bu_clr, f"{{{A_NS}}}srgbClr", val=clean_hex)
                    pPr.append(bu_clr)
                elif bullet_style and bullet_style.color_hex:
                    bu_clr = etree.Element(f"{{{A_NS}}}buClr")
                    etree.SubElement(bu_clr, f"{{{A_NS}}}srgbClr", val=bullet_style.color_hex.upper())
                    pPr.append(bu_clr)
                elif bullet_style and bullet_style.scheme_color:
                    bu_clr = etree.Element(f"{{{A_NS}}}buClr")
                    etree.SubElement(bu_clr, f"{{{A_NS}}}schemeClr", val=bullet_style.scheme_color)
                    pPr.append(bu_clr)
                else:
                    # Inherit text color
                    pPr.append(etree.Element(f"{{{A_NS}}}buClrTx"))

                # Bullet Size
                size_pct = 100000 if level == 0 else (90000 if level == 1 else 85000)
                if bullet_style and bullet_style.size_pct and level == 0:
                    size_pct = bullet_style.size_pct
                pPr.append(etree.Element(f"{{{A_NS}}}buSzPct", val=str(size_pct)))

                # Bullet Font (if template used Wingdings or Symbol, preserve it)
                if bullet_style and bullet_style.font_typeface and level == 0:
                    pPr.append(etree.Element(f"{{{A_NS}}}buFont", typeface=bullet_style.font_typeface))

                # Bullet Character
                pPr.append(etree.Element(f"{{{A_NS}}}buChar", char=bullet_char))
        else:
            # Explicitly disable bullets for plain paragraphs, titles, and card headers
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            pPr.append(etree.Element(f"{{{A_NS}}}buNone"))

        # 4. Strict DrawingML element order sorting
        cls.order_ppr_children(pPr)
