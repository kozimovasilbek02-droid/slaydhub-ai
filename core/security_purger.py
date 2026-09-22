# -*- coding: utf-8 -*-
"""
core/security_purger.py
Enterprise Security & Integrity Engine for PowerPoint Presentations.

Eliminates:
1. Invisible Ghost Action Buttons (transparent overlay click-jacking traps).
2. Spammy & Promotional Hyperlinks (<a:hlinkClick>, <a:hlinkHover>) from SlideCarnival,
   PresentationGO, Slidesgo, Canva, SlideEgg, and other free template authors.
3. Unsafe click actions (e.g. ppaction://program or macro invocations).
4. Ghost blocker shapes (massive empty/no-fill rectangles that intercept presenter clicks).

Operates via zero-copy native lxml XPath queries for maximum performance across 5,000+ templates.
"""

import re
import logging
from typing import Dict, List, Tuple, Any, Optional
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.presentation import Presentation

logger = logging.getLogger(__name__)

PROMO_DOMAINS = [
    "presentationgo", "slidescarnival", "slidesgo", "slideegg",
    "slidesmania", "slidenest", "showeet", "allppt", "powerpointschool",
    "thepopp", "pptmon", "canva.com", "freepik.com", "vecteezy",
    "envato", "graphicriver", "creativemarket", "templatemonster",
    "slidebean", "pitch.com", "slidemodel", "free-power-point-templates",
    "free-powerpoint-templates-design", "presentationmagazine",
    "templateswise", "hislide", "infograpia", "sketchbubble",
    "free-ppt", "freeppt", "slideteam", "slidebazaar"
]

PROMO_TOOLTIP_PATTERNS = [
    r"presentationgo", r"slidescarnival", r"slidesgo", r"slideegg",
    r"slide\s*nest", r"powerpoint", r"free\s*template", r"download",
    r"template\s*by", r"credit", r"watermark", r"logo"
]


class SecurityPurger:
    """
    Scans, detects, and purges invisible ghost click-jacking shapes and spammy
    promotional hyperlinks across Presentation slides, layouts, and masters.
    """

    @classmethod
    def is_promo_url(cls, url: Optional[str], tooltip: Optional[str] = None) -> bool:
        """Determines if a URL or tooltip corresponds to an author/ad watermark or promo site."""
        if url:
            u_clean = url.lower().strip()
            for dom in PROMO_DOMAINS:
                if dom in u_clean:
                    return True

        if tooltip:
            t_clean = tooltip.lower().strip()
            for pat in PROMO_TOOLTIP_PATTERNS:
                if re.search(pat, t_clean):
                    return True

        return False

    @classmethod
    def is_unsafe_action(cls, action: Optional[str]) -> bool:
        """Detects potentially hazardous actions (executables, macros)."""
        if not action:
            return False
        act = action.lower().strip()
        if any(bad in act for bad in ["ppaction://program", "ppaction://macro", "ppaction://ole"]):
            return True
        return False

    @classmethod
    def is_ghost_shape(
        cls,
        shape: Any,
        slide_w_pt: float = 720.0,
        slide_h_pt: float = 405.0
    ) -> Tuple[bool, str]:
        """
        Detects if a shape is an invisible ghost shape or click-jacking trap.
        Criteria:
        1. No visible text or empty text frame.
        2. Not a legitimate media type (picture with image data, table, chart).
        3. No visible fill (a:noFill or solidFill with alpha <= 5%).
        4. No visible outline/border (no border or a:noFill or w=0).
        5. Has an active click/hover hyperlink OR spans an immense transparent area (>40% canvas).
        """
        # Keep groups for recursive child checks
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            return False, "group"

        # Legitimate media with data
        if getattr(shape, "shape_type", None) in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART):
            return False, "media_content"

        elem = getattr(shape, "_element", None)
        if elem is None:
            return False, "no_element"

        # 1. Text check
        if getattr(shape, "has_text_frame", False):
            try:
                tf_text = shape.text_frame.text.strip()
                if tf_text:
                    return False, "has_text"
            except Exception:
                pass

        # 2. Geometry measurements (in points)
        w_pt = 0.0
        h_pt = 0.0
        l_pt = 0.0
        t_pt = 0.0
        try:
            w_pt = shape.width.pt if hasattr(shape.width, "pt") else float(shape.width) / 12700.0
            h_pt = shape.height.pt if hasattr(shape.height, "pt") else float(shape.height) / 12700.0
            l_pt = shape.left.pt if hasattr(shape.left, "pt") else float(shape.left) / 12700.0
            t_pt = shape.top.pt if hasattr(shape.top, "pt") else float(shape.top) / 12700.0
        except Exception:
            pass

        # 3. Fill Inspection via native XPath
        has_explicit_no_fill = bool(elem.xpath('.//p:spPr/a:noFill'))
        has_color_fill = elem.xpath('.//p:spPr/a:solidFill | .//p:spPr/a:gradFill | .//p:spPr/a:blipFill | .//p:spPr/a:pattFill')

        is_transparent_fill = False
        if has_explicit_no_fill or not has_color_fill:
            is_transparent_fill = True
        else:
            # Check for alpha opacity <= 5%
            alphas = elem.xpath('.//p:spPr/a:solidFill//a:alpha/@val')
            if alphas:
                try:
                    if int(alphas[0]) <= 5000:
                        is_transparent_fill = True
                except Exception:
                    pass

        # If it has a visible color fill, it is not an invisible ghost
        if not is_transparent_fill:
            return False, "visible_fill"

        # 4. Border / Outline Inspection
        has_outline = False
        ln_nodes = elem.xpath('.//p:spPr/a:ln')
        if ln_nodes:
            ln = ln_nodes[0]
            if not ln.xpath('./a:noFill'):
                w_attr = ln.get('w')
                if w_attr is None or int(w_attr) > 0:
                    has_outline = True

        if has_outline:
            return False, "visible_outline"

        # 5. Hyperlink and Click Trap Trigger Check
        hlinks = elem.xpath('.//a:hlinkClick | .//a:hlinkHover')
        if hlinks:
            return True, "ghost_action_button_with_hyperlink"

        # Check for massive transparent overlay blocking the canvas
        area_ratio = (w_pt * h_pt) / max(1.0, (slide_w_pt * slide_h_pt))
        if area_ratio >= 0.35:
            return True, f"invisible_click_blocker_overlay_{area_ratio:.2f}"

        return False, "innocuous_empty_shape"

    @classmethod
    def prune_shape_from_xml(cls, shape: Any) -> bool:
        """Directly removes a shape element from its parent XML tree."""
        try:
            elem = getattr(shape, "_element", None)
            if elem is not None:
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)
                    return True
        except Exception as e:
            logger.debug(f"Could not prune ghost shape from XML: {e}")
        return False

    @classmethod
    def purge_shape_hyperlinks(cls, shape: Any, container_part: Any = None) -> int:
        """
        Identifies and removes spammy, promotional, or unsafe hyperlinks (<a:hlinkClick>, <a:hlinkHover>)
        from shape non-visual properties, paragraph runs, and text frames.
        """
        elem = getattr(shape, "_element", None)
        if elem is None:
            return 0

        hlink_nodes = elem.xpath('.//a:hlinkClick | .//a:hlinkHover')
        if not hlink_nodes:
            return 0

        purged = 0
        for node in hlink_nodes:
            should_remove = False

            # Check action attribute
            action = node.get("action")
            if cls.is_unsafe_action(action):
                should_remove = True

            # Check tooltip attribute
            tooltip = node.get("tooltip")
            if cls.is_promo_url(None, tooltip):
                should_remove = True

            # Check target URL via relationship id
            r_id = node.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if not r_id:
                r_id = node.get('r:id') or node.get('id')

            target_url = None
            if r_id and container_part and hasattr(container_part, "rels"):
                try:
                    if r_id in container_part.rels:
                        target_url = container_part.rels[r_id].target_ref
                except Exception:
                    pass

            if target_url and cls.is_promo_url(target_url, tooltip):
                should_remove = True

            # If node has no target and no action, or points to invalidUrl, remove it
            if not target_url and not action and node.get("invalidUrl"):
                should_remove = True

            if should_remove:
                parent = node.getparent()
                if parent is not None:
                    parent.remove(node)
                    purged += 1

        return purged

    @classmethod
    def purge_container(
        cls,
        container: Any,
        container_part: Any = None,
        slide_w_pt: float = 720.0,
        slide_h_pt: float = 405.0
    ) -> Dict[str, int]:
        """
        Recursively sweeps a shapes container (Slide, Layout, or Master),
        pruning invisible ghost buttons and defusing spam hyperlinks.
        """
        stats = {"ghosts_pruned": 0, "hyperlinks_defused": 0}
        try:
            shapes_list = list(container)
        except Exception:
            return stats

        for s in shapes_list:
            # 1. Check if shape is a ghost action shape or invisible blocker
            is_ghost, reason = cls.is_ghost_shape(s, slide_w_pt, slide_h_pt)
            if is_ghost:
                if cls.prune_shape_from_xml(s):
                    stats["ghosts_pruned"] += 1
                    logger.info(f"Purged ghost shape ({reason}) from slide hierarchy.")
                    continue

            # 2. Check for recursive groups
            if getattr(s, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
                try:
                    group_stats = cls.purge_container(s.shapes, container_part, slide_w_pt, slide_h_pt)
                    stats["ghosts_pruned"] += group_stats["ghosts_pruned"]
                    stats["hyperlinks_defused"] += group_stats["hyperlinks_defused"]
                except Exception:
                    pass

            # 3. Defuse any spammy/promotional hyperlinks on the shape or its runs
            defused = cls.purge_shape_hyperlinks(s, container_part)
            stats["hyperlinks_defused"] += defused

        return stats

    @classmethod
    def purge_presentation_security(cls, prs: Presentation) -> Dict[str, int]:
        """
        Master security audit: cleans entire presentation hierarchy
        (Slide Masters, Slide Layouts, Slide Instances, Notes Slides).
        """
        total_stats = {"ghosts_pruned": 0, "hyperlinks_defused": 0}

        # Determine presentation slide dimensions
        slide_w_pt = 720.0
        slide_h_pt = 405.0
        try:
            slide_w_pt = prs.slide_width.pt if hasattr(prs.slide_width, "pt") else float(prs.slide_width) / 12700.0
            slide_h_pt = prs.slide_height.pt if hasattr(prs.slide_height, "pt") else float(prs.slide_height) / 12700.0
        except Exception:
            pass

        # 1. Slide Masters and Layouts
        for master in prs.slide_masters:
            m_part = getattr(master, "part", None)
            m_stats = cls.purge_container(master.shapes, m_part, slide_w_pt, slide_h_pt)
            total_stats["ghosts_pruned"] += m_stats["ghosts_pruned"]
            total_stats["hyperlinks_defused"] += m_stats["hyperlinks_defused"]

            for layout in master.slide_layouts:
                l_part = getattr(layout, "part", None)
                l_stats = cls.purge_container(layout.shapes, l_part, slide_w_pt, slide_h_pt)
                total_stats["ghosts_pruned"] += l_stats["ghosts_pruned"]
                total_stats["hyperlinks_defused"] += l_stats["hyperlinks_defused"]

        # 2. Slide Instances and Notes Slides
        for slide in prs.slides:
            s_part = getattr(slide, "part", None)
            s_stats = cls.purge_container(slide.shapes, s_part, slide_w_pt, slide_h_pt)
            total_stats["ghosts_pruned"] += s_stats["ghosts_pruned"]
            total_stats["hyperlinks_defused"] += s_stats["hyperlinks_defused"]

            if hasattr(slide, "has_notes_slide") and slide.has_notes_slide:
                try:
                    n_part = getattr(slide.notes_slide, "part", None)
                    n_stats = cls.purge_container(slide.notes_slide.shapes, n_part, slide_w_pt, slide_h_pt)
                    total_stats["ghosts_pruned"] += n_stats["ghosts_pruned"]
                    total_stats["hyperlinks_defused"] += n_stats["hyperlinks_defused"]
                except Exception:
                    pass

        logger.info(
            f"SecurityPurger Presentation Audit Complete: "
            f"{total_stats['ghosts_pruned']} ghost shapes pruned, "
            f"{total_stats['hyperlinks_defused']} spam hyperlinks defused."
        )
        return total_stats
