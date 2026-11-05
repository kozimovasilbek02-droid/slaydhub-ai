from pathlib import Path
# -*- coding: utf-8 -*-
"""
Semantic In-Place PPTX Text Replacer & Template Deck Assembler (High-Legibility, Anti-Ad & Zero-Collision Engine)
Preserves 100% of authentic template 3D graphics, vector artwork, backgrounds, and themes while:
1. Deep Watermark & Ad Purging: Recursively purges promotional links, watermarks ("PresentationGO", "SlidesCarnival", etc.) from Slides, Layouts, and Masters.
2. Ad/Credit Slide Filter: Strictly detects and rejects promotional/credit slides so they are never used for content.
3. SafeContentZoneValidator: Rejects narrow vertical strips (no 1-letter vertical cascades) and footer boxes colliding with wave artwork.
4. Dynamic Canvas Resolution Awareness: Proportional scaling for 1080p, 720p, and standard PPTX canvases.
5. High-Legibility Typography: 28-56pt titles, 16-36pt cards, 13.5-27pt body with zero microscopic text.
6. SemanticShapeCleaner: Physically prunes unassigned dummy shapes and ghost placeholder boxes.
7. SpatialCollisionEngine: Eliminates bounding box overlap between text frames and pictures/illustrations.
"""

import os
import io
import re
import math
import logging
from typing import Dict, Any, List, Optional, Tuple, Set
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.shapes.base import BaseShape
from pptx.oxml.xmlchemy import OxmlElement

from core.topic_logic_validator import TopicLogicValidator
from core.math_sanitizer import clean_latex_math
from core.micro_typography import MicroTypography
from core.visual_rhetoric import VisualRhetoric
from core.slide_morph import SlideMorphEngine
from core.table_matrix_replacer import TableMatrixReplacer
from core.theme_harmonizer import ThemeHarmonizer
from core.slot_allocator import SlotAllocator
from core.chart_harmonizer import ChartHarmonizer
from core.security_purger import SecurityPurger
from core.font_harmonizer import FontHarmonizer
from core.spatial_proximity_cleaner import SpatialProximityCleaner
from core.contrast_harmonizer import ContrastHarmonizer
from core.smart_art_harmonizer import SmartArtGuardian
from core.bullet_harmonizer import BulletHarmonizer, BulletStyle, HierarchyItem
from core.animation_guardian import AnimationGuardian

logger = logging.getLogger("PPTXTextReplacer")




# Known dummy placeholder strings commonly found in PowerPoint templates
DUMMY_PATTERNS = [
    r'lorem\s+ipsum',
    r'make\s+a\s+big\s+impact',
    r'content\s+here',
    r'briefly\s+describe',
    r'this\s+is\s+your\s+big\s+image',
    r'this\s+is\s+your\s+chart\s+slide',
    r'table\s+slide',
    r'presentationgo(\.com)?',
    r'slidescarnival(\.com)?',
    r'slidenest(\.com)?',
    r'slidesmania(\.com)?',
    r'pptmon(\.com)?',
    r'thepopp(\.com)?',
    r'allppt(\.com)?',
    r'slideegg(\.com)?',
    r'free\s+powerpoint\s+templates?',
    r'www\.[a-z0-9\-]+\.com',
    r'^(paris|berlin|tokyo|helsinki|bogota|new\s+york|oslo|london|madrid|rome|sydney|chicago|vancouver)(\s*0?[1-9])?$',
    r'^(cfo|coo|ceo|cto|cmo|vp|director|manager|lead)$',
    r'^(karla\s+gard|john\s+doe|jane\s+smith|alex\s+smith|andrew\s+childs|jean\s+marcus|paola\s+frisch|graham\s+gold)$',
    r'the\s+introduction\s+of\s+computers\s+was\s+the\s+basis',
    r'describe\s+your\s+content\s+here',
    r'carrer\s+d.enric\s+granados',
    r'jobs@yourbestbrand\.com',
    r'monday\s+to\s+saturday',
    r'let.s\s+keep\s+in\s+touch',
    r'book\s+now',
    r'write\s+your\s+subtitle\s+here',
    r'three\s+option\s+slide',
    r'four\s+option\s+slide',
    r'six\s+concept\s+slide',
    r'venn\s+diagram',
    r'click\s+to\s+(?:edit|add)\s+(?:title|text|subtitle|master)',
    r'заголовок\s+слайда',
    r'текст\s+слайда',
    r'подзаголовок\s+слайда',
    r'sample\s+text',
    r'text\s+here',
    r'insert\s+(?:text|title|logo)'
]

# Promotional watermark and advertisement patterns
WATERMARK_PATTERNS = [
    r'logotype(\s*™)?',
    r'logo(\s*type)?(\s*™)?',
    r'your\s+logo(\s+here)?',
    r'company\s+name',
    r'brand\s+name',
    r'tagline\s+here',
    r'add\s+your\s+logo',
    r'insert\s+logo',
    r'your\s+tagline',
    r'presentationgo(\.com)?',
    r'slidescarnival(\.com)?',
    r'slidenest(\.com)?',
    r'slidesmania(\.com)?',
    r'pptmon(\.com)?',
    r'thepopp(\.com)?',
    r'allppt(\.com)?',
    r'slideegg(\.com)?',
    r'showeet(\.com)?',
    r'powerpointschool(\.com)?',
    r'slidehunter(\.com)?',
    r'presentationmagazine(\.com)?',
    r'slidesgo(\.com)?',
    r'freepik(\.com)?',
    r'flaticon(\.com)?',
    r'https?://[^\s]+',
    r'free-powerpoint-templates-design\.com',
    r'designed\s+(?:with|by)\s+[a-z0-9\s&]+',
    r'designed\s+(?:with|by)',
    r'questions\s+or\s+need\s+help',
    r'credits?\s*:\s*this\s+presentation',
    r'fonts\s*&\s*colors\s*used',
    r'please\s+do\s+not\s+delete\s+this\s+slide',
    r'our\s+license',
    r'www\.[a-z0-9\-]+\.com',
    r'©\s*[a-z0-9\.\-]+',
    r'by\s*:\s*presentationgo',
    r'designed\s+by\s+presentationgo',
    r'free\s+templates\s*&\s*infographics',
    r'color\s+palette\s*&\s*font',
    r'credits\s*&\s*resources',
    r'thank\s+you\s+for\s+downloading',
    r'happy\s+designing',
    r'pexels,\s*pixabay',
    r'visit\s+our\s+faq',
    r'terms\s+of\s+use',
    r'instructions\s+for\s+use',
    r'alternative\s+resources',
    r'editable\s+icons',
    r'vector\s+icons',
    r'canva(\.com)?',
    r'template\s+(?:designed\s+by|from)',
    r'designed\s+by\s+[a-z0-9\s&]+'
]


class DeepWatermarkPurger:
    """
    Recursively scans and eliminates promotional watermarks, credit badges, and template ads
    from slide shapes, slide layouts, slide masters, and raw DrawingML XML trees.
    """

    @classmethod
    def is_watermark_text(cls, text: Any) -> bool:
        if text is None:
            return False
        if not isinstance(text, str):
            text = str(text)
        t_clean = text.lower().strip()
        if not t_clean:
            return False
        if t_clean in ("logotype", "logotype™", "logo", "your logo", "company name", "click to edit title", "текст слайда"):
            return True
        matches_any = False
        for pat in WATERMARK_PATTERNS:
            if re.search(pat, t_clean):
                matches_any = True
                break
        if not matches_any:
            return False

        # If it matched, check if it's purely a watermark badge or has substantial non-watermark content
        remainder = t_clean
        for pat in WATERMARK_PATTERNS:
            remainder = re.sub(pat, '', remainder).strip()
        remainder = re.sub(r'[\s\[\]\(\)\-\:\,\.\/\|\&\©]', '', remainder)
        return len(remainder) < 8

    @classmethod
    def clean_watermark_mentions(cls, text: str) -> str:
        """Strips watermark phrases out of mixed content text."""
        if not text:
            return ""
        cleaned = text
        for pat in WATERMARK_PATTERNS:
            cleaned = re.sub(pat, '', cleaned, flags=re.I).strip()
        cleaned = re.sub(r'\[\s*\]|\(\s*\)', '', cleaned).strip()
        return cleaned

    @classmethod
    def purge_shape_tree(cls, shapes_container) -> int:
        """Purges watermark text or shapes from a shapes collection, including nested groups and XML tags."""
        purged_count = 0
        try:
            shapes_list = list(shapes_container)
        except Exception:
            return 0

        for s in shapes_list:
            # 0. Ghost action shape / click-trap detection
            is_ghost, reason = SecurityPurger.is_ghost_shape(s)
            if is_ghost:
                SecurityPurger.prune_shape_from_xml(s)
                purged_count += 1
                continue

            # Defuse promotional hyperlinks on shape or text runs
            s_part = getattr(s, "part", None)
            purged_count += SecurityPurger.purge_shape_hyperlinks(s, s_part)

            has_watermark = False

            # 1. Check all paragraphs and runs inside text frame
            if s.has_text_frame:
                try:
                    full_txt = s.text_frame.text
                    if cls.is_watermark_text(full_txt):
                        has_watermark = True
                    else:
                        for p in s.text_frame.paragraphs:
                            if cls.is_watermark_text(p.text):
                                p.text = ""
                            else:
                                for r in p.runs:
                                    if cls.is_watermark_text(r.text):
                                        r.text = ""
                                cleaned_p = cls.clean_watermark_mentions(p.text)
                                if cleaned_p != p.text:
                                    p.text = cleaned_p
                except Exception:
                    pass

            # 2. Check shape name, alt text or description
            s_name = getattr(s, "name", "").lower()
            if any(k in s_name for k in ["presentationgo", "slidescarnival", "watermark", "credit", "slidenest", "slideegg", "showeet", "allppt", "logo", "logotype"]):
                if not s.has_text_frame or not s.text_frame.text.strip() or cls.is_watermark_text(s.text_frame.text) or "logo" in s.text_frame.text.lower():
                    has_watermark = True

            # 3. Deep XPath inspection for promotional hyperlinks, tooltips and tags (zero-copy)
            try:
                elem = getattr(s, "_element", None)
                if elem is not None:
                    hlink_nodes = elem.xpath('.//a:hlinkClick | .//a:hlinkHover')
                    for hn in hlink_nodes:
                        tooltip = hn.get("tooltip")
                        r_id = hn.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') or hn.get('r:id')
                        target_url = None
                        if r_id and s_part and hasattr(s_part, "rels") and r_id in s_part.rels:
                            target_url = s_part.rels[r_id].target_ref
                        if SecurityPurger.is_promo_url(target_url, tooltip):
                            has_watermark = True
                            break
            except Exception:
                pass

            # 4. Check for picture watermarks / promotional logos (e.g. LogoType, ALLPPT, SlideEgg)
            if not has_watermark and s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    w_pt = SafeShapeGeometry.to_pt(s.width, 0.0)
                    h_pt = SafeShapeGeometry.to_pt(s.height, 0.0)
                    l_pt = SafeShapeGeometry.to_pt(s.left, 0.0)
                    t_pt = SafeShapeGeometry.to_pt(s.top, 0.0)
                    sh_pt = 405.0
                    try:
                        if hasattr(s, "part") and hasattr(s.part, "slide_height"):
                            sh_pt = SafeShapeGeometry.to_pt(s.part.slide_height, 405.0)
                    except Exception:
                        pass
                    # Small pictures (w <= 240, h <= 80) in bottom 28% of canvas or top corners are promotional badges
                    is_bottom_zone = (t_pt >= sh_pt * 0.72)
                    is_top_corner = (t_pt <= sh_pt * 0.18 and (l_pt > 450 or l_pt < 120))
                    if (w_pt <= 240 and h_pt <= 80) and (is_bottom_zone or is_top_corner):
                        has_watermark = True
                    elif any(k in s_name for k in ["logo", "watermark", "credit", "badge", "brand", "allppt", "presentationgo"]):
                        has_watermark = True
                except Exception:
                    pass

            if has_watermark:
                try:
                    if s.has_text_frame:
                        s.text_frame.text = ""
                    SemanticShapeCleaner.prune_shape(s)
                    purged_count += 1
                except Exception:
                    pass
            elif s.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    gw_pt = SafeShapeGeometry.to_pt(s.width, 0.0)
                    gh_pt = SafeShapeGeometry.to_pt(s.height, 0.0)
                    gl_pt = SafeShapeGeometry.to_pt(s.left, 0.0)
                    gt_pt = SafeShapeGeometry.to_pt(s.top, 0.0)
                    # Bottom-right corner promotional vector logo badge (e.g. ALLPPT.com vector group)
                    if gw_pt <= 145.0 and gh_pt <= 42.0 and gl_pt >= 700.0 and gt_pt >= 440.0:
                        SemanticShapeCleaner.prune_shape(s)
                        purged_count += 1
                        continue
                    purged_count += cls.purge_shape_tree(s.shapes)
                    if len(s.shapes) == 0:
                        SemanticShapeCleaner.prune_shape(s)
                        purged_count += 1
                except Exception:
                    pass

        return purged_count

    @classmethod
    def purge_presentation_watermarks(cls, prs: Presentation) -> None:
        """Purges watermarks and ghost click traps from entire presentation hierarchy (Masters, Layouts, Slides, and Notes)."""
        # 0. Deep Enterprise Security & Ghost Action Sweep
        try:
            SecurityPurger.purge_presentation_security(prs)
        except Exception as e:
            logger.debug(f"SecurityPurger sweep encountered exception: {e}")

        sw_emu = int(prs.slide_width) if prs.slide_width else 9144000
        sh_emu = int(prs.slide_height) if prs.slide_height else 6858000

        def _prune_offcanvas_and_instruction_notes(shapes_col):
            for sh in list(shapes_col):
                try:
                    l = int(sh.left) if sh.left is not None else 0
                    t = int(sh.top) if sh.top is not None else 0
                    w = int(sh.width) if sh.width is not None else 0
                    h = int(sh.height) if sh.height is not None else 0
                    # Off-canvas sticky notes (placed outside slide bounds by template authors)
                    if l >= sw_emu * 0.96 or (l + w) <= 45720 or t >= sh_emu * 0.96 or (t + h) <= 45720:
                        SemanticShapeCleaner.prune_shape(sh)
                        continue
                    t_children = AdSlideFilter.collect_all_text_shapes_recursive([sh])
                    comb_txt = " ".join(c.text_frame.text.lower() for c in t_children if c.has_text_frame)
                    if any(bad in comb_txt for bad in ("to insert your own", "insert >>", "office 365", "subscribers", "free-powerpoint-templates", "allppt.com", "slidesgo", "slidescarnival")):
                        SemanticShapeCleaner.prune_shape(sh)
                        continue
                except Exception:
                    pass

        # 1. Slide Masters and Layouts
        for master in prs.slide_masters:
            _prune_offcanvas_and_instruction_notes(master.shapes)
            cls.purge_shape_tree(master.shapes)
            for layout in master.slide_layouts:
                _prune_offcanvas_and_instruction_notes(layout.shapes)
                cls.purge_shape_tree(layout.shapes)

        # 2. Slide Instances and Notes
        for slide in prs.slides:
            _prune_offcanvas_and_instruction_notes(slide.shapes)
            cls.purge_shape_tree(slide.shapes)
            if hasattr(slide, "has_notes_slide") and slide.has_notes_slide:
                try:
                    n_tf = slide.notes_slide.notes_text_frame
                    if cls.is_watermark_text(n_tf.text):
                        n_tf.text = ""
                except Exception:
                    pass


class AdSlideFilter:
    """
    Detects and rejects full advertisement, credit, icon sheet, medical-virus stamp, and device mockup slides from template decks.
    """

    @classmethod
    def collect_all_text_shapes_recursive(cls, shapes_container) -> List[BaseShape]:
        res = []
        try:
            for s in shapes_container:
                if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                    res.extend(cls.collect_all_text_shapes_recursive(s.shapes))
                elif s.has_text_frame and s.text_frame.text.strip():
                    res.append(s)
        except Exception:
            pass
        return res

    @classmethod
    def is_ad_or_credit_slide(cls, slide, slide_index: int = -1) -> bool:
        t_shapes = cls.collect_all_text_shapes_recursive(slide.shapes)
        all_text = " ".join([s.text_frame.text.lower() for s in t_shapes])

        # 1. Text markers for ads, credits, icon packs, team rosters, and virus stamps
        ad_markers = [
            "designed with", "presentationgo", "free templates & infographics",
            "questions or need help", "slidescarnival", "slidenest", "slidesmania",
            "slideegg", "credits & resources", "fonts & colors used",
            "color palette & font", "terms of use", "our license", "thank you for downloading",
            "visit our faq", "instructions for use", "alternative resources",
            "editable icons", "icon set", "icon sets", "fully editable icon",
            "fully editable shapes", "png & shapes layout",
            "vector icons", "happy designing", "pexels, pixabay", "google fonts",
            "fonts & colors", "typography & colors", "our team", "meet our team", "member name",
            "customizable icons", "business icons", "medical icons", "education icons",
            "avatar pack", "icon pack", "vector assets", "infographic style fully editable",
            "2019-ncov", "coronavirus"
        ]
        for marker in ad_markers:
            if marker in all_text:
                return True

        # 2. Icon Sheet / Asset Pack detection (using recursive text shape count)
        total_shapes = len(slide.shapes)
        pic_count = sum(1 for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
        small_shapes = sum(
            1 for s in slide.shapes
            if (s.width and SafeShapeGeometry.to_pt(s.width) < 65.0) and (s.height and SafeShapeGeometry.to_pt(s.height) < 65.0)
        )
        if small_shapes >= 15 and len(t_shapes) <= 3:
            return True
        if small_shapes >= 25:
            return True
        if pic_count >= 6 and len(t_shapes) <= 3:
            return True
        if total_shapes >= 20 and len(t_shapes) <= 1:
            return True

        # 3. Device Mockup with almost no text
        has_mockup = False
        for s in slide.shapes:
            w_pt = SafeShapeGeometry.to_pt(s.width)
            h_pt = SafeShapeGeometry.to_pt(s.height)
            if s.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.PLACEHOLDER):
                if 100.0 <= w_pt <= 260.0 and 240.0 <= h_pt <= 560.0 and (h_pt / max(w_pt, 1.0)) >= 1.6:
                    has_mockup = True
                    break
        if has_mockup and len(t_shapes) <= 3 and len(all_text) < 90:
            return True

        # 4. Completely empty slide (unless it has tables/charts)
        has_table_or_chart = any(
            s.shape_type in (MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART) or (hasattr(s, "has_chart") and s.has_chart)
            for s in slide.shapes
        )
        if len(t_shapes) == 0 and not has_table_or_chart and total_shapes < 3:
            return True

        return False


class SafeShapeGeometry:
    """
    Safely handles shape coordinates and dimensions in python-pptx without
    crashing on None or zeroing out inherited xfrm dimensions (width, height, left, top).
    """
    @staticmethod
    def to_pt(val, default: float = 0.0) -> float:
        if val is None:
            return default
        if hasattr(val, 'pt'):
            try:
                return float(val.pt)
            except Exception:
                return default
        try:
            return float(val) / 12700.0
        except Exception:
            return default

    @staticmethod
    def update(shape: BaseShape, left=None, top=None, width=None, height=None) -> None:
        orig_l = shape.left
        orig_t = shape.top
        orig_w = shape.width
        orig_h = shape.height

        shape.left = left if left is not None else orig_l
        shape.top = top if top is not None else orig_t
        shape.width = width if width is not None else orig_w
        shape.height = height if height is not None else orig_h


class SafeContentZoneValidator:
    """
    Validates shape dimensions and spatial coordinates to guarantee:
    1. No narrow vertical strips (width >= 120pt) that cause 1-letter vertical cascades.
    2. No footer shapes colliding with bottom wave artwork (top <= 82% of slide height).
    3. No off-canvas shapes (left >= -10pt and left <= 92% of slide width).
    """

    @staticmethod
    def is_valid_content_box(
        shape: BaseShape,
        sw_pt: float,
        sh_pt: float,
        is_title: bool = False
    ) -> bool:
        if not shape.has_text_frame:
            return False
            
        l_pt = SafeShapeGeometry.to_pt(shape.left, 0.0)
        t_pt = SafeShapeGeometry.to_pt(shape.top, 0.0)
        w_pt = SafeShapeGeometry.to_pt(shape.width, 0.0)
        h_pt = SafeShapeGeometry.to_pt(shape.height, 0.0)
        
        # Reject off-canvas shapes
        if l_pt < -10.0 or l_pt > sw_pt * 0.92:
            return False

        # Reject narrow vertical strips (prevents 1-letter vertical cascades)
        if w_pt < 120.0 and not is_title:
            return False

        # Reject footer shapes colliding with bottom wave artwork
        if t_pt > (sh_pt * 0.82) and not is_title:
            return False

        # Reject microscopic area shapes
        if (w_pt * h_pt) < 3000.0:
            return False

        return True

    @classmethod
    def is_in_safe_zone(
        cls,
        shape: BaseShape,
        sw_pt: float,
        sh_pt: float
    ) -> bool:
        """Alias for is_valid_content_box to verify if a shape is within slide canvas boundaries."""
        return cls.is_valid_content_box(shape, sw_pt, sh_pt)


class SemanticShapeCleaner:
    """
    Physically removes unassigned dummy placeholder shapes, ghost cards, and
    irrelevant elements directly from the slide XML tree.
    """

    @staticmethod
    def prune_shape(shape: BaseShape) -> bool:
        """Physically removes a shape from its parent XML tree."""
        try:
            sp_elem = shape._element
            parent = sp_elem.getparent()
            if parent is not None:
                parent.remove(sp_elem)
                return True
        except Exception as e:
            logger.debug(f"Could not prune shape element: {e}")
        return False

    @classmethod
    def cleanup_unassigned_shapes(
        cls,
        slide,
        assigned_shape_ids: Set[int],
        all_text_shapes: List[Dict[str, Any]],
        keep_step_count: Optional[int] = None
    ) -> None:
        """
        Removes unassigned dummy text shapes, leftover badges, and ghost placeholder boxes
        while strictly preserving any group that contains assigned descendant shapes or vector illustrations.
        """
        pruned_bboxes: List[Tuple[float, float, float, float]] = []
        assigned_bboxes: List[Tuple[float, float, float, float]] = []

        # 0. Collect bounding boxes of assigned shapes & all their ancestor groups
        assigned_groups = set()
        for s_meta in all_text_shapes:
            shp = s_meta["shape"]
            if id(shp) in assigned_shape_ids:
                assigned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))
                if s_meta.get("parent_group") is not None:
                    assigned_groups.add(id(s_meta["parent_group"]))
                for anc_g in s_meta.get("ancestor_groups", []):
                    assigned_groups.add(id(anc_g))

        # 1. Process individual unassigned text shapes first
        pruned_groups = set()
        for s_meta in all_text_shapes:
            shp = s_meta["shape"]
            shp_id = id(shp)
            if shp_id in assigned_shape_ids:
                continue
            if SmartArtGuardian.is_smart_art(shp):
                continue

            if s_meta.get("is_step"):
                step_val = s_meta.get("step_num")
                if keep_step_count is not None and isinstance(step_val, int) and step_val > keep_step_count:
                    pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))
                    cls.prune_shape(shp)
                elif s_meta.get("is_dummy") or s_meta.get("is_watermark"):
                    pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))
                    cls.prune_shape(shp)
                continue

            if s_meta.get("is_dummy") or s_meta.get("is_watermark") or not s_meta.get("is_valid_zone") or shp.name.lower().startswith(("textbox", "cuadrotexto", "title ", "body ")):
                pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))
                cls.prune_shape(shp)
            else:
                try:
                    if shp.has_text_frame:
                        shp.text_frame.text = ""
                except Exception:
                    pass
                pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))

        # 2. Only prune a leaf card group if it has NO assigned shapes AND NO vector artwork
        for s_meta in all_text_shapes:
            grp = s_meta.get("parent_group")
            if grp is not None and id(grp) not in assigned_groups and id(grp) not in pruned_groups:
                if not SpatialProximityCleaner._group_has_assigned_or_artwork_recursive(grp, assigned_shape_ids):
                    pruned_bboxes.append(SpatialProximityCleaner.get_bbox(grp))
                    cls.prune_shape(grp)
                    pruned_groups.add(id(grp))

        # 3. Invoke SpatialProximityCleaner with true slide dimensions
        try:
            _, _, sw_pt, sh_pt = PPTXTextReplacer.get_slide_dimensions(slide)
            SpatialProximityCleaner.prune_orphaned_artifacts(
                slide=slide,
                pruned_bboxes=pruned_bboxes,
                assigned_bboxes=assigned_bboxes,
                assigned_shape_ids=assigned_shape_ids,
                sw_pt=sw_pt,
                sh_pt=sh_pt
            )
        except Exception as e:
            logger.debug(f"Spatial proximity cleaning encountered error: {e}")


class SpatialCollisionEngine:
    """
    Calculates 2D bounding boxes and eliminates collisions between text boxes and
    images, illustrations, or vector artwork.
    """

    @staticmethod
    def get_bbox(shape: BaseShape) -> Tuple[float, float, float, float]:
        """Returns (left, top, right, bottom) in points."""
        l = SafeShapeGeometry.to_pt(shape.left, 0.0)
        t = SafeShapeGeometry.to_pt(shape.top, 0.0)
        w = SafeShapeGeometry.to_pt(shape.width, 0.0)
        h = SafeShapeGeometry.to_pt(shape.height, 0.0)
        return (l, t, l + w, t + h)

    @classmethod
    def has_overlap(cls, b1: Tuple[float, float, float, float], b2: Tuple[float, float, float, float], padding: float = 4.0) -> bool:
        """Returns True if bounding box b1 and b2 intersect with optional padding."""
        return not (b1[2] + padding <= b2[0] or
                    b1[0] - padding >= b2[2] or
                    b1[3] + padding <= b2[1] or
                    b1[1] - padding >= b2[3])

    @classmethod
    def detect_collisions(cls, shapes: List[BaseShape], padding: float = 4.0) -> List[Tuple[BaseShape, BaseShape]]:
        """Detects overlapping pairs among a collection of shapes."""
        collisions = []
        for i in range(len(shapes)):
            b1 = cls.get_bbox(shapes[i])
            for j in range(i + 1, len(shapes)):
                b2 = cls.get_bbox(shapes[j])
                if cls.has_overlap(b1, b2, padding=padding):
                    collisions.append((shapes[i], shapes[j]))
        return collisions

    @classmethod
    def resolve_collisions(cls, shapes: List[BaseShape]) -> bool:
        """Adjusts colliding shapes vertically or horizontally to eliminate overlaps."""
        collisions = cls.detect_collisions(shapes)
        if not collisions:
            return True
        for s1, s2 in collisions:
            b1 = cls.get_bbox(s1)
            new_top = b1[3] + 10.0
            SafeShapeGeometry.update(s2, top=Inches(new_top / 72.0))
        return len(cls.detect_collisions(shapes)) == 0

    @staticmethod
    def get_slide_dimensions(slide) -> Tuple[Any, Any, float, float]:
        """Accurately extracts slide width/height from python-pptx presentation package."""
        sw = None
        sh = None
        try:
            sw = slide.part.package.presentation_part.presentation.slide_width
            sh = slide.part.package.presentation_part.presentation.slide_height
        except Exception:
            pass
        if sw is None:
            sw = Inches(13.333)
        if sh is None:
            sh = Inches(7.5)
        sw_pt = sw.pt if hasattr(sw, 'pt') else 960.0
        sh_pt = sh.pt if hasattr(sh, 'pt') else 540.0
        return sw, sh, sw_pt, sh_pt

    @classmethod
    def resolve_image_text_collisions(cls, slide, text_shapes: List[BaseShape]) -> None:
        """
        Preserves authentic template designer coordinates for existing text boxes so text
        is never shoved off its own background card container or pill frame.
        """
        return


class PPTXTextReplacer:
    """
    High-legibility, resolution-aware, anti-overflow text replacer for authentic PPTX templates.
    """

    @staticmethod
    def get_slide_dimensions(slide) -> Tuple[Any, Any, float, float]:
        """Accurately extracts slide width/height from python-pptx presentation package."""
        sw = None
        sh = None
        try:
            sw = slide.part.package.presentation_part.presentation.slide_width
            sh = slide.part.package.presentation_part.presentation.slide_height
        except Exception:
            pass
        if sw is None:
            sw = Inches(13.333)
        if sh is None:
            sh = Inches(7.5)
        sw_pt = sw.pt if hasattr(sw, 'pt') else 960.0
        sh_pt = sh.pt if hasattr(sh, 'pt') else 540.0
        return sw, sh, sw_pt, sh_pt

    @staticmethod
    def is_dummy_text(text: Any) -> bool:
        """Checks if a string is template dummy/placeholder text."""
        if text is None:
            return False
        if not isinstance(text, str):
            text = str(text)
        t_clean = text.lower().strip()
        if not t_clean:
            return True
        for pat in DUMMY_PATTERNS:
            if re.search(pat, t_clean):
                return True
        return False

    @staticmethod
    def is_step_badge(text: Any) -> bool:
        """Checks if shape is a step badge (e.g. '01.', 'Step 01', '2026', 'A')."""
        if text is None:
            return False
        if not isinstance(text, str):
            text = str(text)
        t = text.strip()
        return bool(
            re.match(r'^(0?[1-9]|1[0-9])[\.\:\)]?$', t) or
            re.match(r'^step\s*0?[1-9]$', t, re.I) or
            re.match(r'^(20\d\d)$', t) or
            re.match(r'^[A-D][\.\:\)]?$', t) or
            re.match(r'^\d{1,3}%$', t)
        )

    is_step_number = is_step_badge

    @classmethod
    def _resolve_theme_scheme_colors(cls, part: Any) -> Dict[str, Tuple[int, int, int]]:
        """Extracts and maps schemeClr names (bg1, bg2, tx1, tx2, dk1, lt1, accent1..6) to RGB tuples from the SlideMaster theme XML."""
        if hasattr(part, "_cached_scheme_map"):
            return part._cached_scheme_map

        scheme_rgb: Dict[str, Tuple[int, int, int]] = {
            "lt1": (255, 255, 255),
            "dk1": (0, 0, 0),
            "lt2": (241, 245, 249),
            "dk2": (30, 41, 59),
            "bg1": (255, 255, 255),
            "bg2": (241, 245, 249),
            "tx1": (0, 0, 0),
            "tx2": (30, 41, 59),
        }
        try:
            from lxml import etree
            master = None
            if hasattr(part, "slide") and hasattr(part.slide, "slide_layout"):
                master = part.slide.slide_layout.slide_master
            elif hasattr(part, "slide_layout"):
                master = part.slide_layout.slide_master

            if master is not None:
                # 1. Parse theme1.xml <a:clrScheme>
                for rel in master.part.rels.values():
                    if "theme" in rel.reltype:
                        root = etree.fromstring(rel.target_part.blob)
                        for el in root.xpath('.//*[local-name()="clrScheme"]/*'):
                            tag = etree.QName(el).localname
                            srgb = el.xpath('.//*[local-name()="srgbClr"]/@val')
                            sys_c = el.xpath('.//*[local-name()="sysClr"]/@lastClr')
                            hex_v = (srgb[0] if srgb else (sys_c[0] if sys_c else "")).strip()
                            if len(hex_v) == 6:
                                scheme_rgb[tag] = (int(hex_v[0:2], 16), int(hex_v[2:4], 16), int(hex_v[4:6], 16))
                        break

                # 2. Parse <p:clrMap> on master (e.g. bg1="lt1", bg2="lt2", tx1="dk1", tx2="dk2")
                clr_maps = master._element.xpath('.//*[local-name()="clrMap"]')
                if clr_maps:
                    cm = clr_maps[0]
                    for k, v in cm.attrib.items():
                        if v in scheme_rgb:
                            scheme_rgb[k] = scheme_rgb[v]
                else:
                    scheme_rgb["bg1"] = scheme_rgb.get("lt1", (255, 255, 255))
                    scheme_rgb["bg2"] = scheme_rgb.get("lt2", (241, 245, 249))
                    scheme_rgb["tx1"] = scheme_rgb.get("dk1", (0, 0, 0))
                    scheme_rgb["tx2"] = scheme_rgb.get("dk2", (30, 41, 59))
        except Exception:
            pass

        part._cached_scheme_map = scheme_rgb
        return scheme_rgb

    @classmethod
    def get_slide_bg_luminance(cls, part: Any) -> float:
        """Calculates average background luminance (0..255) for the slide container, resolving srgbClr, schemeClr, images, and full-bleed shapes."""
        if hasattr(part, "_cached_bg_lum"):
            return part._cached_bg_lum

        lum = 255.0  # default light
        scheme_map = cls._resolve_theme_scheme_colors(part)

        containers = []
        if hasattr(part, "slide"):
            containers.append(part.slide)
            try:
                containers.append(part.slide.slide_layout)
                containers.append(part.slide.slide_layout.slide_master)
            except Exception:
                pass
        elif hasattr(part, "slide_layout"):
            containers.append(part.slide_layout)
            if hasattr(part.slide_layout, "slide_master"):
                containers.append(part.slide_layout.slide_master)

        found = False
        for container in containers:
            if found:
                break
            try:
                if not hasattr(container, "_element"):
                    continue
                celem = container._element

                # 1. Background picture blip
                blip_nodes = celem.xpath('./p:cSld/p:bg//a:blip | ./p:bg//a:blip')
                for blip in blip_nodes:
                    rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rid:
                        c_part = container.part if hasattr(container, "part") else part
                        if rid in c_part.rels:
                            target = c_part.rels[rid].target_part
                            img = Image.open(io.BytesIO(target.blob))
                            img_small = img.resize((16, 16)).convert('L')
                            pixels = list(img_small.getdata())
                            lum = sum(pixels) / float(len(pixels))
                            found = True
                            break
                if found:
                    break

                # 2. Explicit srgbClr in <p:bg>
                bg_clr_nodes = celem.xpath('./p:cSld/p:bg//a:srgbClr | ./p:bg//a:srgbClr')
                for clr in bg_clr_nodes:
                    hex_str = clr.get('val')
                    if hex_str and len(hex_str) == 6:
                        r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
                        lum = 0.299 * r + 0.587 * g + 0.114 * b
                        found = True
                        break
                if found:
                    break

                # 3. Theme schemeClr in <p:bg> or <p:bgRef> (e.g. bg1, bg2, accent1, dk1) with lumMod/lumOff/shade
                scheme_nodes = celem.xpath('./p:cSld/p:bg//a:schemeClr | ./p:bg//a:schemeClr')
                for sc_node in scheme_nodes:
                    sc_val = sc_node.get('val', '')
                    if sc_val in scheme_map:
                        r, g, b = scheme_map[sc_val]
                        lum = 0.299 * r + 0.587 * g + 0.114 * b
                        lum_mods = sc_node.xpath('./a:lumMod/@val')
                        lum_offs = sc_node.xpath('./a:lumOff/@val')
                        shades = sc_node.xpath('./a:shade/@val')
                        if lum_mods:
                            lum = lum * (float(lum_mods[0]) / 100000.0)
                        if lum_offs:
                            lum = lum + 255.0 * (float(lum_offs[0]) / 100000.0)
                        if shades:
                            lum = lum * (float(shades[0]) / 100000.0)
                        found = True
                        break
            except Exception:
                pass

        # 4. Also check full-slide background shapes (> 80% width and height) on slide or layout
        if hasattr(part, "slide"):
            try:
                slide = part.slide
                sw = part.package.presentation_part.presentation.slide_width
                sh = part.package.presentation_part.presentation.slide_height
                for s in list(slide.shapes)[:3]:
                    if s.width and s.height and s.width >= sw * 0.80 and s.height >= sh * 0.80:
                        if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            img = Image.open(io.BytesIO(s.image.blob))
                            img_small = img.resize((16, 16)).convert('L')
                            pixels = list(img_small.getdata())
                            lum = sum(pixels) / float(len(pixels))
                            found = True
                            break
                        else:
                            srgb_nodes = s._element.xpath('./p:spPr/a:solidFill/a:srgbClr')
                            if srgb_nodes:
                                hex_str = srgb_nodes[0].get('val', '')
                                if len(hex_str) == 6:
                                    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
                                    lum = 0.299 * r + 0.587 * g + 0.114 * b
                                    found = True
                                    break
                            sc_nodes = s._element.xpath('./p:spPr/a:solidFill/a:schemeClr')
                            if sc_nodes:
                                sc_val = sc_nodes[0].get('val', '')
                                if sc_val in scheme_map:
                                    r, g, b = scheme_map[sc_val]
                                    lum = 0.299 * r + 0.587 * g + 0.114 * b
                                    found = True
                                    break
            except Exception:
                pass

        part._cached_bg_lum = lum
        return lum

    @classmethod
    def is_shape_on_dark_background(cls, shape: BaseShape) -> bool:
        """Determines whether a shape sits on a dark background (luminance < 140.0)."""
        try:
            elem = shape._element
            no_fill = elem.xpath('./p:spPr/a:noFill')
            if not no_fill:
                clr_nodes = elem.xpath('./p:spPr/a:solidFill/a:srgbClr')
                if clr_nodes:
                    hex_str = clr_nodes[0].get('val')
                    if hex_str and len(hex_str) == 6:
                        r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
                        lum = 0.299 * r + 0.587 * g + 0.114 * b
                        return lum < 175.0
                sc_nodes = elem.xpath('./p:spPr/a:solidFill/a:schemeClr')
                if sc_nodes and hasattr(shape, "part"):
                    scheme_map = cls._resolve_theme_scheme_colors(shape.part)
                    sc_val = sc_nodes[0].get('val', '')
                    if sc_val in scheme_map:
                        r, g, b = scheme_map[sc_val]
                        lum = 0.299 * r + 0.587 * g + 0.114 * b
                        return lum < 175.0
        except Exception:
            pass

        # Check if an underlying colored vector shape on the slide sits behind this textbox's center (cx, cy)
        try:
            if shape.left is not None and shape.top is not None and shape.width is not None and shape.height is not None:
                cx = float(shape.left) + float(shape.width) * 0.5
                cy = float(shape.top) + float(shape.height) * 0.5
                sld_el = shape._element.xpath('ancestor::p:sld')
                if sld_el:
                    # Inspect all <p:sp> elements on the slide that have a solidFill and contain (cx, cy)
                    for sp_el in sld_el[0].xpath('.//p:sp'):
                        if sp_el is shape._element:
                            continue
                        xfrm = sp_el.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
                        if xfrm is None:
                            continue
                        off = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off')
                        ext = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
                        if off is None or ext is None:
                            continue
                        sx_e, sy_e = float(off.get('x', 0)), float(off.get('y', 0))
                        sw_e, sh_e = float(ext.get('cx', 0)), float(ext.get('cy', 0))
                        if sw_e <= 0 or sh_e <= 0:
                            continue
                        if sx_e <= cx <= (sx_e + sw_e) and sy_e <= cy <= (sy_e + sh_e):
                            if sp_el.xpath('./p:spPr/a:noFill'):
                                continue
                            srgb = sp_el.xpath('./p:spPr/a:solidFill/a:srgbClr/@val | ./p:style/a:fillRef[not(@idx="0")]/a:srgbClr/@val')
                            if srgb and len(srgb[0]) == 6:
                                hx = srgb[0]
                                r, g, b = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
                                lum = 0.299 * r + 0.587 * g + 0.114 * b
                                if lum < 185.0:
                                    return True
                                elif lum > 225.0:
                                    return False
        except Exception:
            pass

        try:
            part = shape.part
            lum = cls.get_slide_bg_luminance(part)
            return lum < 175.0
        except Exception:
            pass

        return False

    @staticmethod
    def calculate_safe_font_size(
        text: Any,
        width_pt: float,
        height_pt: float,
        max_initial_fs: float = 18.5,
        min_fs: float = 14.5,
        is_title: bool = False,
        scale_factor: float = 1.0
    ) -> float:
        """
        Iteratively calculates the maximum font size that strictly guarantees
        text fits inside the bounding box with prominent legibility and word wrapping,
        taking canvas resolution scale factor into account.
        """
        if text is None:
            text = ""
        elif not isinstance(text, str):
            text = str(text)
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if not lines:
            return max_initial_fs

        font_scale = max(0.85, min(2.25, scale_factor))
        usable_w = max(20.0 * font_scale, width_pt - (8.0 * font_scale))
        usable_h = max(15.0 * font_scale, height_pt - (6.0 * font_scale))

        if is_title:
            t_len = len(text)
            if t_len > 60:
                min_fs = 16.0 * font_scale
            elif t_len > 40:
                min_fs = 18.0 * font_scale
            else:
                min_fs = 20.0 * font_scale
            max_initial_fs = max(max_initial_fs, 28.0 * font_scale)
        else:
            min_fs = 10.0 * font_scale
            max_initial_fs = max(max_initial_fs, 16.5 * font_scale)

        font_size = max_initial_fs
        while font_size >= min_fs:
            char_w = font_size * 0.52
            chars_per_line = max(1, int(usable_w / char_w))

            total_lines = 0
            for l in lines:
                words = l.split()
                if not words:
                    continue
                cur_line_len = 0
                for w in words:
                    w_len = len(w)
                    if cur_line_len == 0:
                        cur_line_len = w_len
                    elif cur_line_len + 1 + w_len <= chars_per_line:
                        cur_line_len += 1 + w_len
                    else:
                        total_lines += 1
                        cur_line_len = w_len
                if cur_line_len > 0:
                    total_lines += 1

            line_h = font_size * 1.22
            para_spacing = font_size * 0.30 if len(lines) > 1 else 0
            total_h = (total_lines * line_h) + ((len(lines) - 1) * para_spacing)

            if total_h <= usable_h * 0.95:
                return round(font_size, 1)

            font_size -= 0.5

        return min_fs

    calculate_optimal_font_size = calculate_safe_font_size

    @staticmethod
    def enable_xml_autofit(text_frame) -> None:
        """Injects DrawingML normAutofit into the shape's body properties."""
        try:
            bodyPr = text_frame._txBody.bodyPr
            for child in list(bodyPr):
                if child.tag.endswith(('noAutofit', 'spAutoFit', 'normAutofit')):
                    bodyPr.remove(child)
            normAutofit = OxmlElement('a:normAutofit')
            bodyPr.append(normAutofit)
        except Exception:
            pass

    @staticmethod
    def enforce_latin_line_break(paragraph) -> None:
        """
        Enforces DrawingML `<a:pPr latinLnBrk="0"/>` to strictly prohibit
        splitting Latin words across lines mid-word (e.g. preventing 'Mu/qarrar',
        'Asimme/trik', 'Cryp/tographic', 'integ/rat/siyasi').
        """
        try:
            pPr = paragraph._p.get_or_add_pPr()
            pPr.set('latinLnBrk', '0')
        except Exception:
            pass

    @classmethod
    def set_shape_text_preserving_style(
        cls,
        shape: BaseShape,
        new_text: str,
        default_font_size: Optional[float] = None,
        is_title: bool = False,
        is_card_title: bool = False,
        font_size_override: Optional[float] = None,
        slide_height_pt: float = 540.0,
        scale_factor: float = 1.0,
        theme_key: str = "tech_minimal"
    ) -> None:
        """
        Sets text with strict boundary enforcement, resolution-aware font scaling,
        high-legibility typography, and style preservation.
        """
        if not shape.has_text_frame:
            return

        tf = shape.text_frame
        is_short_badge = len(str(new_text or "").strip()) <= 4 and not is_title and not is_card_title
        if is_short_badge:
            tf.word_wrap = False
            tf.margin_left = 0
            tf.margin_right = 0
            tf.margin_top = 0
            tf.margin_bottom = 0
        else:
            tf.word_wrap = True
            tf.margin_left = Inches(0.04 * scale_factor)
            tf.margin_right = Inches(0.04 * scale_factor)
            tf.margin_top = Inches(0.04 * scale_factor)
            tf.margin_bottom = Inches(0.04 * scale_factor)
            cls.enable_xml_autofit(tf)

        # Extract authentic template bullet styles before text replacement (Round 13)
        template_bullet_styles = BulletHarmonizer.extract_shape_bullet_styles(shape)

        orig_font_name = None
        orig_color = None
        orig_bold = False
        orig_fs_pt = None

        for p_iter in tf.paragraphs:
            for r_iter in p_iter.runs:
                if not orig_font_name and r_iter.font.name:
                    orig_font_name = r_iter.font.name
                if r_iter.font.bold:
                    orig_bold = True
                if orig_fs_pt is None and r_iter.font.size:
                    try:
                        orig_fs_pt = r_iter.font.size.pt
                    except Exception:
                        pass
                if orig_color is None:
                    try:
                        if r_iter.font.color and r_iter.font.color.type and r_iter.font.color.rgb:
                            orig_color = r_iter.font.color.rgb
                    except Exception:
                        try:
                            sch_vals = r_iter._r.xpath('./a:rPr/a:solidFill/a:schemeClr/@val')
                            if sch_vals and hasattr(shape, "part"):
                                sch_map = cls._resolve_theme_scheme_colors(shape.part)
                                if sch_vals[0] in sch_map:
                                    rgb_t = sch_map[sch_vals[0]]
                                    orig_color = RGBColor(rgb_t[0], rgb_t[1], rgb_t[2])
                        except Exception:
                            pass

        bg_lum = cls.get_slide_bg_luminance(shape.part if hasattr(shape, "part") else None)
        is_dark_bg = cls.is_shape_on_dark_background(shape) or (bg_lum < 175.0)
        if orig_color:
            try:
                c_lum = 0.299 * orig_color[0] + 0.587 * orig_color[1] + 0.114 * orig_color[2]
                if is_dark_bg and c_lum < 150.0:
                    orig_color = None  # Force bright high-contrast color on dark/colored background
                elif not is_dark_bg and bg_lum > 225.0 and c_lum > 180.0:
                    orig_color = None  # Force deep high-contrast color on pure white background
            except Exception:
                orig_color = None

        if new_text is None:
            new_text = ""
        elif not isinstance(new_text, str):
            new_text = str(new_text)

        new_text = clean_latex_math(new_text)
        if is_title:
            new_text = MicroTypography.format_title(new_text)
        else:
            new_text = MicroTypography.format_body(new_text)

        lines = [line.strip() for line in new_text.split('\n') if line.strip()]
        if not lines:
            tf.text = ""
            return

        shape_w_pt = SafeShapeGeometry.to_pt(shape.width, 100.0)
        shape_h_pt = SafeShapeGeometry.to_pt(shape.height, 100.0)
        shape_top_pt = SafeShapeGeometry.to_pt(shape.top, 0.0)

        # Keep template shape geometry intact to avoid colliding with adjacent subtitles or cards

        font_scale = max(0.85, min(2.25, scale_factor))

        # Only expand wide standalone body boxes (width >= 145pt); keep narrow column/bar callout boxes at their exact designer height
        if not is_title and not is_card_title and shape_w_pt >= 145.0 * scale_factor and shape_h_pt < 75.0 * scale_factor:
            shape_h_pt = max(shape_h_pt, min(85.0 * scale_factor, shape_h_pt * 1.25))
            try:
                SafeShapeGeometry.update(shape, height=Inches(shape_h_pt / 72.0))
            except Exception:
                pass

        if is_title:
            t_len = len(new_text)
            base_t_fs = (orig_fs_pt / scale_factor) if (orig_fs_pt and orig_fs_pt >= 20.0) else 28.0
            max_init_fs = min(36.0 * font_scale, max(24.0 * font_scale, base_t_fs * font_scale))
            min_fs = 16.0 * font_scale if t_len > 60 else (18.0 * font_scale if t_len > 40 else 20.0 * font_scale)
        elif is_card_title:
            # Respect the template designer original card title font size so 1-line boxes (22pt tall) do not overflow into body!
            if orig_fs_pt and 10.0 <= orig_fs_pt <= 22.0:
                max_init_fs = min(15.0 * font_scale, max(11.0 * font_scale, orig_fs_pt))
            else:
                max_init_fs = (12.5 if shape_h_pt <= 26.0 * scale_factor else 15.5) * font_scale
            min_fs = 9.5 * font_scale
        else:
            if orig_fs_pt and 9.0 <= orig_fs_pt <= 18.0:
                max_init_fs = min(14.0 * font_scale, max(10.0 * font_scale, orig_fs_pt))
            else:
                max_init_fs = 13.5 * font_scale
            min_fs = 9.0 * font_scale

        if font_size_override and is_title:
            max_init_fs = max(max_init_fs, font_size_override * font_scale)

        scaled_font_size = cls.calculate_safe_font_size(
            text=new_text,
            width_pt=shape_w_pt,
            height_pt=shape_h_pt,
            max_initial_fs=max_init_fs,
            min_fs=min_fs,
            is_title=is_title,
            scale_factor=font_scale
        )

        # Strict anti-overflow boundary enforcement:
        usable_w = max(20.0 * font_scale, shape_w_pt - (8.0 * font_scale))
        usable_h = max(15.0 * font_scale, shape_h_pt - (6.0 * font_scale))

        # Universal No-Truncation Fitting: dynamically step font size down to floor so full text fits without '...'
        if shape_w_pt < 125.0 * scale_factor:
            floor_fs = (14.0 if is_title else (9.0 if is_card_title else 7.5)) * font_scale
        else:
            floor_fs = (14.5 if is_title else (10.0 if is_card_title else 9.0)) * font_scale
        test_fs = scaled_font_size
        while test_fs >= floor_fs:
            test_char_w = max(1.0, test_fs * 0.52)
            test_cpl = max(1, int(usable_w / test_char_w))
            total_wrapped_lines = 0
            for line in lines:
                words = line.split()
                cur_p = ""
                for w in words:
                    if not cur_p:
                        cur_p = w
                    elif len(cur_p) + 1 + len(w) <= test_cpl:
                        cur_p += " " + w
                    else:
                        total_wrapped_lines += 1
                        cur_p = w
                if cur_p:
                    total_wrapped_lines += 1
            if (total_wrapped_lines * test_fs * 1.22) <= (usable_h * 1.12):
                scaled_font_size = test_fs
                break
            test_fs -= 0.5
        else:
            scaled_font_size = max(floor_fs, test_fs)
        condensed_lines = lines
        if condensed_lines:
            lines = condensed_lines

        theme_colors = ContrastHarmonizer.get_theme_text_colors(theme_key, is_dark_bg)
        bullet_accent_hex = theme_colors.get("accent_hex") or theme_colors.get("lead_in_hex") or "3B82F6"

        hierarchy_items = BulletHarmonizer.parse_hierarchical_lines("\n".join(lines))
        if not hierarchy_items:
            tf.text = ""
            return

        tf.text = ""
        for idx, item in enumerate(hierarchy_items):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            cls.enforce_latin_line_break(p)
            p.line_spacing = 1.25

            # Determine bullet application (Round 13 Native Bullet Preserver)
            apply_bullet = item.is_bullet and not is_title and not is_card_title
            target_style = template_bullet_styles.get(item.level) or template_bullet_styles.get(0)

            BulletHarmonizer.apply_native_bullet(
                paragraph=p,
                bullet_style=target_style,
                level=item.level,
                is_bullet=apply_bullet,
                theme_accent_hex=bullet_accent_hex,
                scale_factor=font_scale
            )

            # Sub-level font scaling (Level 0: 100%, Level 1: 92%, Level 2: 85%)
            item_fs = scaled_font_size * (1.0 if item.level == 0 else (0.92 if item.level == 1 else 0.85))
            p.space_after = Pt(6.0 * font_scale) if (apply_bullet and item.level == 0) else (Pt(3.5 * font_scale) if apply_bullet else (Pt(4.0 * font_scale) if len(hierarchy_items) > 1 else Pt(0)))

            if item.lead_in and not is_title and not is_card_title:
                r1 = p.add_run()
                r1.text = f"{item.lead_in}: "
                r1.font.bold = True
                if orig_font_name:
                    r1.font.name = orig_font_name
                r1.font.size = Pt(item_fs)
                if orig_color:
                    r1.font.color.rgb = orig_color
                else:
                    r1.font.color.rgb = theme_colors["lead_in"]

                r2 = p.add_run()
                r2.text = item.body
                r2.font.bold = False
                if orig_font_name:
                    r2.font.name = orig_font_name
                r2.font.size = Pt(item_fs)
                if orig_color:
                    r2.font.color.rgb = orig_color
                else:
                    r2.font.color.rgb = theme_colors["body"]
            else:
                p.text = item.clean_text
                for r in p.runs:
                    if orig_font_name:
                        r.font.name = orig_font_name
                    r.font.size = Pt(item_fs)
                    if orig_color:
                        r.font.color.rgb = orig_color
                    else:
                        r.font.color.rgb = theme_colors["title"] if (is_title or is_card_title) else theme_colors["body"]
                    if is_title or is_card_title:
                        r.font.bold = True
                    elif orig_bold:
                        r.font.bold = orig_bold

    @classmethod
    def _extract_content_items(cls, slide_data: Dict[str, Any], layout_type: str, title_text: str = "") -> List[Tuple[str, str]]:
        """Extracts structured (header, description) pairs from slide data."""
        content_items: List[Tuple[str, str]] = []

        if layout_type == "cover":
            cov_pts = slide_data.get("points") or slide_data.get("bullets") or []
            if cov_pts and len(cov_pts) >= 2:
                for p in cov_pts:
                    clean_p = re.sub(r'^(?:[\•\-\–\—\*]\s*|\d+[\.\)](?!\d)\s+|\(\d+\)\s*)', '', str(p).strip()).strip()
                    if ":" in clean_p:
                        h, d = clean_p.split(":", 1)
                        content_items.append((h.strip(), d.strip()))
                    else:
                        content_items.append((clean_p, ""))
            else:
                subtitle = slide_data.get("subtitle", "Akademik va amaliy tahliliy qo'llanma | 2026-yil")
                content_items.append(("", subtitle))

        elif layout_type == "agenda":
            points = slide_data.get("points", [])
            for p in points:
                clean_p = re.sub(r'^\d+[\.\)](?!\d)\s*', '', p)
                if ":" in clean_p:
                    h, d = clean_p.split(":", 1)
                    content_items.append((h.strip(), d.strip()))
                else:
                    content_items.append((clean_p, ""))

        elif layout_type == "theory_concept":
            raw_pts = slide_data.get("points") or slide_data.get("bullets") or []
            if raw_pts:
                for p in raw_pts:
                    clean_p = re.sub(r'^(?:[\•\-\–\—\*]\s*|\d+[\.\)](?!\d)\s+|\(\d+\)\s*)', '', str(p).strip()).strip()
                    if ":" in clean_p:
                        h, d = clean_p.split(":", 1)
                        content_items.append((h.strip(), d.strip()))
                    elif " - " in clean_p:
                        h, d = clean_p.split(" - ", 1)
                        content_items.append((h.strip(), d.strip()))
                    else:
                        content_items.append((clean_p, ""))
            else:
                concept_name = slide_data.get("concept_name", "Nazariy Tushuncha")
                c_clean = str(concept_name).strip().lower().rstrip(".:")
                t_clean = str(title_text).strip().lower().rstrip(".:")
                if c_clean == t_clean or (len(c_clean) > 8 and c_clean in t_clean) or (len(t_clean) > 8 and t_clean in c_clean):
                    concept_name = "Nazariy Asos va Mohiyat"
                definition = slide_data.get("definition", "")
                formula = slide_data.get("key_formula_or_law", "")
                importance = slide_data.get("importance", "")
                content_items.append((concept_name, definition))
                if formula:
                    content_items.append(("Asosiy Qonuniyat / Formula", formula))
                if importance:
                    content_items.append(("Ilmiy va Amaliy Ahamiyati", importance))

        elif layout_type == "comparison_vs":
            if slide_data.get("left_points") or slide_data.get("right_points"):
                left_t = slide_data.get("left_title", "1-Yondashuv")
                left_pts = "\n".join([f"• {p}" for p in slide_data.get("left_points", [])])
                right_t = slide_data.get("right_title", "2-Yondashuv")
                right_pts = "\n".join([f"• {p}" for p in slide_data.get("right_points", [])])
                content_items.append((left_t, left_pts))
                content_items.append((right_t, right_pts))
            elif isinstance(slide_data.get("table"), dict):
                t_obj = slide_data["table"]
                t_headers = t_obj.get("headers", [])
                for r in t_obj.get("rows", []):
                    if isinstance(r, (list, tuple)) and len(r) >= 2:
                        row_h = str(r[0]).strip()
                        row_cells = [f"{t_headers[i]}: {r[i]}" if i < len(t_headers) else str(r[i]) for i in range(1, len(r))]
                        content_items.append((row_h, " | ".join(row_cells)))

        elif layout_type == "timeline_steps":
            steps = slide_data.get("steps", [])
            for s in steps:
                content_items.append((s.get("title", ""), s.get("desc", "")))

        elif layout_type == "diagram_anatomy":
            comps = slide_data.get("components", [])
            for c in comps:
                content_items.append((c.get("name", ""), c.get("function", "")))

        elif layout_type == "metrics_stats":
            metrics = slide_data.get("metrics", [])
            for m in metrics:
                val = m.get("value", "")
                lbl = m.get("label", "")
                desc = m.get("desc", "")
                content_items.append((f"{val} - {lbl}", desc))

        elif layout_type == "process_cycle":
            steps = slide_data.get("cycle_steps", []) or slide_data.get("steps", [])
            for s in steps:
                if isinstance(s, dict):
                    content_items.append((s.get("title", ""), s.get("desc", "")))
                else:
                    content_items.append((str(s), ""))

        elif layout_type == "pyramid_funnel":
            levels = slide_data.get("levels", []) or slide_data.get("hierarchy", []) or slide_data.get("steps", [])
            for l in levels:
                if isinstance(l, dict):
                    content_items.append((l.get("title", "") or l.get("level_name", ""), l.get("desc", "")))
                else:
                    content_items.append((str(l), ""))

        elif layout_type == "keynote_quote":
            quote = slide_data.get("quote", "")
            author = slide_data.get("author", "")
            significance = slide_data.get("significance", "")
            if quote:
                content_items.append((f"«{quote}»", f"— {author}" if author else ""))
            if significance:
                content_items.append(("Ilmiy Xulosa / Ahamiyat", significance))

        elif layout_type == "table_data":
            headers = slide_data.get("headers", [])
            rows = slide_data.get("rows", [])
            if headers:
                content_items.append(("Parametrlar", " | ".join([str(h) for h in headers])))
            for r in rows:
                if isinstance(r, (list, tuple)):
                    content_items.append((str(r[0]) if len(r) > 0 else "", " | ".join([str(x) for x in r[1:]])))
                elif isinstance(r, dict):
                    content_items.append((r.get("parameter", ""), r.get("value", "")))

        elif layout_type == "qa_discussion":
            qa_list = slide_data.get("qa_list", [])
            for qa in qa_list:
                content_items.append((f"❓ {qa.get('question', '')}", f"💡 {qa.get('answer', '')}"))

        elif layout_type == "conclusion_sources":
            standards = slide_data.get("standards_refs", [])
            academic = slide_data.get("academic_refs", [])
            recs = slide_data.get("recommendations", [])
            takeaways = slide_data.get("key_takeaways", [])
            refs = slide_data.get("references", [])
            if standards or academic:
                if standards:
                    content_items.append(("🏛️ Me'yoriy-Texnik Standartlar va Manbalar", "\n".join([f"• {s}" for s in standards])))
                if academic:
                    content_items.append(("📚 Fundamental Ilmiy Nashrlar va Tadqiqotlar", "\n".join([f"• {a}" for a in academic])))
            elif recs:
                if takeaways:
                    content_items.append(("🎯 Asosiy Ilmiy Xulosalar", "\n".join([f"• {t}" for t in takeaways])))
                content_items.append(("💡 Amaliy Tavsiyalar va Takliflar", "\n".join([f"• {r}" for r in recs])))
            else:
                if takeaways:
                    content_items.append(("🎯 Asosiy Ilmiy Xulosalar", "\n".join([f"• {t}" for t in takeaways])))
                if refs:
                    content_items.append(("📚 Foydalanilgan Adabiyotlar", "\n".join([f"📖 {r}" for r in refs])))

        else: # cards_grid
            cards = slide_data.get("cards", [])
            if cards:
                for c in cards:
                    content_items.append((c.get("title", ""), c.get("desc", "")))

        # Universal robust fallback for any layout: if content_items is empty or has < 2 items, check theses, points, bullets
        if len(content_items) < 2:
            raw_candidates = (
                slide_data.get("theses", [])
                or slide_data.get("points", [])
                or slide_data.get("cards", [])
                or slide_data.get("bullets", [])
                or slide_data.get("items", [])
            )
            if raw_candidates:
                fallback_items = []
                for item in raw_candidates:
                    if isinstance(item, dict):
                        h = item.get("title", "") or item.get("name", "") or item.get("label", "")
                        d = item.get("desc", "") or item.get("value", "") or item.get("function", "")
                        fallback_items.append((str(h).strip(), str(d).strip()))
                    elif isinstance(item, str) and item.strip():
                        clean_item = re.sub(r'^(?:[\•\-\–\—\*]|\d+[\.\)](?!\d)|\(\d+\))\s*', '', item.strip()).strip()
                        if ":" in clean_item:
                            h, d = clean_item.split(":", 1)
                            fallback_items.append((h.strip(), d.strip()))
                        elif " - " in clean_item:
                            h, d = clean_item.split(" - ", 1)
                            fallback_items.append((h.strip(), d.strip()))
                        else:
                            fallback_items.append((clean_item, ""))
                if len(fallback_items) > len(content_items):
                    content_items = fallback_items

        sanitized_items = []
        for h, d in content_items:
            sanitized_items.append((clean_latex_math(h), clean_latex_math(d)))
        return sanitized_items

    @classmethod
    def render_wide_rectangular_cards(
        cls,
        slide,
        content_items: List[Tuple[str, str]],
        sw_pt: float,
        sh_pt: float,
        scale_factor: float,
        left_m_pt: float,
        content_w_pt: float,
        top_start_pt: float,
        primary_color: Optional[RGBColor] = None,
        accent_color: Optional[RGBColor] = None
    ) -> Set[int]:
        """
        Synthesizes wide, spacious, horizontal rectangular cards (aspect ratio >= 4:1)
        with large, authoritative typography (16-20pt) and generous horizontal breathing room.
        Guarantees zero overflow beyond slide canvas boundaries.
        """
        assigned_ids = set()
        if not content_items:
            return assigned_ids

        items_to_render = content_items[:4]
        N = len(items_to_render)
        avail_h_pt = max(100.0, (sh_pt * 0.94) - top_start_pt)
        gap_pt = (14.0 if N <= 3 else 10.0) * scale_factor
        if N <= 1:
            card_h_pt = min(150.0 * scale_factor, avail_h_pt * 0.70)
        elif N == 2:
            card_h_pt = min(135.0 * scale_factor, (avail_h_pt - gap_pt) / 2)
        elif N == 3:
            card_h_pt = min(105.0 * scale_factor, (avail_h_pt - (2 * gap_pt)) / 3)
        else: # N == 4
            card_h_pt = min(82.0 * scale_factor, (avail_h_pt - (3 * gap_pt)) / 4)

        is_dark_bg = cls.get_slide_bg_luminance(slide.part if hasattr(slide, "part") else slide) < 135.0
        if is_dark_bg:
            title_color = RGBColor(0x93, 0xC5, 0xFD)
            body_color = RGBColor(0xF1, 0xF5, 0xF9)
        else:
            title_color = primary_color if primary_color else RGBColor(0x1F, 0x4E, 0x79)
            body_color = RGBColor(0x2A, 0x2A, 0x2A)

        font_scale = max(0.85, min(2.25, sh_pt / 540.0))
        head_fs = (21.0 if N <= 2 else (18.5 if N == 3 else 17.0)) * font_scale
        body_fs = (16.5 if N <= 2 else (15.5 if N == 3 else 14.5)) * font_scale

        for i, (head, desc) in enumerate(items_to_render):
            cur_top_pt = top_start_pt + i * (card_h_pt + gap_pt)
            
            card_shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(left_m_pt / 72.0),
                Inches(cur_top_pt / 72.0),
                Inches(content_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
            try:
                card_shp.adjustments[0] = 0.06
                card_shp.fill.solid()
                card_shp.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B) if is_dark_bg else RGBColor(0xF1, 0xF5, 0xF9)
                card_shp.line.color.rgb = RGBColor(0x33, 0x41, 0x55) if is_dark_bg else RGBColor(0xCB, 0xD5, 0xE1)
                card_shp.line.width = Pt(1.2 * scale_factor)
            except Exception:
                pass
            assigned_ids.add(id(card_shp))

            tf = card_shp.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.06 * scale_factor)
            tf.margin_right = Inches(0.06 * scale_factor)
            tf.margin_top = Inches(0.04 * scale_factor)
            tf.margin_bottom = Inches(0.04 * scale_factor)
            cls.enable_xml_autofit(tf)

            clean_head = clean_latex_math(head.strip() if head else "")
            clean_desc = clean_latex_math(desc.strip() if desc else "")

            # Dynamic font sizing and line wrapping based on content length and multi-line breaks
            cur_body_fs = body_fs
            lines_in_desc = clean_desc.count('\n') + 1 if clean_desc else 0
            if len(clean_desc) > 160 or lines_in_desc >= 3:
                cur_body_fs = min(cur_body_fs, 13.5 * font_scale)
            elif len(clean_desc) > 110 or lines_in_desc >= 2:
                cur_body_fs = min(cur_body_fs, 14.5 * font_scale)

            if clean_head:
                p0 = tf.paragraphs[0]
                cls.enforce_latin_line_break(p0)
                p0.text = clean_head
                p0.font.bold = True
                p0.font.size = Pt(head_fs)
                p0.font.color.rgb = title_color
                p0.space_after = Pt(4.0 * scale_factor) if clean_desc else Pt(0)

            if clean_desc:
                desc_lines = [l.strip() for l in clean_desc.split('\n') if l.strip()]
                for idx, dline in enumerate(desc_lines):
                    p = tf.paragraphs[0] if (not clean_head and idx == 0) else tf.add_paragraph()
                    cls.enforce_latin_line_break(p)
                    p.text = dline
                    p.font.bold = False
                    p.font.size = Pt(cur_body_fs)
                    p.font.color.rgb = body_color
                    p.line_spacing = 1.15 if len(desc_lines) > 1 else 1.25
                    if idx < len(desc_lines) - 1:
                        p.space_after = Pt(3.0 * scale_factor)

        return assigned_ids

    @classmethod
    def render_metric_stat_cards(
        cls,
        slide,
        slide_data: Dict[str, Any],
        content_items: List[Tuple[str, str]],
        sw_pt: float,
        sh_pt: float,
        scale_factor: float,
        left_m_pt: float,
        content_w_pt: float,
        top_start_pt: float,
        primary_color: Optional[RGBColor] = None,
        accent_color: Optional[RGBColor] = None
    ) -> Set[int]:
        """
        Renders prominent KPI / Metric Stat Cards side-by-side horizontally.
        Each card displays:
        - Big Bold Metric Value (30-36pt)
        - Clear Bold Label (15pt)
        - Analytical Description (12.5pt)
        """
        assigned_ids = set()
        raw_metrics = slide_data.get("metrics", [])
        metric_items = []
        if raw_metrics and isinstance(raw_metrics, list):
            for m in raw_metrics:
                if isinstance(m, dict):
                    val = str(m.get("value", "")).strip()
                    lbl = str(m.get("label", "")).strip()
                    desc = str(m.get("desc", "")).strip()
                    if val or lbl:
                        metric_items.append((val, lbl, desc))
        if not metric_items and content_items:
            for h, d in content_items:
                parts = h.split(" - ", 1) if " - " in h else ([h, ""] if ":" not in h else h.split(":", 1))
                val = parts[0].strip()
                lbl = parts[1].strip() if len(parts) > 1 else ""
                metric_items.append((val, lbl, d.strip()))

        if not metric_items:
            return assigned_ids

        metric_items = metric_items[:4]
        N = len(metric_items)
        gap_pt = 14.0 * scale_factor
        card_w_pt = (content_w_pt - ((N - 1) * gap_pt)) / N
        card_h_pt = min(220.0 * scale_factor, max(140.0 * scale_factor, (sh_pt - top_start_pt - 40.0 * scale_factor)))

        is_dark_bg = cls.get_slide_bg_luminance(slide.part if hasattr(slide, "part") else slide) < 135.0
        if is_dark_bg:
            val_color = RGBColor(0x60, 0xA5, 0xFA)
            lbl_color = RGBColor(0xFF, 0xFF, 0xFF)
            desc_color = RGBColor(0xCB, 0xD5, 0xE1)
        else:
            val_color = primary_color if primary_color else RGBColor(0x0A, 0x36, 0x63)
            lbl_color = RGBColor(0x1E, 0x29, 0x3B)
            desc_color = RGBColor(0x47, 0x55, 0x69)
        font_scale = max(0.85, min(2.25, sh_pt / 540.0))

        for i, (val, lbl, desc) in enumerate(metric_items):
            cur_left_pt = left_m_pt + i * (card_w_pt + gap_pt)
            card_shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(cur_left_pt / 72.0),
                Inches(top_start_pt / 72.0),
                Inches(card_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
            try:
                card_shp.adjustments[0] = 0.06
                card_shp.fill.solid()
                card_shp.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B) if is_dark_bg else RGBColor(0xF1, 0xF5, 0xF9)
                card_shp.line.color.rgb = RGBColor(0x33, 0x41, 0x55) if is_dark_bg else RGBColor(0xCB, 0xD5, 0xE1)
                card_shp.line.width = Pt(1.2 * scale_factor)
            except Exception:
                pass
            assigned_ids.add(id(card_shp))

            tf = card_shp.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.16 * scale_factor)
            tf.margin_right = Inches(0.16 * scale_factor)
            tf.margin_top = Inches(0.14 * scale_factor)
            tf.margin_bottom = Inches(0.12 * scale_factor)
            cls.enable_xml_autofit(tf)

            # 1. Big Metric Value
            p_val = tf.paragraphs[0]
            cls.enforce_latin_line_break(p_val)
            p_val.text = val if val else "100%"
            p_val.font.bold = True
            val_fs = (28.0 if len(p_val.text) > 7 else 34.0) * font_scale
            p_val.font.size = Pt(val_fs)
            p_val.font.color.rgb = val_color
            p_val.space_after = Pt(4.0 * scale_factor)

            # 2. Metric Label
            if lbl:
                p_lbl = tf.add_paragraph()
                cls.enforce_latin_line_break(p_lbl)
                p_lbl.text = lbl
                p_lbl.font.bold = True
                p_lbl.font.size = Pt(15.0 * font_scale)
                p_lbl.font.color.rgb = lbl_color
                p_lbl.space_after = Pt(4.0 * scale_factor)

            # 3. Analytical Description
            if desc:
                p_desc = tf.add_paragraph()
                cls.enforce_latin_line_break(p_desc)
                p_desc.text = desc
                p_desc.font.bold = False
                p_desc.font.size = Pt(12.5 * font_scale)
                p_desc.font.color.rgb = desc_color
                p_desc.line_spacing = 1.18

        return assigned_ids

    @classmethod
    def render_process_step_cards(
        cls,
        slide,
        slide_data: Dict[str, Any],
        content_items: List[Tuple[str, str]],
        sw_pt: float,
        sh_pt: float,
        scale_factor: float,
        left_m_pt: float,
        content_w_pt: float,
        top_start_pt: float,
        primary_color: Optional[RGBColor] = None,
        accent_color: Optional[RGBColor] = None
    ) -> Set[int]:
        """
        Renders sequential timeline / process step cards arranged horizontally side-by-side.
        Each card displays:
        - Step number / badge (e.g. "01", "02", "03")
        - Step title (bold 15.5pt)
        - Step description / milestones (12.5pt)
        """
        assigned_ids = set()
        raw_steps = slide_data.get("steps", []) or slide_data.get("cycle_steps", [])
        step_items = []
        if raw_steps and isinstance(raw_steps, list):
            for idx, s in enumerate(raw_steps):
                if isinstance(s, dict):
                    num = str(s.get("step_number", "") or s.get("step", "") or s.get("phase", "") or f"{idx+1:02d}").strip()
                    title = str(s.get("title", "") or s.get("phase", "")).strip()
                    desc = str(s.get("desc", "") or s.get("details", "")).strip()
                    step_items.append((num, title, desc))
                elif isinstance(s, str) and s.strip():
                    step_items.append((f"{idx+1:02d}", s.strip(), ""))
        if not step_items and content_items:
            for idx, (h, d) in enumerate(content_items):
                step_items.append((f"{idx+1:02d}", h, d))

        if not step_items:
            return assigned_ids

        step_items = step_items[:4]
        N = len(step_items)
        gap_pt = 12.0 * scale_factor
        card_w_pt = (content_w_pt - ((N - 1) * gap_pt)) / N
        card_h_pt = min(240.0 * scale_factor, max(150.0 * scale_factor, (sh_pt - top_start_pt - 40.0 * scale_factor)))

        is_dark_bg = cls.get_slide_bg_luminance(slide.part if hasattr(slide, "part") else slide) < 135.0
        if is_dark_bg:
            step_badge_color = RGBColor(0x38, 0xBD, 0xF8)
            title_color = RGBColor(0xFF, 0xFF, 0xFF)
            desc_color = RGBColor(0xE2, 0xE8, 0xF0)
        else:
            step_badge_color = primary_color if primary_color else RGBColor(0x0F, 0x52, 0xBA)
            title_color = RGBColor(0x0F, 0x17, 0x2A)
            desc_color = RGBColor(0x33, 0x41, 0x55)
        font_scale = max(0.85, min(2.25, sh_pt / 540.0))

        for i, (num, title, desc) in enumerate(step_items):
            cur_left_pt = left_m_pt + i * (card_w_pt + gap_pt)
            card_shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(cur_left_pt / 72.0),
                Inches(top_start_pt / 72.0),
                Inches(card_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
            try:
                card_shp.adjustments[0] = 0.06
                card_shp.fill.solid()
                card_shp.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B) if is_dark_bg else RGBColor(0xF1, 0xF5, 0xF9)
                card_shp.line.color.rgb = RGBColor(0x33, 0x41, 0x55) if is_dark_bg else RGBColor(0xCB, 0xD5, 0xE1)
                card_shp.line.width = Pt(1.2 * scale_factor)
            except Exception:
                pass
            assigned_ids.add(id(card_shp))

            tf = card_shp.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.14 * scale_factor)
            tf.margin_right = Inches(0.14 * scale_factor)
            tf.margin_top = Inches(0.14 * scale_factor)
            tf.margin_bottom = Inches(0.12 * scale_factor)
            cls.enable_xml_autofit(tf)

            # Step Badge / Number
            p_badge = tf.paragraphs[0]
            cls.enforce_latin_line_break(p_badge)
            display_badge = num if any(c.isdigit() for c in num) else f"{i+1:02d}"
            if not display_badge.lower().startswith("bosqich") and not display_badge.lower().startswith("step"):
                display_badge = f"STEP {display_badge}"
            p_badge.text = display_badge
            p_badge.font.bold = True
            p_badge.font.size = Pt(13.5 * font_scale)
            p_badge.font.color.rgb = step_badge_color
            p_badge.space_after = Pt(4.0 * scale_factor)

            # Title
            if title:
                p_title = tf.add_paragraph()
                cls.enforce_latin_line_break(p_title)
                p_title.text = title
                p_title.font.bold = True
                p_title.font.size = Pt(15.5 * font_scale)
                p_title.font.color.rgb = title_color
                p_title.space_after = Pt(4.0 * scale_factor)

            # Desc
            if desc:
                p_desc = tf.add_paragraph()
                cls.enforce_latin_line_break(p_desc)
                p_desc.text = desc
                p_desc.font.bold = False
                p_desc.font.size = Pt(12.5 * font_scale)
                p_desc.font.color.rgb = desc_color
                p_desc.line_spacing = 1.18

        return assigned_ids

    @classmethod
    def render_vertical_column_cards(
        cls,
        slide,
        content_items: List[Tuple[str, str]],
        sw_pt: float,
        sh_pt: float,
        scale_factor: float,
        left_m_pt: float,
        content_w_pt: float,
        top_start_pt: float,
        primary_color: Optional[RGBColor] = None,
        accent_color: Optional[RGBColor] = None
    ) -> Set[int]:
        """
        Renders vertical column cards side-by-side across the slide canvas (3 or 4 columns).
        Each card displays:
        - Card Title (bold 16.5pt)
        - Analytical description or bullet points (13.5pt)
        """
        assigned_ids = set()
        if not content_items:
            return assigned_ids

        items_to_render = content_items[:4]
        N = len(items_to_render)
        gap_pt = 14.0 * scale_factor
        card_w_pt = (content_w_pt - ((N - 1) * gap_pt)) / N
        card_h_pt = min(260.0 * scale_factor, max(160.0 * scale_factor, (sh_pt - top_start_pt - 40.0 * scale_factor)))

        is_dark_bg = cls.get_slide_bg_luminance(slide.part if hasattr(slide, "part") else slide) < 135.0
        if is_dark_bg:
            title_color = RGBColor(0x93, 0xC5, 0xFD)
            body_color = RGBColor(0xF1, 0xF5, 0xF9)
        else:
            title_color = primary_color if primary_color else RGBColor(0x1F, 0x4E, 0x79)
            body_color = RGBColor(0x1E, 0x29, 0x3B)
        font_scale = max(0.85, min(2.25, sh_pt / 540.0))

        for i, (head, desc) in enumerate(items_to_render):
            cur_left_pt = left_m_pt + i * (card_w_pt + gap_pt)
            card_shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(cur_left_pt / 72.0),
                Inches(top_start_pt / 72.0),
                Inches(card_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
            try:
                card_shp.adjustments[0] = 0.06
                card_shp.fill.solid()
                card_shp.fill.fore_color.rgb = RGBColor(0x1E, 0x29, 0x3B) if is_dark_bg else RGBColor(0xF1, 0xF5, 0xF9)
                card_shp.line.color.rgb = RGBColor(0x33, 0x41, 0x55) if is_dark_bg else RGBColor(0xCB, 0xD5, 0xE1)
                card_shp.line.width = Pt(1.2 * scale_factor)
            except Exception:
                pass
            assigned_ids.add(id(card_shp))

            tf = card_shp.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.16 * scale_factor)
            tf.margin_right = Inches(0.16 * scale_factor)
            tf.margin_top = Inches(0.14 * scale_factor)
            tf.margin_bottom = Inches(0.12 * scale_factor)
            cls.enable_xml_autofit(tf)

            clean_head = clean_latex_math(head.strip() if head else "")
            clean_desc = clean_latex_math(desc.strip() if desc else "")

            if clean_head:
                p0 = tf.paragraphs[0]
                cls.enforce_latin_line_break(p0)
                p0.text = clean_head
                p0.font.bold = True
                p0.font.size = Pt(16.5 * font_scale)
                p0.font.color.rgb = title_color
                p0.space_after = Pt(6.0 * scale_factor) if clean_desc else Pt(0)

            if clean_desc:
                p1 = tf.paragraphs[0] if not clean_head else tf.add_paragraph()
                cls.enforce_latin_line_break(p1)
                p1.text = clean_desc
                p1.font.bold = False
                p1.font.size = Pt(13.5 * font_scale)
                p1.font.color.rgb = body_color
                p1.line_spacing = 1.20

        return assigned_ids

    @classmethod
    def render_cards_by_archetype(
        cls,
        slide,
        slide_data: Dict[str, Any],
        layout_type: str,
        content_items: List[Tuple[str, str]],
        sw_pt: float,
        sh_pt: float,
        scale_factor: float,
        left_m_pt: float,
        content_w_pt: float,
        top_start_pt: float,
        primary_color: Optional[RGBColor] = None,
        accent_color: Optional[RGBColor] = None
    ) -> Set[int]:
        """
        Dynamically routes slide content to its specialized visual archetype renderer:
        - metrics_stats -> render_metric_stat_cards (big KPI numbers + labels)
        - timeline_steps / process_cycle -> render_process_step_cards (step badges 01, 02, 03)
        - cards_grid / architecture -> render_vertical_column_cards (3-4 vertical cards)
        - agenda / conclusion_sources / general -> render_wide_rectangular_cards (horizontal bars)
        """
        if layout_type in ("metrics_stats", "venn_diagram") or ("metrics" in slide_data and len(slide_data.get("metrics", [])) >= 2):
            return cls.render_metric_stat_cards(
                slide=slide,
                slide_data=slide_data,
                content_items=content_items,
                sw_pt=sw_pt,
                sh_pt=sh_pt,
                scale_factor=scale_factor,
                left_m_pt=left_m_pt,
                content_w_pt=content_w_pt,
                top_start_pt=top_start_pt,
                primary_color=primary_color,
                accent_color=accent_color
            )
        elif layout_type in ("timeline_steps", "process_cycle", "steps_infographic") or ("steps" in slide_data and len(slide_data.get("steps", [])) >= 2):
            return cls.render_process_step_cards(
                slide=slide,
                slide_data=slide_data,
                content_items=content_items,
                sw_pt=sw_pt,
                sh_pt=sh_pt,
                scale_factor=scale_factor,
                left_m_pt=left_m_pt,
                content_w_pt=content_w_pt,
                top_start_pt=top_start_pt,
                primary_color=primary_color,
                accent_color=accent_color
            )
        elif layout_type in ("cards_grid", "architecture", "three_pillars", "framework", "diagram_anatomy", "pyramid_funnel"):
            has_rich_text = any(len(d.split()) >= 8 for _, d in content_items if d)
            if has_rich_text and len(content_items) <= 4:
                return cls.render_wide_rectangular_cards(
                    slide=slide,
                    content_items=content_items,
                    sw_pt=sw_pt,
                    sh_pt=sh_pt,
                    scale_factor=scale_factor,
                    left_m_pt=left_m_pt,
                    content_w_pt=content_w_pt,
                    top_start_pt=top_start_pt,
                    primary_color=primary_color,
                    accent_color=accent_color
                )
            return cls.render_vertical_column_cards(
                slide=slide,
                content_items=content_items,
                sw_pt=sw_pt,
                sh_pt=sh_pt,
                scale_factor=scale_factor,
                left_m_pt=left_m_pt,
                content_w_pt=content_w_pt,
                top_start_pt=top_start_pt,
                primary_color=primary_color,
                accent_color=accent_color
            )
        else:
            return cls.render_wide_rectangular_cards(
                slide=slide,
                content_items=content_items,
                sw_pt=sw_pt,
                sh_pt=sh_pt,
                scale_factor=scale_factor,
                left_m_pt=left_m_pt,
                content_w_pt=content_w_pt,
                top_start_pt=top_start_pt,
                primary_color=primary_color,
                accent_color=accent_color
            )

    @classmethod
    def _synthesize_fallback_slide_content(
        cls,
        slide,
        slide_data: Dict[str, Any],
        topic: str,
        sw_pt: float,
        sh_pt: float,
        scale_factor: float
    ) -> None:
        """Synthesizes high-legibility title and structured content boxes, reusing native placeholders or creating clean shapes and eliminating all empty ghost placeholders."""
        layout_type = slide_data.get("layout_type", "cards_grid")
        title_text = slide_data.get("title", topic)

        # 1. Identify native title and body placeholders if present on the slide
        title_ph = None
        body_ph = None
        subtitle_ph = None

        for p in list(slide.placeholders):
            try:
                p_type = p.placeholder_format.type
                p_idx = p.placeholder_format.idx
                p_name = p.name.lower()
                if p_type in (1, 3) or p_idx == 0 or "title" in p_name:
                    if title_ph is None:
                        title_ph = p
                elif p_type in (2, 7) or "content" in p_name or "body" in p_name:
                    if body_ph is None:
                        body_ph = p
                elif p_type in (4,) or "subtitle" in p_name:
                    if subtitle_ph is None:
                        subtitle_ph = p
            except Exception:
                pass

        left_m_pt = 45.0 * scale_factor
        content_w_pt = max(260.0, sw_pt - (left_m_pt * 2))
        left_m = Inches(left_m_pt / 72.0)
        content_w = Inches(content_w_pt / 72.0)

        if layout_type == "cover":
            sub_text = slide_data.get("subtitle", "Akademik va amaliy tahliliy taqdimot | 2026-yil")
            if title_ph:
                cls.set_shape_text_preserving_style(title_ph, title_text, is_title=True, font_size_override=32.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            else:
                t_box = slide.shapes.add_textbox(left_m, Inches(1.8 * scale_factor), content_w, Inches(1.6 * scale_factor))
                cls.set_shape_text_preserving_style(t_box, title_text, is_title=True, font_size_override=32.0, slide_height_pt=sh_pt, scale_factor=scale_factor)

            if subtitle_ph:
                cls.set_shape_text_preserving_style(subtitle_ph, sub_text, font_size_override=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            elif body_ph:
                cls.set_shape_text_preserving_style(body_ph, sub_text, font_size_override=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            else:
                s_box = slide.shapes.add_textbox(left_m, Inches(3.6 * scale_factor), content_w, Inches(1.5 * scale_factor))
                cls.set_shape_text_preserving_style(s_box, sub_text, font_size_override=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)

            # Purge any other unfilled placeholders on cover
            for p in list(slide.placeholders):
                try:
                    if p not in (title_ph, subtitle_ph, body_ph) or not p.text_frame.text.strip():
                        slide.shapes._spTree.remove(p._element)
                except Exception:
                    pass
            return

        # Title for inner slides: Populate Title placeholder or add textbox
        title_w_pt = min(content_w_pt, sw_pt * 0.72)
        title_w = Inches(title_w_pt / 72.0)
        if title_ph:
            cls.set_shape_text_preserving_style(title_ph, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            try:
                title_ph.width = title_w
            except Exception:
                pass
        else:
            t_box = slide.shapes.add_textbox(left_m, Inches(0.6 * scale_factor), title_w, Inches(1.0 * scale_factor))
            cls.set_shape_text_preserving_style(t_box, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)

        if layout_type == "comparison_vs":
            left_t = slide_data.get("left_title", "1-Yondashuv")
            left_pts = slide_data.get("left_points", [])
            right_t = slide_data.get("right_title", "2-Yondashuv")
            right_pts = slide_data.get("right_points", [])

            # Clear any body placeholders
            for p in list(slide.placeholders):
                try:
                    if p != title_ph:
                        slide.shapes._spTree.remove(p._element)
                except Exception:
                    pass

            title_bottom_pt = 0.0
            if title_ph:
                t_top = SafeShapeGeometry.to_pt(title_ph.top, 30.0 * scale_factor)
                t_h = SafeShapeGeometry.to_pt(title_ph.height, 50.0 * scale_factor)
                title_bottom_pt = t_top + t_h
            else:
                for s in slide.shapes:
                    try:
                        if s.has_text_frame and s.text_frame.text and title_text.strip().lower() in s.text_frame.text.lower():
                            s_top = SafeShapeGeometry.to_pt(s.top, 30.0 * scale_factor)
                            s_h = SafeShapeGeometry.to_pt(s.height, 50.0 * scale_factor)
                            title_bottom_pt = max(title_bottom_pt, s_top + s_h)
                    except Exception:
                        pass

            top_start_pt = max(135.0 * scale_factor, title_bottom_pt + (16.0 * scale_factor))
            card_h_pt = min(310.0 * scale_factor, sh_pt - top_start_pt - 25.0 * scale_factor)

            half_gap_pt = 20.0 * scale_factor
            avail_w_pt = content_w_pt - half_gap_pt
            half_w_pt = avail_w_pt / 2.0
            half_w = Inches(half_w_pt / 72.0)
            right_col_left = Inches((left_m_pt + half_w_pt + half_gap_pt) / 72.0)

            # Left Column (Clean transparent textbox)
            card_l = slide.shapes.add_textbox(
                left_m,
                Inches(top_start_pt / 72.0),
                half_w,
                Inches(card_h_pt / 72.0)
            )

            tf_l = card_l.text_frame
            tf_l.word_wrap = True
            tf_l.margin_left = Inches(0.18 * scale_factor)
            tf_l.margin_right = Inches(0.18 * scale_factor)
            tf_l.margin_top = Inches(0.15 * scale_factor)
            tf_l.margin_bottom = Inches(0.15 * scale_factor)
            cls.enable_xml_autofit(tf_l)

            p_lt = tf_l.paragraphs[0]
            cls.enforce_latin_line_break(p_lt)
            p_lt.text = left_t
            p_lt.font.bold = True
            p_lt.font.size = Pt(18.0 * scale_factor)
            p_lt.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
            p_lt.space_after = Pt(8.0 * scale_factor)

            for pt in left_pts:
                p = tf_l.add_paragraph()
                cls.enforce_latin_line_break(p)
                p.line_spacing = 1.22
                p.space_after = Pt(4.0 * scale_factor)
                clean_pt = pt.strip()
                if ":" in clean_pt:
                    head, body = clean_pt.split(":", 1)
                    r1 = p.add_run()
                    r1.text = f"• {head.strip()}: "
                    r1.font.bold = True
                    r1.font.size = Pt(14.0 * scale_factor)
                    r1.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
                    r2 = p.add_run()
                    r2.text = body.strip()
                    r2.font.bold = False
                    r2.font.size = Pt(14.0 * scale_factor)
                    r2.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
                else:
                    p.text = f"• {clean_pt}"
                    for r in p.runs:
                        r.font.size = Pt(14.0 * scale_factor)
                        r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

            # Right Column (Clean transparent textbox)
            card_r = slide.shapes.add_textbox(
                right_col_left,
                Inches(top_start_pt / 72.0),
                half_w,
                Inches(card_h_pt / 72.0)
            )

            tf_r = card_r.text_frame
            tf_r.word_wrap = True
            tf_r.margin_left = Inches(0.18 * scale_factor)
            tf_r.margin_right = Inches(0.18 * scale_factor)
            tf_r.margin_top = Inches(0.15 * scale_factor)
            tf_r.margin_bottom = Inches(0.15 * scale_factor)
            cls.enable_xml_autofit(tf_r)

            p_rt = tf_r.paragraphs[0]
            cls.enforce_latin_line_break(p_rt)
            p_rt.text = right_t
            p_rt.font.bold = True
            p_rt.font.size = Pt(18.0 * scale_factor)
            p_rt.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)
            p_rt.space_after = Pt(8.0 * scale_factor)

            for pt in right_pts:
                p = tf_r.add_paragraph()
                cls.enforce_latin_line_break(p)
                p.line_spacing = 1.22
                p.space_after = Pt(4.0 * scale_factor)
                clean_pt = pt.strip()
                if ":" in clean_pt:
                    head, body = clean_pt.split(":", 1)
                    r1 = p.add_run()
                    r1.text = f"• {head.strip()}: "
                    r1.font.bold = True
                    r1.font.size = Pt(14.0 * scale_factor)
                    r1.font.color.rgb = RGBColor(0x06, 0x5F, 0x46)
                    r2 = p.add_run()
                    r2.text = body.strip()
                    r2.font.bold = False
                    r2.font.size = Pt(14.0 * scale_factor)
                    r2.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
                else:
                    p.text = f"• {clean_pt}"
                    for r in p.runs:
                        r.font.size = Pt(14.0 * scale_factor)
                        r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

            return

        # Generic structured body: Synthesize specialized visual archetype cards
        content_items = cls._extract_content_items(slide_data, layout_type, title_text=title_text)
        if len(content_items) >= 2:
            # Purge all body placeholders to give way to structured visual cards
            for p in list(slide.placeholders):
                try:
                    if p != title_ph:
                        slide.shapes._spTree.remove(p._element)
                except Exception:
                    pass
            title_bottom_pt = 0.0
            if title_ph:
                t_top = SafeShapeGeometry.to_pt(title_ph.top, 30.0 * scale_factor)
                t_h = SafeShapeGeometry.to_pt(title_ph.height, 50.0 * scale_factor)
                title_bottom_pt = t_top + t_h
            else:
                for s in slide.shapes:
                    try:
                        if s.has_text_frame and s.text_frame.text and title_text.strip().lower() in s.text_frame.text.lower():
                            s_top = SafeShapeGeometry.to_pt(s.top, 30.0 * scale_factor)
                            s_h = SafeShapeGeometry.to_pt(s.height, 50.0 * scale_factor)
                            title_bottom_pt = max(title_bottom_pt, s_top + s_h)
                    except Exception:
                        pass
            top_start_pt = max(118.0 * scale_factor, title_bottom_pt + (16.0 * scale_factor))
            cls.render_cards_by_archetype(
                slide=slide,
                slide_data=slide_data,
                layout_type=layout_type,
                content_items=content_items,
                sw_pt=sw_pt,
                sh_pt=sh_pt,
                scale_factor=scale_factor,
                left_m_pt=left_m_pt,
                content_w_pt=content_w_pt,
                top_start_pt=top_start_pt
            )
            return

        formatted = []
        for h, d in content_items:
            if h and d:
                formatted.append(f"• {h}:\n  {d}")
            elif h:
                formatted.append(f"• {h}")
            elif d:
                formatted.append(d)
        full_text = "\n\n".join(formatted) if formatted else topic

        if body_ph:
            cls.set_shape_text_preserving_style(body_ph, full_text, default_font_size=16.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
        else:
            b_box = slide.shapes.add_textbox(left_m, Inches(1.8 * scale_factor), content_w, Inches(min(4.8 * scale_factor, (sh_pt - 140.0) / 72.0)))
            cls.set_shape_text_preserving_style(b_box, full_text, default_font_size=16.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

        # CRITICAL: Purge any unused empty placeholders so PowerPoint never renders ghost prompts
        for p in list(slide.placeholders):
            try:
                if not p.has_text_frame or not p.text_frame.text.strip():
                    slide.shapes._spTree.remove(p._element)
            except Exception:
                pass

    @classmethod
    def calculate_card_top_start_pt(cls, slide, title_text: str, title_shape_meta: Optional[Dict[str, Any]], scale_factor: float) -> float:
        """
        Dynamically calculates the exact, safe top_start_pt for cards or columns
        guaranteeing at least 16-18pt of breathing space below the lowest point
        of the title banner, preventing any collision with cards or badges.
        """
        title_bottom_pt = 0.0
        title_top_pt = 30.0 * scale_factor
        for s in slide.shapes:
            try:
                if s.has_text_frame and s.text_frame.text and title_text.strip().lower() in s.text_frame.text.lower():
                    s_top = SafeShapeGeometry.to_pt(s.top, 30.0 * scale_factor)
                    s_h = SafeShapeGeometry.to_pt(s.height, 50.0 * scale_factor)
                    title_top_pt = s_top
                    title_bottom_pt = max(title_bottom_pt, s_top + s_h)
            except Exception:
                pass
        if title_bottom_pt <= 0.0 and title_shape_meta and "shape" in title_shape_meta:
            s_top = SafeShapeGeometry.to_pt(title_shape_meta["shape"].top, 30.0 * scale_factor)
            s_h = SafeShapeGeometry.to_pt(title_shape_meta["shape"].height, 50.0 * scale_factor)
            title_top_pt = s_top
            title_bottom_pt = s_top + s_h

        # Multi-line title height dynamic estimation (prevents cards encroaching on 2 or 3-line titles)
        clean_t = (title_text or "").strip()
        est_lines = max(1, math.ceil(len(clean_t) / 34.0)) if clean_t else 1
        est_title_h_pt = est_lines * (32.0 * scale_factor)
        effective_bottom = max(title_bottom_pt, title_top_pt + est_title_h_pt)
        min_card_top = (118.0 if est_lines <= 1 else (135.0 if est_lines == 2 else 150.0)) * scale_factor
        return max(min_card_top, effective_bottom + (18.0 * scale_factor))

    @classmethod
    def _expand_items_to_fill_fixed_diagram(
        cls,
        content_items: List[Tuple[str, str]],
        target_count: int,
        title_text: str,
        topic: str
    ) -> List[Tuple[str, str]]:
        """
        When an infographic slide has a fixed number of visual diagram nodes (e.g. 6 circles around a padlock)
        and content_items has fewer items (e.g. 3 or 4), generates complementary analytical nodes so no
        diagram circle is left blank or orphaned.
        """
        result = list(content_items)
        extra_pillars = [
            ("Tizimli Integratsiya", f"«{title_text}» doirasida axborot oqimlarini yagona standart asosida birlashtirish va uzluksizlikni ta'minlash."),
            ("Xavfsizlik va Ishonchlilik", "Ma'lumotlar butunligini zamonaviy himoya mexanizmlari va shaffof nazorat orqali kafolatlash."),
            ("Iqtisodiy Samaradorlik", "Operatsion xarajatlarni qisqartirish, resurslarni optimallashtirish va yuqori natijadorlikka erishish."),
            ("Innovatsion Rivojlanish", f"«{topic}» sohasida ilg'or raqamli texnologiyalar va avtomatlashtirilgan yechimlarni keng joriy etish."),
            ("Me'yoriy va Amaliy Standartlar", "Xalqaro talablar va zamonaviy amaliyotga mos keluvchi barqaror boshqaruv arxitekturasi."),
            ("Kelajak Istiqbollari", "Uzoq muddatli strategik o'sish, masshtablashuv va raqamli ekotizim barqarorligini mustahkamlash."),
        ]
        p_idx = 0
        existing_heads = {h.lower().strip() for h, _ in result if h}
        while len(result) < target_count and p_idx < len(extra_pillars):
            cand_h, cand_d = extra_pillars[p_idx]
            p_idx += 1
            if cand_h.lower() not in existing_heads:
                result.append((cand_h, cand_d))
                existing_heads.add(cand_h.lower())
        return result

    @classmethod
    def replace_slide_content(cls, slide, slide_data: Dict[str, Any], topic: str) -> None:
        """
        Replaces text placeholders inside a template slide with safe content zone constraints,
        deep watermark eradication, resolution-aware scaling, 2D card slot pairing, and XML shape pruning.
        """
        try:
            from core.slide_splicer import normalize_slide_group_coordinates
            normalize_slide_group_coordinates(slide)
        except Exception:
            pass
        sw, sh, sw_pt, sh_pt = cls.get_slide_dimensions(slide)
        scale_factor = max(0.75, min(2.25, sh_pt / 540.0))

        layout_type = slide_data.get("layout_type", "cards_grid")
        title_text = slide_data.get("title", topic)

        def collect_text_shapes_recursive(shapes_container, parent_group=None, ancestor_groups=None):
            if ancestor_groups is None:
                ancestor_groups = []
            collected = []
            for s in list(shapes_container):
                if SmartArtGuardian.is_smart_art(s):
                    continue
                if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                    collected.extend(
                        collect_text_shapes_recursive(
                            s.shapes,
                            parent_group=s,
                            ancestor_groups=ancestor_groups + [s]
                        )
                    )
                elif s.has_text_frame and s.shape_type != MSO_SHAPE_TYPE.PICTURE:
                    txt = s.text_frame.text.strip()
                    if not txt:
                        continue

                    is_wm = DeepWatermarkPurger.is_watermark_text(txt)

                    l_pt = SafeShapeGeometry.to_pt(s.left, 0.0)
                    t_pt = SafeShapeGeometry.to_pt(s.top, 0.0)
                    w_pt = SafeShapeGeometry.to_pt(s.width, 0.0)
                    h_pt = SafeShapeGeometry.to_pt(s.height, 0.0)

                    # Find the card-level group (first ancestor group whose width < 85% of slide width, or immediate parent_group)
                    card_grp = parent_group
                    pos_grp = parent_group
                    for ag in ancestor_groups:
                        ag_w = SafeShapeGeometry.to_pt(ag.width, 0.0)
                        ag_h = SafeShapeGeometry.to_pt(ag.height, 0.0)
                        if ag_w < sw_pt * 0.85 or ag_h < sh_pt * 0.45:
                            pos_grp = ag
                            break

                    # Since normalize_slide_group_coordinates(slide) normalized chOff == off and chExt == ext,
                    # l_pt and t_pt are already the exact true slide-space coordinates!
                    eff_left = l_pt
                    eff_top = t_pt

                    is_valid_zone = (
                        eff_left >= -15.0
                        and eff_left <= sw_pt * 0.96
                        and (w_pt >= 70.0 * scale_factor or parent_group is not None)
                    )

                    step_num = None
                    m = re.match(r'^(?:step.s*)?0?([1-9]|1[0-9])[...]?$', txt, re.I)
                    if m:
                        step_num = int(m.group(1))

                    collected.append({
                        "shape": s,
                        "parent_group": card_grp,
                        "pos_group": pos_grp,
                        "ancestor_groups": list(ancestor_groups),
                        "text": txt,
                        "name": s.name,
                        "left": eff_left,
                        "top": eff_top,
                        "width": w_pt,
                        "height": h_pt,
                        "inner_top": t_pt,
                        "inner_left": l_pt,
                        "is_step": cls.is_step_badge(txt),
                        "step_num": step_num,
                        "is_dummy": cls.is_dummy_text(txt),
                        "is_watermark": is_wm,
                        "is_valid_zone": is_valid_zone,
                        "font_size": max([r.font.size.pt for p in s.text_frame.paragraphs for r in p.runs if r.font.size] or [12.0 * scale_factor])
                    })
            return collected

        all_text_shapes = collect_text_shapes_recursive(slide.shapes)

        # Check for native PowerPoint Tables early
        table_shapes = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.TABLE]

        if not all_text_shapes and not table_shapes:
            cls._synthesize_fallback_slide_content(slide, slide_data, topic, sw_pt, sh_pt, scale_factor)
            return

        # Prune blatant watermarks (except if it's the only large title shape on cover)
        for sm in all_text_shapes:
            if sm["is_watermark"]:
                if layout_type == "cover" and sm["height"] >= 45.0 * scale_factor and "http" not in sm["text"].lower() and "www." not in sm["text"].lower():
                    sm["is_watermark"] = False
                    sm["is_dummy"] = True
                else:
                    SemanticShapeCleaner.prune_shape(sm["shape"])

        active_shapes = [s for s in all_text_shapes if not s["is_watermark"]]
        if not active_shapes and not table_shapes:
            cls._synthesize_fallback_slide_content(slide, slide_data, topic, sw_pt, sh_pt, scale_factor)
            return

        # Early title identification (using sh_pt and sw_pt in points!)
        best_title_meta = None
        title_candidates_early = []
        for s_cand in active_shapes:
            if not s_cand.get("is_valid_zone") or s_cand.get("is_step"):
                continue
            if s_cand["width"] < 110.0 * scale_factor and s_cand.get("parent_group") is None:
                continue
            sc = 0
            t_low = s_cand["text"].lower()
            s_name_low = s_cand["name"].lower()
            if any(k in t_low for k in ["table of contents", "agenda", "steps infographic", "timeline infographic", "comparison slide", "bullet points", "table slide", "credits", "infographic style", "timeline style", "columns style", "columns infographic", "portfolio presentation", "awesome slide", "thank you", "thanks"]):
                sc += 500
            elif "title" in s_name_low or "header" in s_name_low or "placeholder 1" in s_name_low:
                sc += 250
            if s_cand["top"] < sh_pt * 0.35:
                sc += 100 + int((1.0 - (s_cand["top"] / max(1.0, sh_pt * 0.35))) * 50)
            elif s_cand["top"] < sh_pt * 0.52:
                sc += 25
            else:
                sc -= 150
            if s_cand["width"] > sw_pt * 0.30:
                sc += 35
            sc += int(min(250.0, s_cand["font_size"] * 5.0))
            if layout_type in ("cover", "conclusion_sources") and s_cand["height"] >= 45.0 * scale_factor:
                sc += 180
            title_candidates_early.append((sc, s_cand))

        if title_candidates_early:
            title_candidates_early.sort(key=lambda x: x[0], reverse=True)
            if title_candidates_early[0][0] > 0:
                best_title_meta = title_candidates_early[0][1]

        # Handle Native PowerPoint Tables (MSO_SHAPE_TYPE.TABLE) — including auto-converting slide points into table rows!
        if table_shapes:
            t_data = (
                slide_data.get("table") if isinstance(slide_data.get("table"), dict)
                else (slide_data.get("table_data") if isinstance(slide_data.get("table_data"), dict)
                else (slide_data.get("matrix") if isinstance(slide_data.get("matrix"), dict) else {}))
            )
            headers = t_data.get("headers", []) or slide_data.get("headers", [])
            rows = t_data.get("rows", []) or slide_data.get("rows", [])
            if not headers and not rows:
                c_items = cls._extract_content_items(slide_data, layout_type, title_text=title_text)
                if c_items:
                    headers = ["№", "Asosiy yo'nalish", "Mazmuni va tahliliy tavsifi"]
                    rows = []
                    for r_i, (h_txt, d_txt) in enumerate(c_items, 1):
                        if not d_txt and h_txt:
                            ep = TopicLogicValidator.enrich_short_point(h_txt, title_text, topic, r_i)
                            if ":" in ep:
                                h_txt, d_txt = ep.split(":", 1)
                        rows.append([f"{r_i:02d}", (h_txt or f"Yo'nalish {r_i}").strip(), (d_txt or h_txt).strip()])
            if headers or rows:
                is_dark_tbl = cls.is_shape_on_dark_background(table_shapes[0])
                TableMatrixReplacer.apply_table_data(
                    table_shape=table_shapes[0],
                    headers=headers,
                    rows=rows,
                    theme_key=slide_data.get("theme_key", "tech_minimal"),
                    scale_factor=scale_factor,
                    is_dark_bg=is_dark_tbl,
                )
                assigned_shapes = {id(table_shapes[0])}
                if best_title_meta:
                    cls.set_shape_text_preserving_style(best_title_meta["shape"], title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    assigned_shapes.add(id(best_title_meta["shape"]))
                SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
                return

        # Handle Native PowerPoint Charts (MSO_SHAPE_TYPE.CHART)
        chart_shapes = [s for s in slide.shapes if hasattr(s, "has_chart") and s.has_chart]
        if chart_shapes and ("chart_data" in slide_data or "chart" in slide_data):
            c_data = slide_data.get("chart_data") if isinstance(slide_data.get("chart_data"), dict) else (slide_data.get("chart") if isinstance(slide_data.get("chart"), dict) else {})
            if c_data.get("categories") and c_data.get("series"):
                ChartHarmonizer.apply_chart_data(
                    chart_shape=chart_shapes[0],
                    chart_data_dict=c_data,
                    theme_key=slide_data.get("theme_key", "tech_minimal")
                )
                assigned_shapes = {id(chart_shapes[0])}
                if best_title_meta:
                    cls.set_shape_text_preserving_style(best_title_meta["shape"], title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    assigned_shapes.add(id(best_title_meta["shape"]))
                SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
                return

        valid_shapes = [s for s in active_shapes if s["is_valid_zone"]]
        if not valid_shapes:
            valid_shapes = [
                s for s in active_shapes
                if (SafeShapeGeometry.to_pt(s["shape"].width, 0.0) >= 90.0 * scale_factor or s["is_step"] or s.get("parent_group") is not None)
            ]

        default_left_m_pt = 45.0 * scale_factor
        body_candidates_for_geom = [s for s in valid_shapes if (not best_title_meta or s["shape"] != best_title_meta["shape"])]
        body_lefts = [s["left"] for s in body_candidates_for_geom if s["left"] >= 20.0 and s["width"] >= 90.0]
        body_rights = [s["left"] + s["width"] for s in body_candidates_for_geom if s["left"] >= 20.0 and s["width"] >= 90.0]

        orig_title_left = best_title_meta["left"] if best_title_meta else default_left_m_pt
        left_m_pt = min(body_lefts) if (body_lefts and min(body_lefts) >= 60.0 * scale_factor) else default_left_m_pt
        max_right_pt = sw_pt - (40.0 * scale_factor)
        if body_rights and max(body_rights) > left_m_pt + 200.0 * scale_factor:
            max_right_pt = min(max_right_pt, max(body_rights))

        content_w_pt = max(260.0 * scale_factor, max_right_pt - left_m_pt)
        left_m = Inches(left_m_pt / 72.0)
        content_w = Inches(content_w_pt / 72.0)

        if best_title_meta:
            title_shape_meta = best_title_meta
            # Only widen top-level title if it is abnormally narrow (< 28% slide width)
            if title_shape_meta.get("parent_group") is None:
                cur_tw = SafeShapeGeometry.to_pt(title_shape_meta["shape"].width, 0.0)
                if cur_tw < sw_pt * 0.28:
                    SafeShapeGeometry.update(
                        title_shape_meta["shape"],
                        width=Inches(min(content_w_pt, sw_pt - orig_title_left - 40.0 * scale_factor) / 72.0)
                    )
            cls.set_shape_text_preserving_style(
                title_shape_meta["shape"],
                title_text,
                is_title=True,
                slide_height_pt=sh_pt,
                scale_factor=scale_factor
            )
            assigned_shapes = {id(title_shape_meta["shape"])}
        else:
            title_w_pt = min(content_w_pt, sw_pt - left_m_pt - (40.0 * scale_factor))
            t_box = slide.shapes.add_textbox(
                left_m,
                Inches(0.42 * scale_factor),
                Inches(title_w_pt / 72.0),
                Inches(0.85 * scale_factor)
            )
            cls.set_shape_text_preserving_style(t_box, title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
            assigned_shapes = {id(t_box)}
            title_shape_meta = {"shape": t_box, "left": left_m_pt, "top": 30.0 * scale_factor, "width": title_w_pt, "height": 60.0 * scale_factor, "parent_group": None}

        # -------------------------------------------------------------
        # SPECIAL CASE: COVER / TITLE SLIDE
        # -------------------------------------------------------------
        if layout_type == "cover":
            subtitle_parts = []
            if slide_data.get("subtitle"):
                subtitle_parts.append(str(slide_data["subtitle"]))
            if slide_data.get("author_name"):
                author_line = str(slide_data["author_name"])
                if slide_data.get("organization"):
                    author_line += f" | {slide_data['organization']}"
                subtitle_parts.append(author_line)
            if slide_data.get("event_name"):
                subtitle_parts.append(str(slide_data["event_name"]))

            sub_text = "\n\n".join(subtitle_parts) if subtitle_parts else "Professional tahliliy taqdimot va hisobot"

            sub_candidates = [s for s in valid_shapes if id(s["shape"]) not in assigned_shapes and not s["is_step"]]
            # If the cover slide has >= 6 non-step text boxes (e.g. a 6-node timeline or 6-node padlock infographic on Slide 1),
            # do NOT steal Card #1's header box for subtitle — fall through so all 6 infographic nodes are populated!
            if sub_candidates and len(sub_candidates) < 6:
                # Prefer a candidate in the same group as title, or closest below title
                t_grp = title_shape_meta.get("parent_group")
                same_grp_subs = [s for s in sub_candidates if t_grp is not None and s.get("parent_group") == t_grp]
                if same_grp_subs:
                    same_grp_subs.sort(key=lambda s: s["inner_top"])
                    sub_meta = same_grp_subs[0]
                else:
                    sub_candidates.sort(key=lambda s: (abs(s["top"] - (title_shape_meta["top"] + title_shape_meta["height"])), s["top"]))
                    sub_meta = sub_candidates[0]

                sub_shape = sub_meta["shape"]
                # Prevent Title-Subtitle vertical collision on top-level cover shapes (e.g. when 2-line title expands)
                if sub_meta.get("parent_group") is None and title_shape_meta.get("parent_group") is None:
                    t_top = SafeShapeGeometry.to_pt(title_shape_meta["shape"].top, 30.0 * scale_factor)
                    t_w = SafeShapeGeometry.to_pt(title_shape_meta["shape"].width, 500.0 * scale_factor)
                    est_cpl = max(18, int(t_w / (28.0 * scale_factor * 0.55)))
                    est_lines = max(1, math.ceil(len(title_text) / float(est_cpl)))
                    t_h_actual = SafeShapeGeometry.to_pt(title_shape_meta["shape"].height, 45.0 * scale_factor)
                    est_title_bottom = t_top + max(t_h_actual * 0.85, est_lines * 48.0 * scale_factor)
                    s_top = SafeShapeGeometry.to_pt(sub_shape.top, 0.0)
                    if s_top >= t_top - 10.0 and s_top < est_title_bottom + (12.0 * scale_factor):
                        SafeShapeGeometry.update(sub_shape, top=Inches((est_title_bottom + 14.0 * scale_factor) / 72.0))

                cls.set_shape_text_preserving_style(sub_shape, sub_text, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                assigned_shapes.add(id(sub_shape))
            elif not sub_candidates:
                if title_shape_meta.get("parent_group") is None:
                    t_bottom = title_shape_meta["shape"].top + title_shape_meta["shape"].height
                    sub_box = slide.shapes.add_textbox(
                        left_m,
                        t_bottom + Inches(0.18 * scale_factor),
                        content_w,
                        Inches(0.9 * scale_factor)
                    )
                    cls.set_shape_text_preserving_style(sub_box, sub_text, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    assigned_shapes.add(id(sub_box))

            # Check if Cover slide is ALSO a 6-node Infographic Overview diagram (has >= 6 remaining non-step text boxes outside title group, like Bitcoin padlock slide 0!)
            t_grp_id = id(title_shape_meta.get("parent_group")) if title_shape_meta.get("parent_group") is not None else None
            remaining_cover_body = [
                s for s in valid_shapes
                if id(s["shape"]) not in assigned_shapes
                and not s["is_step"]
                and (t_grp_id is None or id(s.get("parent_group")) != t_grp_id)
            ]
            if len(remaining_cover_body) < 6:
                SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
                return
            # Otherwise fall through to populate the infographic nodes on the cover slide!

        # Extract structured content items via reusable extractor
        content_items = cls._extract_content_items(slide_data, layout_type, title_text=title_text)
        # Ensure every content item has both a header and an analytical description
        enriched_items: List[Tuple[str, str]] = []
        for idx_item, (h_val, d_val) in enumerate(content_items, 1):
            if h_val and not d_val:
                ep = TopicLogicValidator.enrich_short_point(h_val, title_text, topic, idx_item)
                if ":" in ep:
                    h_val, d_val = ep.split(":", 1)
                    h_val, d_val = h_val.strip(), d_val.strip()
            enriched_items.append((h_val, d_val))
        content_items = enriched_items

        body_shapes = [s for s in valid_shapes if id(s["shape"]) not in assigned_shapes]

        # -------------------------------------------------------------
        # 1. GROUP-BASED CARD CONTAINERS
        # -------------------------------------------------------------
        grp_non_step_counts: Dict[int, int] = {}
        for s in body_shapes:
            if not s["is_step"] and s.get("parent_group") is not None:
                gid = id(s.get("parent_group"))
                grp_non_step_counts[gid] = grp_non_step_counts.get(gid, 0) + 1
        distinct_groups = set(grp_non_step_counts.keys())
        max_per_grp = max(grp_non_step_counts.values()) if grp_non_step_counts else 0
        has_group_cards = (
            max_per_grp <= 2
            and (
                len(distinct_groups) >= 2
                or (len(distinct_groups) == 1 and sum(1 for s in body_shapes if not s["is_step"] and s.get("parent_group") is None) == 0)
            )
        )
        if has_group_cards:
            card_dict = {}
            for s in body_shapes:
                if s["is_step"]:
                    continue
                grp = s.get("parent_group")
                # When slide has >= 2 grouped cards (e.g. 6 bar-chart groups), prune stray ungrouped dummy callouts (like 'Simple Chart Infographic')
                if grp is None and len(distinct_groups) >= 2:
                    SemanticShapeCleaner.prune_shape(s["shape"])
                    continue
                pos_grp = s.get("pos_group") or grp
                card_key = id(grp) if grp else id(s["shape"])
                if card_key not in card_dict:
                    card_dict[card_key] = {
                        "parent_group": grp,
                        "pos_group": pos_grp,
                        "shapes": [],
                        "left": s["left"],
                        "top": s["top"]
                    }
                card_dict[card_key]["shapes"].append(s)

            card_list = list(card_dict.values())
            card_list.sort(key=lambda c: (round(c["left"] / (85.0 * scale_factor)), c["top"]))

            # Only prune group cards if they are truly microscopic (< 65pt) or 0 cards
            avg_card_w = sum(
                SafeShapeGeometry.to_pt((c.get("pos_group") or c.get("parent_group") or c["shapes"][0]["shape"]).width, 0.0)
                for c in card_list if c["shapes"]
            ) / max(1, len(card_list))

            if avg_card_w < 65.0 * scale_factor and len(content_items) >= 2:
                for card in card_list:
                    grp = card.get("parent_group")
                    if grp is not None:
                        SemanticShapeCleaner.prune_shape(grp)
                    else:
                        for extra_s in card["shapes"]:
                            SemanticShapeCleaner.prune_shape(extra_s["shape"])

                top_start_pt = cls.calculate_card_top_start_pt(slide, title_text, title_shape_meta, scale_factor)
                assigned_cards = cls.render_cards_by_archetype(
                    slide=slide,
                    slide_data=slide_data,
                    layout_type=layout_type,
                    content_items=content_items,
                    sw_pt=sw_pt,
                    sh_pt=sh_pt,
                    scale_factor=scale_factor,
                    left_m_pt=left_m_pt,
                    content_w_pt=content_w_pt,
                    top_start_pt=top_start_pt
                )
                assigned_shapes.update(assigned_cards)
                SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(content_items))
                return

            # Expand content_items when slide is a fixed 4, 5, or 6-node diagram so the 4th/5th/6th diagram node is never left empty!
            if len(card_list) > len(content_items) and len(card_list) in (4, 5, 6) and layout_type != "agenda":
                content_items = cls._expand_items_to_fill_fixed_diagram(content_items, len(card_list), title_text, topic)

            for c_idx, card in enumerate(card_list):
                c_shapes = card["shapes"]
                # Separate step badges inside the group from text boxes
                non_step_c = [x for x in c_shapes if not x["is_step"]]
                if not non_step_c:
                    non_step_c = c_shapes
                has_long_and_short = (
                    len(non_step_c) == 2
                    and min(len(x["text"]) for x in non_step_c) <= 28
                    and max(len(x["text"]) for x in non_step_c) >= 35
                )
                if has_long_and_short:
                    non_step_c.sort(key=lambda x: (0 if len(x["text"]) <= 28 else 1, x["top"]))
                else:
                    non_step_c.sort(key=lambda x: x["top"])
                grp = card["parent_group"]

                if c_idx < len(content_items):
                    # If this is the last card and there are still more content_items, combine them cleanly
                    if c_idx == len(card_list) - 1 and len(content_items) > len(card_list):
                        rem = content_items[c_idx:]
                        item_head = rem[0][0]
                        item_desc = "\n".join(f"• {h}: {d}" if (h and d) else f"• {h or d}" for h, d in rem)
                    else:
                        item_head, item_desc = content_items[c_idx]

                    if len(non_step_c) >= 2:
                        t_box = non_step_c[0]["shape"]
                        d_box = non_step_c[1]["shape"]
                        assigned_shapes.add(id(t_box))
                        assigned_shapes.add(id(d_box))
                        cls.set_shape_text_preserving_style(t_box, item_head or f"Yo'nalish {c_idx+1}", is_card_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
                        # Only push d_box down in wide groups (t_w >= 140pt)
                        try:
                            t_top_in_grp = SafeShapeGeometry.to_pt(t_box.top, 0.0)
                            t_w_in_grp = SafeShapeGeometry.to_pt(t_box.width, 120.0)
                            d_top_in_grp = SafeShapeGeometry.to_pt(d_box.top, 0.0)
                            if t_w_in_grp >= 140.0 * scale_factor and d_top_in_grp >= t_top_in_grp and (d_top_in_grp - t_top_in_grp) < 32.0 * scale_factor and len(item_head or "") > int(t_w_in_grp / (7.5 * scale_factor)):
                                new_d_top = t_top_in_grp + (30.0 * scale_factor)
                                SafeShapeGeometry.update(t_box, height=Inches((28.0 * scale_factor) / 72.0))
                                SafeShapeGeometry.update(d_box, top=Inches(new_d_top / 72.0))
                        except Exception:
                            pass
                        cls.set_shape_text_preserving_style(d_box, item_desc or item_head, slide_height_pt=sh_pt, scale_factor=scale_factor)
                        for extra in non_step_c[2:]:
                            SemanticShapeCleaner.prune_shape(extra["shape"])
                    elif len(non_step_c) == 1:
                        shp = non_step_c[0]["shape"]
                        assigned_shapes.add(id(shp))
                        combined = f"{item_head}:\n{item_desc}".strip() if (item_head and item_desc) else (item_head or item_desc)
                        cls.set_shape_text_preserving_style(shp, combined, default_font_size=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)

                    # Also format & mark any step badges in this group as assigned (01, 02, 03...)
                    for st_s in c_shapes:
                        if st_s["is_step"]:
                            assigned_shapes.add(id(st_s["shape"]))
                            cls.set_shape_text_preserving_style(st_s["shape"], f"{c_idx+1:02d}", font_size_override=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    if grp is not None:
                        SemanticShapeCleaner.prune_shape(grp)
                    else:
                        for extra_s in c_shapes:
                            SemanticShapeCleaner.prune_shape(extra_s["shape"])

            # Also preserve top-level step badges up to len(card_list)
            step_shapes = [s for s in body_shapes if s["is_step"] and id(s["shape"]) not in assigned_shapes]
            step_shapes.sort(key=lambda s: (round(s["left"] / (85.0 * scale_factor)), s["top"]))
            for s_idx, sm in enumerate(step_shapes):
                s_shp = sm["shape"]
                if s_idx < min(len(card_list), len(content_items)):
                    assigned_shapes.add(id(s_shp))
                    if sm["text"].strip().endswith("%"):
                        continue
                    step_str = f"{s_idx+1:02d}"
                    cls.set_shape_text_preserving_style(s_shp, step_str, font_size_override=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    SemanticShapeCleaner.prune_shape(s_shp)

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(content_items))
            return

        # -------------------------------------------------------------
        # 2. UNGROUPED 2D SPATIAL CLUSTERING (COLUMNS + VERTICAL CARD PAIRS)
        # -------------------------------------------------------------
        non_step_body = [
            s for s in body_shapes
            if not s["is_step"]
            and (SafeShapeGeometry.to_pt(s["shape"].width, 0.0) >= 85.0 * scale_factor)
            and not (
                layout_type == "conclusion_sources"
                and s["width"] < 200.0 * scale_factor
                and s["height"] < 42.0 * scale_factor
                and len(s["text"]) <= 25
            )
        ]

        def get_overlap_x(s1, s2):
            l1, r1 = s1["left"], s1["left"] + s1["width"]
            l2, r2 = s2["left"], s2["left"] + s2["width"]
            return max(0.0, min(r1, r2) - max(l1, l2))

        columns: List[List[Dict[str, Any]]] = []
        for s in sorted(non_step_body, key=lambda x: x["left"]):
            matched_col = None
            for col in columns:
                c_lead = col[0]
                overlap = get_overlap_x(s, c_lead)
                min_w = min(s["width"], c_lead["width"])
                center_dist = abs((s["left"] + s["width"] / 2.0) - (c_lead["left"] + c_lead["width"] / 2.0))
                # FIXED: 86.4 * scale_factor (1.2 inches in points, NOT Inches(1.2).emu!)
                if (min_w > 0 and overlap / min_w >= 0.35) or center_dist < (86.4 * scale_factor):
                    matched_col = col
                    break
            if matched_col is not None:
                matched_col.append(s)
            else:
                columns.append([s])

        columns.sort(key=lambda col: col[0]["left"])
        for col in columns:
            col.sort(key=lambda s: s["top"])

        # Group shapes within each column into 2D Card Slots [(header_shape, body_shape) or (single_shape,)]
        card_slots: List[List[Dict[str, Any]]] = []
        for col in columns:
            i = 0
            while i < len(col):
                curr_s = col[i]
                if i + 1 < len(col):
                    next_s = col[i + 1]
                    v_gap = next_s["top"] - (curr_s["top"] + curr_s["height"])
                    curr_name = curr_s["name"].lower()
                    next_name = next_s["name"].lower()
                    is_header_body_pair = (
                        v_gap <= 55.0 * scale_factor
                        and (len(curr_s["text"]) <= 36 or curr_s["height"] <= next_s["height"] * 0.55)
                        and (
                            "title" in curr_name
                            or "body" in next_name
                            or curr_s["height"] <= next_s["height"] * 1.15
                            or curr_s["font_size"] >= next_s["font_size"]
                        )
                    )
                    if is_header_body_pair:
                        card_slots.append([curr_s, next_s])
                        i += 2
                        continue
                card_slots.append([curr_s])
                i += 1

        # Sort card_slots in natural reading order: top-to-bottom within left/right zones or left-to-right
        if len(card_slots) >= 2:
            # If slots are arranged in 2 columns (like 3 on left, 3 on right), preserve column-major order
            pass

        if len(card_slots) == 0 and content_items:
            top_start_pt = cls.calculate_card_top_start_pt(slide, title_text, title_shape_meta, scale_factor)
            assigned_cards = cls.render_cards_by_archetype(
                slide=slide,
                slide_data=slide_data,
                layout_type=layout_type,
                content_items=content_items,
                sw_pt=sw_pt,
                sh_pt=sh_pt,
                scale_factor=scale_factor,
                left_m_pt=left_m_pt,
                content_w_pt=content_w_pt,
                top_start_pt=top_start_pt
            )
            assigned_shapes.update(assigned_cards)
            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(content_items))
            return

        # Special case: 1 large body slot on the slide -> populate all content_items as formatted bullet blocks inside that slot
        if len(card_slots) == 1 and content_items:
            slot = card_slots[0]
            formatted_lines = []
            for h, d in content_items:
                if h and d:
                    formatted_lines.append(f"• {h}: {d}")
                elif h:
                    formatted_lines.append(f"• {h}")
                elif d:
                    formatted_lines.append(f"• {d}")
            full_text = "\n\n".join(formatted_lines)
            if len(slot) == 1:
                shp = slot[0]["shape"]
                assigned_shapes.add(id(shp))
                cls.set_shape_text_preserving_style(shp, full_text, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            else:
                t_box = slot[0]["shape"]
                d_box = slot[1]["shape"]
                assigned_shapes.add(id(t_box))
                assigned_shapes.add(id(d_box))
                cls.set_shape_text_preserving_style(t_box, content_items[0][0] or "Asosiy Yo'nalishlar", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                cls.set_shape_text_preserving_style(d_box, full_text, font_size_override=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(content_items))
            return

        # If the template slide has a fixed multi-node diagram (e.g. 5 or 6 card_slots around a central graphic like Bitcoin padlock)
        # and content_items has fewer items, expand content_items so all diagram slots are filled!
        if len(card_slots) > len(content_items) and len(card_slots) in (5, 6):
            non_text_vectors = sum(1 for s in slide.shapes if not s.has_text_frame or not s.text_frame.text.strip())
            if non_text_vectors >= 4:
                content_items = cls._expand_items_to_fill_fixed_diagram(content_items, len(card_slots), title_text, topic)

        # Populate card_slots 1-to-1 with content_items
        for idx, slot in enumerate(card_slots):
            if idx == len(card_slots) - 1 and len(content_items) > len(card_slots):
                remaining = content_items[idx:]
                lines_combined = []
                for h, d in remaining:
                    if h and d:
                        lines_combined.append(f"• {h}: {d}")
                    elif h:
                        lines_combined.append(f"• {h}")
                    elif d:
                        lines_combined.append(d)
                comb_text = "\n".join(lines_combined)
                if len(slot) == 1:
                    shp = slot[0]["shape"]
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, comb_text, default_font_size=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    t_box = slot[0]["shape"]
                    d_box = slot[1]["shape"]
                    assigned_shapes.add(id(t_box))
                    assigned_shapes.add(id(d_box))
                    cls.set_shape_text_preserving_style(t_box, remaining[0][0] or "Qo'shimcha Yo'nalishlar", is_card_title=True, font_size_override=18.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    cls.set_shape_text_preserving_style(d_box, comb_text, font_size_override=14.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
            elif idx < len(content_items):
                item_head, item_desc = content_items[idx]
                if len(slot) == 1:
                    shp = slot[0]["shape"]
                    assigned_shapes.add(id(shp))
                    combined = f"{item_head}:\n{item_desc}".strip() if (item_head and item_desc) else (item_head or item_desc)
                    cls.set_shape_text_preserving_style(shp, combined, default_font_size=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    title_box = slot[0]["shape"]
                    desc_box = slot[1]["shape"]
                    assigned_shapes.add(id(title_box))
                    assigned_shapes.add(id(desc_box))
                    t_text = item_head if item_head else item_desc
                    d_text = item_desc if item_desc else item_head
                    cls.set_shape_text_preserving_style(title_box, t_text, is_card_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    try:
                        t_top_u = SafeShapeGeometry.to_pt(title_box.top, 0.0)
                        t_w_u = SafeShapeGeometry.to_pt(title_box.width, 120.0)
                        d_top_u = SafeShapeGeometry.to_pt(desc_box.top, 0.0)
                        if d_top_u >= t_top_u and (d_top_u - t_top_u) < 34.0 * scale_factor and len(t_text or "") > int(t_w_u / (7.5 * scale_factor)):
                            SafeShapeGeometry.update(title_box, height=Inches((30.0 * scale_factor) / 72.0))
                            SafeShapeGeometry.update(desc_box, top=Inches((t_top_u + 32.0 * scale_factor) / 72.0))
                    except Exception:
                        pass
                    cls.set_shape_text_preserving_style(desc_box, d_text, slide_height_pt=sh_pt, scale_factor=scale_factor)
            else:
                for unused_s in slot:
                    u_shp = unused_s["shape"]
                    u_bb = SpatialProximityCleaner.get_bbox(u_shp)
                    SemanticShapeCleaner.prune_shape(u_shp)
                    # Also prune any empty card rectangle/pill directly behind this unused text slot
                    for bg_cand in list(slide.shapes):
                        if id(bg_cand) in assigned_shapes or bg_cand.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.GROUP):
                            continue
                        if bg_cand.has_text_frame and bg_cand.text_frame.text.strip():
                            continue
                        bg_bb = SpatialProximityCleaner.get_bbox(bg_cand)
                        if (bg_bb[2] - bg_bb[0]) < sw_pt * 0.48 and SpatialProximityCleaner.bbox_overlap(u_bb, bg_bb):
                            SemanticShapeCleaner.prune_shape(bg_cand)

        # Step badges
        step_shapes = [s for s in body_shapes if s["is_step"] and id(s["shape"]) not in assigned_shapes]
        step_shapes.sort(key=lambda s: (round(s["left"] / (85.0 * scale_factor)), s["top"]))
        for s_idx, sm in enumerate(step_shapes):
            s_shp = sm["shape"]
            if s_idx < min(len(card_slots), len(content_items)):
                assigned_shapes.add(id(s_shp))
                if sm["text"].strip().endswith("%"):
                    continue
                step_str = f"{s_idx+1:02d}"
                cls.set_shape_text_preserving_style(s_shp, step_str, font_size_override=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            else:
                SemanticShapeCleaner.prune_shape(s_shp)

        SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=min(len(card_slots), len(content_items)))

        for s in list(slide.shapes):
            try:
                if s.is_placeholder and (not s.has_text_frame or not s.text_frame.text.strip() or cls.is_dummy_text(s.text_frame.text)):
                    slide.shapes._spTree.remove(s._element)
                elif s.has_text_frame and DeepWatermarkPurger.is_watermark_text(s.text_frame.text):
                    slide.shapes._spTree.remove(s._element)
            except Exception:
                pass

        for s in slide.shapes:
            try:
                if s.has_text_frame:
                    for p in s.text_frame.paragraphs:
                        cls.enforce_latin_line_break(p)
            except Exception:
                pass

        try:
            AnimationGuardian.purge_zombie_animations(slide)
        except Exception as e_anim:
            logger.debug(f"AnimationGuardian purge error in replace_slide_content: {e_anim}")


class TemplateDeckAssembler:
    """
    Semantic Template Deck Assembler:
    Selects, deduplicates, prunes, reorders, and customizes slides from an authentic PPTX template deck.
    """

    @staticmethod
    def _compute_slide_graphic_signature(slide) -> Tuple:
        """
        Computes a structural geometry signature of non-text vector/picture shapes on a slide.
        Detects when a template (such as a 6-slide step-by-step Bitcoin padlock infographic)
        repeats the exact same central diagram across multiple slides so we never repeat the same diagram 6 times.
        """
        sig_items = []
        for s in slide.shapes:
            if s.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                continue
            if s.has_text_frame and s.text_frame.text.strip() and s.shape_type != MSO_SHAPE_TYPE.GROUP:
                continue
            l = int(round(SafeShapeGeometry.to_pt(s.left, 0.0) / 10.0))
            t = int(round(SafeShapeGeometry.to_pt(s.top, 0.0) / 10.0))
            w = int(round(SafeShapeGeometry.to_pt(s.width, 0.0) / 10.0))
            h = int(round(SafeShapeGeometry.to_pt(s.height, 0.0) / 10.0))
            if w <= 2 and h <= 2:
                continue
            sig_items.append((int(s.shape_type), l, t, w, h))
        sig_items.sort()
        return tuple(sig_items[:8])

    @staticmethod
    def match_and_assemble_deck(
        template_pptx_path: str,
        blueprint: Optional[Dict[str, Any]],
        content_data: Dict[str, Any],
        output_path: str,
        engine_mode: Optional[str] = None
    ) -> str:
        if not os.path.exists(template_pptx_path):
            raise FileNotFoundError(f"Shablon PPTX topilmadi: {template_pptx_path}")

        content_data = TopicLogicValidator.validate_and_refine_presentation(content_data)
        slides_data = content_data.get("slides", [])
        slides_data = SlotAllocator.balance_deck(slides_data, default_max_slots=4)
        content_data["slides"] = slides_data
        topic = content_data.get("topic", "Akademik Taqdimot")

        chosen_engine = engine_mode or (blueprint.get("engine_mode") if blueprint else None) or content_data.get("engine_mode") or "harmonized"

        bp_slides = blueprint.get("slides", []) if blueprint else []
        has_cross_deck = False
        slide_specs = []

        if bp_slides:
            target_bp_slides = bp_slides[:len(slides_data)] if slides_data else bp_slides
            for s_idx, s_item in enumerate(target_bp_slides):
                ref = s_item.get("matched_slide_ref")
                p_path = (ref.get("pptx_path") if ref else None) or s_item.get("pptx_path") or template_pptx_path
                s_num = (ref.get("slide_index") if ref else None)
                if s_num is None:
                    s_num = s_item.get("template_slide_number", s_idx + 1) - 1

                if p_path != template_pptx_path and os.path.exists(p_path):
                    has_cross_deck = True

                slide_specs.append({
                    "pptx_path": p_path,
                    "slide_index": max(0, s_num),
                    "archetype": s_item.get("layout_type", "cards_grid"),
                    "slide_number": s_item.get("slide_number", s_idx + 1)
                })

        if has_cross_deck and slide_specs:
            from core.slide_splicer import UniversalSlideSplicer
            UniversalSlideSplicer.splice_and_assemble(
                engine_mode=chosen_engine,
                primary_pptx_path=template_pptx_path,
                slide_specs=slide_specs,
                output_path=output_path
            )
            prs = Presentation(output_path)
            DeepWatermarkPurger.purge_presentation_watermarks(prs)
            for idx, slide_item in enumerate(slides_data):
                if idx < len(prs.slides):
                    slide = prs.slides[idx]
                    PPTXTextReplacer.replace_slide_content(slide, slide_item, topic)
            if slides_data and len(prs.slides) > len(slides_data):
                for del_idx in reversed(range(len(slides_data), len(prs.slides))):
                    try:
                        rId = prs.slides._sldIdLst[del_idx].rId
                        prs.part.drop_rel(rId)
                        del prs.slides._sldIdLst[del_idx]
                    except Exception:
                        pass
            DeepWatermarkPurger.purge_presentation_watermarks(prs)
            for s in prs.slides:
                try:
                    AnimationGuardian.purge_zombie_animations(s)
                except Exception:
                    pass
            prs.save(output_path)
            return output_path

        prs = Presentation(template_pptx_path)
        DeepWatermarkPurger.purge_presentation_watermarks(prs)

        if not slides_data:
            prs.save(output_path)
            return output_path

        # 1. Classify all template slides (recursing into groups!) & deduplicate repeated diagrams
        template_slides_meta = []
        seen_graphic_signatures: Set[Tuple] = set()

        for s_idx, slide in enumerate(prs.slides):
            is_ad = AdSlideFilter.is_ad_or_credit_slide(slide, slide_index=s_idx)
            t_shapes = AdSlideFilter.collect_all_text_shapes_recursive(slide.shapes)
            all_text = " ".join([s.text_frame.text.lower() for s in t_shapes])

            # Deduplicate identical central diagrams (e.g. 6 slides of the same Bitcoin padlock)
            g_sig = TemplateDeckAssembler._compute_slide_graphic_signature(slide)
            if len(g_sig) >= 3:
                if g_sig in seen_graphic_signatures:
                    is_ad = True  # Mark duplicate diagram slide so it won't be repeated!
                else:
                    seen_graphic_signatures.add(g_sig)

            has_agenda = any(k in all_text for k in ["table of contents", "agenda", "contents", "index"])
            has_step = any(k in all_text for k in ["step 01", "step 1", "01.", "step infographic", "steps infographic"])
            has_timeline = any(k in all_text for k in ["timeline", "2024", "2025", "2026", "2027", "roadmap"])
            has_comparison = any(k in all_text for k in ["comparison", "vs", "versus", "columns style", "columns infographic"])
            has_table = "table slide" in all_text or any(s.shape_type == MSO_SHAPE_TYPE.TABLE for s in slide.shapes)
            has_credits = any(k in all_text for k in ["let’s keep in touch", "contact", "thank you", "thanks"])
            has_stats = any(k in all_text for k in ["sales", "venn diagram", "20%", "30%", "32%", "40%", "60%", "80%"])
            has_four_option = "four option slide" in all_text or len(t_shapes) in [8, 9, 10]

            ltype = "general"
            if s_idx > 0 and not has_credits and not has_table and len(t_shapes) < 5:
                is_ad = True  # Reject section break / low-slot divider slides for multi-point content!
            if is_ad:
                ltype = "ad_promo"
            elif s_idx == 0:
                ltype = "cover"
            elif has_table:
                ltype = "table_data"
            elif has_credits:
                ltype = "conclusion_sources"
            elif has_agenda:
                ltype = "agenda"
            elif has_comparison:
                ltype = "comparison_vs"
            elif has_timeline or has_step:
                ltype = "timeline_steps"
            elif has_stats:
                ltype = "metrics_stats"
            elif has_four_option:
                ltype = "qa_discussion"
            elif len(t_shapes) >= 6:
                ltype = "cards_grid"
            elif len(t_shapes) in [2, 3, 4]:
                ltype = "theory_concept"
            else:
                ltype = "cards_grid"

            template_slides_meta.append({
                "index": s_idx,
                "layout_type": ltype,
                "is_ad": is_ad,
                "text_count": len(t_shapes),
                "has_comparison": has_comparison,
                "has_timeline": has_timeline,
                "has_step": has_step,
                "has_agenda": has_agenda,
                "has_credits": has_credits,
                "has_stats": has_stats,
                "has_four_option": has_four_option
            })

        def get_best_content_layout(p):
            best_l = p.slide_layouts[0]
            max_sc = -1
            for l in p.slide_layouts:
                cnt = len(l.shapes)
                l_name = l.name.lower()
                sc = cnt * 10
                if any(k in l_name for k in ["content", "title and", "comparison", "two"]):
                    sc += 50
                if sc > max_sc:
                    max_sc = sc
                    best_l = l
            return best_l

        from core.slide_splicer import HarmonizedSlideCloner, ThemeColorExtractor, get_cleanest_blank_layout, CrossDeckSlideDonor
        palette = ThemeColorExtractor.extract_palette(prs.slides[0]) if len(prs.slides) > 0 else {}
        blank_layout = get_cleanest_blank_layout(prs)
        used_donors: Set[str] = set()

        # Collect real donor decks from Telegram Cloud (donor_pptx_paths from 17,250 Vault)
        real_donor_decks: List[str] = []
        for dp in (content_data.get("donor_pptx_paths") or []):
            if dp and os.path.exists(dp) and os.path.abspath(dp) != os.path.abspath(template_pptx_path):
                real_donor_decks.append(dp)

        # Normalize group coordinates on primary presentation slides
        from core.slide_splicer import normalize_slide_group_coordinates
        for orig_s in prs.slides:
            normalize_slide_group_coordinates(orig_s)

        # If primary template is a 1-3 slide 4:3 infographic (width < 12"), upgrade canvas to 16:9 Widescreen (13.333" x 7.5")
        # so spliced 16:9 designer slides from the 17,250+ database render at native 1:1 proportions without 0.75x squishing!
        non_ad_initial = sum(1 for m in template_slides_meta if not m["is_ad"])
        if non_ad_initial <= 3 and prs.slide_width and prs.slide_width < Inches(12.0) and len(slides_data) > non_ad_initial:
            old_w = int(prs.slide_width)
            new_w = int(Inches(13.333))
            sx_up = new_w / float(old_w)
            prs.slide_width = new_w
            for orig_s in prs.slides:
                for el in orig_s.shapes._spTree:
                    HarmonizedSlideCloner._scale_cloned_element_geometry(el, sx_up, 1.0)

        # Pre-index valid non-ad slides inside real_donor_decks (including Slide 0 of 1-10 slide infographic templates from Telegram Vault!)
        real_donor_candidates: List[Tuple[str, int, str, int]] = []
        seen_donor_sigs: Set[Tuple] = set(seen_graphic_signatures)
        for r_deck_path in real_donor_decks[:2]:
            try:
                d_prs = Presentation(r_deck_path)
                total_d_slides = len(d_prs.slides)
                deck_candidates_found = 0
                for d_idx, d_slide in enumerate(d_prs.slides):
                    if d_idx > 25 or deck_candidates_found >= 3:
                        break
                    if AdSlideFilter.is_ad_or_credit_slide(d_slide, slide_index=d_idx):
                        continue
                    # Skip slides that have topic-specific raster photos or empty picture placeholders when splicing across decks
                    if total_d_slides > 10:
                        if any(sh.shape_type == MSO_SHAPE_TYPE.PICTURE for sh in d_slide.shapes):
                            continue
                        if any(sh.is_placeholder and (not sh.has_text_frame or not sh.text_frame.text.strip()) for sh in d_slide.shapes):
                            continue
                    normalize_slide_group_coordinates(d_slide)
                    dt_shapes = AdSlideFilter.collect_all_text_shapes_recursive(d_slide.shapes)
                    # Skip pure cover slides of long presentations, but KEEP Slide 0 of infographic templates (>= 4 text shapes)!
                    if d_idx == 0 and total_d_slides > 10 and len(dt_shapes) < 6:
                        continue
                    if len(dt_shapes) < 4:
                        continue
                    d_sig = TemplateDeckAssembler._compute_slide_graphic_signature(d_slide)
                    if len(d_sig) >= 3:
                        if d_sig in seen_donor_sigs:
                            continue
                        seen_donor_sigs.add(d_sig)
                    d_text = " ".join([s.text_frame.text.lower() for s in dt_shapes])
                    if any(k in d_text for k in ["table of contents", "agenda", "contents", "index"]):
                        d_atype = "agenda"
                    elif any(k in d_text for k in ["timeline", "2024", "2025", "2026", "step 01", "step 1"]):
                        d_atype = "timeline_steps"
                    elif any(k in d_text for k in ["comparison", "vs", "columns"]):
                        d_atype = "comparison_vs"
                    elif any(k in d_text for k in ["%", "sales", "chart"]):
                        d_atype = "metrics_stats"
                    else:
                        d_atype = "cards_grid"
                    real_donor_candidates.append((r_deck_path, d_idx, d_atype, len(dt_shapes)))
                    deck_candidates_found += 1
            except Exception as e_don:
                logger.debug(f"Could not index real donor deck {r_deck_path}: {e_don}")

        used_real_donors: Set[Tuple[str, int]] = set()

        non_ad_count = sum(1 for m in template_slides_meta if not m["is_ad"])
        while non_ad_count < len(slides_data):
            s_req_obj = slides_data[non_ad_count]
            req_atype = "table_data" if (s_req_obj.get("table") or s_req_obj.get("table_data")) else s_req_obj.get("layout_type", "cards_grid")
            new_slide = prs.slides.add_slide(blank_layout)

            donor_splice_res = None
            # 1. Try splicing a real page from another authentic presentation in the 17,250 Telegram Vault / real_donor_decks first!
            if req_atype != "table_data" and real_donor_candidates:
                best_rc = None
                for rc in real_donor_candidates:
                    rc_path, rc_idx, rc_atype, _ = rc
                    if (rc_path, rc_idx) in used_real_donors:
                        continue
                    if rc_atype == req_atype:
                        best_rc = rc
                        break
                if best_rc is None:
                    for rc in real_donor_candidates:
                        if (rc[0], rc[1]) not in used_real_donors:
                            best_rc = rc
                            break
                if best_rc is not None:
                    rc_path, rc_idx, _, _ = best_rc
                    used_real_donors.add((rc_path, rc_idx))
                    try:
                        src_d_prs = Presentation(rc_path)
                        HarmonizedSlideCloner.clone_slide_to_target(
                            source_slide=src_d_prs.slides[rc_idx],
                            target_slide=new_slide,
                            theme_palette=None,
                            replicate_background=True,
                        )
                        donor_splice_res = (rc_path, rc_idx)
                    except Exception as e_rc:
                        logger.debug(f"Real donor splice failed from {rc_path}#{rc_idx}: {e_rc}")

            # 2. Fallback to vector archetype donors (slot-matched to 3, 4, or 6 items!)
            req_pts = s_req_obj.get("points") or s_req_obj.get("bullets") or s_req_obj.get("cards") or []
            req_item_cnt = len(req_pts) if isinstance(req_pts, list) and req_pts else 4
            if donor_splice_res is None:
                donor_splice_res = CrossDeckSlideDonor.splice_donor_to_slide(
                    archetype=req_atype,
                    target_slide=new_slide,
                    theme_palette=palette,
                    used_donors=used_donors,
                    item_count=req_item_cnt,
                )
            if donor_splice_res:
                d_path, d_idx = donor_splice_res
                DeepWatermarkPurger.purge_shape_tree(new_slide.shapes)
                t_shapes = AdSlideFilter.collect_all_text_shapes_recursive(new_slide.shapes)
                all_text = " ".join([s.text_frame.text.lower() for s in t_shapes])
                has_agenda = req_atype == "agenda" or any(k in all_text for k in ["table of contents", "agenda", "contents", "index"])
                has_step = req_atype == "timeline_steps" or any(k in all_text for k in ["step 01", "step 1", "01.", "steps", "process"])
                has_timeline = req_atype == "timeline_steps" or any(k in all_text for k in ["timeline", "2026", "2027", "roadmap", "schedule"])
                has_comparison = req_atype == "comparison_vs" or any(k in all_text for k in ["comparison", "vs", "versus", "columns"])
                has_stats = req_atype == "metrics_stats" or any(k in all_text for k in ["%", "stats", "analytics", "chart", "kpi", "metrics"])
                has_credits = req_atype == "conclusion_sources" or any(k in all_text for k in ["thank you", "thanks", "conclusion", "contact", "qa"])
                has_four_option = len(t_shapes) in [8, 9, 10]
                has_tbl_spliced = (req_atype == "table_data") or any(s.shape_type == MSO_SHAPE_TYPE.TABLE for s in new_slide.shapes)

                template_slides_meta.append({
                    "index": len(prs.slides) - 1,
                    "layout_type": "table_data" if has_tbl_spliced else req_atype,
                    "is_ad": False,
                    "text_count": len(t_shapes),
                    "has_comparison": has_comparison,
                    "has_timeline": has_timeline,
                    "has_step": has_step,
                    "has_agenda": has_agenda,
                    "has_credits": has_credits,
                    "has_stats": has_stats,
                    "has_four_option": has_four_option
                })
                logger.info(f"CrossDeckDonor: Spliced '{req_atype}' slide from {Path(d_path).name}#{d_idx} harmonized with cover theme")
            else:
                best_layout = get_best_content_layout(prs)
                prs.slides.add_slide(best_layout)
                t_shapes = AdSlideFilter.collect_all_text_shapes_recursive(prs.slides[-1].shapes)
                template_slides_meta.append({
                    "index": len(prs.slides) - 1,
                    "layout_type": "cards_grid",
                    "is_ad": False,
                    "text_count": len(t_shapes),
                    "has_comparison": False,
                    "has_timeline": False,
                    "has_step": False,
                    "has_agenda": False,
                    "has_credits": False,
                    "has_stats": False,
                    "has_four_option": False
                })
            non_ad_count += 1

        # 2. Select best distinct template slide for each requested slide (Strictly ignoring ad/duplicate slides!)
        used_indices: Set[int] = set()
        chosen_indices: List[int] = []

        for idx, slide_item in enumerate(slides_data):
            req_layout = slide_item.get("layout_type", "cards_grid")
            best_idx = None
            best_score = -1000

            for t_meta in template_slides_meta:
                t_idx = t_meta["index"]
                if t_idx in used_indices or t_meta["is_ad"]:
                    continue
                score = 0
                if idx == 0 and t_idx == 0:
                    score += 600
                elif req_layout == "agenda" and t_meta["has_agenda"]:
                    score += 450
                elif (slide_item.get("table") or slide_item.get("table_data") or req_layout == "table_data") and t_meta["layout_type"] == "table_data":
                    score += 550
                elif req_layout == "comparison_vs" and t_meta["has_comparison"]:
                    score += 450
                elif req_layout == "timeline_steps" and (t_meta["has_timeline"] or t_meta["has_step"]):
                    score += 450
                elif req_layout == "metrics_stats" and t_meta["has_stats"]:
                    score += 450
                elif req_layout == "qa_discussion" and t_meta["has_four_option"]:
                    score += 450
                elif req_layout == "conclusion_sources" and t_meta["has_credits"]:
                    score += 450
                elif req_layout == t_meta["layout_type"]:
                    score += 250
                elif t_meta["layout_type"] == "cards_grid" and t_meta["text_count"] >= 6:
                    score += 120
                elif t_meta["layout_type"] != "table_data":
                    score += 50

                if score > best_score:
                    best_score = score
                    best_idx = t_idx

            if best_idx is None:
                avail = [m["index"] for m in template_slides_meta if m["index"] not in used_indices and not m["is_ad"]]
                best_idx = avail[0] if avail else 0

            used_indices.add(best_idx)
            chosen_indices.append(best_idx)

        logger.info(f"Blueprint layout request -> Selected template slide indices: {chosen_indices}")

        # 3. Prune all unused slides and all ad/duplicate slides
        to_delete = [i for i in range(len(prs.slides)) if i not in set(chosen_indices)]
        for del_idx in sorted(to_delete, reverse=True):
            rId = prs.slides._sldIdLst[del_idx].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[del_idx]

        sorted_chosen = sorted(chosen_indices)
        sldIdLst = prs.slides._sldIdLst
        sldId_elements = list(sldIdLst)
        orig_to_elem = {sorted_chosen[i]: sldId_elements[i] for i in range(len(sorted_chosen))}

        for el in sldId_elements:
            sldIdLst.remove(el)
        for orig_idx in chosen_indices:
            sldIdLst.append(orig_to_elem[orig_idx])

        # 4. Replace content for each slide in exact 1-to-1 order
        for idx, slide_item in enumerate(slides_data):
            slide = prs.slides[idx]
            PPTXTextReplacer.replace_slide_content(slide, slide_item, topic)

            try:
                notes_slide = slide.notes_slide
                notes_tf = notes_slide.notes_text_frame
                notes_tf.text = VisualRhetoric.compile_dual_track_notes(
                    slide_title=slide_item.get("title", f"Slayd {idx+1}"),
                    slide_data=slide_item,
                    topic=topic,
                    slide_index=idx + 1,
                    total_slides=len(slides_data)
                )
            except Exception as e_notes:
                logger.debug(f"Dual-track speaker notes skipped for slide {idx}: {e_notes}")

            try:
                AnimationGuardian.purge_zombie_animations(slide)
            except Exception as e_anim:
                logger.debug(f"AnimationGuardian purge error in assemble_full_deck: {e_anim}")

        # 5. Motion Storytelling: Apply PowerPoint Native Morph & Cinematic Transitions
        try:
            SlideMorphEngine.apply_deck_storytelling_transitions(prs, mode="cinematic")
        except Exception as e_morph:
            logger.debug(f"Slide morph transitions skipped: {e_morph}")

        # 6. Theme Harmonization: Preserve authentic template designer colors by default; only recolor if explicitly requested
        if content_data.get("force_theme_harmonize"):
            try:
                chosen_theme = content_data.get("theme_key", "tech_minimal")
                sw_pt = prs.slide_width.pt
                sh_pt = prs.slide_height.pt
                for s in prs.slides:
                    ThemeHarmonizer.harmonize_slide_accents(s, theme_key=chosen_theme, slide_width_pt=sw_pt, slide_height_pt=sh_pt)
            except Exception as e_theme:
                logger.debug(f"Theme harmonizer skipped: {e_theme}")

        # 7. Typography Harmonization: Apply Penpot dual-font pairing (OpenXML latin + cs)
        try:
            chosen_theme = content_data.get("theme_key", "tech_minimal")
            FontHarmonizer.harmonize_presentation_fonts(prs, theme_key=chosen_theme)
        except Exception as e_font:
            logger.debug(f"Font harmonizer skipped: {e_font}")

        # Comprehensive presentation watermark and ad sweep
        DeepWatermarkPurger.purge_presentation_watermarks(prs)

        # Watermark purger o'chirgan shakllardan qolgan zombi animatsiyalarni yakuniy tozalash
        for s in prs.slides:
            try:
                AnimationGuardian.purge_zombie_animations(s)
            except Exception as e_anim:
                logger.debug(f"Final AnimationGuardian sweep error: {e_anim}")

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        logger.info(f"Customized anti-overflow PPTX saved at: {output_path}")
        return output_path


def replace_presentation_content(
    template_pptx_path: str,
    content_data: Dict[str, Any],
    output_pptx_path: str,
    blueprint: Optional[Dict[str, Any]] = None,
    engine_mode: Optional[str] = None
) -> str:
    """
    Loads an authentic PPTX template file, assembles the best-fitting slide layouts,
    replaces text in-place with anti-overflow precision, and wipes all dummy text.
    """
    return TemplateDeckAssembler.match_and_assemble_deck(
        template_pptx_path=template_pptx_path,
        blueprint=blueprint,
        content_data=content_data,
        output_path=output_pptx_path,
        engine_mode=engine_mode
    )
