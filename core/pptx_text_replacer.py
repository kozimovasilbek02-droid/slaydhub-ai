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
    r'free\s+powerpoint\s+templates?',
    r'designed\s+with\s+.*?by',
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
                    group_has_wm = False
                    for child in s.shapes:
                        if child.has_text_frame and cls.is_watermark_text(child.text_frame.text):
                            group_has_wm = True
                            break
                    if group_has_wm:
                        SemanticShapeCleaner.prune_shape(s)
                        purged_count += 1
                    else:
                        purged_count += cls.purge_shape_tree(s.shapes)
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

        # 1. Slide Masters and Layouts
        for master in prs.slide_masters:
            cls.purge_shape_tree(master.shapes)
            for layout in master.slide_layouts:
                cls.purge_shape_tree(layout.shapes)

        # 2. Slide Instances and Notes
        for slide in prs.slides:
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
    Detects and rejects full advertisement / credit slides from template decks.
    """

    @classmethod
    def is_ad_or_credit_slide(cls, slide) -> bool:
        t_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        all_text = " ".join([s.text_frame.text.lower() for s in t_shapes])
        
        ad_markers = [
            "designed with",
            "presentationgo",
            "free templates & infographics",
            "questions or need help",
            "slidescarnival",
            "slidenest",
            "slidesmania",
            "allppt",
            "slideegg",
            "credits & resources",
            "fonts & colors used",
            "color palette & font",
            "terms of use",
            "our license",
            "thank you for downloading",
            "free powerpoint templates",
            "visit our faq",
            "instructions for use",
            "alternative resources",
            "editable icons",
            "icon set",
            "vector icons",
            "happy designing",
            "pexels, pixabay"
        ]
        for marker in ad_markers:
            if marker in all_text:
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
        Removes unassigned dummy text shapes, leftover badges, and ghost placeholder boxes.
        Also uses SpatialProximityCleaner to prune orphaned vector badges, connectors, and dead arrows.
        """
        pruned_bboxes: List[Tuple[float, float, float, float]] = []
        assigned_bboxes: List[Tuple[float, float, float, float]] = []

        # 0. Collect bounding boxes of assigned shapes
        for s_meta in all_text_shapes:
            shp = s_meta["shape"]
            if id(shp) in assigned_shape_ids:
                assigned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))

        # 1. Collect groups where at least one child shape was assigned
        assigned_groups = set()
        for s_meta in all_text_shapes:
            if id(s_meta["shape"]) in assigned_shape_ids and s_meta.get("parent_group") is not None:
                assigned_groups.add(id(s_meta["parent_group"]))

        # 2. Prune unassigned groups entirely (removes cards, background badges, and dummy text)
        pruned_groups = set()
        for s_meta in all_text_shapes:
            grp = s_meta.get("parent_group")
            if grp is not None and id(grp) not in assigned_groups and id(grp) not in pruned_groups:
                pruned_bboxes.append(SpatialProximityCleaner.get_bbox(grp))
                cls.prune_shape(grp)
                pruned_groups.add(id(grp))

        # 3. Process individual standalone shapes
        for s_meta in all_text_shapes:
            shp = s_meta["shape"]
            shp_id = id(shp)
            if shp_id in assigned_shape_ids:
                continue
            if SmartArtGuardian.is_smart_art(shp):
                continue
            if s_meta.get("parent_group") is not None and id(s_meta["parent_group"]) in pruned_groups:
                continue

            # If it's a step badge beyond the active count, prune it
            if s_meta.get("is_step"):
                step_val = s_meta.get("step_num")
                if keep_step_count is not None and isinstance(step_val, int) and step_val > keep_step_count:
                    pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))
                    cls.prune_shape(shp)
                elif s_meta.get("is_dummy") or s_meta.get("is_watermark"):
                    pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))
                    cls.prune_shape(shp)
                continue

            # If it's an unassigned text shape, watermark, or dummy placeholder, prune it
            if s_meta.get("is_dummy") or s_meta.get("is_watermark") or not s_meta.get("is_valid_zone") or shp.name.lower().startswith(("textbox", "cuadrotexto")):
                pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))
                cls.prune_shape(shp)
            else:
                try:
                    if shp.has_text_frame:
                        shp.text_frame.text = ""
                except Exception:
                    pass
                pruned_bboxes.append(SpatialProximityCleaner.get_bbox(shp))

        # 4. Invoke SpatialProximityCleaner to eliminate orphaned vector badges and dead connectors
        try:
            sw_pt = 720.0
            sh_pt = 405.0
            prs_part = getattr(slide, "part", None)
            if prs_part and hasattr(prs_part, "package"):
                prs_obj = getattr(prs_part.package, "presentation", None)
                if prs_obj:
                    sw_pt = getattr(prs_obj.slide_width, "pt", 720.0)
                    sh_pt = getattr(prs_obj.slide_height, "pt", 405.0)

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
        Identifies image/graphic shapes on the slide and repositions or constrains
        text shapes to prevent overlapping.
        """
        sw, sh, sw_pt, sh_pt = cls.get_slide_dimensions(slide)

        # Find picture and large foreground illustration shapes (excluding full-slide backgrounds)
        visual_shapes = []
        for s in slide.shapes:
            w_pt = SafeShapeGeometry.to_pt(s.width, 0.0)
            h_pt = SafeShapeGeometry.to_pt(s.height, 0.0)
            
            # Ignore full-slide background graphics/cards
            if w_pt > sw_pt * 0.82 and h_pt > sh_pt * 0.82:
                continue

            if s.shape_type == MSO_SHAPE_TYPE.PICTURE:
                visual_shapes.append(s)
            elif not s.has_text_frame and w_pt > 80 and h_pt > 80 and (w_pt < sw_pt * 0.7 or h_pt < sh_pt * 0.7):
                visual_shapes.append(s)

        if not visual_shapes:
            return

        for txt_shp in text_shapes:
            if not txt_shp.has_text_frame:
                continue
            t_box = cls.get_bbox(txt_shp)

            for v_shp in visual_shapes:
                v_box = cls.get_bbox(v_shp)
                if cls.has_overlap(t_box, v_box):
                    v_center_x = (v_box[0] + v_box[2]) / 2.0

                    if v_center_x > sw_pt * 0.45:
                        max_allowed_w = max(120.0, v_box[0] - t_box[0] - 12.0)
                        if t_box[2] > v_box[0]:
                            SafeShapeGeometry.update(txt_shp, width=Inches(max_allowed_w / 72.0))
                    else:
                        new_left = v_box[2] + 12.0
                        if t_box[0] < new_left:
                            avail_w = max(120.0, sw_pt - new_left - 30.0)
                            SafeShapeGeometry.update(txt_shp, left=Inches(new_left / 72.0), width=Inches(avail_w / 72.0))

                    t_box = cls.get_bbox(txt_shp)


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
            re.match(r'^[A-D][\.\:\)]?$', t)
        )

    is_step_number = is_step_badge

    @classmethod
    def get_slide_bg_luminance(cls, part: Any) -> float:
        """Calculates average background luminance (0..255) for the slide container, with caching."""
        if hasattr(part, "_cached_bg_lum"):
            return part._cached_bg_lum

        lum = 255.0  # default light
        containers = []
        if hasattr(part, "slide"):
            containers.append(part.slide)
        if hasattr(part, "slide_layout"):
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
                # Native XPath check for background picture blip
                blip_nodes = celem.xpath('.//p:bg//a:blip | .//p:bgPr//a:blip')
                for blip in blip_nodes:
                    rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rid:
                        c_part = container.part if hasattr(container, "part") else part
                        if rid in c_part.rels:
                            target = c_part.rels[rid].target_part
                            img = Image.open(io.BytesIO(target.blob))
                            img_small = img.resize((15, 15)).convert('L')
                            pixels = list(img_small.get_flattened_data()) if hasattr(img_small, 'get_flattened_data') else list(img_small.getdata())
                            lum = sum(pixels) / float(len(pixels))
                            found = True
                            break
                if found:
                    break

                # Native XPath check for solid background color
                bg_clr_nodes = celem.xpath('.//p:bg//a:srgbClr | .//p:bgPr//a:srgbClr')
                for clr in bg_clr_nodes:
                    hex_str = clr.get('val')
                    if hex_str and len(hex_str) == 6:
                        r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
                        lum = 0.299 * r + 0.587 * g + 0.114 * b
                        found = True
                        break
            except Exception:
                pass

        if not found and hasattr(part, "slide"):
            try:
                slide = part.slide
                sw = part.package.presentation_part.presentation.slide_width
                sh = part.package.presentation_part.presentation.slide_height
                for s in slide.shapes:
                    if s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.width >= sw * 0.90 and s.height >= sh * 0.90:
                        img = Image.open(io.BytesIO(s.image.blob))
                        img_small = img.resize((15, 15)).convert('L')
                        pixels = list(img_small.get_flattened_data()) if hasattr(img_small, 'get_flattened_data') else list(img_small.getdata())
                        lum = sum(pixels) / float(len(pixels))
                        found = True
                        break
            except Exception:
                pass

        part._cached_bg_lum = lum
        return lum

    @classmethod
    def is_shape_on_dark_background(cls, shape: BaseShape) -> bool:
        """Determines whether a shape sits on a dark background (luminance < 135.0)."""
        try:
            elem = shape._element
            # Native lxml XPath check for shape fill color without string serialization
            no_fill = elem.xpath('.//p:spPr/a:noFill')
            if not no_fill:
                clr_nodes = elem.xpath('.//p:spPr//a:srgbClr')
                if clr_nodes:
                    hex_str = clr_nodes[0].get('val')
                    if hex_str and len(hex_str) == 6:
                        r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
                        lum = 0.299 * r + 0.587 * g + 0.114 * b
                        return lum < 135.0
        except Exception:
            pass

        try:
            part = shape.part
            lum = cls.get_slide_bg_luminance(part)
            return lum < 135.0
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

        font_scale = max(0.95, min(1.10, scale_factor))
        usable_w = max(20.0 * font_scale, width_pt - (8.0 * font_scale))
        usable_h = max(15.0 * font_scale, height_pt - (6.0 * font_scale))

        if is_title:
            t_len = len(text)
            if t_len > 60:
                min_fs = 18.0 * font_scale
            elif t_len > 40:
                min_fs = 20.0 * font_scale
            else:
                min_fs = 22.0 * font_scale
            max_initial_fs = max(max_initial_fs, 30.0 * font_scale)
        else:
            min_fs = max(min_fs, 14.0 * font_scale)
            max_initial_fs = max(max_initial_fs, 17.5 * font_scale)

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

        if tf.paragraphs and tf.paragraphs[0].runs:
            r0 = tf.paragraphs[0].runs[0]
            orig_font_name = r0.font.name
            orig_bold = r0.font.bold
            try:
                if r0.font.color and r0.font.color.type:
                    orig_color = r0.font.color.rgb
            except Exception:
                pass

        is_dark_bg = cls.is_shape_on_dark_background(shape)
        if orig_color:
            try:
                c_lum = 0.299 * orig_color[0] + 0.587 * orig_color[1] + 0.114 * orig_color[2]
                if is_dark_bg and c_lum < 135.0:
                    orig_color = None  # Force contrast light color on dark background
                elif not is_dark_bg and c_lum > 215.0:
                    orig_color = None  # Force contrast dark color on light background
            except Exception:
                pass

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

        # Safe dynamic vertical expansion when space is available on slide
        if shape_top_pt + shape_h_pt + (25.0 * scale_factor) < (slide_height_pt * 0.85):
            if is_title and shape_h_pt < 95.0 * scale_factor:
                extra_h_pt = min(40.0 * scale_factor, (slide_height_pt * 0.30) - shape_top_pt - shape_h_pt)
                if extra_h_pt > 0:
                    shape_h_pt += extra_h_pt
                    SafeShapeGeometry.update(shape, height=Inches(shape_h_pt / 72.0))
            elif not is_title and len(lines) > 2:
                extra_h_pt = min(50.0 * scale_factor, (slide_height_pt * 0.85) - (shape_top_pt + shape_h_pt))
                if extra_h_pt > 0:
                    shape_h_pt += extra_h_pt
                    SafeShapeGeometry.update(shape, height=Inches(shape_h_pt / 72.0))

        font_scale = max(0.95, min(1.10, scale_factor))

        if is_title:
            t_len = len(new_text)
            max_init_fs = 32.0 * font_scale
            min_fs = 18.0 * font_scale if t_len > 60 else (20.0 * font_scale if t_len > 40 else 22.0 * font_scale)
        elif is_card_title:
            max_init_fs = 22.0 * font_scale
            min_fs = 16.5 * font_scale
        else:
            max_init_fs = 18.5 * font_scale
            min_fs = 14.5 * font_scale

        if font_size_override:
            max_init_fs = max(max_init_fs, font_size_override * font_scale)
            if not is_title:
                min_fs = min(min_fs, font_size_override * font_scale)
        elif default_font_size:
            max_init_fs = max(max_init_fs, default_font_size * font_scale)

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

        if is_title:
            # Title invariant: NEVER truncate with '...'!
            # Dynamically calculate exact font size down to 16pt so all lines fit completely
            test_fs = scaled_font_size
            while test_fs >= (16.0 * font_scale):
                test_char_w = max(1.0, test_fs * 0.52)
                test_cpl = max(1, int(usable_w / test_char_w))
                total_title_lines = 0
                for line in lines:
                    words = line.split()
                    cur_p = ""
                    for w in words:
                        if not cur_p:
                            cur_p = w
                        elif len(cur_p) + 1 + len(w) <= test_cpl:
                            cur_p += " " + w
                        else:
                            total_title_lines += 1
                            cur_p = w
                    if cur_p:
                        total_title_lines += 1
                if (total_title_lines * test_fs * 1.25) <= (usable_h * 1.05):
                    scaled_font_size = test_fs
                    break
                test_fs -= 0.5
            else:
                scaled_font_size = max(15.0 * font_scale, test_fs)
            condensed_lines = lines
        else:
            char_w = max(1.0, scaled_font_size * 0.52)
            chars_per_line = max(1, int(usable_w / char_w))
            max_lines_allowed = max(1, int(usable_h / (scaled_font_size * 1.22)))

            condensed_lines = []
            cur_lines_count = 0
            for line in lines:
                words = line.split()
                line_parts = []
                cur_p = ""
                for w in words:
                    if not cur_p:
                        cur_p = w
                    elif len(cur_p) + 1 + len(w) <= chars_per_line:
                        cur_p += " " + w
                    else:
                        line_parts.append(cur_p)
                        cur_p = w
                if cur_p:
                    line_parts.append(cur_p)
                
                if cur_lines_count + len(line_parts) <= max_lines_allowed:
                    condensed_lines.append(line)
                    cur_lines_count += len(line_parts)
                else:
                    remaining_lines = max_lines_allowed - cur_lines_count
                    if remaining_lines > 0:
                        fitting_parts = line_parts[:remaining_lines]
                        if len(line_parts) > remaining_lines:
                            fitting_parts[-1] = re.sub(r'[\s\.\,\;\:]+$', '', fitting_parts[-1]) + "..."
                        condensed_lines.append(" ".join(fitting_parts))
                        cur_lines_count += len(fitting_parts)
                    break
        
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
            subtitle = slide_data.get("subtitle", "Akademik va amaliy tahliliy qo'llanma | 2026-yil")
            content_items.append(("", subtitle))

        elif layout_type == "agenda":
            points = slide_data.get("points", [])
            for p in points:
                clean_p = re.sub(r'^\d+[\.\)]\s*', '', p)
                if ":" in clean_p:
                    h, d = clean_p.split(":", 1)
                    content_items.append((h.strip(), d.strip()))
                else:
                    content_items.append((clean_p, ""))

        elif layout_type == "theory_concept":
            concept_name = slide_data.get("concept_name", "Nazariy Tushuncha")
            # Prevent title duplication: if concept_name echoes the slide title, replace with analytical heading
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
            left_t = slide_data.get("left_title", "1-Yondashuv")
            left_pts = "\n".join([f"• {p}" for p in slide_data.get("left_points", [])])
            right_t = slide_data.get("right_title", "2-Yondashuv")
            right_pts = "\n".join([f"• {p}" for p in slide_data.get("right_points", [])])
            content_items.append((left_t, left_pts))
            content_items.append((right_t, right_pts))

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
                        clean_item = re.sub(r'^(?:[\•\-\–\—\*]|\d+[\.\)]|\(\d+\))\s*', '', item.strip()).strip()
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
            card_h_pt = min(135.0, (avail_h_pt - gap_pt) / 2)
        elif N == 3:
            card_h_pt = min(100.0, (avail_h_pt - (2 * gap_pt)) / 3)
        else: # N == 4
            card_h_pt = min(78.0, (avail_h_pt - (3 * gap_pt)) / 4)

        is_dark_bg = cls.get_slide_bg_luminance(slide.part if hasattr(slide, "part") else slide) < 135.0
        if is_dark_bg:
            title_color = RGBColor(0x93, 0xC5, 0xFD)
            body_color = RGBColor(0xF1, 0xF5, 0xF9)
        else:
            title_color = primary_color if primary_color else RGBColor(0x1F, 0x4E, 0x79)
            body_color = RGBColor(0x2A, 0x2A, 0x2A)

        font_scale = max(0.95, min(1.10, sh_pt / 540.0))
        head_fs = (21.0 if N <= 2 else (18.5 if N == 3 else 17.0)) * font_scale
        body_fs = (16.5 if N <= 2 else (15.5 if N == 3 else 14.5)) * font_scale

        for i, (head, desc) in enumerate(items_to_render):
            cur_top_pt = top_start_pt + i * (card_h_pt + gap_pt)
            
            card_shp = slide.shapes.add_textbox(
                Inches(left_m_pt / 72.0),
                Inches(cur_top_pt / 72.0),
                Inches(content_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
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
        font_scale = max(0.95, min(1.10, sh_pt / 540.0))

        for i, (val, lbl, desc) in enumerate(metric_items):
            cur_left_pt = left_m_pt + i * (card_w_pt + gap_pt)
            card_shp = slide.shapes.add_textbox(
                Inches(cur_left_pt / 72.0),
                Inches(top_start_pt / 72.0),
                Inches(card_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
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
        font_scale = max(0.95, min(1.10, sh_pt / 540.0))

        for i, (num, title, desc) in enumerate(step_items):
            cur_left_pt = left_m_pt + i * (card_w_pt + gap_pt)
            card_shp = slide.shapes.add_textbox(
                Inches(cur_left_pt / 72.0),
                Inches(top_start_pt / 72.0),
                Inches(card_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
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
        font_scale = max(0.95, min(1.10, sh_pt / 540.0))

        for i, (head, desc) in enumerate(items_to_render):
            cur_left_pt = left_m_pt + i * (card_w_pt + gap_pt)
            card_shp = slide.shapes.add_textbox(
                Inches(cur_left_pt / 72.0),
                Inches(top_start_pt / 72.0),
                Inches(card_w_pt / 72.0),
                Inches(card_h_pt / 72.0)
            )
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
    def replace_slide_content(cls, slide, slide_data: Dict[str, Any], topic: str) -> None:
        """
        Replaces text placeholders inside a template slide with safe content zone constraints,
        deep watermark eradication, resolution-aware scaling, and XML shape pruning.
        """
        sw, sh, sw_pt, sh_pt = cls.get_slide_dimensions(slide)
        scale_factor = max(0.75, min(1.33, sh_pt / 540.0))

        layout_type = slide_data.get("layout_type", "cards_grid")
        title_text = slide_data.get("title", topic)

        def collect_text_shapes_recursive(shapes_container, parent_group=None):
            collected = []
            for s in list(shapes_container):
                if SmartArtGuardian.is_smart_art(s):
                    logger.debug(f"SmartArtGuardian: Shielding SmartArt '{getattr(s, 'name', '')}' from standard text replacement.")
                    continue
                if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                    collected.extend(collect_text_shapes_recursive(s.shapes, parent_group=s))
                elif s.has_text_frame and s.shape_type != MSO_SHAPE_TYPE.PICTURE:
                    txt = s.text_frame.text.strip()
                    if not txt:
                        continue
                    
                    is_wm = DeepWatermarkPurger.is_watermark_text(txt)
                    
                    l_val = s.left if s.left is not None else (parent_group.left if parent_group else 0)
                    t_val = s.top if s.top is not None else (parent_group.top if parent_group else 0)
                    w_val = s.width if s.width is not None else (parent_group.width if parent_group else 0)
                    h_val = s.height if s.height is not None else (parent_group.height if parent_group else 0)
                    
                    l_pt = SafeShapeGeometry.to_pt(l_val, 0.0)
                    t_pt = SafeShapeGeometry.to_pt(t_val, 0.0)
                    w_pt = SafeShapeGeometry.to_pt(w_val, 0.0)
                    h_pt = SafeShapeGeometry.to_pt(h_val, 0.0)

                    eff_left = SafeShapeGeometry.to_pt(parent_group.left, l_pt) if parent_group and parent_group.left is not None else l_pt
                    eff_top = SafeShapeGeometry.to_pt(parent_group.top, t_pt) if parent_group and parent_group.top is not None else t_pt

                    is_valid_zone = (eff_left >= -10.0 and eff_left <= sw_pt * 0.95 and (w_pt >= 80.0 or parent_group is not None))
                    
                    step_num = None
                    m = re.match(r'^(?:step\s*)?0?([1-9]|1[0-9])[\.\:\)]?$', txt, re.I)
                    if m:
                        step_num = int(m.group(1))

                    collected.append({
                        "shape": s,
                        "parent_group": parent_group,
                        "text": txt,
                        "name": s.name,
                        "left": eff_left,
                        "top": eff_top,
                        "width": w_pt,
                        "height": h_pt,
                        "inner_top": t_pt,
                        "is_step": cls.is_step_badge(txt),
                        "step_num": step_num,
                        "is_dummy": cls.is_dummy_text(txt),
                        "is_watermark": is_wm,
                        "is_valid_zone": is_valid_zone,
                        "font_size": max([r.font.size.pt for p in s.text_frame.paragraphs for r in p.runs if r.font.size] or [12.0 * scale_factor])
                    })
            return collected

        all_text_shapes = collect_text_shapes_recursive(slide.shapes)

        if not all_text_shapes:
            cls._synthesize_fallback_slide_content(slide, slide_data, topic, sw_pt, sh_pt, scale_factor)
            return

        # Prune blatant watermarks immediately from the slide
        for sm in all_text_shapes:
            if sm["is_watermark"]:
                SemanticShapeCleaner.prune_shape(sm["shape"])

        # Filter active shapes to only valid non-watermark shapes
        active_shapes = [s for s in all_text_shapes if not s["is_watermark"]]
        if not active_shapes:
            cls._synthesize_fallback_slide_content(slide, slide_data, topic, sw_pt, sh_pt, scale_factor)
            return

        # Resolve image vs text collisions before filling text
        SpatialCollisionEngine.resolve_image_text_collisions(slide, [s["shape"] for s in active_shapes])

        # -------------------------------------------------------------
        # 1. SPECIALIZED LAYOUT HANDLERS FOR HIGHEST VISUAL PERFECTION
        # -------------------------------------------------------------
        shape_names = {s["name"]: s["shape"] for s in active_shapes}

        # Handle Agenda Slide
        if layout_type == "agenda" and ("CuadroTexto 24" in shape_names or any("table of contents" in s["text"].lower() for s in active_shapes)):
            pts = slide_data.get("points", [])
            assigned_shapes = set()
            for s in active_shapes:
                shp = s["shape"]
                name = s["name"]
                t = s["text"].lower()
                if name == "CuadroTexto 24" or "table of contents" in t:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 5" and len(pts) > 0:
                    assigned_shapes.add(id(shp))
                    clean_t = re.sub(r'^\d+[\.\)]\s*', '', pts[0]).split(":", 1)[0]
                    cls.set_shape_text_preserving_style(shp, clean_t, is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 4" and len(pts) > 0:
                    assigned_shapes.add(id(shp))
                    clean_d = pts[0].split(":", 1)[1].strip() if ":" in pts[0] else pts[0]
                    cls.set_shape_text_preserving_style(shp, clean_d, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 11" and len(pts) > 1:
                    assigned_shapes.add(id(shp))
                    clean_t = re.sub(r'^\d+[\.\)]\s*', '', pts[1]).split(":", 1)[0]
                    cls.set_shape_text_preserving_style(shp, clean_t, is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 10" and len(pts) > 1:
                    assigned_shapes.add(id(shp))
                    clean_d = pts[1].split(":", 1)[1].strip() if ":" in pts[1] else pts[1]
                    cls.set_shape_text_preserving_style(shp, clean_d, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 16" and len(pts) > 2:
                    assigned_shapes.add(id(shp))
                    clean_t = re.sub(r'^\d+[\.\)]\s*', '', pts[2]).split(":", 1)[0]
                    cls.set_shape_text_preserving_style(shp, clean_t, is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 15" and len(pts) > 2:
                    assigned_shapes.add(id(shp))
                    clean_d = pts[2].split(":", 1)[1].strip() if ":" in pts[2] else pts[2]
                    cls.set_shape_text_preserving_style(shp, clean_d, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 21" and len(pts) > 3:
                    assigned_shapes.add(id(shp))
                    clean_t = re.sub(r'^\d+[\.\)]\s*', '', pts[3]).split(":", 1)[0]
                    cls.set_shape_text_preserving_style(shp, clean_t, is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "CuadroTexto 20" and len(pts) > 3:
                    assigned_shapes.add(id(shp))
                    clean_d = pts[3].split(":", 1)[1].strip() if ":" in pts[3] else pts[3]
                    cls.set_shape_text_preserving_style(shp, clean_d, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(pts))
            return

        # Handle Theory & Concept (Circular Photo Slide)
        if layout_type == "theory_concept" and len(active_shapes) <= 4 and any("big image" in s["text"].lower() for s in active_shapes):
            c_name = slide_data.get("concept_name", "Nazariy Tushuncha")
            definition = slide_data.get("definition", "")
            formula = slide_data.get("key_formula_or_law", "")
            importance = slide_data.get("importance", "")
            
            sorted_by_top = sorted(active_shapes, key=lambda s: s["top"])
            upper_shp = sorted_by_top[0]["shape"]
            lower_shp = sorted_by_top[1]["shape"] if len(sorted_by_top) > 1 else None
            
            assigned_shapes = {id(upper_shp)}
            cls.set_shape_text_preserving_style(upper_shp, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            if lower_shp:
                assigned_shapes.add(id(lower_shp))
                body_lines = [
                    f"• {c_name}:",
                    f"  {definition}",
                    f"",
                    f"• Asosiy Qonuniyat / Formula:",
                    f"  {formula}",
                    f"",
                    f"• Ilmiy va Amaliy Ahamiyati:",
                    f"  {importance}"
                ]
                cls.set_shape_text_preserving_style(lower_shp, "\n".join(body_lines), font_size_override=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            
            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # Handle Native PowerPoint Tables (MSO_SHAPE_TYPE.TABLE)
        table_shapes = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.TABLE]
        if table_shapes and ("table_data" in slide_data or "rows" in slide_data or "matrix" in slide_data):
            t_data = slide_data.get("table_data") if isinstance(slide_data.get("table_data"), dict) else (slide_data.get("matrix") if isinstance(slide_data.get("matrix"), dict) else {})
            headers = t_data.get("headers", []) or slide_data.get("headers", [])
            rows = t_data.get("rows", []) or slide_data.get("rows", [])
            if headers or rows:
                TableMatrixReplacer.apply_table_data(
                    table_shape=table_shapes[0],
                    headers=headers,
                    rows=rows,
                    theme_key=slide_data.get("theme_key", "tech_minimal")
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

        # Handle Comparison VS Slide (Slide with Diamonds & Sub-cards)

        if layout_type == "comparison_vs" and ("TextBox 30" in shape_names and "TextBox 31" in shape_names):
            left_t = slide_data.get("left_title", "Klassik Asimmetrik Kriptografiya")
            left_pts = slide_data.get("left_points", [])
            right_t = slide_data.get("right_title", "Post-Kvant Kriptografiyasi (PQC)")
            right_pts = slide_data.get("right_points", [])

            assigned_shapes = set()
            for s in active_shapes:
                shp = s["shape"]
                name = s["name"]
                if name == "TextBox 15":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 30":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, left_t, is_card_title=True, font_size_override=19.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 10":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "Faktorizatsiya & Logarifm", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 11":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, left_pts[0] if len(left_pts) > 0 else "Faktorizatsiyaga tayanadi", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 7":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "Kvant Zaifligi", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 18":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, left_pts[1] if len(left_pts) > 1 else "Kvant kompyuterlarida buziladi", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 19":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "Kalit Parametri", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 20":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, left_pts[2] if len(left_pts) > 2 else "Kichik kalit va yuqori tezlik", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 31":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, right_t, is_card_title=True, font_size_override=19.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 23":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "Panjaralar & Xeshlar", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 24":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, right_pts[0] if len(right_pts) > 0 else "Panjaralarga asoslangan", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 26":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "Mutlaq Himoya", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 27":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, right_pts[1] if len(right_pts) > 1 else "Kvant hujumlariga chidamli", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 28":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "Resurs Talabi", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 29":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, right_pts[2] if len(right_pts) > 2 else "Kattaroq kalit hajmi talab etiladi", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # Handle Timeline Steps (Staircase Infographic)
        if layout_type == "timeline_steps" and ("TextBox 49" in shape_names and "TextBox 36" in shape_names):
            steps = slide_data.get("steps", [])
            assigned_shapes = set()
            for s in active_shapes:
                shp = s["shape"]
                name = s["name"]
                if name == "TextBox 1":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 49" and len(steps) > 0:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[0].get("title", ""), is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 48" and len(steps) > 0:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[0].get("desc", ""), font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 51" and len(steps) > 1:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[1].get("title", ""), is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 50" and len(steps) > 1:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[1].get("desc", ""), font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 53" and len(steps) > 2:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[2].get("title", ""), is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 52" and len(steps) > 2:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[2].get("desc", ""), font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 55" and len(steps) > 3:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[3].get("title", ""), is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 54" and len(steps) > 3:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, steps[3].get("desc", ""), font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(steps))
            return

        # Handle Q&A Discussion (4 Cards)
        if layout_type == "qa_discussion" and ("TextBox 7" in shape_names and "TextBox 6" in shape_names):
            qa_list = slide_data.get("qa_list", [])
            assigned_shapes = set()
            for s in active_shapes:
                shp = s["shape"]
                name = s["name"]
                if name == "TextBox 14":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 7":
                    assigned_shapes.add(id(shp))
                    q_text = qa_list[0].get("question", "SNDL Hujumi nima?") if len(qa_list) > 0 else "SNDL Hujumi"
                    cls.set_shape_text_preserving_style(shp, f"❓ {q_text}", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 6":
                    assigned_shapes.add(id(shp))
                    a_text = qa_list[0].get("answer", "") if len(qa_list) > 0 else ""
                    cls.set_shape_text_preserving_style(shp, f"💡 {a_text}", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 9":
                    assigned_shapes.add(id(shp))
                    q_text = qa_list[1].get("question", "Qachon PQC ga o'tish kerak?") if len(qa_list) > 1 else "Migratsiya Vaqti"
                    cls.set_shape_text_preserving_style(shp, f"❓ {q_text}", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 15" or name == "TextBox 8":
                    assigned_shapes.add(id(shp))
                    a_text = qa_list[1].get("answer", "") if len(qa_list) > 1 else ""
                    cls.set_shape_text_preserving_style(shp, f"💡 {a_text}", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 11":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "❓ Simmetrik Shifrlash Xavfsizmi?", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 16":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "💡 AES-256 kvant hisoblashlarida ham samarali 128-bit xavfsizlikni to'liq kafolatlaydi.", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 13":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "❓ Asosiy PQC Standartlari Qaysilar?", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 17":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "💡 NIST tasdiqlagan ML-KEM (Kyber) kalit almashish va ML-DSA (Dilithium) raqamli imzo standartlaridir.", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # Handle Metrics & Stats (Venn / Stats)
        if layout_type == "metrics_stats" and ("TextBox 36" in shape_names and "TextBox 37" in shape_names):
            metrics = slide_data.get("metrics", [])
            assigned_shapes = set()
            for s in active_shapes:
                shp = s["shape"]
                name = s["name"]
                if name == "TextBox 2":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, title_text, is_title=True, font_size_override=28.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 36" and len(metrics) > 0:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, f"{metrics[0].get('value', '')} - {metrics[0].get('label', '')}", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 37" and len(metrics) > 0:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, metrics[0].get("desc", ""), font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 38" and len(metrics) > 1:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, f"{metrics[1].get('value', '')} - {metrics[1].get('label', '')}", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 39" and len(metrics) > 1:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, metrics[1].get("desc", ""), font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 40" and len(metrics) > 2:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, f"{metrics[2].get('value', '')} - {metrics[2].get('label', '')}", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 41" and len(metrics) > 2:
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, metrics[2].get("desc", ""), font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 42":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "256-bit - Simmetrik Himoya", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                elif name == "TextBox 43":
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, "AES-256 algoritmi kvant hujumlariga qarshi 128-bitlik xavfsizlik darajasini to'liq saqlab qoladi.", font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # -------------------------------------------------------------
        # 2. GENERAL SPATIAL CLUSTERING FOR ANY TEMPLATE DECK
        # -------------------------------------------------------------
        # Filter valid on-canvas shapes (strict width floor >= 120pt for content boxes)
        valid_shapes = [s for s in active_shapes if s["is_valid_zone"]]
        if not valid_shapes:
            valid_shapes = [
                s for s in active_shapes 
                if (SafeShapeGeometry.to_pt(s["shape"].width, 0.0) >= 120.0 or s["is_step"])
            ]

        # Identify Title Shape (Upper Banner with width >= 120pt)
        title_candidates = []
        for s in valid_shapes:
            if s["is_step"]:
                continue
            w_pt = SafeShapeGeometry.to_pt(s["shape"].width, 0.0)
            if w_pt < 120.0:
                continue

            score = 0
            t_low = s["text"].lower()
            s_name_low = s["name"].lower()
            
            if any(k in t_low for k in ["table of contents", "agenda", "contents", "steps infographic", "timeline infographic", "comparison slide", "bullet points", "table slide", "credits"]):
                score += 500
            elif "title" in s_name_low or "header" in s_name_low or "cuadrotexto 24" in s_name_low or "textbox 1" in s_name_low or "textbox 15" in s_name_low:
                score += 250

            if s["top"] < sh * 0.35:
                score += 100
                score += int((1.0 - (s["top"] / (sh * 0.35))) * 50)
            elif s["top"] < sh * 0.50:
                score += 20
            else:
                score -= 150 # Harsh penalty for title placed at bottom

            if s["width"] > sw * 0.35:
                score += 30
            if s["font_size"] >= (24 * scale_factor):
                score += 25

            title_candidates.append((score, s))

        title_candidates.sort(key=lambda x: x[0], reverse=True)
        default_left_m_pt = 45.0 * scale_factor

        best_title_meta = title_candidates[0][1] if (title_candidates and title_candidates[0][0] > 0) else None
        orig_title_left = SafeShapeGeometry.to_pt(best_title_meta["shape"].left, 0.0) if best_title_meta else 0.0

        # Body candidate shapes (excluding the selected title shape)
        body_candidates_for_geom = [s for s in valid_shapes if (not best_title_meta or s["shape"] != best_title_meta["shape"])]
        body_lefts = [s["left"] for s in body_candidates_for_geom if s["left"] >= 20.0 and s["width"] >= 100.0]
        body_rights = [s["left"] + s["width"] for s in body_candidates_for_geom if s["left"] >= 20.0 and s["width"] >= 100.0]

        # Background zone awareness: If template body placeholders or title are shifted right (indicating a left sidebar / color stripe), respect it!
        if body_lefts:
            min_body_l = min(body_lefts)
            if min_body_l >= 70.0 * scale_factor:
                left_m_pt = min_body_l
            elif orig_title_left >= 70.0 * scale_factor:
                left_m_pt = orig_title_left
            else:
                left_m_pt = default_left_m_pt
        elif orig_title_left >= 70.0 * scale_factor:
            left_m_pt = orig_title_left
        else:
            left_m_pt = default_left_m_pt

        # Calculate max right boundary to ensure content never overflows the canvas
        max_right_pt = sw_pt - (35.0 * scale_factor)
        if body_rights:
            max_body_r = max(body_rights)
            if max_body_r > left_m_pt + 200.0:
                max_right_pt = min(max_right_pt, max_body_r)

        content_w_pt = max(260.0, max_right_pt - left_m_pt)
        left_m = Inches(left_m_pt / 72.0)
        content_w = Inches(content_w_pt / 72.0)

        if best_title_meta:
            title_shape_meta = best_title_meta
            cls.set_shape_text_preserving_style(title_shape_meta["shape"], title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
            t_top_pt = SafeShapeGeometry.to_pt(title_shape_meta["shape"].top, 28.0)
            if t_top_pt > 45.0 * scale_factor:
                title_shape_meta["shape"].top = Inches((26.0 * scale_factor) / 72.0)
            cur_h_pt = SafeShapeGeometry.to_pt(title_shape_meta["shape"].height, 55.0)
            title_shape_meta["shape"].height = Inches(max(cur_h_pt, 62.0 * scale_factor) / 72.0)

            # Preserve title left if the template designer positioned it away from sidebars
            if orig_title_left >= 70.0 * scale_factor:
                title_shape_meta["shape"].left = Inches(orig_title_left / 72.0)
                title_w_pt = min(content_w_pt, sw_pt - orig_title_left - (35.0 * scale_factor))
                title_shape_meta["shape"].width = Inches(title_w_pt / 72.0)
            else:
                title_shape_meta["shape"].left = left_m
                title_w_pt = min(content_w_pt, sw_pt - left_m_pt - (35.0 * scale_factor))
                title_shape_meta["shape"].width = Inches(title_w_pt / 72.0)

            assigned_shapes = {id(title_shape_meta["shape"])}
        else:
            # Create a clean, wide title banner dynamically fitting canvas width
            title_w_pt = min(content_w_pt, sw_pt - left_m_pt - (35.0 * scale_factor))
            title_w = Inches(title_w_pt / 72.0)
            t_box = slide.shapes.add_textbox(
                left_m,
                Inches(0.38 * scale_factor),
                title_w,
                Inches(0.85 * scale_factor)
            )
            cls.set_shape_text_preserving_style(t_box, title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
            assigned_shapes = {id(t_box)}
            title_shape_meta = {"shape": t_box}

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

            sub_text = "\n\n".join(subtitle_parts) if subtitle_parts else topic

            # Find a suitable subtitle box from remaining valid shapes
            sub_candidates = [s for s in valid_shapes if id(s["shape"]) not in assigned_shapes and not s["is_step"]]
            if sub_candidates:
                sub_candidates.sort(key=lambda s: s["top"])
                sub_shape = sub_candidates[0]["shape"]
                cls.set_shape_text_preserving_style(sub_shape, sub_text, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                assigned_shapes.add(id(sub_shape))
            else:
                t_bottom = title_shape_meta["shape"].top + title_shape_meta["shape"].height
                sub_box = slide.shapes.add_textbox(
                    left_m,
                    t_bottom + Inches(0.2 * scale_factor),
                    content_w,
                    Inches(1.2 * scale_factor)
                )
                cls.set_shape_text_preserving_style(sub_box, sub_text, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                assigned_shapes.add(id(sub_box))

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # Extract structured content items via reusable extractor
        content_items = cls._extract_content_items(slide_data, layout_type, title_text=title_text)

        # Shape Grouping into Cards / Columns in Safe Zone
        body_shapes = [s for s in valid_shapes if s["shape"] != title_shape_meta["shape"]]

        # Check if template has card groups (shapes organized inside GroupShape containers)
        has_group_cards = any(s.get("parent_group") is not None for s in body_shapes)
        if has_group_cards:
            card_dict = {}
            for s in body_shapes:
                if s["is_step"]:
                    continue
                grp = s.get("parent_group")
                card_key = id(grp) if grp else id(s["shape"])
                if card_key not in card_dict:
                    card_dict[card_key] = {
                        "parent_group": grp,
                        "shapes": [],
                        "left": s["left"],
                        "top": s["top"]
                    }
                card_dict[card_key]["shapes"].append(s)

            # Sort cards logically: reading order (left-to-right columns, top-to-bottom within column)
            card_list = list(card_dict.values())
            card_list.sort(key=lambda c: (int(c["left"] > sw_pt * 0.45), c["top"]))

            # Check if group cards are cramped narrow slivers (width < 250pt) or insufficient for all content items
            avg_card_w = sum(SafeShapeGeometry.to_pt(c["parent_group"].width if c.get("parent_group") else c["shapes"][0]["shape"].width, 0.0) for c in card_list if c["shapes"]) / max(1, len(card_list))
            if (avg_card_w < 250.0 or len(card_list) < len(content_items)) and len(content_items) >= 2:
                # Prune cramped template cards and synthesize visual archetype cards across content_w_pt
                for card in card_list:
                    grp = card.get("parent_group")
                    if grp is not None:
                        SemanticShapeCleaner.prune_shape(grp)
                    else:
                        for extra_s in card["shapes"]:
                            SemanticShapeCleaner.prune_shape(extra_s["shape"])

                # Ensure title banner is guaranteed present and legible on canvas
                clean_target_title = re.sub(r'[\s\u00A0]+', ' ', title_text).strip().lower()
                has_title = False
                if title_shape_meta and title_shape_meta.get("shape") in list(slide.shapes):
                    has_title = True
                    assigned_shapes.add(id(title_shape_meta["shape"]))
                else:
                    for s in slide.shapes:
                        try:
                            if s.has_text_frame and s.text_frame.text:
                                norm_st = re.sub(r'[\s\u00A0]+', ' ', s.text_frame.text).strip().lower()
                                if clean_target_title in norm_st or (len(clean_target_title) > 15 and clean_target_title[:15] in norm_st):
                                    has_title = True
                                    assigned_shapes.add(id(s))
                                    break
                        except Exception:
                            pass
                if not has_title:
                    t_box = slide.shapes.add_textbox(left_m, Inches(0.55 * scale_factor), content_w, Inches(0.85 * scale_factor))
                    cls.set_shape_text_preserving_style(t_box, title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    assigned_shapes.add(id(t_box))

                top_start_pt = cls.calculate_card_top_start_pt(slide, title_text, title_shape_meta, scale_factor)

                # Clean orphaned diagram vectors and empty decorative shapes in the content canvas so new archetype cards have a clean, unblemished background
                for s in list(slide.shapes):
                    if s.shape_type != MSO_SHAPE_TYPE.PLACEHOLDER and (not s.has_text_frame or not s.text_frame.text.strip()):
                        s_top = SafeShapeGeometry.to_pt(s.top, 0.0)
                        s_w = SafeShapeGeometry.to_pt(s.width, 0.0)
                        s_h = SafeShapeGeometry.to_pt(s.height, 0.0)
                        if s_top >= top_start_pt - 25.0 and s_w < sw_pt * 0.85 and s_h < sh_pt * 0.85:
                            SemanticShapeCleaner.prune_shape(s)
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

            # Populate cards 1-to-1 with content_items with large typography
            for c_idx, card in enumerate(card_list):
                c_shapes = card["shapes"]
                c_shapes.sort(key=lambda x: x.get("inner_top", x["top"]))
                grp = card["parent_group"]

                if c_idx < len(content_items):
                    item_head, item_desc = content_items[c_idx]
                    if len(c_shapes) >= 2:
                        t_box = c_shapes[0]["shape"]
                        d_box = c_shapes[1]["shape"]
                        assigned_shapes.add(id(t_box))
                        assigned_shapes.add(id(d_box))
                        cls.set_shape_text_preserving_style(t_box, item_head or "Xulosa", is_card_title=True, font_size_override=18.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                        cls.set_shape_text_preserving_style(d_box, item_desc or item_head, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                        for extra in c_shapes[2:]:
                            SemanticShapeCleaner.prune_shape(extra["shape"])
                    elif len(c_shapes) == 1:
                        shp = c_shapes[0]["shape"]
                        assigned_shapes.add(id(shp))
                        combined = f"{item_head}\n{item_desc}".strip() if (item_head and item_desc) else (item_head or item_desc)
                        cls.set_shape_text_preserving_style(shp, combined, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    # Prune excess unassigned card group!
                    if grp is not None:
                        SemanticShapeCleaner.prune_shape(grp)
                    else:
                        for extra_s in c_shapes:
                            SemanticShapeCleaner.prune_shape(extra_s["shape"])

            # Clean any remaining unassigned shapes
            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(content_items))

            # Remove empty placeholders
            for s in list(slide.shapes):
                try:
                    if s.is_placeholder and (not s.has_text_frame or not s.text_frame.text.strip() or cls.is_dummy_text(s.text_frame.text)):
                        slide.shapes._spTree.remove(s._element)
                    elif s.has_text_frame and DeepWatermarkPurger.is_watermark_text(s.text_frame.text):
                        slide.shapes._spTree.remove(s._element)
                except Exception:
                    pass

            # Visual Enhancement: Inject topic-relevant scientific illustration or diagram into slide
            try:
                from core.image_fetcher import AcademicImageFetcher
                img_p = AcademicImageFetcher.get_image_for_slide(topic, title_text, layout_type)
                if img_p:
                    AcademicImageFetcher.inject_image_into_slide(slide, img_p, sw_pt, sh_pt, scale_factor, layout_type=layout_type)
            except Exception as e:
                logger.debug(f"Image injection error: {e}")
            return
        non_step_body = [
            s for s in body_shapes 
            if not s["is_step"] and (SafeShapeGeometry.to_pt(s["shape"].width, 0.0) >= 120.0)
        ]

        # True spatial column clustering based on horizontal bounding box overlap and proximity
        def get_overlap_x(s1, s2):
            l1, r1 = s1["left"], s1["left"] + s1["width"]
            l2, r2 = s2["left"], s2["left"] + s2["width"]
            return max(0, min(r1, r2) - max(l1, l2))

        columns: List[List[Dict[str, Any]]] = []
        for s in sorted(non_step_body, key=lambda x: x["left"]):
            matched_col = None
            for col in columns:
                c_lead = col[0]
                overlap = get_overlap_x(s, c_lead)
                min_w = min(s["width"], c_lead["width"])
                center_dist = abs((s["left"] + s["width"] / 2) - (c_lead["left"] + c_lead["width"] / 2))
                # Horizontal overlap >= 35% or centers within 1.2 inches
                if (min_w > 0 and overlap / min_w >= 0.35) or center_dist < (Inches(1.2).emu * scale_factor):
                    matched_col = col
                    break
            if matched_col is not None:
                matched_col.append(s)
            else:
                columns.append([s])

        # Sort columns strictly left-to-right
        columns.sort(key=lambda col: col[0]["left"])

        # Within each column, sort shapes strictly top-to-bottom
        for col in columns:
            col.sort(key=lambda s: s["top"])

        # Comparison layout specific handling: requires 2 distinct side-by-side columns
        if layout_type == "comparison_vs":
            left_t = slide_data.get("left_title")
            right_t = slide_data.get("right_title")
            left_pts_raw = slide_data.get("left_points")
            right_pts_raw = slide_data.get("right_points")

            if not left_t and content_items and len(content_items) >= 2:
                left_t = content_items[0][0]
                left_pts_raw = [content_items[0][1]] if content_items[0][1] else [content_items[0][0]]
                right_t = content_items[1][0]
                right_pts_raw = [content_items[1][1]] if content_items[1][1] else [content_items[1][0]]

            if not left_t:
                left_t = "1-Yondashuv"
            if not right_t:
                right_t = "2-Yondashuv"
            if not left_pts_raw:
                left_pts_raw = []
            if not right_pts_raw:
                right_pts_raw = []

            left_pts = "\n".join([f"• {p}" if not p.startswith("•") else p for p in left_pts_raw])
            right_pts = "\n".join([f"• {p}" if not p.startswith("•") else p for p in right_pts_raw])

            col_widths_ok = (
                len(columns) >= 2 and
                SafeShapeGeometry.to_pt(columns[0][0]["shape"].width, 0.0) >= (200.0 * scale_factor) and
                SafeShapeGeometry.to_pt(columns[1][0]["shape"].width, 0.0) >= (200.0 * scale_factor)
            )

            if col_widths_ok:
                # Column 0 = Left
                col_l = columns[0]
                if len(col_l) == 1:
                    assigned_shapes.add(id(col_l[0]["shape"]))
                    cls.set_shape_text_preserving_style(col_l[0]["shape"], f"{left_t}\n\n{left_pts}", default_font_size=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    assigned_shapes.add(id(col_l[0]["shape"]))
                    assigned_shapes.add(id(col_l[1]["shape"]))
                    cls.set_shape_text_preserving_style(col_l[0]["shape"], left_t, is_card_title=True, font_size_override=18.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    cls.set_shape_text_preserving_style(col_l[1]["shape"], left_pts, font_size_override=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    for extra in col_l[2:]:
                        SemanticShapeCleaner.prune_shape(extra["shape"])

                # Column 1 = Right
                col_r = columns[1]
                if len(col_r) == 1:
                    assigned_shapes.add(id(col_r[0]["shape"]))
                    cls.set_shape_text_preserving_style(col_r[0]["shape"], f"{right_t}\n\n{right_pts}", default_font_size=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    assigned_shapes.add(id(col_r[0]["shape"]))
                    assigned_shapes.add(id(col_r[1]["shape"]))
                    cls.set_shape_text_preserving_style(col_r[0]["shape"], right_t, is_card_title=True, font_size_override=18.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    cls.set_shape_text_preserving_style(col_r[1]["shape"], right_pts, font_size_override=15.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    for extra in col_r[2:]:
                        SemanticShapeCleaner.prune_shape(extra["shape"])

                # Prune any unused columns 2..N
                for u_col in columns[2:]:
                    for s in u_col:
                        SemanticShapeCleaner.prune_shape(s["shape"])
            else:
                # Template columns are narrow/cramped or absent: prune them and synthesize 2 balanced wide comparison cards
                for s in non_step_body:
                    SemanticShapeCleaner.prune_shape(s["shape"])
                for col in columns:
                    for s in col:
                        SemanticShapeCleaner.prune_shape(s["shape"])

                # Dynamically calculate side-by-side balanced cards fitting canvas width
                half_gap_pt = 24.0 * scale_factor
                avail_w_pt = content_w_pt - half_gap_pt
                half_w = Inches((avail_w_pt / 2.0) / 72.0)
                right_col_left = Inches((left_m_pt + (avail_w_pt / 2.0) + half_gap_pt) / 72.0)
                comp_top_pt = cls.calculate_card_top_start_pt(slide, title_text, title_shape_meta, scale_factor)
                card_body_h = Inches(min(320.0 * scale_factor, sh_pt - comp_top_pt - 45.0 * scale_factor) / 72.0)

                lt_box = slide.shapes.add_textbox(left_m, Inches(comp_top_pt / 72.0), half_w, Inches(0.55 * scale_factor))
                assigned_shapes.add(id(lt_box))
                cls.set_shape_text_preserving_style(lt_box, left_t, is_card_title=True, font_size_override=19.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

                lb_box = slide.shapes.add_textbox(left_m, Inches((comp_top_pt + 44.0 * scale_factor) / 72.0), half_w, card_body_h)
                assigned_shapes.add(id(lb_box))
                cls.set_shape_text_preserving_style(lb_box, left_pts, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

                rt_box = slide.shapes.add_textbox(right_col_left, Inches(comp_top_pt / 72.0), half_w, Inches(0.55 * scale_factor))
                assigned_shapes.add(id(rt_box))
                cls.set_shape_text_preserving_style(rt_box, right_t, is_card_title=True, font_size_override=19.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

                rb_box = slide.shapes.add_textbox(right_col_left, Inches((comp_top_pt + 44.0 * scale_factor) / 72.0), half_w, card_body_h)
                assigned_shapes.add(id(rb_box))
                cls.set_shape_text_preserving_style(rb_box, right_pts, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # General layouts
        if len(columns) <= 1:
            if len(content_items) >= 2:
                # User preference: render specialized visual archetype cards across content_w_pt
                for col in columns:
                    for s in col:
                        SemanticShapeCleaner.prune_shape(s["shape"])

                clean_target_title = re.sub(r'[\s\u00A0]+', ' ', title_text).strip().lower()
                has_title = False
                if title_shape_meta and title_shape_meta.get("shape") in list(slide.shapes):
                    has_title = True
                    assigned_shapes.add(id(title_shape_meta["shape"]))
                else:
                    for s in slide.shapes:
                        try:
                            if s.has_text_frame and s.text_frame.text:
                                norm_st = re.sub(r'[\s\u00A0]+', ' ', s.text_frame.text).strip().lower()
                                if clean_target_title in norm_st or (len(clean_target_title) > 15 and clean_target_title[:15] in norm_st):
                                    has_title = True
                                    assigned_shapes.add(id(s))
                                    break
                        except Exception:
                            pass
                if not has_title:
                    t_box = slide.shapes.add_textbox(left_m, Inches(0.55 * scale_factor), content_w, Inches(0.85 * scale_factor))
                    cls.set_shape_text_preserving_style(t_box, title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    assigned_shapes.add(id(t_box))

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
                SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
                return
            elif len(columns) == 1 and content_items:
                col_shapes = columns[0]
                formatted_lines = []
                for h, d in content_items:
                    if h and d:
                        formatted_lines.append(f"• {h}:\n  {d}")
                    elif h:
                        formatted_lines.append(f"• {h}")
                    elif d:
                        formatted_lines.append(d)
                full_text = "\n\n".join(formatted_lines)

                if len(col_shapes) == 1:
                    shp = col_shapes[0]["shape"]
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, full_text, default_font_size=16.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    t_box = col_shapes[0]["shape"]
                    d_box = col_shapes[1]["shape"]
                    assigned_shapes.add(id(t_box))
                    assigned_shapes.add(id(d_box))
                    first_head = content_items[0][0] if content_items[0][0] else "Asosiy Konsept"
                    cls.set_shape_text_preserving_style(t_box, first_head, is_card_title=True, font_size_override=19.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    cls.set_shape_text_preserving_style(d_box, full_text, font_size_override=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    for extra in col_shapes[2:]:
                        SemanticShapeCleaner.prune_shape(extra["shape"])
            elif not columns and content_items:
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
            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # Check if columns are cramped narrow slivers (e.g. columns with avg width < 160pt)
        avg_col_w = sum(SafeShapeGeometry.to_pt(c[0]["shape"].width, 0.0) for c in columns) / max(1, len(columns)) if columns else 0.0
        if ((len(columns) >= 3 and avg_col_w < 150.0) or (len(columns) == 2 and avg_col_w < 180.0)) and len(content_items) >= 2:
            for col in columns:
                for s in col:
                    SemanticShapeCleaner.prune_shape(s["shape"])

            clean_target_title = re.sub(r'[\s\u00A0]+', ' ', title_text).strip().lower()
            has_title = False
            if title_shape_meta and title_shape_meta.get("shape") in list(slide.shapes):
                has_title = True
                assigned_shapes.add(id(title_shape_meta["shape"]))
            else:
                for s in slide.shapes:
                    try:
                        if s.has_text_frame and s.text_frame.text:
                            norm_st = re.sub(r'[\s\u00A0]+', ' ', s.text_frame.text).strip().lower()
                            if clean_target_title in norm_st or (len(clean_target_title) > 15 and clean_target_title[:15] in norm_st):
                                has_title = True
                                assigned_shapes.add(id(s))
                                break
                    except Exception:
                        pass
            if not has_title:
                t_box = slide.shapes.add_textbox(left_m, Inches(0.55 * scale_factor), content_w, Inches(0.85 * scale_factor))
                cls.set_shape_text_preserving_style(t_box, title_text, is_title=True, slide_height_pt=sh_pt, scale_factor=scale_factor)
                assigned_shapes.add(id(t_box))

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
            SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes)
            return

        # Map content items 1-to-1 to columns, condensing overflow into last column if items > columns
        for idx in range(len(columns)):
            col_shapes = columns[idx]
            if idx == len(columns) - 1 and len(content_items) > len(columns):
                remaining = content_items[idx:]
                lines_combined = []
                for h, d in remaining:
                    if h and d:
                        lines_combined.append(f"• {h}:\n  {d}")
                    elif h:
                        lines_combined.append(f"• {h}")
                    elif d:
                        lines_combined.append(d)
                comb_text = "\n\n".join(lines_combined)

                if len(col_shapes) == 1:
                    shp = col_shapes[0]["shape"]
                    assigned_shapes.add(id(shp))
                    cls.set_shape_text_preserving_style(shp, comb_text, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    t_box = col_shapes[0]["shape"]
                    d_box = col_shapes[1]["shape"]
                    assigned_shapes.add(id(t_box))
                    assigned_shapes.add(id(d_box))
                    cls.set_shape_text_preserving_style(t_box, remaining[0][0] if remaining[0][0] else "Qo'shimcha Xulosalar", is_card_title=True, font_size_override=19.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    cls.set_shape_text_preserving_style(d_box, comb_text, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                for extra in col_shapes[2:]:
                    SemanticShapeCleaner.prune_shape(extra["shape"])
            elif idx < len(content_items):
                item_head, item_desc = content_items[idx]
                if len(col_shapes) == 1:
                    shp = col_shapes[0]["shape"]
                    assigned_shapes.add(id(shp))
                    combined = f"{item_head}\n{item_desc}".strip() if (item_head and item_desc) else (item_head or item_desc)
                    cls.set_shape_text_preserving_style(shp, combined, default_font_size=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
                else:
                    title_box = col_shapes[0]["shape"]
                    desc_box = col_shapes[1]["shape"]
                    assigned_shapes.add(id(title_box))
                    assigned_shapes.add(id(desc_box))

                    t_text = item_head if item_head else item_desc
                    d_text = item_desc if item_head else ""

                    cls.set_shape_text_preserving_style(title_box, t_text, is_card_title=True, font_size_override=19.5, slide_height_pt=sh_pt, scale_factor=scale_factor)
                    cls.set_shape_text_preserving_style(desc_box, d_text, font_size_override=15.5, slide_height_pt=sh_pt, scale_factor=scale_factor)

                    for extra_shp in col_shapes[2:]:
                        SemanticShapeCleaner.prune_shape(extra_shp["shape"])

        # Crucial: Physically prune unused columns to eliminate leftover placeholder text
        if len(columns) > len(content_items):
            for unused_col in columns[len(content_items):]:
                for s in unused_col:
                    SemanticShapeCleaner.prune_shape(s["shape"])

        # Also isolate step badges so they only receive step numbers, never long text (preventing 1-letter cascades)
        step_shapes = [s for s in body_shapes if s["is_step"]]
        step_shapes.sort(key=lambda s: s["left"])
        for s_idx, sm in enumerate(step_shapes):
            s_shp = sm["shape"]
            if s_idx < len(content_items):
                assigned_shapes.add(id(s_shp))
                step_str = f"{s_idx+1:02d}" if sm.get("step_num") is not None else sm["text"]
                cls.set_shape_text_preserving_style(s_shp, step_str, font_size_override=16.0, slide_height_pt=sh_pt, scale_factor=scale_factor)
            else:
                SemanticShapeCleaner.prune_shape(s_shp)

        # Safely prune all remaining unused dummy text shapes, footer boxes, and narrow strips
        SemanticShapeCleaner.cleanup_unassigned_shapes(slide, assigned_shapes, all_text_shapes, keep_step_count=len(content_items))

        # CRITICAL: Physical sweep for ANY remaining empty or dummy placeholders to eliminate "Click to edit title" and "Текст слайда"
        for s in list(slide.shapes):
            try:
                if s.is_placeholder and (not s.has_text_frame or not s.text_frame.text.strip() or cls.is_dummy_text(s.text_frame.text)):
                    slide.shapes._spTree.remove(s._element)
                elif s.has_text_frame and DeepWatermarkPurger.is_watermark_text(s.text_frame.text):
                    slide.shapes._spTree.remove(s._element)
            except Exception:
                pass

        # CRITICAL: Universal enforcement of latinLnBrk="0" across all paragraphs on slide to prevent mid-word wrapping
        for s in slide.shapes:
            try:
                if s.has_text_frame:
                    for p in s.text_frame.paragraphs:
                        cls.enforce_latin_line_break(p)
            except Exception:
                pass

        # Visual Enhancement: Inject topic-relevant scientific illustration or diagram into slide
        try:
            from core.image_fetcher import AcademicImageFetcher
            img_p = AcademicImageFetcher.get_image_for_slide(topic, title_text, layout_type)
            if img_p:
                AcademicImageFetcher.inject_image_into_slide(slide, img_p, sw_pt, sh_pt, scale_factor, layout_type=layout_type)
        except Exception as e:
            logger.debug(f"Image injection error: {e}")


class TemplateDeckAssembler:
    """
    Semantic Template Deck Assembler:
    Selects, prunes, reorders, and customizes slides from an authentic PPTX template deck.
    """

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

        # Run Topic Logic Validator to sanitize and structure slide text
        content_data = TopicLogicValidator.validate_and_refine_presentation(content_data)

        slides_data = content_data.get("slides", [])
        # Balance content density to prevent slot crowding and font shrinkage
        slides_data = SlotAllocator.balance_deck(slides_data, default_max_slots=4)
        content_data["slides"] = slides_data
        topic = content_data.get("topic", "Akademik Taqdimot")


        chosen_engine = engine_mode or (blueprint.get("engine_mode") if blueprint else None) or content_data.get("engine_mode") or "harmonized"

        # Check if blueprint contains multi-template cross-deck slide specifications
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

        # If cross-deck slides exist, execute UniversalSlideSplicer with selected engine
        if has_cross_deck and slide_specs:
            logger.info(f"Assembling multi-archetype presentation using engine '{chosen_engine}' across {len(slide_specs)} slides...")
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
            # Prune any excess slides beyond len(slides_data)
            if slides_data and len(prs.slides) > len(slides_data):
                for del_idx in reversed(range(len(slides_data), len(prs.slides))):
                    try:
                        rId = prs.slides._sldIdLst[del_idx].rId
                        prs.part.drop_rel(rId)
                        del prs.slides._sldIdLst[del_idx]
                    except Exception as e:
                        logger.warning(f"Failed to drop excess slide {del_idx}: {e}")
            # Final sweep of entire customized presentation
            DeepWatermarkPurger.purge_presentation_watermarks(prs)
            prs.save(output_path)
            logger.info(f"Successfully assembled and customized multi-deck PPTX at {output_path}")
            return output_path

        prs = Presentation(template_pptx_path)

        # 0. Deep Purge of all promotional watermarks from Masters, Layouts, and Slides
        DeepWatermarkPurger.purge_presentation_watermarks(prs)

        if not slides_data:
            prs.save(output_path)
            return output_path

        # 1. Classify all template slides & identify ad/credit slides
        template_slides_meta = []
        for s_idx, slide in enumerate(prs.slides):
            is_ad = AdSlideFilter.is_ad_or_credit_slide(slide)
            t_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
            all_text = " ".join([s.text_frame.text.lower() for s in t_shapes])
            
            has_agenda = any(k in all_text for k in ["table of contents", "agenda", "contents", "index"])
            has_step = any(k in all_text for k in ["step 01", "step 1", "01.", "step infographic", "steps infographic"])
            has_timeline = any(k in all_text for k in ["timeline", "2026", "2027", "2028", "2029", "2030", "roadmap"])
            has_comparison = any(k in all_text for k in ["comparison", "vs", "versus"])
            has_table = "table slide" in all_text
            has_credits = any(k in all_text for k in ["credits", "fonts", "images", "let’s keep in touch", "contact", "thank you", "thanks"])
            has_stats = any(k in all_text for k in ["sales", "venn diagram", "20%", "37%", "strengths", "weaknesses"])
            has_four_option = "four option slide" in all_text or len(t_shapes) in [8, 9, 10]
            
            ltype = "general"
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
            elif len(t_shapes) >= 10:
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

        # Ensure presentation has at least as many non-ad slides as requested
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

        non_ad_count = sum(1 for m in template_slides_meta if not m["is_ad"])
        while non_ad_count < len(slides_data):
            best_layout = get_best_content_layout(prs)
            prs.slides.add_slide(best_layout)
            t_shapes = [s for s in prs.slides[-1].shapes if s.has_text_frame and s.text_frame.text.strip()]
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

        # 2. Select best distinct template slide for each requested slide (Strictly ignoring ad slides!)
        used_indices: Set[int] = set()
        chosen_indices: List[int] = []
        
        for idx, slide_item in enumerate(slides_data):
            req_layout = slide_item.get("layout_type", "cards_grid")
            best_idx = None
            best_score = -1000
            
            for t_meta in template_slides_meta:
                t_idx = t_meta["index"]
                if t_idx in used_indices or t_meta["is_ad"]:
                    continue # NEVER pick an ad slide!
                score = 0
                if idx == 0 and t_idx == 0:
                    score += 500
                elif req_layout == "agenda" and t_meta["has_agenda"]:
                    score += 400
                elif req_layout == "comparison_vs" and t_meta["has_comparison"]:
                    score += 400
                elif req_layout == "timeline_steps" and (t_meta["has_timeline"] or t_meta["has_step"]):
                    score += 400
                elif req_layout == "metrics_stats" and t_meta["has_stats"]:
                    score += 400
                elif req_layout == "qa_discussion" and t_meta["has_four_option"]:
                    score += 400
                elif req_layout == "conclusion_sources" and t_meta["has_credits"]:
                    score += 400
                elif req_layout == t_meta["layout_type"]:
                    score += 200
                elif t_meta["layout_type"] == "cards_grid" and t_meta["text_count"] >= 6:
                    score += 100
                elif t_meta["layout_type"] != "table_data":
                    score += 50
                    
                if score > best_score:
                    best_score = score
                    best_idx = t_idx
                    
            if best_idx is None:
                avail = [i for i, m in enumerate(template_slides_meta) if i not in used_indices and not m["is_ad"]]
                best_idx = avail[0] if avail else 0
                
            used_indices.add(best_idx)
            chosen_indices.append(best_idx)
            
        logger.info(f"Blueprint layout request -> Selected template slide indices: {chosen_indices}")
        
        # 3. Prune all unused slides and all ad slides
        to_delete = [i for i in range(len(prs.slides)) if i not in set(chosen_indices)]
        for del_idx in sorted(to_delete, reverse=True):
            rId = prs.slides._sldIdLst[del_idx].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[del_idx]
            
        # Reorder slides in prs.slides._sldIdLst to match chosen_indices exactly
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

            # Dual-Track Deck: Compile rich speaker notes (speech script + evidence + Q&A)
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

        # 5. Motion Storytelling: Apply PowerPoint Native Morph & Cinematic Transitions
        try:
            SlideMorphEngine.apply_deck_storytelling_transitions(prs, mode="cinematic")
        except Exception as e_morph:
            logger.debug(f"Slide morph transitions skipped: {e_morph}")

        # 6. Theme Harmonization: Harmonize vector accent shapes to Penpot palette
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
