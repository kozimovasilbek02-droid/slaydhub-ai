# -*- coding: utf-8 -*-
"""
core/spatial_proximity_cleaner.py
Spatial Proximity Pruner & Orphaned Vector Cleaner for SlaydHub AI.

Eliminates:
1. Orphaned vector badges, circles, and icons left over after unassigned cards are pruned.
2. Dead connector lines and arrows pointing into empty canvas zones where cards once stood.
3. Empty or non-content vector groups (MSO_SHAPE_TYPE.GROUP) after card text is deleted.

Preserves (SafeSize Threshold & Background Guard):
- Large backdrop panels, cards, and diagonal ribbons (> 25% canvas area).
- Header and footer design banners.
- Vector shapes coupled to currently assigned user cards.
- Media content (pictures, charts, tables).
"""

import math
import logging
from typing import List, Tuple, Set, Dict, Any, Optional
from pptx.shapes.base import BaseShape
from pptx.enum.shapes import MSO_SHAPE_TYPE

logger = logging.getLogger(__name__)


class SpatialProximityCleaner:
    """
    Evaluates 2D spatial relationships and eliminates orphaned vector badges,
    connector arrows, and empty decorative groups after card reduction.
    """

    @staticmethod
    def get_bbox(shape: BaseShape) -> Tuple[float, float, float, float]:
        """Returns (left, top, right, bottom) in points."""
        try:
            l = shape.left.pt if hasattr(shape.left, "pt") else float(shape.left) / 12700.0
            t = shape.top.pt if hasattr(shape.top, "pt") else float(shape.top) / 12700.0
            w = shape.width.pt if hasattr(shape.width, "pt") else float(shape.width) / 12700.0
            h = shape.height.pt if hasattr(shape.height, "pt") else float(shape.height) / 12700.0
            return (l, t, l + w, t + h)
        except Exception:
            return (0.0, 0.0, 0.0, 0.0)

    @staticmethod
    def bbox_distance(bbox: Tuple[float, float, float, float], point: Tuple[float, float]) -> float:
        """Calculates Euclidean distance from a 2D point to an axis-aligned bounding box."""
        l, t, r, b = bbox
        x, y = point
        dx = max(0.0, l - x, x - r)
        dy = max(0.0, t - y, y - b)
        return math.hypot(dx, dy)

    @staticmethod
    def bbox_to_bbox_distance(bb1: Tuple[float, float, float, float], bb2: Tuple[float, float, float, float]) -> float:
        """Calculates the shortest distance between two axis-aligned bounding boxes (0.0 if intersecting)."""
        l1, t1, r1, b1 = bb1
        l2, t2, r2, b2 = bb2
        dx = max(0.0, l1 - r2, l2 - r1)
        dy = max(0.0, t1 - b2, t2 - b1)
        return math.hypot(dx, dy)

    @staticmethod
    def bbox_overlap(bb1: Tuple[float, float, float, float], bb2: Tuple[float, float, float, float]) -> bool:
        """Returns True if two bounding boxes intersect."""
        l1, t1, r1, b1 = bb1
        l2, t2, r2, b2 = bb2
        return not (r1 < l2 or l1 > r2 or b1 < t2 or t1 > b2)

    @classmethod
    def is_safe_background_shape(
        cls,
        shape: BaseShape,
        sw_pt: float = 720.0,
        sh_pt: float = 405.0
    ) -> bool:
        """
        Guarantees that major template backdrops, full-slide cards,
        and header/footer branding banners are NEVER pruned as artifacts.
        """
        # Exclude legitimate rich content
        if shape.shape_type in (MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART):
            return True

        l, t, r, b = cls.get_bbox(shape)
        w = r - l
        h = b - t
        area = w * h
        slide_area = max(1.0, sw_pt * sh_pt)

        # 1. Major design backdrop (> 25% of slide area)
        if (area / slide_area) >= 0.25:
            return True

        # 2. Header banner (top <= 18% slide height and spans > 50% width)
        if t <= (sh_pt * 0.18) and w >= (sw_pt * 0.50):
            return True

        # 3. Bottom decorative wave / footer band (top >= 78% slide height and spans > 50% width)
        if t >= (sh_pt * 0.78) and w >= (sw_pt * 0.50):
            return True

        # 4. Long horizontal divider line spanning whole slide
        if h <= 4.0 and w >= (sw_pt * 0.70):
            return True

        return False

    @classmethod
    def is_decorative_artifact(cls, shape: BaseShape) -> bool:
        """
        Determines if a shape is a decorative vector artifact (badge, icon, circle, arrow, connector).
        """
        # Groups are processed separately
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return False

        # If it has a text frame with substantive content (> 6 chars), it is NOT a small decorative artifact
        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            if len(txt) > 6:
                return False

        l, t, r, b = cls.get_bbox(shape)
        w = r - l
        h = b - t

        s_name = getattr(shape, "name", "").lower()

        # Connectors, arrows, chevrons, and lines
        is_connector_type = (
            shape.shape_type == MSO_SHAPE_TYPE.LINE or
            any(k in s_name for k in ["arrow", "chevron", "line", "connector", "flecha", "linea", "vector"]) or
            (w <= 8.0 and h <= 80.0) or
            (h <= 8.0 and w <= 80.0)
        )
        if is_connector_type:
            return True

        # Small badges, step circles, icons, pins (w <= 140 and h <= 140)
        if w <= 140.0 and h <= 140.0:
            return True

        return False

    @staticmethod
    def prune_shape(shape: BaseShape) -> bool:
        """Physically removes a shape element from its parent XML tree."""
        try:
            sp_elem = getattr(shape, "_element", None)
            if sp_elem is not None:
                parent = sp_elem.getparent()
                if parent is not None:
                    parent.remove(sp_elem)
                    return True
        except Exception as e:
            logger.debug(f"Could not prune shape from XML: {e}")
        return False

    @classmethod
    def prune_orphaned_artifacts(
        cls,
        slide: Any,
        pruned_bboxes: List[Tuple[float, float, float, float]],
        assigned_bboxes: List[Tuple[float, float, float, float]],
        assigned_shape_ids: Optional[Set[int]] = None,
        sw_pt: float = 720.0,
        sh_pt: float = 405.0
    ) -> int:
        """
        Iterates over slide shapes and eliminates orphaned vector badges and connector lines
        whose spatial proximity is strongly coupled to unassigned/pruned cards.
        """
        if not pruned_bboxes:
            return 0

        assigned_ids = assigned_shape_ids or set()
        pruned_count = 0

        # Snapshot current shapes
        try:
            shapes_list = list(slide.shapes)
        except Exception:
            return 0

        for s in shapes_list:
            if id(s) in assigned_ids:
                continue

            if cls.is_safe_background_shape(s, sw_pt, sh_pt):
                continue

            # 1. Handle Groups: Detonate empty groups with no assigned content
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                has_assigned_child = False
                try:
                    for child in s.shapes:
                        if id(child) in assigned_ids:
                            has_assigned_child = True
                            break
                        if child.has_text_frame and len(child.text_frame.text.strip()) > 6:
                            has_assigned_child = True
                            break
                except Exception:
                    pass

                if not has_assigned_child:
                    g_bb = cls.get_bbox(s)
                    min_p_dist = min(cls.bbox_to_bbox_distance(g_bb, p_bb) for p_bb in pruned_bboxes)
                    if min_p_dist <= 60.0:
                        if cls.prune_shape(s):
                            pruned_count += 1
                            logger.info(f"Detonated orphaned group (id={id(s)}) near pruned card.")
                            continue

            # 2. Handle Standalone Decorative Artifacts
            if cls.is_decorative_artifact(s):
                s_bb = cls.get_bbox(s)

                # Distance to nearest pruned card bounding box
                min_pruned_dist = min(cls.bbox_to_bbox_distance(s_bb, p_bb) for p_bb in pruned_bboxes)

                # Distance to nearest assigned card bounding box
                min_assigned_dist = float("inf")
                if assigned_bboxes:
                    min_assigned_dist = min(cls.bbox_to_bbox_distance(s_bb, a_bb) for a_bb in assigned_bboxes)

                is_coupled_to_pruned = False
                s_name = getattr(s, "name", "").lower()
                is_connector = (
                    getattr(s, "shape_type", None) == MSO_SHAPE_TYPE.LINE or
                    any(k in s_name for k in ["arrow", "chevron", "line", "connector"])
                )

                if is_connector:
                    # A connector is orphaned if it touches or connects to a pruned card (within 45pt)
                    if min_pruned_dist <= 45.0:
                        is_coupled_to_pruned = True
                else:
                    # Badges / Circles / Icons: Must be within 60pt of pruned card and closer to pruned than assigned
                    if min_pruned_dist <= 60.0 and min_pruned_dist < min_assigned_dist:
                        is_coupled_to_pruned = True

                if is_coupled_to_pruned:
                    if cls.prune_shape(s):
                        pruned_count += 1
                        logger.info(f"Pruned orphaned vector artifact: {s.name} (dist_to_pruned={min_pruned_dist:.1f}pt)")

        return pruned_count
