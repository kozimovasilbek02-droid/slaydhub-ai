# -*- coding: utf-8 -*-
"""
core/spatial_proximity_cleaner.py
Spatial Proximity Pruner & Orphaned Vector Cleaner for SlaydHub AI.
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
        l, t, r, b = bbox
        x, y = point
        dx = max(0.0, l - x, x - r)
        dy = max(0.0, t - y, y - b)
        return math.hypot(dx, dy)

    @staticmethod
    def bbox_to_bbox_distance(bb1: Tuple[float, float, float, float], bb2: Tuple[float, float, float, float]) -> float:
        l1, t1, r1, b1 = bb1
        l2, t2, r2, b2 = bb2
        dx = max(0.0, l1 - r2, l2 - r1)
        dy = max(0.0, t1 - b2, t2 - b1)
        return math.hypot(dx, dy)

    @staticmethod
    def bbox_overlap(bb1: Tuple[float, float, float, float], bb2: Tuple[float, float, float, float]) -> bool:
        l1, t1, r1, b1 = bb1
        l2, t2, r2, b2 = bb2
        return not (r1 < l2 or l1 > r2 or b1 < t2 or t1 > b2)

    @classmethod
    def is_safe_background_shape(
        cls,
        shape: BaseShape,
        sw_pt: float = 960.0,
        sh_pt: float = 540.0,
    ) -> bool:
        if shape.shape_type in (MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.PICTURE):
            return True

        l, t, r, b = cls.get_bbox(shape)
        w = r - l
        h = b - t
        area = w * h
        slide_area = max(1.0, sw_pt * sh_pt)

        if (area / slide_area) >= 0.20:
            return True
        if t <= (sh_pt * 0.18) and w >= (sw_pt * 0.45):
            return True
        if t >= (sh_pt * 0.78) and w >= (sw_pt * 0.45):
            return True
        if h <= 6.0 and w >= (sw_pt * 0.65):
            return True

        return False

    @classmethod
    def is_decorative_artifact(cls, shape: BaseShape, scale_factor: float = 1.0) -> bool:
        if shape.shape_type in (MSO_SHAPE_TYPE.GROUP, MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART):
            return False

        if shape.has_text_frame:
            txt = shape.text_frame.text.strip()
            if len(txt) > 4:
                return False

        l, t, r, b = cls.get_bbox(shape)
        w = r - l
        h = b - t
        s_name = getattr(shape, "name", "").lower()

        is_connector_type = (
            shape.shape_type == MSO_SHAPE_TYPE.LINE
            or any(k in s_name for k in ["arrow", "chevron", "line", "connector", "flecha", "linea"])
            or (w <= 8.0 * scale_factor and h <= 80.0 * scale_factor)
            or (h <= 8.0 * scale_factor and w <= 80.0 * scale_factor)
        )
        if is_connector_type:
            return True

        if w <= 95.0 * scale_factor and h <= 95.0 * scale_factor:
            return True

        return False

    @staticmethod
    def prune_shape(shape: BaseShape) -> bool:
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
    def _group_has_assigned_or_artwork_recursive(cls, grp, assigned_ids: Set[int]) -> bool:
        """
        Recursively checks if a group or any of its nested sub-groups contains:
        - an assigned shape (id in assigned_ids)
        - non-empty substantive text (> 4 chars)
        - a picture, chart, or table
        - a multi-element vector illustration (>= 4 non-text vector shapes)
        """
        if id(grp) in assigned_ids:
            return True
        vector_count = 0
        has_any_text_frame = False
        try:
            for child in grp.shapes:
                if id(child) in assigned_ids:
                    return True
                if child.shape_type == MSO_SHAPE_TYPE.GROUP:
                    if cls._group_has_assigned_or_artwork_recursive(child, assigned_ids):
                        return True
                elif child.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART, MSO_SHAPE_TYPE.TABLE):
                    return True
                elif child.has_text_frame and child.text_frame.text.strip():
                    has_any_text_frame = True
                    if len(child.text_frame.text.strip()) > 4:
                        return True
                elif child.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
                    has_any_text_frame = True
                else:
                    vector_count += 1
        except Exception:
            return True

        # Any group with >= 2 non-text vector shapes is a vector illustration / icon / graphic -> preserve!
        if vector_count >= 2:
            return True
        return False

    @classmethod
    def prune_orphaned_artifacts(
        cls,
        slide: Any,
        pruned_bboxes: List[Tuple[float, float, float, float]],
        assigned_bboxes: List[Tuple[float, float, float, float]],
        assigned_shape_ids: Optional[Set[int]] = None,
        sw_pt: float = 960.0,
        sh_pt: float = 540.0,
    ) -> int:
        if not pruned_bboxes:
            return 0

        assigned_ids = assigned_shape_ids or set()
        pruned_count = 0
        scale_factor = max(1.0, sh_pt / 540.0)

        try:
            shapes_list = list(slide.shapes)
        except Exception:
            return 0

        for s in shapes_list:
            if id(s) in assigned_ids:
                continue

            if cls.is_safe_background_shape(s, sw_pt, sh_pt):
                continue

            # 1. Handle Groups safely (recursing into nested sub-groups and protecting artwork)
            if s.shape_type == MSO_SHAPE_TYPE.GROUP:
                if cls._group_has_assigned_or_artwork_recursive(s, assigned_ids):
                    continue
                g_bb = cls.get_bbox(s)
                min_p_dist = min(cls.bbox_to_bbox_distance(g_bb, p_bb) for p_bb in pruned_bboxes)
                min_a_dist = min((cls.bbox_to_bbox_distance(g_bb, a_bb) for a_bb in assigned_bboxes), default=float("inf"))
                if min_p_dist <= 35.0 * scale_factor and min_p_dist < min_a_dist:
                    if cls.prune_shape(s):
                        pruned_count += 1
                        continue

            # 2. Handle Standalone Decorative Artifacts
            if cls.is_decorative_artifact(s, scale_factor=scale_factor):
                s_bb = cls.get_bbox(s)
                min_pruned_dist = min(cls.bbox_to_bbox_distance(s_bb, p_bb) for p_bb in pruned_bboxes)
                min_assigned_dist = min((cls.bbox_to_bbox_distance(s_bb, a_bb) for a_bb in assigned_bboxes), default=float("inf"))

                if min_pruned_dist <= 30.0 * scale_factor and (min_pruned_dist * 1.5) < min_assigned_dist:
                    if cls.prune_shape(s):
                        pruned_count += 1

        return pruned_count
