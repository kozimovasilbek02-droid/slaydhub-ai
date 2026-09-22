# -*- coding: utf-8 -*-
"""
core/slot_allocator.py
Unified Content Density & Slot Capacity Scoring Engine for SlaydHub AI.
Dual-Brain Collaborative Module (Antigravity & Claude Code).

Features:
1. SlotCapacityScorer (Claude): Measures shape area (>3000 pt²), filters micro-badges, scores templates.
2. Template Matcher (Claude): Optimal slot-count matching (1000+ points for exact match).
3. Polymorphic Clustering (Claude + Antigravity): Safely merges overflow points into pairs without KeyError.
4. Continuation Slide Engine (Antigravity): Seamlessly splits dense slides into Part 1/Part 2 slides.
5. Deck Balancer: Integrates with TemplateDeckAssembler in core/pptx_text_replacer.py.
"""

import math
import copy
import logging
from typing import List, Dict, Any, Tuple, Optional, Union

logger = logging.getLogger("SlotAllocator")


class SlotAllocator:
    """
    Unified Slot Allocator & Content Density Optimizer.
    """

    DEFAULT_MIN_SLOT_CAPACITY = 3
    DEFAULT_MAX_SLOT_CAPACITY = 6
    MIN_SLOT_AREA_PT2 = 3000.0

    @classmethod
    def _extract_item_pair(cls, item: Any) -> Tuple[str, str]:
        """
        Polymorphic item extractor. Safely handles dicts, tuples, objects, or strings.
        Prevents KeyError: 0 when cards are formatted as dictionaries.
        """
        if isinstance(item, dict):
            title = str(item.get("title") or item.get("header") or item.get("name") or "")
            desc = str(item.get("description") or item.get("desc") or item.get("content") or item.get("text") or "")
            return title, desc
        elif isinstance(item, (tuple, list)):
            title = str(item[0]) if len(item) > 0 else ""
            desc = str(item[1]) if len(item) > 1 else ""
            return title, desc
        else:
            return str(item), ""

    @classmethod
    def calculate_slide_capacity(cls, text_shapes: List[Any], sw_pt: float = 960.0, sh_pt: float = 540.0) -> Dict[str, Any]:
        """
        Claude's Capacity Scoring Engine:
        Evaluates a template slide to determine how many independent content slots it has,
        filtering out microscopic badges and headers.
        """
        valid_slots = []
        total_area = 0.0

        for sp in text_shapes:
            try:
                # Use native python-pptx pt conversion if available, else EMU to pt (1 pt = 12700 EMU)
                w = getattr(sp.width, "pt", float(sp.width) / 12700.0) if hasattr(sp, "width") else 0.0
                h = getattr(sp.height, "pt", float(sp.height) / 12700.0) if hasattr(sp, "height") else 0.0
                area = w * h

                # Skip small shapes (badges, numbers, icons)
                if area < cls.MIN_SLOT_AREA_PT2:
                    continue

                # Skip header band (top 15% of slide)
                top_pt = getattr(sp.top, "pt", float(sp.top) / 12700.0) if hasattr(sp, "top") else 0.0
                if top_pt < (sh_pt * 0.15):
                    continue

                valid_slots.append(sp)
                total_area += area
            except Exception:
                continue

        slot_count = len(valid_slots)
        return {
            "slot_count": slot_count,
            "total_area_pt2": total_area,
            "avg_slot_area": (total_area / slot_count) if slot_count > 0 else 0.0
        }

    @classmethod
    def match_template(cls, content_items: List[Any], templates: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Claude's Template Matcher:
        Scores templates and picks the best one where template slot_count >= len(content_items).
        """
        req_slots = len(content_items)
        best_match = None
        best_score = -float("inf")

        for tpl in templates:
            cap = tpl.get("capacity", {})
            t_slots = cap.get("slot_count", 0)
            if t_slots == 0:
                continue

            if t_slots == req_slots:
                score = 1000 + cap.get("avg_slot_area", 0)
            elif t_slots > req_slots:
                score = 500 - (t_slots - req_slots) * 10
            else:
                score = -1000 * (req_slots - t_slots)

            if score > best_score:
                best_score = score
                best_match = tpl

        needs_spillover = best_match["capacity"]["slot_count"] < req_slots if best_match and "capacity" in best_match else True

        return {
            "template": best_match,
            "needs_spillover": needs_spillover,
            "available_slots": best_match["capacity"]["slot_count"] if best_match and "capacity" in best_match else 0
        }

    @classmethod
    def cluster_content(cls, content_items: List[Any], max_slots: int) -> List[Dict[str, str]]:
        """
        Claude's Semantic Clustering (Hardened against KeyError):
        Groups dense items into pairs while safely handling polymorphic item types (dicts, tuples, strings).
        """
        if len(content_items) <= max_slots or max_slots <= 0:
            return content_items

        chunk_size = math.ceil(len(content_items) / max_slots)
        clustered: List[Dict[str, str]] = []

        for i in range(0, len(content_items), chunk_size):
            chunk = content_items[i:i + chunk_size]
            pairs = [cls._extract_item_pair(c) for c in chunk]

            # Merge headings and descriptions cleanly
            merged_heading = " & ".join([p[0] for p in pairs if p[0]])
            desc_lines = []
            for p in pairs:
                if p[0] and p[1]:
                    desc_lines.append(f"• {p[0]}:\n  {p[1]}")
                elif p[1]:
                    desc_lines.append(f"• {p[1]}")
                elif p[0]:
                    desc_lines.append(f"• {p[0]}")
            merged_desc = "\n\n".join(desc_lines)

            clustered.append({"title": merged_heading, "description": merged_desc})

        return clustered[:max_slots]

    @classmethod
    def paginate_content(cls, content_items: List[Any], max_slots: int) -> List[List[Any]]:
        """
        Claude's Pagination Engine:
        Splits dense items into pages of size max_slots.
        """
        if max_slots <= 0 or len(content_items) <= max_slots:
            return [content_items]

        pages = []
        for i in range(0, len(content_items), max_slots):
            pages.append(content_items[i:i + max_slots])
        return pages

    @classmethod
    def count_content_items(cls, slide_data: Dict[str, Any]) -> int:
        """
        Counts the logical content items (cards, steps, bullet points, takeaways) in slide_data.
        """
        for key in ["cards", "items", "points", "steps", "key_takeaways", "pillars", "milestones"]:
            val = slide_data.get(key)
            if isinstance(val, list) and len(val) > 0:
                return len(val)

        left_pts = slide_data.get("left_points", [])
        right_pts = slide_data.get("right_points", [])
        if left_pts or right_pts:
            return len(left_pts) + len(right_pts)

        content = slide_data.get("content", "")
        if isinstance(content, list):
            return len(content)
        elif isinstance(content, str) and content.strip():
            lines = [l.strip() for l in content.split("\n") if l.strip()]
            bullet_lines = [l for l in lines if l.startswith(("-", "*", "•", "1.", "2.", "3.", "4."))]
            if len(bullet_lines) >= 2:
                return len(bullet_lines)
            return max(1, len(lines))

        return 1

    @classmethod
    def partition_into_continuation_slides(
        cls,
        slide_data: Dict[str, Any],
        max_slots_per_slide: int = 4
    ) -> List[Dict[str, Any]]:
        """
        Antigravity's Continuation Slide Generator:
        Splits a single dense slide into two or more sequential continuation slides
        with numbered sub-titles ("Part 1 of 2", "Part 2 of 2").
        """
        target_key = None
        for key in ["cards", "items", "points", "steps", "key_takeaways", "pillars", "milestones"]:
            if isinstance(slide_data.get(key), list) and len(slide_data[key]) > 0:
                target_key = key
                break

        if not target_key:
            return [slide_data]

        all_items = slide_data[target_key]
        total_items = len(all_items)

        if total_items <= max_slots_per_slide:
            return [slide_data]

        num_slides = (total_items + max_slots_per_slide - 1) // max_slots_per_slide
        chunk_size = (total_items + num_slides - 1) // num_slides

        result_slides = []
        base_title = slide_data.get("title", "Taqdimot Slaydi")

        for part_idx in range(num_slides):
            start_i = part_idx * chunk_size
            end_i = min(start_i + chunk_size, total_items)
            part_items = all_items[start_i:end_i]

            if not part_items:
                continue

            new_slide = copy.deepcopy(slide_data)
            new_slide[target_key] = part_items
            new_slide["title"] = f"{base_title} ({part_idx + 1}/{num_slides}-Qism)"
            new_slide["is_continuation"] = True
            new_slide["continuation_part"] = part_idx + 1
            new_slide["continuation_total"] = num_slides

            result_slides.append(new_slide)

        logger.info(f"SlotAllocator: Partitioned '{base_title}' ({total_items} items) into {len(result_slides)} continuation slides.")
        return result_slides

    @classmethod
    def balance_deck(
        cls,
        slides_data: List[Dict[str, Any]],
        default_max_slots: int = 4
    ) -> List[Dict[str, Any]]:
        """
        Deck Balancer for TemplateDeckAssembler:
        Scans all slides and expands dense ones into continuation slides.
        """
        balanced = []
        for s_item in slides_data:
            item_count = cls.count_content_items(s_item)
            if item_count > default_max_slots and not s_item.get("is_cover", False) and s_item.get("layout_type") not in ["table_data", "cover"]:
                parts = cls.partition_into_continuation_slides(s_item, max_slots_per_slide=default_max_slots)
                balanced.extend(parts)
            else:
                balanced.append(s_item)

        logger.info(f"SlotAllocator: Balanced deck from {len(slides_data)} to {len(balanced)} slides.")
        return balanced
