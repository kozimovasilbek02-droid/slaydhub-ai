# -*- coding: utf-8 -*-
"""
core/animation_guardian.py
Zombie Animation Target & Timing Tree Purger (Dual-Brain Round 15).
Detects and purges dangling animation behaviors, time nodes, and build lists
pointing to pruned or deleted shapes in OpenXML <p:timing> (§19.5).
Eliminates PowerPoint slideshow repair prompts, frozen animation clicks,
and OpenXML timing schema corruption.
"""

import logging
from typing import Set, Dict, Optional, List
import lxml.etree as etree

logger = logging.getLogger("AnimationGuardian")

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Standard animation behavior tags in DrawingML/PresentationML
ANIM_BEHAVIOR_TAGS = {
    "set", "anim", "animMotion", "animClr", "animScale",
    "animRot", "animEffect", "cmd", "cBhvr"
}


class AnimationGuardian:
    """
    Guarantees OpenXML timing tree health by purging zombie animation targets
    and remapping animation shape identifiers across slides.
    """

    @classmethod
    def collect_valid_shape_ids(cls, slide) -> Set[str]:
        """
        Extracts the complete set of existing, active shape IDs across the slide,
        including shapes inside group shapes, tables, charts, and connectors.
        Every valid shape in OpenXML has a <p:cNvPr id="..."> element.
        """
        valid_ids: Set[str] = set()
        try:
            tree = slide.element
            # Matches all cNvPr elements across normal shapes, pictures, groups, connectors, etc.
            cnvprs = tree.xpath('.//*[local-name()="cNvPr"]')
            for elem in cnvprs:
                sid = elem.get('id')
                if sid:
                    valid_ids.add(str(sid).strip())
        except Exception as e:
            logger.debug(f"Failed to collect valid shape IDs: {e}")
        return valid_ids

    @classmethod
    def _prune_empty_containers(cls, node: etree._Element) -> None:
        """
        Recursively cleans up empty container elements (childTnLst, par, seq, cTn)
        whose children were removed during zombie animation pruning.
        """
        if node is None or len(node) > 0:
            return

        parent = node.getparent()
        if parent is None:
            return

        tag = node.tag.split("}")[-1]
        if tag in ("childTnLst", "bldLst", "stCondLst", "endCondLst"):
            parent.remove(node)
            # Recurse up to check if parent became empty
            cls._prune_empty_containers(parent)
        elif tag in ("par", "seq", "cTn") and len(parent) == 1 and parent.tag.split("}")[-1] == "childTnLst":
            parent.remove(node)
            cls._prune_empty_containers(parent)

    @classmethod
    def purge_zombie_animations(cls, slide) -> int:
        """
        Scans <p:timing> and purges all animation behaviors, time nodes, and build lists
        targeting shapes that no longer exist on the slide.
        If all animations were removed, removes <p:timing> to return the slide to a clean state.
        Returns the number of purged zombie animation nodes.
        """
        timings = slide.element.xpath('./*[local-name()="timing"]')
        if not timings:
            return 0

        timing_elem = timings[0]
        valid_ids = cls.collect_valid_shape_ids(slide)
        purged_count = 0

        # --- STEP 1: Purge Zombie Build List Items (<p:bldLst>) ---
        bld_lists = timing_elem.xpath('.//*[local-name()="bldLst"]')
        for bld_lst in bld_lists:
            for child in list(bld_lst):
                spid = child.get('spid')
                if spid and spid.strip() not in valid_ids:
                    logger.debug(f"Purging zombie bldLst item targeting spid='{spid}'")
                    bld_lst.remove(child)
                    purged_count += 1
            if len(bld_lst) == 0:
                parent = bld_lst.getparent()
                if parent is not None:
                    parent.remove(bld_lst)

        # --- STEP 2: Purge Zombie Animation Targets (<p:spTgt>) ---
        sp_tgts = timing_elem.xpath('.//*[local-name()="spTgt"]')
        for sp_tgt in sp_tgts:
            spid = sp_tgt.get('spid')
            if not spid or spid.strip() in valid_ids:
                continue

            purged_count += 1
            logger.debug(f"Found zombie animation target spid='{spid}'")

            # Locate the enclosing animation behavior or time node
            # Hierarchy: spTgt -> tgtEl -> cBhvr -> [set, anim, animMotion, ...] -> childTnLst -> cTn -> par
            curr = sp_tgt
            target_to_remove = None

            while curr is not None and curr != timing_elem:
                tag = curr.tag.split("}")[-1]
                if tag in ANIM_BEHAVIOR_TAGS and tag != "cBhvr":
                    target_to_remove = curr
                    break
                elif tag in ("par", "seq"):
                    target_to_remove = curr
                    break
                curr = curr.getparent()

            if target_to_remove is not None:
                parent = target_to_remove.getparent()
                if parent is not None:
                    parent.remove(target_to_remove)
                    cls._prune_empty_containers(parent)

        # --- STEP 3: Clean Root Timing if Zero Active Animations Remain ---
        # Check if any animated behaviors or targets remain
        remaining_targets = timing_elem.xpath('.//*[local-name()="spTgt"]')
        remaining_behaviors = timing_elem.xpath('.//*[local-name()="cBhvr"]')
        remaining_blds = timing_elem.xpath('.//*[@spid]')

        if len(remaining_targets) == 0 and len(remaining_behaviors) == 0 and len(remaining_blds) == 0:
            logger.debug("No active animations remaining in <p:timing>. Safely removing empty timing node.")
            slide.element.remove(timing_elem)

        return purged_count

    @classmethod
    def remap_animation_target_ids(cls, slide, id_map: Dict[str, str]) -> int:
        """
        Updates animation target identifiers (spid="old_id" -> spid="new_id")
        when shapes are renumbered (e.g. during cross-deck splicing or collision avoidance).
        Returns the number of remapped targets.
        """
        if not id_map:
            return 0

        timings = slide.element.xpath('./*[local-name()="timing"]')
        if not timings:
            return 0

        timing_elem = timings[0]
        remapped_count = 0

        for elem in timing_elem.xpath('.//*[@spid]'):
            spid = elem.get('spid')
            if spid in id_map:
                new_id = str(id_map[spid])
                elem.set('spid', new_id)
                remapped_count += 1

        return remapped_count
