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
    "animRot", "animEffect", "cmd", "cBhvr", "audio", "video"
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
    def _has_active_animation_payload(cls, node: etree._Element) -> bool:
        """Tugun ichida kamida bitta tirik animatsiya harakati (cBhvr yoki bldP) borligini tekshiradi."""
        for el in node.iter():
            local = el.tag.split("}")[-1]
            if local in ("cBhvr", "anim", "animEffect", "animMotion", "animClr", "animScale", "animRot", "set", "cmd", "audio", "video"):
                return True
        return False

    @classmethod
    def _prune_empty_containers(cls, node: Optional[etree._Element], root_stop: Optional[etree._Element] = None) -> None:
        """
        ECMA-376 §19.5 ga muvofiq ichida hech qanday animatsiya xulq-atvori (cBhvr) qolmagan
        childTnLst, cTn, par, seq, excl va bldLst konteynerlarini yuqoriga qarab rekursiv tozalaydi.
        """
        curr = node
        while curr is not None and curr != root_stop:
            parent = curr.getparent()
            if parent is None:
                break

            tag = curr.tag.split("}")[-1]
            if tag in (
                "childTnLst", "bldLst", "stCondLst", "endCondLst",
                "prevCondLst", "nextCondLst", "tgtEl", "cond"
            ):
                if len(curr) == 0:
                    parent.remove(curr)
                    curr = parent
                    continue
                break

            if tag in ("cTn", "par", "seq", "excl"):
                if not cls._has_active_animation_payload(curr):
                    parent.remove(curr)
                    curr = parent
                    continue
                break

            break

    @classmethod
    def purge_zombie_animations(cls, slide) -> int:
        """
        <p:timing> ichidagi o'chirilgan shakllarga ishora qiluvchi barcha <p:spTgt>, <p:subSpTgt>
        va <p:bldLst> elementlarini hamda ulardan qolgan bo'sh <p:par>/<p:cTn> qobiqlarini tozalaydi.
        """
        timings = slide.element.xpath('./*[local-name()="timing"]')
        if not timings:
            return 0

        timing_elem = timings[0]
        valid_ids = cls.collect_valid_shape_ids(slide)
        purged_count = 0

        # 1-QADAM: <p:bldLst> ichidagi zombi elementlarni tozalash
        for bld_lst in timing_elem.xpath('.//*[local-name()="bldLst"]'):
            for child in list(bld_lst):
                spid = (child.get("spid") or "").strip()
                if not spid or spid not in valid_ids:
                    bld_lst.remove(child)
                    purged_count += 1
            if len(bld_lst) == 0 and bld_lst.getparent() is not None:
                bld_lst.getparent().remove(bld_lst)

        # 2-QADAM: <p:spTgt>, <p:subSpTgt> va <p:inkTgt> zombi nishonlarini tozalash
        for sp_tgt in list(timing_elem.xpath('.//*[local-name()="spTgt" or local-name()="inkTgt"]')):
            spid = (sp_tgt.get("spid") or "").strip()
            for st in list(sp_tgt.xpath('./*[local-name()="subSpTgt"]')):
                if (st.get("spid") or "").strip() not in valid_ids:
                    sp_tgt.remove(st)
                    purged_count += 1

            if spid and spid in valid_ids:
                continue

            purged_count += 1
            curr = sp_tgt
            target_to_remove = sp_tgt

            while curr is not None and curr != timing_elem:
                tag = curr.tag.split("}")[-1]
                if (tag in ANIM_BEHAVIOR_TAGS and tag != "cBhvr") or tag == "cond":
                    target_to_remove = curr
                    break
                if tag in ("par", "seq"):
                    target_to_remove = curr
                    break
                curr = curr.getparent()

            parent = target_to_remove.getparent()
            if parent is not None:
                parent.remove(target_to_remove)
                cls._prune_empty_containers(parent, root_stop=timing_elem)

        # 3-QADAM: Agar aktiv animatsiya qolmagan bo'lsa, <p:timing> ni butunlay olib tashlash
        if not cls._has_active_animation_payload(timing_elem):
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
