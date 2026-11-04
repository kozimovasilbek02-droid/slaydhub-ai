# -*- coding: utf-8 -*-
"""
engines/slide_pruner.py
Fast OpenXML-level slide trimming engine.
Reduces a 30-50 slide authentic presentation template down to the exact requested slide count (e.g., 5, 8, 10, 15, 20)
while preserving the Cover (Slide 1), Agenda (Slide 2), Diverse Content Archetypes, and Conclusion (Last Slide).
Zero external runtime dependencies beyond python-pptx.
"""

import os
import logging
from typing import List, Set, Optional, Dict, Any
from pptx import Presentation

logger = logging.getLogger(__name__)


class SlidePruner:
    """OpenXML darajasida ortiqcha slaydlarni toza o'chirib, kerakli sonni saqlab qoluvchi dvigatel."""

    def __init__(self, pptx_path: str):
        self.pptx_path = pptx_path
        self.prs = Presentation(pptx_path)

    @property
    def total_slides(self) -> int:
        return len(self.prs.slides)

    def prune_to_count(self, target_count: int, agenda_roles: Optional[List[str]] = None) -> Presentation:
        """
        Shablondagi slaydlarni maqsadli 'target_count' soniga moslab qisqartiradi.
        """
        total = len(self.prs.slides)
        if target_count >= total:
            logger.info(f"Shablon allaqachon {total} ta slayd, qisqartirish shart emas (so'ralgan: {target_count}).")
            return self.prs

        retained_indices = self._calculate_retained_indices(total, target_count)
        self._remove_unwanted_slides(retained_indices)
        logger.info(f"Slaydlar {total} tadan {len(self.prs.slides)} taga qisqartirildi (Target: {target_count}).")
        return self.prs

    def _calculate_retained_indices(self, total: int, target_count: int) -> Set[int]:
        """
        Muqova (0), Reja (1), Xulosa va oraliq arxetiplarni reklama, ikonka to'plami va bo'sh maketlarsiz tanlaydi.
        """
        from core.pptx_text_replacer import AdSlideFilter
        valid_indices = []
        for i in range(total):
            try:
                if not AdSlideFilter.is_ad_or_credit_slide(self.prs.slides[i]):
                    valid_indices.append(i)
            except Exception:
                valid_indices.append(i)

        if not valid_indices:
            valid_indices = list(range(total))

        if target_count >= len(valid_indices):
            return set(valid_indices)

        if target_count <= 1:
            return {valid_indices[0]}
        if target_count == 2:
            return {valid_indices[0], valid_indices[-1]}

        retained: Set[int] = {valid_indices[0], valid_indices[-1]}
        middle_candidates = valid_indices[1:-1]
        needed_middle = target_count - 2

        if needed_middle > 0 and middle_candidates:
            step = max(1, len(middle_candidates) // needed_middle)
            selected = middle_candidates[::step][:needed_middle]
            if len(selected) < needed_middle:
                remaining = [idx for idx in middle_candidates if idx not in selected]
                selected.extend(remaining[: (needed_middle - len(selected))])
            retained.update(selected)

        return retained

    def _remove_unwanted_slides(self, retained_indices: Set[int]) -> None:
        """
        OpenXML sldIdLst elementlarini orqadan oldinga qarab o'chiradi, rId bog'lamalarini
        (drop_rel) uzadi hamda p14:sectionLst va custShowLst dagi zombi sldId larni tozalaydi.
        """
        sldIdLst = self.prs.slides._sldIdLst
        total = len(self.prs.slides)
        removed_ids: Set[str] = set()
        removed_rids: Set[str] = set()

        for i in range(total - 1, -1, -1):
            if i not in retained_indices:
                try:
                    sldId = sldIdLst[i]
                    sid_val = sldId.get("id")
                    rId = sldId.rId
                    if sid_val:
                        removed_ids.add(str(sid_val))
                    if rId:
                        removed_rids.add(str(rId))
                    sldIdLst.remove(sldId)
                    if rId:
                        self.prs.part.drop_rel(rId)
                except Exception as exc:
                    logger.warning(f"Slayd {i} ni o'chirishda xatolik: {exc}")

        # presentation.xml ichidagi sectionLst va custShowLst dan o'chirilgan slaydlarni tozalash
        try:
            pres_elm = self.prs._element
            for sec_sld in pres_elm.xpath('.//*[local-name()="section"]//*[local-name()="sldId"]'):
                if str(sec_sld.get("id", "")) in removed_ids and sec_sld.getparent() is not None:
                    sec_sld.getparent().remove(sec_sld)
            for cust_sld in pres_elm.xpath('.//*[local-name()="custShow"]//*[local-name()="sld"]'):
                r_id = cust_sld.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if r_id in removed_rids and cust_sld.getparent() is not None:
                    cust_sld.getparent().remove(cust_sld)
        except Exception as exc:
            logger.debug(f"sectionLst/custShowLst tozalash ogohlantirish: {exc}")

    def save(self, out_path: str) -> str:
        self.prs.save(out_path)
        return out_path


def prune_presentation_file(src_pptx_path: str, target_count: int, out_pptx_path: str) -> str:
    """Qulay funksiya: Faylni oladi, qisqartiradi va yangi manzilga saqlaydi."""
    pruner = SlidePruner(src_pptx_path)
    pruner.prune_to_count(target_count)
    return pruner.save(out_pptx_path)
