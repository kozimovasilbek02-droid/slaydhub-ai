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
        Muqova (0), Reja (1), Xulosa (total-1) va oraliq arxetiplarni proporsional tanlaydi.
        """
        if target_count <= 1:
            return {0}
        if target_count == 2:
            return {0, total - 1}

        retained: Set[int] = {0, total - 1}

        # Agar kamida 4 ta slayd so'ralsa, 2-slaydni (odatda Agenda/Reja) saqlashga harakat qilamiz
        middle_candidates = list(range(1, total - 1))
        needed_middle = target_count - 2

        if needed_middle > 0 and middle_candidates:
            # Oraliq slaydlardan teng qadam bilan eng xilma-xil maketlarni tanlaymiz
            step = max(1, len(middle_candidates) // needed_middle)
            selected = middle_candidates[::step][:needed_middle]
            
            # Agar bo'sh o'rin qolsa, qolganlaridan to'ldiramiz
            if len(selected) < needed_middle:
                remaining = [idx for idx in middle_candidates if idx not in selected]
                selected.extend(remaining[: (needed_middle - len(selected))])

            retained.update(selected)

        return retained

    def _remove_unwanted_slides(self, retained_indices: Set[int]) -> None:
        """
        OpenXML sldIdLst ichidagi sldId elementlarini orqadan oldinga (reverse) qarab o'chiradi.
        Bu indeks siljishi (index shift) muammolarini 100% oldini oladi.
        """
        sldIdLst = self.prs.slides._sldIdLst
        total = len(self.prs.slides)

        for i in range(total - 1, -1, -1):
            if i not in retained_indices:
                try:
                    sldIdLst.remove(sldIdLst[i])
                except Exception as exc:
                    logger.warning(f"Slayd {i} ni o'chirishda xatolik: {exc}")

    def save(self, out_path: str) -> str:
        self.prs.save(out_path)
        return out_path


def prune_presentation_file(src_pptx_path: str, target_count: int, out_pptx_path: str) -> str:
    """Qulay funksiya: Faylni oladi, qisqartiradi va yangi manzilga saqlaydi."""
    pruner = SlidePruner(src_pptx_path)
    pruner.prune_to_count(target_count)
    return pruner.save(out_pptx_path)
