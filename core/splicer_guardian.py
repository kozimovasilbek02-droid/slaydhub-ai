# -*- coding: utf-8 -*-
"""
core/splicer_guardian.py
Cross-Deck Slide Splicing & Safe Transplanter Guardian (Dual-Brain Round 14).
Resolves Dangling "rId" pointers, Shape ID Collisions, and safely maps media
between fragmented OpenXML slide parts during slide duplication or merging.
"""

import copy
import logging
from typing import Any, Dict, Optional

logger = logging.getLogger("SplicerGuardian")

# Common OpenXML namespaces
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class SplicerGuardian:
    """
    Safely clones slides or shapes while remapping all associated
    image/media relationships and rewriting unique geometric IDs.
    """

    @classmethod
    def _remap_relationships(
        cls,
        copied_elem: Any,
        source_part: Any,
        target_part: Any,
        rel_map: Optional[Dict[str, str]] = None
    ) -> None:
        """
        Deep scans an OpenXML element and remaps all r:embed, r:id, and r:link
        pointers so they point to proper relationships in the target_part.
        Uses rel_map cache to prevent duplicate relationship registrations.
        Safely strips unresolvable dangling rIds to eliminate PowerPoint repair prompts.
        """
        if rel_map is None:
            rel_map = {}

        attr_tags = [f"{{{R_NS}}}embed", f"{{{R_NS}}}id", f"{{{R_NS}}}link"]

        for elem in copied_elem.iter():
            for attr in attr_tags:
                rid = elem.get(attr)
                if not rid:
                    continue

                if rid in rel_map:
                    elem.set(attr, rel_map[rid])
                    continue

                # Locate relationship in source_part, slide_layout, or slide_master
                source_rel = None
                if hasattr(source_part, "rels") and rid in source_part.rels:
                    source_rel = source_part.rels[rid]
                elif hasattr(source_part, "slide_layout") and hasattr(source_part.slide_layout, "part") and rid in source_part.slide_layout.part.rels:
                    source_rel = source_part.slide_layout.part.rels[rid]
                elif hasattr(source_part, "slide_master") and hasattr(source_part.slide_master, "part") and rid in source_part.slide_master.part.rels:
                    source_rel = source_part.slide_master.part.rels[rid]

                if source_rel is not None:
                    try:
                        if source_rel.is_external:
                            new_rId = target_part.relate_to(source_rel.target_ref, source_rel.reltype, is_external=True)
                        else:
                            new_rId = target_part.relate_to(source_rel.target_part, source_rel.reltype, is_external=False)

                        rel_map[rid] = new_rId
                        elem.set(attr, new_rId)
                    except Exception as e:
                        logger.debug(f"Failed to relate {rid}: {e}")
                        # Clean dangling attribute to prevent corrupt XML
                        del elem.attrib[attr]
                else:
                    # Dangling rId not found in source relationships -> strip to prevent PowerPoint repair crash
                    logger.debug(f"Stripping unresolvable dangling {attr}='{rid}'")
                    del elem.attrib[attr]

    @classmethod
    def safe_clone_shape_with_rels(
        cls,
        shape_elem: Any,
        source_part: Any,
        target_part: Any,
        rel_map: Optional[Dict[str, str]] = None
    ) -> Any:
        """
        Copies an XML shape node, fixes its nested relationship dependencies (like pictures and blipFills),
        and readies it for attachment to a new slide part.
        """
        cloned_elem = copy.deepcopy(shape_elem)
        cls._remap_relationships(cloned_elem, source_part, target_part, rel_map=rel_map)
        return cloned_elem

    @classmethod
    def assign_unique_shape_ids(cls, slide) -> None:
        """
        Scans all graphic frame identifiers (<p:cNvPr id="...">) uniformly across a slide.
        In case of duplicate shape IDs, regenerates non-colliding incremental IDs starting from max_id + 1.
        Crucial for avoiding animation crashes and grouping selection bugs.
        """
        try:
            tree = slide.element
            existing_ids = []

            # Extract all current IDs to find a safe boundary
            cnvprs = tree.xpath('.//p:cNvPr')
            for cNvPr in cnvprs:
                try:
                    val = int(cNvPr.get('id', '0'))
                    existing_ids.append(val)
                except ValueError:
                    pass

            max_id = max(max(existing_ids, default=1), 1)
            seen = set()
            id_map: Dict[str, str] = {}

            for cNvPr in cnvprs:
                try:
                    sid = int(cNvPr.get('id', '0'))
                    if sid in seen or sid <= 0:
                        max_id += 1
                        old_str = str(sid)
                        new_str = str(max_id)
                        cNvPr.set('id', new_str)
                        seen.add(max_id)
                        if sid > 0:
                            id_map[old_str] = new_str
                    else:
                        seen.add(sid)
                except ValueError:
                    max_id += 1
                    cNvPr.set('id', str(max_id))
                    seen.add(max_id)

            # Remap animation target spids if shapes were renumbered
            if id_map:
                try:
                    from core.animation_guardian import AnimationGuardian
                    AnimationGuardian.remap_animation_target_ids(slide, id_map)
                except Exception as e:
                    logger.debug(f"Animation remapping error: {e}")
        except Exception as e:
            logger.debug(f"Failed to assign unique IDs: {e}")

    @classmethod
    def replicate_slide_background(cls, source_slide, target_slide) -> None:
        """
        Copies the background markup (<p:bg>) exactly from the source to target slide.
        If source slide has no explicit <p:bg>, inspects its slide_layout and slide_master.
        This fixes the problem where dynamic continuation slides lose the author's original
        gradient, pattern, or heavily styled texture backgrounds.
        """
        try:
            source_csld = source_slide.element.xpath('./p:cSld')
            target_csld = target_slide.element.xpath('./p:cSld')

            if not source_csld or not target_csld:
                return

            # 1. Safely fetch source background element
            source_bg = source_csld[0].xpath('./p:bg')
            bg_source_part = source_slide.part

            # 2. Fallback: inspect layout or master if slide has no direct <p:bg>
            if not source_bg and hasattr(source_slide, "slide_layout"):
                layout_csld = source_slide.slide_layout.element.xpath('./p:cSld')
                if layout_csld:
                    source_bg = layout_csld[0].xpath('./p:bg')
                    bg_source_part = source_slide.slide_layout.part

            if not source_bg and hasattr(source_slide, "slide_layout") and hasattr(source_slide.slide_layout, "slide_master"):
                master_csld = source_slide.slide_layout.slide_master.element.xpath('./p:cSld')
                if master_csld:
                    source_bg = master_csld[0].xpath('./p:bg')
                    bg_source_part = source_slide.slide_layout.slide_master.part

            if not source_bg:
                return

            cloned_bg = cls.safe_clone_shape_with_rels(source_bg[0], bg_source_part, target_slide.part)

            target_bg = target_csld[0].xpath('./p:bg')
            if target_bg:
                # Replace existing background
                target_csld[0].replace(target_bg[0], cloned_bg)
            else:
                # Insert at the very top of cSld to properly render behind internal elements
                target_csld[0].insert(0, cloned_bg)

        except Exception as e:
            logger.debug(f"Failed to replicate background: {e}")
