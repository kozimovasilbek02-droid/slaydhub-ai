# -*- coding: utf-8 -*-
"""
core/slide_splicer.py
Multi-Deck Slide Splicing & Assembly Engines for SlaydHub AI.
Implements 3 Distinct Splicing Engines:
1. HarmonizedSlideCloner: Deep XML cloning with Theme Harmonizer & resolution scaling (colors, fonts & coordinates match target deck).
2. MultiMasterSlideSplicer: Preserves each slide's native background, textures, and master styling.
3. ModularComponentStacker: Injects high-impact graphic & infographic components into base deck.
"""

import os
import tempfile
from pathlib import Path
import io
import copy
import logging
from typing import Dict, Any, List, Optional, Tuple

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

from core.splicer_guardian import SplicerGuardian

logger = logging.getLogger("SlideSplicer")


class ThemeColorExtractor:
    """
    Extracts dominant color palette and font pairings from a template slide (recursing into groups).
    """

    @classmethod
    def _walk_shapes(cls, shapes_container):
        for shp in shapes_container:
            if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from cls._walk_shapes(shp.shapes)
            else:
                yield shp

    @classmethod
    def extract_palette(cls, slide) -> Dict[str, Any]:
        fonts = []
        colors = []
        shape_colors = []

        for shp in cls._walk_shapes(slide.shapes):
            if shp.has_text_frame:
                for p in shp.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.name:
                            fonts.append(r.font.name)
                        try:
                            if r.font.color and r.font.color.type and r.font.color.rgb:
                                colors.append(r.font.color.rgb)
                        except Exception:
                            pass
            try:
                srgb_nodes = shp._element.xpath('./p:spPr/a:solidFill/a:srgbClr')
                for node in srgb_nodes:
                    val = node.get('val', '')
                    if len(val) == 6 and val.upper() not in ('FFFFFF', '000000', 'F8FAFC', '191919', '1B1B1B'):
                        shape_colors.append(RGBColor(int(val[0:2], 16), int(val[2:4], 16), int(val[4:6], 16)))
            except Exception:
                pass

        primary_font = fonts[0] if fonts else "Arial"
        primary_color = shape_colors[0] if shape_colors else (colors[0] if colors else RGBColor(0x1F, 0x4E, 0x79))
        accent_color = shape_colors[1] if len(shape_colors) > 1 else (colors[1] if len(colors) > 1 else RGBColor(0x02, 0x84, 0xC7))

        return {
            "primary_font": primary_font,
            "primary_color": primary_color,
            "accent_color": accent_color,
        }


def get_cleanest_blank_layout(prs: Presentation):
    """Finds the layout with 0 placeholders or fewest placeholders to prevent ghost prompts."""
    for layout in prs.slide_layouts:
        name = layout.name.lower()
        if ("blank" in name or "пустой" in name) and len(layout.placeholders) == 0:
            return layout
    for layout in prs.slide_layouts:
        if len(layout.placeholders) == 0:
            return layout
    return min(prs.slide_layouts, key=lambda l: len(l.placeholders))


class HarmonizedSlideCloner:
    """
    Engine 1: Deep XML Shape Cloning with Master Theme Harmonization & Resolution Scaling.
    Clones vector shapes, diagrams, and layout blocks from candidate templates,
    scales coordinates when source and target slide dimensions differ (e.g. 13.3x7.5 -> 26.7x15.0),
    and harmonizes font families and accent colors to match the cover slide.
    """

    @staticmethod
    def _scale_cloned_element_geometry(elem_copy, sx: float, sy: float) -> None:
        """Scales outer xfrm (off x,y and ext cx,cy) and font sizes when slide dimensions differ."""
        if abs(sx - 1.0) <= 0.03 and abs(sy - 1.0) <= 0.03:
            return
        try:
            # 1. Scale outer spPr/xfrm or grpSpPr/xfrm (do NOT scale inner chOff/chExt of groups)
            outer_xfrms = elem_copy.xpath('./p:spPr/a:xfrm | ./p:grpSpPr/a:xfrm | ./p:xfrm')
            for xfrm in outer_xfrms:
                for off in xfrm.xpath('./a:off'):
                    if 'x' in off.attrib:
                        off.set('x', str(int(round(int(off.get('x', '0')) * sx))))
                    if 'y' in off.attrib:
                        off.set('y', str(int(round(int(off.get('y', '0')) * sy))))
                for ext in xfrm.xpath('./a:ext'):
                    if 'cx' in ext.attrib:
                        ext.set('cx', str(int(round(int(ext.get('cx', '0')) * sx))))
                    if 'cy' in ext.attrib:
                        ext.set('cy', str(int(round(int(ext.get('cy', '0')) * sy))))

            # 2. Scale font sizes (a:rPr and a:defRPr) by sy
            for rpr in elem_copy.xpath('.//a:rPr | .//a:defRPr | .//a:endParaRPr'):
                if 'sz' in rpr.attrib:
                    try:
                        old_sz = int(rpr.get('sz'))
                        rpr.set('sz', str(int(round(old_sz * sy))))
                    except Exception:
                        pass
        except Exception as e:
            logger.debug(f"Geometry scaling warning: {e}")

    @classmethod
    def clone_slide_to_target(
        cls,
        source_slide,
        target_slide,
        theme_palette: Optional[Dict[str, Any]] = None,
        replicate_background: bool = True,
    ) -> None:
        """
        Clones all shapes, vector artwork, diagrams, and pictures from source_slide to target_slide.
        """
        # Clear any existing placeholder shapes on target slide to prevent ghost prompts
        for old_s in list(target_slide.shapes):
            try:
                target_slide.shapes._spTree.remove(old_s._element)
            except Exception:
                pass
        for p in list(target_slide.placeholders):
            try:
                target_slide.shapes._spTree.remove(p._element)
            except Exception:
                pass

        try:
            from core.animation_guardian import AnimationGuardian
            AnimationGuardian.purge_zombie_animations(target_slide)
        except Exception:
            pass

        # 1. Replicate background only when requested (for cross-deck donors, keep target deck's native background)
        if replicate_background:
            SplicerGuardian.replicate_slide_background(source_slide, target_slide)

        # Compute resolution scale factors between source and target presentations
        sx, sy = 1.0, 1.0
        try:
            src_prs = source_slide.part.package.presentation_part.presentation
            tgt_prs = target_slide.part.package.presentation_part.presentation
            if src_prs.slide_width and tgt_prs.slide_width:
                sx = float(tgt_prs.slide_width) / float(src_prs.slide_width)
            if src_prs.slide_height and tgt_prs.slide_height:
                sy = float(tgt_prs.slide_height) / float(src_prs.slide_height)
        except Exception:
            pass

        # 2. Clone each shape safely with full media relationship remapping
        rel_map = {}
        for shp in source_slide.shapes:
            try:
                elem_copy = SplicerGuardian.safe_clone_shape_with_rels(
                    shp._element,
                    source_slide.part,
                    target_slide.part,
                    rel_map=rel_map,
                )

                if hasattr(elem_copy, "spPr"):
                    spPr = elem_copy.spPr
                    if spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm") is None:
                        if shp.left is not None and shp.top is not None and shp.width is not None and shp.height is not None:
                            xfrm_elem = parse_xml(
                                f'<a:xfrm {nsdecls("a")}><a:off x="{shp.left}" y="{shp.top}"/><a:ext cx="{shp.width}" cy="{shp.height}"/></a:xfrm>'
                            )
                            spPr.insert(0, xfrm_elem)

                cls._scale_cloned_element_geometry(elem_copy, sx, sy)
                target_slide.shapes._spTree.append(elem_copy)
            except Exception as e:
                logger.debug(f"Could not clone shape {shp.name}: {e}")

        # 3. Re-assign unique shape IDs across target slide to prevent ID collision corruption
        SplicerGuardian.assign_unique_shape_ids(target_slide)

        # 4. Apply Theme Harmonization if palette is provided
        if theme_palette:
            cls.apply_theme_harmony(target_slide, theme_palette)

    @classmethod
    def apply_theme_harmony(cls, slide, theme_palette: Dict[str, Any]) -> None:
        """
        Harmonizes text font families and key accent colors across cloned shapes (including groups).
        """
        target_font = theme_palette.get("primary_font", "Arial")
        target_color = theme_palette.get("primary_color")

        for shp in ThemeColorExtractor._walk_shapes(slide.shapes):
            if not shp.has_text_frame:
                continue
            for p in shp.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name:
                        r.font.name = target_font
                    if target_color and r.font.bold:
                        try:
                            # Preserve white/light text inside colored step badges (luminance > 210)
                            cur_rgb = r.font.color.rgb if (r.font.color and r.font.color.type) else None
                            if cur_rgb and (0.299 * cur_rgb[0] + 0.587 * cur_rgb[1] + 0.114 * cur_rgb[2]) > 210.0:
                                pass
                            else:
                                r.font.color.rgb = target_color
                        except Exception:
                            pass


class MultiMasterSlideSplicer:
    """
    Engine 2: Multi-Master Presentation Splicing.
    Assembles a dynamic presentation where each slide keeps its original
    Slide Master, background textures, and distinctive graphic artwork.
    """

    @classmethod
    def assemble_multi_master_deck(
        cls,
        base_prs: Presentation,
        slide_specs: List[Dict[str, Any]],
        output_path: str,
    ) -> str:
        for del_idx in reversed(range(1, len(base_prs.slides))):
            try:
                rId = base_prs.slides._sldIdLst[del_idx].rId
                base_prs.part.drop_rel(rId)
                del base_prs.slides._sldIdLst[del_idx]
            except Exception:
                pass

        blank_layout = get_cleanest_blank_layout(base_prs)
        specs_to_splice = (
            slide_specs[1:]
            if slide_specs and (slide_specs[0].get("slide_number") == 1 or slide_specs[0].get("archetype") == "cover")
            else slide_specs
        )

        for spec in specs_to_splice:
            src_pptx = spec.get("pptx_path")
            src_idx = spec.get("slide_index", 0)

            if not src_pptx or not os.path.exists(src_pptx):
                base_prs.slides.add_slide(blank_layout)
                continue

            try:
                src_prs = Presentation(src_pptx)
                if src_idx < len(src_prs.slides):
                    src_slide = src_prs.slides[src_idx]
                    new_slide = base_prs.slides.add_slide(blank_layout)
                    HarmonizedSlideCloner.clone_slide_to_target(src_slide, new_slide, theme_palette=None)
                else:
                    base_prs.slides.add_slide(blank_layout)
            except Exception as e:
                logger.warning(f"Error splicing slide from {src_pptx}: {e}")
                base_prs.slides.add_slide(blank_layout)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        base_prs.save(output_path)
        return output_path


class ModularComponentStacker:
    """
    Engine 3: Modular Infographic Component Injection.
    """

    @classmethod
    def inject_components(
        cls,
        target_slide,
        source_component_slide,
        vertical_offset_pt: float = 0.0,
    ) -> None:
        for shp in source_component_slide.shapes:
            is_title = False
            if shp.has_text_frame:
                if "title" in shp.name.lower() or shp.top < Inches(1.2).emu:
                    is_title = True

            if is_title:
                continue

            try:
                elem_copy = SplicerGuardian.safe_clone_shape_with_rels(
                    shp._element,
                    source_component_slide.part,
                    target_slide.part,
                )
                if vertical_offset_pt != 0.0 and hasattr(elem_copy, "spPr"):
                    spPr = elem_copy.spPr
                    xfrm = spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
                    if xfrm is not None:
                        off = xfrm.find("{http://schemas.openxmlformats.org/drawingml/2006/main}off")
                        if off is not None and "y" in off.attrib:
                            try:
                                off.set("y", str(int(off.attrib["y"]) + int(vertical_offset_pt * 12700)))
                            except ValueError:
                                pass
                target_slide.shapes._spTree.append(elem_copy)
            except Exception as e:
                logger.debug(f"Could not inject component {shp.name}: {e}")

        SplicerGuardian.assign_unique_shape_ids(target_slide)


class UniversalSlideSplicer:
    """
    Master Dispatcher for the 3 Splicing Engines.
    """

    @classmethod
    def splice_and_assemble(
        cls,
        engine_mode: str,
        primary_pptx_path: str,
        slide_specs: List[Dict[str, Any]],
        output_path: str,
    ) -> str:
        if not os.path.exists(primary_pptx_path):
            raise FileNotFoundError(f"Primary template not found: {primary_pptx_path}")

        prs = Presentation(primary_pptx_path)

        if engine_mode == "multi_master":
            return MultiMasterSlideSplicer.assemble_multi_master_deck(prs, slide_specs, output_path)

        for del_idx in reversed(range(1, len(prs.slides))):
            try:
                rId = prs.slides._sldIdLst[del_idx].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[del_idx]
            except Exception:
                pass

        palette = ThemeColorExtractor.extract_palette(prs.slides[0]) if len(prs.slides) > 0 else {}
        blank_layout = get_cleanest_blank_layout(prs)

        specs_to_splice = (
            slide_specs[1:]
            if slide_specs and (slide_specs[0].get("slide_number") == 1 or slide_specs[0].get("archetype") == "cover")
            else slide_specs
        )

        for spec in specs_to_splice:
            src_pptx = spec.get("pptx_path")
            src_idx = spec.get("slide_index", 0)

            if not src_pptx or not os.path.exists(src_pptx):
                prs.slides.add_slide(blank_layout)
                continue

            try:
                src_prs = Presentation(src_pptx)
                if src_idx < len(src_prs.slides):
                    src_slide = src_prs.slides[src_idx]
                    target_slide = prs.slides.add_slide(blank_layout)

                    if engine_mode == "modular_stacker":
                        ModularComponentStacker.inject_components(target_slide, src_slide)
                    else:
                        HarmonizedSlideCloner.clone_slide_to_target(src_slide, target_slide, theme_palette=palette)
                else:
                    prs.slides.add_slide(blank_layout)
            except Exception as e:
                logger.warning(f"Error splicing slide from {src_pptx}: {e}")
                prs.slides.add_slide(blank_layout)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        return output_path


class CrossDeckSlideDonor:
    """
    17,250 ta shablon va andozalardan mavzu rejasiga eng mos keluvchi
    boy strukturali betlarni (infografika, taqqoslash, jarayon, SWOT, metrika)
    ajratib oluvchi va asosiy taqdimotga rang/shrift/o'lcham uyg'unligida ulovchi donor tizimi.
    """
    _ARCHETYPE_INDEX: Optional[Dict[str, List[Tuple[str, int]]]] = None

    @classmethod
    def _ensure_index(cls) -> None:
        if cls._ARCHETYPE_INDEX is not None:
            return
        cls._ARCHETYPE_INDEX = {}
        curr_dir = Path(__file__).resolve().parent.parent
        bundled_dir = curr_dir / "templates"
        index_file = bundled_dir / "donor_archetype_index.json"

        if index_file.exists():
            try:
                import json
                with open(index_file, "r", encoding="utf-8") as f:
                    raw_data = json.load(f)
                for atype, items in raw_data.items():
                    cls._ARCHETYPE_INDEX[atype] = []
                    for fn, s_num in items:
                        full_p = str(bundled_dir / fn)
                        if os.path.exists(full_p):
                            cls._ARCHETYPE_INDEX[atype].append((full_p, s_num))
                logger.info(
                    f"CrossDeckDonor: {sum(len(v) for v in cls._ARCHETYPE_INDEX.values())} ta toza donor slayd indeksi yuklandi."
                )
                return
            except Exception as e:
                logger.warning(f"Error loading donor index JSON: {e}")

    @classmethod
    def splice_donor_to_slide(
        cls,
        archetype: str,
        target_slide: Any,
        theme_palette: Optional[Dict[str, Any]],
        used_donors: set,
    ) -> Optional[Tuple[str, int]]:
        """
        Tanlangan arxetip bo'yicha donor slaydni topib, target_slide o'lchamiga moslab (sx, sy)
        va asosiy taqdimot fonini buzmasdan (replicate_background=False) nusxalaydi.
        """
        import gc
        cls._ensure_index()
        if not cls._ARCHETYPE_INDEX:
            return None

        candidates = list(cls._ARCHETYPE_INDEX.get(archetype, []))
        for fallback_atype in ("cards_grid", "timeline_steps", "theory_concept", "comparison_vs"):
            if fallback_atype != archetype:
                candidates.extend(cls._ARCHETYPE_INDEX.get(fallback_atype, []))

        for p_path, s_idx in candidates:
            donor_key = f"{p_path}#{s_idx}"
            if donor_key in used_donors:
                continue
            used_donors.add(donor_key)
            try:
                donor_prs = Presentation(p_path)
                if s_idx < len(donor_prs.slides):
                    donor_slide = donor_prs.slides[s_idx]
                    HarmonizedSlideCloner.clone_slide_to_target(
                        donor_slide,
                        target_slide,
                        theme_palette=theme_palette,
                        replicate_background=False,
                    )
                    del donor_prs
                    gc.collect()
                    return p_path, s_idx
            except Exception as e:
                logger.debug(f"Failed to clone donor slide {donor_key}: {e}")
                continue

        return None

    @classmethod
    def get_donor_slide(cls, archetype: str, used_donors: set) -> Optional[Tuple[Any, str, int]]:
        cls._ensure_index()
        if not cls._ARCHETYPE_INDEX:
            return None

        candidates = cls._ARCHETYPE_INDEX.get(archetype, [])
        if not candidates:
            candidates = cls._ARCHETYPE_INDEX.get("cards_grid", []) or cls._ARCHETYPE_INDEX.get("general", [])

        for p_path, s_idx in candidates:
            donor_key = f"{p_path}#{s_idx}"
            if donor_key in used_donors:
                continue
            used_donors.add(donor_key)
            try:
                donor_prs = Presentation(p_path)
                if s_idx < len(donor_prs.slides):
                    return donor_prs.slides[s_idx], p_path, s_idx
            except Exception:
                continue

        return None
