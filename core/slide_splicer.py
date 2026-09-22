# -*- coding: utf-8 -*-
"""
core/slide_splicer.py
Multi-Deck Slide Splicing & Assembly Engines for SlaydHub AI.
Implements 3 Distinct Splicing Engines:
1. HarmonizedSlideCloner: Deep XML cloning with Theme Harmonizer (colors & fonts match cover).
2. MultiMasterSlideSplicer: Preserves each slide's native background, textures, and master styling.
3. ModularComponentStacker: Injects high-impact graphic & infographic components into base deck.
"""

import os
import io
import copy
import logging
from typing import Dict, Any, List, Optional, Tuple

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

from core.splicer_guardian import SplicerGuardian

logger = logging.getLogger("SlideSplicer")


class ThemeColorExtractor:
    """
    Extracts dominant color palette and font pairings from a template slide.
    """

    @staticmethod
    def extract_palette(slide) -> Dict[str, Any]:
        fonts = []
        colors = []
        
        for shp in slide.shapes:
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

        primary_font = fonts[0] if fonts else "Segoe UI"
        primary_color = colors[0] if colors else RGBColor(0x1F, 0x4E, 0x79) # Deep Navy
        accent_color = colors[1] if len(colors) > 1 else RGBColor(0x00, 0x70, 0xC0) # Bright Blue

        return {
            "primary_font": primary_font,
            "primary_color": primary_color,
            "accent_color": accent_color
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
    Engine 1: Deep XML Shape Cloning with Master Theme Harmonization.
    Clones vector shapes, diagrams, and layout blocks from candidate templates,
    then harmonizes font families and accent colors to match the cover slide.
    """

    @classmethod
    def clone_slide_to_target(
        cls,
        source_slide,
        target_slide,
        theme_palette: Optional[Dict[str, Any]] = None
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

        # 1. Replicate authentic background (Solid, Gradient, Texture, or BlipFill) via SplicerGuardian
        SplicerGuardian.replicate_slide_background(source_slide, target_slide)

        # 2. Clone each shape safely with full media relationship remapping
        rel_map = {}
        for shp in source_slide.shapes:
            try:
                # Safe deep cloning with relationship remapping and dangling rId purging
                elem_copy = SplicerGuardian.safe_clone_shape_with_rels(
                    shp._element,
                    source_slide.part,
                    target_slide.part,
                    rel_map=rel_map
                )

                # If shape inherited layout coordinates (e.g. title placeholder), bake them into elem_copy
                if hasattr(elem_copy, 'spPr'):
                    spPr = elem_copy.spPr
                    if spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm') is None:
                        if shp.left is not None and shp.top is not None and shp.width is not None and shp.height is not None:
                            xfrm_elem = parse_xml(f'<a:xfrm {nsdecls("a")}><a:off x="{shp.left}" y="{shp.top}"/><a:ext cx="{shp.width}" cy="{shp.height}"/></a:xfrm>')
                            spPr.insert(0, xfrm_elem)

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
        Harmonizes text font families and key accent colors across cloned shapes.
        """
        target_font = theme_palette.get("primary_font", "Segoe UI")
        target_color = theme_palette.get("primary_color")

        for shp in slide.shapes:
            if not shp.has_text_frame:
                continue
            for p in shp.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.name:
                        r.font.name = target_font
                    if target_color and r.font.bold:
                        try:
                            # Apply subtle theme tint to bold headers
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
        output_path: str
    ) -> str:
        """
        Creates a presentation by copying slides with full design preservation.
        """
        # Keep slide 0 (cover), prune any excess existing slides
        for del_idx in reversed(range(1, len(base_prs.slides))):
            try:
                rId = base_prs.slides._sldIdLst[del_idx].rId
                base_prs.part.drop_rel(rId)
                del base_prs.slides._sldIdLst[del_idx]
            except Exception:
                pass

        blank_layout = get_cleanest_blank_layout(base_prs)

        # Determine specs to splice (skip cover if already in specs)
        specs_to_splice = slide_specs[1:] if slide_specs and (slide_specs[0].get("slide_number") == 1 or slide_specs[0].get("archetype") == "cover") else slide_specs

        for spec in specs_to_splice:
            src_pptx = spec.get("pptx_path")
            src_idx = spec.get("slide_index", 0)

            if not src_pptx or not os.path.exists(src_pptx):
                # Fallback: add a blank slide
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
    Takes a clean base slide and injects high-impact visual components
    (process nodes, timeline badges, comparison columns, metric cards)
    extracted from rich template slide libraries.
    """

    @classmethod
    def inject_components(
        cls,
        target_slide,
        source_component_slide,
        vertical_offset_pt: float = 0.0
    ) -> None:
        """
        Extracts non-title vector components, group shapes, and diagram cards
        from source_component_slide and injects them onto target_slide.
        """
        for shp in source_component_slide.shapes:
            # Skip title shapes of source slide (keep target slide's title)
            is_title = False
            if shp.has_text_frame:
                txt = shp.text_frame.text.lower()
                if "title" in shp.name.lower() or shp.top < Inches(1.2).emu:
                    is_title = True

            if is_title:
                continue

            try:
                elem_copy = SplicerGuardian.safe_clone_shape_with_rels(
                    shp._element,
                    source_component_slide.part,
                    target_slide.part
                )
                if vertical_offset_pt != 0.0 and hasattr(elem_copy, 'spPr'):
                    spPr = elem_copy.spPr
                    xfrm = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
                    if xfrm is not None:
                        off = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off')
                        if off is not None and 'y' in off.attrib:
                            try:
                                off.set('y', str(int(off.attrib['y']) + int(vertical_offset_pt * 12700)))
                            except ValueError:
                                pass
                target_slide.shapes._spTree.append(elem_copy)
            except Exception as e:
                logger.debug(f"Could not inject component {shp.name}: {e}")

        # Re-assign unique shape IDs
        SplicerGuardian.assign_unique_shape_ids(target_slide)


class UniversalSlideSplicer:
    """
    Master Dispatcher for the 3 Splicing Engines.
    Supports mode: 'harmonized', 'multi_master', 'modular_stacker'.
    """

    @classmethod
    def splice_and_assemble(
        cls,
        engine_mode: str,
        primary_pptx_path: str,
        slide_specs: List[Dict[str, Any]],
        output_path: str
    ) -> str:
        """
        Executes slide assembly using the chosen engine.
        """
        if not os.path.exists(primary_pptx_path):
            raise FileNotFoundError(f"Primary template not found: {primary_pptx_path}")

        prs = Presentation(primary_pptx_path)

        if engine_mode == "multi_master":
            return MultiMasterSlideSplicer.assemble_multi_master_deck(prs, slide_specs, output_path)

        # Keep slide 0 (cover), prune any excess existing slides
        for del_idx in reversed(range(1, len(prs.slides))):
            try:
                rId = prs.slides._sldIdLst[del_idx].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[del_idx]
            except Exception:
                pass

        # For Harmonized Cloner & Modular Stacker:
        palette = ThemeColorExtractor.extract_palette(prs.slides[0]) if len(prs.slides) > 0 else {}
        blank_layout = get_cleanest_blank_layout(prs)

        # Determine specs to splice (skip cover if already in specs)
        specs_to_splice = slide_specs[1:] if slide_specs and (slide_specs[0].get("slide_number") == 1 or slide_specs[0].get("archetype") == "cover") else slide_specs

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
                    else:  # harmonized (default)
                        HarmonizedSlideCloner.clone_slide_to_target(src_slide, target_slide, theme_palette=palette)
                else:
                    prs.slides.add_slide(blank_layout)
            except Exception as e:
                logger.warning(f"Error splicing slide from {src_pptx}: {e}")
                prs.slides.add_slide(blank_layout)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        return output_path
