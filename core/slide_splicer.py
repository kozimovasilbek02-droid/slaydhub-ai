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
import tempfile
from pathlib import Path
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

        # Round 15 Zero-Defect: Purge any zombie animation targets left on target_slide
        try:
            from core.animation_guardian import AnimationGuardian
            AnimationGuardian.purge_zombie_animations(target_slide)
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



class CrossDeckSlideDonor:
    """
    17,250 ta shablon va andozalardan mavzu rejasiga eng mos keluvchi
    boy strukturali betlarni (infografika, taqqoslash, jarayon, SWOT, metrika)
    ajratib oluvchi va asosiy taqdimotga rang/shrift uyg'unligida ulovchi donor tizimi.
    Render 512MB RAM Zero-Hang: Taqdimotlarni xotirada ushlab turmaydi, donor_archetype_index.json
    orqali xotira sarfini 0 MB ga tushiradi va kerakli slaydni on-demand nusxalab darhol RAMni tozalaydi.
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

        # 1. Tezkor va 0 MB RAM: Tayyor oldindan kompilyatsiya qilingan JSON indeksni yuklash
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
                logger.info(f"CrossDeckDonor: {sum(len(v) for v in cls._ARCHETYPE_INDEX.values())} ta toza donor slayd indeksi yuklandi (0 MB RAM).")
                return
            except Exception as e:
                logger.warning(f"Error loading donor index JSON: {e}")

        # 2. Dinamik fallback (faqat JSON bo'lmaganda):
        from core.pptx_text_replacer import AdSlideFilter
        import gc
        candidate_files = list(bundled_dir.glob("*.pptx")) if bundled_dir.exists() else []

        for pptx_file in candidate_files:
            if pptx_file.name.startswith(("~$", "pruned_")):
                continue
            p_str = str(pptx_file)
            try:
                prs = Presentation(p_str)
                for s_idx, s in enumerate(prs.slides):
                    if AdSlideFilter.is_ad_or_credit_slide(s):
                        continue
                    t_shapes = [shp for shp in s.shapes if shp.has_text_frame and shp.text_frame.text.strip()]
                    all_text = " ".join([shp.text_frame.text.lower() for shp in t_shapes])

                    # Ikonkalar to'plami yoki bo'sh maketlarni chetlab o'tish
                    pic_count = sum(1 for shp in s.shapes if shp.shape_type == MSO_SHAPE_TYPE.PICTURE)
                    small_shapes = sum(1 for shp in s.shapes if (shp.width and shp.width.pt < 65) and (shp.height and shp.height.pt < 65))
                    if small_shapes >= 8 or (pic_count >= 5 and len(t_shapes) <= 3) or (len(s.shapes) >= 15 and len(t_shapes) <= 1):
                        continue

                    has_agenda = any(k in all_text for k in ["table of contents", "agenda", "contents", "index", "reja"])
                    has_step = any(k in all_text for k in ["step 01", "step 1", "01.", "steps", "process", "bosqich"])
                    has_timeline = any(k in all_text for k in ["timeline", "2026", "2027", "roadmap", "schedule", "xronologiya"])
                    has_comparison = any(k in all_text for k in ["comparison", "vs", "versus", "pros", "cons", "taqqoslash"])
                    has_stats = any(k in all_text for k in ["%", "stats", "analytics", "chart", "kpi", "metrics", "foiz"])
                    has_credits = any(k in all_text for k in ["thank you", "thanks", "conclusion", "contact", "qa", "xulosa", "savollar"])

                    atype = "general"
                    if s_idx == 0:
                        atype = "cover"
                    elif has_agenda:
                        atype = "agenda"
                    elif has_comparison:
                        atype = "comparison_vs"
                    elif has_timeline or has_step:
                        atype = "timeline_steps"
                    elif has_stats:
                        atype = "metrics_stats"
                    elif has_credits:
                        atype = "conclusion_sources"
                    elif len(t_shapes) >= 6:
                        atype = "cards_grid"
                    elif len(t_shapes) in [2, 3, 4]:
                        atype = "theory_concept"
                    else:
                        atype = "cards_grid"

                    cls._ARCHETYPE_INDEX.setdefault(atype, []).append((p_str, s_idx))
                del prs
                gc.collect()
            except Exception as e:
                logger.debug(f"Donor index error for {pptx_file}: {e}")

    @classmethod
    def splice_donor_to_slide(
        cls,
        archetype: str,
        target_slide: Any,
        theme_palette: Optional[Dict[str, Any]],
        used_donors: set
    ) -> Optional[Tuple[str, int]]:
        """
        Tanlangan arxetip bo'yicha donor slaydni topib, to'g'ridan-to'g'ri target_slide ga
        nusxalaydi va donor taqdimotni xotiradan darhol tozalaydi (Zero Memory Leak).
        """
        import gc
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
                    donor_slide = donor_prs.slides[s_idx]
                    HarmonizedSlideCloner.clone_slide_to_target(donor_slide, target_slide, theme_palette=theme_palette)
                    del donor_prs
                    gc.collect()
                    return p_path, s_idx
            except Exception as e:
                logger.debug(f"Failed to clone donor slide {donor_key}: {e}")
                continue

        return None

    @classmethod
    def get_donor_slide(cls, archetype: str, used_donors: set) -> Optional[Tuple[Any, str, int]]:
        """Eski interfeys bilan moslik uchun (on-demand ochadi)."""
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
