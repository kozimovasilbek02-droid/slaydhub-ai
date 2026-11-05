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
from pptx.shapes.autoshape import Shape as _PptxAutoShape
_orig_shape_type_fget = _PptxAutoShape.shape_type.fget
def _safe_shape_type(self):
    try:
        return _orig_shape_type_fget(self)
    except NotImplementedError:
        return MSO_SHAPE_TYPE.TEXT_BOX if getattr(self, "has_text_frame", False) else MSO_SHAPE_TYPE.AUTO_SHAPE
_PptxAutoShape.shape_type = property(_safe_shape_type)

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

from core.splicer_guardian import SplicerGuardian

logger = logging.getLogger("SlideSplicer")

import colorsys
from pptx.oxml.ns import qn

def normalize_slide_group_coordinates(slide) -> None:
    """
    Normalizes <p:grpSpPr><a:xfrm> across all groups on a slide so chOff == off and chExt == ext,
    giving every child shapeinside groups its exact true slide-space (left, top, width, height).
    """
    def _norm_grp(grp_el):
        grpSpPr = grp_el.find(qn("p:grpSpPr"))
        if grpSpPr is None:
            return
        xfrm = grpSpPr.find(qn("a:xfrm"))
        if xfrm is None:
            return
        off = xfrm.find(qn("a:off"))
        ext = xfrm.find(qn("a:ext"))
        chOff = xfrm.find(qn("a:chOff"))
        chExt = xfrm.find(qn("a:chExt"))
        if None in (off, ext, chOff, chExt):
            return
        ox, oy = int(off.get("x", 0)), int(off.get("y", 0))
        ex, ey = int(ext.get("cx", 0)), int(ext.get("cy", 0))
        cx_off, cy_off = int(chOff.get("x", 0)), int(chOff.get("y", 0))
        cx_ext, cy_ext = int(chExt.get("cx", 0)), int(chExt.get("cy", 0))
        if cx_ext <= 0 or cy_ext <= 0 or ex <= 0 or ey <= 0:
            return
        if (ox, oy, ex, ey) != (cx_off, cy_off, cx_ext, cy_ext):
            sx = ex / float(cx_ext)
            sy = ey / float(cy_ext)
            for child in grp_el:
                tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                c_spPr = None
                if tag == "grpSp":
                    c_spPr = child.find(qn("p:grpSpPr"))
                elif tag in ("sp", "pic", "cxnSp", "graphicFrame"):
                    c_spPr = child.find(qn("p:spPr"))
                    if c_spPr is None:
                        c_spPr = child.find(qn("p:xfrm"))
                if c_spPr is not None:
                    c_xfrm = c_spPr if c_spPr.tag.endswith("xfrm") else c_spPr.find(qn("a:xfrm"))
                    if c_xfrm is not None:
                        c_off = c_xfrm.find(qn("a:off"))
                        c_ext = c_xfrm.find(qn("a:ext"))
                        if c_off is not None:
                            x = int(c_off.get("x", 0))
                            y = int(c_off.get("y", 0))
                            c_off.set("x", str(int(round(ox + (x - cx_off) * sx))))
                            c_off.set("y", str(int(round(oy + (y - cy_off) * sy))))
                        if c_ext is not None:
                            w = int(c_ext.get("cx", 0))
                            h = int(c_ext.get("cy", 0))
                            c_ext.set("cx", str(int(round(w * sx))))
                            c_ext.set("cy", str(int(round(h * sy))))
            chOff.set("x", str(ox))
            chOff.set("y", str(oy))
            chExt.set("cx", str(ex))
            chExt.set("cy", str(ey))
        for child in grp_el:
            if child.tag.endswith("grpSp"):
                _norm_grp(child)

    try:
        for el in slide.shapes._spTree:
            if el.tag.endswith("grpSp"):
                _norm_grp(el)
    except Exception:
        pass


def extract_slide_theme_color_map(slide) -> Dict[str, str]:
    theme_rgb = {
        "dk1": "000000", "lt1": "FFFFFF",
        "dk2": "1F497D", "lt2": "EEECE1",
        "accent1": "4F81BD", "accent2": "C0504D",
        "accent3": "9BBB59", "accent4": "8064A2",
        "accent5": "4BACC6", "accent6": "F79646",
        "hlink": "0000FF", "folHlink": "800080",
        "black": "000000", "white": "FFFFFF",
    }
    clr_map = {
        "bg1": "lt1", "tx1": "dk1",
        "bg2": "lt2", "tx2": "dk2",
        "accent1": "accent1", "accent2": "accent2",
        "accent3": "accent3", "accent4": "accent4",
        "accent5": "accent5", "accent6": "accent6",
        "hlink": "hlink", "folHlink": "folHlink",
    }
    try:
        master = slide.slide_layout.slide_master
        m_clrMap = master._element.find(qn("p:clrMap"))
        if m_clrMap is not None:
            for k, v in m_clrMap.attrib.items():
                clr_map[k.split("}")[-1] if "}" in k else k] = v
        s_clrMapOvr = slide._element.find(qn("p:clrMapOvr"))
        if s_clrMapOvr is not None:
            ovr = s_clrMapOvr.find(qn("a:overrideClrMapping"))
            if ovr is not None:
                for k, v in ovr.attrib.items():
                    clr_map[k.split("}")[-1] if "}" in k else k] = v
        for rel in master.part.rels.values():
            if "theme" in rel.reltype:
                t_el = parse_xml(rel.target_part.blob.decode("utf-8", errors="ignore"))
                clrScheme = t_el.find(".//" + qn("a:clrScheme"))
                if clrScheme is not None:
                    for child in clrScheme:
                        cname = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                        srgb = child.find(qn("a:srgbClr"))
                        sysClr = child.find(qn("a:sysClr"))
                        if srgb is not None and srgb.get("val"):
                            theme_rgb[cname] = srgb.get("val").upper()
                        elif sysClr is not None and sysClr.get("lastClr"):
                            theme_rgb[cname] = sysClr.get("lastClr").upper()
                break
    except Exception:
        pass
    final_map = dict(theme_rgb)
    for alias, target in clr_map.items():
        if target in theme_rgb:
            final_map[alias] = theme_rgb[target]
    return final_map


def _apply_color_modifiers(hex_str: str, scheme_el) -> str:
    try:
        r = int(hex_str[0:2], 16) / 255.0
        g = int(hex_str[2:4], 16) / 255.0
        b = int(hex_str[4:6], 16) / 255.0
        h, l, s = colorsys.rgb_to_hls(r, g, b)
        for mod in scheme_el:
            mtag = mod.tag.split("}")[-1] if "}" in mod.tag else mod.tag
            val = int(mod.get("val", "100000")) / 100000.0
            if mtag == "lumMod":
                l = l * val
            elif mtag == "lumOff":
                l = l + val
            elif mtag == "tint":
                l = l + (1.0 - l) * (1.0 - val)
            elif mtag == "shade":
                l = l * val
        l = max(0.0, min(1.0, l))
        r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
        return f"{int(round(r2*255)):02X}{int(round(g2*255)):02X}{int(round(b2*255)):02X}"
    except Exception:
        return hex_str


def bake_scheme_colors_in_element(root_el, color_map: Dict[str, str]) -> None:
    if not color_map or root_el is None:
        return
    nsa = nsdecls("a")
    mod_fn = globals().get("_apply_color_modifiers") or globals().get("apply_color_modifiers")
    for sc in list(root_el.xpath(".//a:schemeClr")):
        val = sc.get("val", "")
        if val in color_map:
            hex_c = mod_fn(color_map[val], sc)
            alpha_el = sc.find(qn("a:alpha"))
            if alpha_el is not None and alpha_el.get("val"):
                aval = alpha_el.get("val")
                new_el = parse_xml(f'<a:srgbClr {nsa} val="{hex_c}"><a:alpha val="{aval}"/></a:srgbClr>')
            else:
                new_el = parse_xml(f'<a:srgbClr {nsa} val="{hex_c}"/>')
            p = sc.getparent()
            if p is not None:
                p.replace(sc, new_el)



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

            # 1b. Scale table column widths (<a:gridCol w="..."/>) by sx so 16:9 tables fit 4:3 slides
            for gcol in elem_copy.xpath('.//a:tblGrid/a:gridCol'):
                if 'w' in gcol.attrib:
                    try:
                        gcol.set('w', str(int(round(int(gcol.get('w', '0')) * sx))))
                    except Exception:
                        pass

            # 2. Scale font sizes (a:rPr and a:defRPr) by min(sx, sy) so narrow badges ('30%', '2025') never wrap when splicing 16:9 -> 4:3
            f_scale = min(sx, sy)
            for rpr in elem_copy.xpath('.//a:rPr | .//a:defRPr | .//a:endParaRPr'):
                if 'sz' in rpr.attrib:
                    try:
                        old_sz = int(rpr.get('sz'))
                        rpr.set('sz', str(int(round(old_sz * f_scale))))
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

        # 1. Replicate background when requested, and hide target SlideMaster shapes (e.g. corner icons) on cross-deck splices
        if replicate_background:
            SplicerGuardian.replicate_slide_background(source_slide, target_slide)
        try:
            target_slide._element.set("showMasterSp", "0")
        except Exception:
            pass

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

        # Normalize group coordinates and extract source slide theme color map before cloning
        normalize_slide_group_coordinates(source_slide)
        src_color_map = extract_slide_theme_color_map(source_slide)
        bake_scheme_colors_in_element(target_slide._element.find(qn("p:cSld")), src_color_map)

        # 2. Clone each shape safely with full media relationship remapping
        rel_map = {}
        for shp in source_slide.shapes:
            try:
                # Skip empty picture/content placeholders so they never render as empty gray boxes
                if shp.is_placeholder and (not shp.has_text_frame or not shp.text_frame.text.strip()):
                    continue

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
                                f'<a:xfrm {nsdecls("a")}><a:off x="{int(shp.left)}" y="{int(shp.top)}"/><a:ext cx="{int(shp.width)}" cy="{int(shp.height)}"/></a:xfrm>'
                            )
                            spPr.insert(0, xfrm_elem)

                # Convert cloned text placeholder into standalone TextBox so it never breaks on target blank_layout
                if shp.is_placeholder:
                    spPr_el = elem_copy.find(qn("p:spPr"))
                    if spPr_el is not None and spPr_el.find(qn("a:prstGeom")) is None and spPr_el.find(qn("a:custGeom")) is None:
                        spPr_el.append(parse_xml(f"<a:prstGeom {nsdecls('a')} prst=\"rect\"><a:avLst/></a:prstGeom>"))
                    nvSpPr = elem_copy.find(qn("p:nvSpPr"))
                    if nvSpPr is not None:
                        cNvPr = nvSpPr.find(qn("p:cNvPr"))
                        if cNvPr is not None and "Placeholder" in cNvPr.get("name", ""):
                            cNvPr.set("name", "Title Header")
                        cNvSpPr = nvSpPr.find(qn("p:cNvSpPr"))
                        if cNvSpPr is not None:
                            cNvSpPr.set("txBox", "1")
                        nvPr = nvSpPr.find(qn("p:nvPr"))
                        if nvPr is not None:
                            for ph_el in list(nvPr.findall(qn("p:ph"))):
                                nvPr.remove(ph_el)

                bake_scheme_colors_in_element(elem_copy, src_color_map)
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
        item_count: int = 4,
    ) -> Optional[Tuple[str, int]]:
        """
        Tanlangan arxetip va punktlar soniga (3, 4 yoki 6 ta) mos donor slaydni topib,
        target_slide o'lchamiga moslab (sx, sy) nusxalaydi.
        """
        import gc
        cls._ensure_index()
        if not cls._ARCHETYPE_INDEX:
            return None

        primary_cands = list(cls._ARCHETYPE_INDEX.get(archetype, []))
        # Match slot capacity to item_count:
        # - For agenda with >= 5 items, prefer agenda_2.pptx (6-slot 2x3 grid)
        # - For 3-point content slides, prefer _2.pptx (3-slot layout) so there is never an empty 4th card!
        if len(primary_cands) >= 2:
            if archetype == "agenda" and item_count >= 5:
                primary_cands = [c for c in primary_cands if "_2" in c[0]] + [c for c in primary_cands if "_2" not in c[0]]
            elif archetype in ("cards_grid", "timeline_steps", "metrics_stats") and item_count == 3:
                primary_cands = [c for c in primary_cands if "_2" in c[0]] + [c for c in primary_cands if "_2" not in c[0]]
            elif archetype in ("cards_grid", "timeline_steps", "metrics_stats") and item_count >= 4:
                primary_cands = [c for c in primary_cands if "_1" in c[0]] + [c for c in primary_cands if "_1" not in c[0]]

        fallback_cands = []
        for fallback_atype in ("cards_grid", "timeline_steps", "theory_concept", "comparison_vs"):
            if fallback_atype != archetype:
                fb_list = list(cls._ARCHETYPE_INDEX.get(fallback_atype, []))
                if item_count == 3:
                    fb_list = [c for c in fb_list if "_2" in c[0] or "theory_concept" in c[0]] + [c for c in fb_list if "_2" not in c[0] and "theory_concept" not in c[0]]
                fallback_cands.extend(fb_list)

        candidates = primary_cands + fallback_cands
        # If all primary_cands were already used in a long deck, allow reusing the slot-matched primary candidate
        chosen_pair = None
        for p_path, s_idx in candidates:
            donor_key = f"{p_path}#{s_idx}"
            if donor_key not in used_donors:
                used_donors.add(donor_key)
                chosen_pair = (p_path, s_idx)
                break
        if chosen_pair is None and primary_cands:
            chosen_pair = primary_cands[0]

        for p_path, s_idx in ([chosen_pair] if chosen_pair else []):
            try:
                donor_prs = Presentation(p_path)
                if s_idx < len(donor_prs.slides):
                    donor_slide = donor_prs.slides[s_idx]
                    HarmonizedSlideCloner.clone_slide_to_target(
                        donor_slide,
                        target_slide,
                        theme_palette=None,
                        replicate_background=True,
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
