# -*- coding: utf-8 -*-
import os
import time
import re
import logging
from pathlib import Path
from typing import Dict, Any, Optional, List
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

logger = logging.getLogger(__name__)


def find_best_authentic_template(topic: str) -> Optional[str]:
    """17,250 ta shablon ichidan AI Smart Selector orqali mavzuga eng mos haqiqiy shablonni tanlaydi."""
    try:
        from core.template_catalog import GLOBAL_TEMPLATE_CATALOG
        top_matches = GLOBAL_TEMPLATE_CATALOG.smart_search(topic, top_k=5)
        for cand in top_matches:
            tpl_id = cand.get("id")
            resolved = GLOBAL_TEMPLATE_CATALOG.resolve_template_pptx_path(tpl_id)
            if resolved and os.path.exists(resolved):
                logger.info(f"Smart Template Match: '{topic}' -> '{cand.get('title')}' ({resolved})")
                return resolved
    except Exception as e:
        logger.debug(f"Smart template selection error: {e}")

    curr_dir = Path(__file__).resolve().parent.parent
    bundled_dir = curr_dir / "templates"
    if bundled_dir.exists():
        # Prefer universal Business / Tech templates over Coronavirus/Medical as default fallback
        t_cat = "biz"
        try:
            from core.template_catalog import GLOBAL_TEMPLATE_CATALOG
            t_cat = GLOBAL_TEMPLATE_CATALOG.detect_category(topic)
        except Exception:
            pass
        if t_cat == "med":
            preferred = [
                "07_Modern_Healthcare_Solutions.pptx",
                "11_Humanity_Heartbeat.pptx",
                "16_Dental_Health_Care.pptx",
                "20_Capsules_Medication_Balancing.pptx",
                "2019 Business Plan PowerPoint Templates.pptx",
            ]
        elif t_cat == "tech":
            preferred = [
                "2019 Business Plan PowerPoint Templates.pptx",
                "09_Digital_Healthcare.pptx",
                "07_Modern_Healthcare_Solutions.pptx",
            ]
        else:
            preferred = [
                "2019 Business Plan PowerPoint Templates.pptx",
                "09_Digital_Healthcare.pptx",
                "07_Modern_Healthcare_Solutions.pptx",
            ]
        for pref_fn in preferred:
            cand_p = bundled_dir / pref_fn
            if cand_p.exists():
                return str(cand_p)
        all_bundled = [p for p in bundled_dir.glob("*.pptx") if "coronavirus" not in p.name.lower()]
        if all_bundled:
            return str(all_bundled[0])

    return None


# ── 1. PENPOT PALITRALARI ─────────────────────────────────────────────────────
PENPOT_PALETTES = {
    "academic_navy": {
        "dark_bg": RGBColor(0x0F, 0x17, 0x2A),
        "light_bg": RGBColor(0xF8, 0xFA, 0xFC),
        "card_bg": RGBColor(0xEE, 0xF2, 0xF6),
        "card_border": RGBColor(0xCB, 0xD5, 0xE1),
        "primary": RGBColor(0x1E, 0x3A, 0x8A),
        "accent": RGBColor(0x02, 0x84, 0xC7),
        "text_dark": RGBColor(0x0F, 0x17, 0x2A),
        "text_muted": RGBColor(0x47, 0x55, 0x69),
        "text_light": RGBColor(0xF8, 0xFA, 0xFC),
    },
    "emerald_science": {
        "dark_bg": RGBColor(0x06, 0x4E, 0x3B),
        "light_bg": RGBColor(0xF0, 0xFD, 0xF4),
        "card_bg": RGBColor(0xDC, 0xFC, 0xE7),
        "card_border": RGBColor(0x86, 0xEF, 0xAC),
        "primary": RGBColor(0x04, 0x78, 0x57),
        "accent": RGBColor(0x10, 0xB9, 0x81),
        "text_dark": RGBColor(0x06, 0x4E, 0x3B),
        "text_muted": RGBColor(0x37, 0x41, 0x51),
        "text_light": RGBColor(0xF0, 0xFD, 0xF4),
    },
    "tech_minimal": {
        "dark_bg": RGBColor(0x18, 0x18, 0x1B),
        "light_bg": RGBColor(0xFA, 0xFA, 0xFA),
        "card_bg": RGBColor(0xF4, 0xF4, 0xF5),
        "card_border": RGBColor(0xE4, 0xE4, 0xE7),
        "primary": RGBColor(0x27, 0x27, 0x2A),
        "accent": RGBColor(0x3B, 0x82, 0xF6),
        "text_dark": RGBColor(0x18, 0x18, 0x1B),
        "text_muted": RGBColor(0x52, 0x52, 0x5B),
        "text_light": RGBColor(0xFA, 0xFA, 0xFA),
    },
    "warm_philosophy": {
        "dark_bg": RGBColor(0x45, 0x1A, 0x03),
        "light_bg": RGBColor(0xFF, 0xFB, 0xEB),
        "card_bg": RGBColor(0xFE, 0xF3, 0xC7),
        "card_border": RGBColor(0xFC, 0xD3, 0x4D),
        "primary": RGBColor(0x78, 0x35, 0x0F),
        "accent": RGBColor(0xD9, 0x77, 0x06),
        "text_dark": RGBColor(0x45, 0x1A, 0x03),
        "text_muted": RGBColor(0x57, 0x53, 0x4E),
        "text_light": RGBColor(0xFF, 0xFB, 0xEB),
    },
    "cyber_dark": {
        "dark_bg": RGBColor(0x03, 0x07, 0x12),
        "light_bg": RGBColor(0x11, 0x18, 0x27),
        "card_bg": RGBColor(0x1F, 0x29, 0x37),
        "card_border": RGBColor(0x37, 0x41, 0x51),
        "primary": RGBColor(0x06, 0xB6, 0xD4),
        "accent": RGBColor(0x14, 0xB8, 0xA6),
        "text_dark": RGBColor(0xF9, 0xFA, 0xFB),
        "text_muted": RGBColor(0x9C, 0xA3, 0xAF),
        "text_light": RGBColor(0xF9, 0xFA, 0xFB),
    },
}

UZ_ORPHAN_PREPOSITIONS_RE = re.compile(
    r"\b(va|ham|bilan|uchun|yoki|ammo|lekin|biroq|to|har|bir|shu|bu|u|o'sha|barcha)\s+",
    re.IGNORECASE,
)


def apply_knuth_plass_uz_orphans(text: str) -> str:
    if not text:
        return ""
    return UZ_ORPHAN_PREPOSITIONS_RE.sub(lambda m: f"{m.group(1)}\u00A0", str(text))


def apply_anti_word_break(paragraph) -> None:
    try:
        pPr = paragraph._p.get_or_add_pPr()
        pPr.set("latinLnBrk", "0")
        pPr.set("eaLnBrk", "0")
        pPr.set("hangingPunct", "1")
    except Exception:
        pass


def apply_p16_morph_transition(slide) -> None:
    MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    P16_NS = "http://schemas.microsoft.com/office/powerpoint/2015/main"
    trans_xml = (
        f'<p:transition {nsdecls("p")} xmlns:mc="{MC_NS}" spd="med" advClick="1">'
        f'  <mc:AlternateContent>'
        f'    <mc:Choice Requires="p16" xmlns:p16="{P16_NS}">'
        f'      <p16:morph option="byObject"/>'
        f'    </mc:Choice>'
        f'    <mc:Fallback><p:fade/></mc:Fallback>'
        f'  </mc:AlternateContent>'
        f'</p:transition>'
    )
    try:
        slide._element.append(parse_xml(trans_xml))
    except Exception:
        pass


def _generate_synthetic_penpot_pptx(spec: Dict[str, Any], out_path: str) -> str:
    """To'liq multi-mavzuli, multi-arxetipli professional .pptx fayl generatsiyasi (Penpot fallback)."""
    from core.topic_logic_validator import TopicLogicValidator
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    theme_name = spec.get("theme", spec.get("recommended_theme", "academic_navy"))
    palette = PENPOT_PALETTES.get(theme_name, PENPOT_PALETTES["academic_navy"])
    topic = spec.get("title") or spec.get("topic") or "Ilmiy Taqdimot"

    slides_data = spec.get("slides", [])
    if not slides_data:
        slides_data = [
            {
                "heading": "Kirish va Dolzarblik",
                "archetype": "cards_grid",
                "points": ["Asosiy maqsad va yo'nalishlar", "Rivojlanish tendensiyalari", "Amaliy ahamiyat"],
                "speaker_note": "Kirish nutqi",
            }
        ]

    # ── SLAYD 1: DARK HERO COVER ──────────────────────────────────────────────
    s1 = prs.slides.add_slide(blank_layout)
    apply_p16_morph_transition(s1)

    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = palette["dark_bg"]
    bg1.line.fill.background()

    pill = s1.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.2), Inches(3.2), Inches(0.45)
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = palette["primary"]
    pill.line.fill.background()
    p_tf = pill.text_frame
    p_tf.word_wrap = True
    p_p = p_tf.paragraphs[0]
    p_p.text = "ILMIY-AMALIY TAQDIMOT"
    p_p.font.name = "Arial"
    p_p.font.size = Pt(11)
    p_p.font.bold = True
    p_p.font.color.rgb = palette["text_light"]
    p_p.alignment = PP_ALIGN.CENTER

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(11.333), Inches(2.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1_t = tf1.paragraphs[0]
    p1_t.text = apply_knuth_plass_uz_orphans(topic)
    p1_t.font.name = "Arial"
    p1_t.font.size = Pt(38 if len(p1_t.text) < 65 else 30)
    p1_t.font.bold = True
    p1_t.font.color.rgb = palette["text_light"]
    apply_anti_word_break(p1_t)

    subtitle_text = spec.get("subtitle") or spec.get("description") or "Professional tahliliy taqdimot va hisobot"
    p1_sub = tf1.add_paragraph()
    p1_sub.text = apply_knuth_plass_uz_orphans(subtitle_text)
    p1_sub.font.name = "Arial"
    p1_sub.font.size = Pt(18)
    p1_sub.font.color.rgb = palette["accent"]
    p1_sub.space_before = Pt(16)
    apply_anti_word_break(p1_sub)

    s1.notes_slide.notes_text_frame.text = f"Muqova: {topic}. Taqdimot rejasiga kirish."

    # ── SLAYD 2: MANTIQIY REJA (AGENDA) ───────────────────────────────────────
    s2 = prs.slides.add_slide(blank_layout)
    apply_p16_morph_transition(s2)

    bg2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = palette["light_bg"]
    bg2.line.fill.background()

    top_bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = palette["accent"]
    top_bar.line.fill.background()

    s2_tb = s2.shapes.add_textbox(Inches(1.0), Inches(0.7), Inches(11.333), Inches(0.9))
    s2_tf = s2_tb.text_frame
    s2_tf.word_wrap = True
    p2_head = s2_tf.paragraphs[0]
    p2_head.text = "Mundarija va Asosiy Yo'nalishlar"
    p2_head.font.name = "Arial"
    p2_head.font.size = Pt(28)
    p2_head.font.bold = True
    p2_head.font.color.rgb = palette["text_dark"]
    apply_anti_word_break(p2_head)

    agenda_items = [slide.get("heading") or slide.get("title") or f"Mavzu {i+1}" for i, slide in enumerate(slides_data)]
    for i, item in enumerate(agenda_items[:6]):
        row_idx = i % 3
        col_idx = i // 3
        c_left = Inches(1.0 + col_idx * 5.8)
        c_top = Inches(1.9 + row_idx * 1.65)
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, c_left, c_top, Inches(5.4), Inches(1.4))
        card.fill.solid()
        card.fill.fore_color.rgb = palette["card_bg"]
        card.line.color.rgb = palette["card_border"]
        card.line.width = Pt(1.2)
        ctf = card.text_frame
        ctf.word_wrap = True
        cp = ctf.paragraphs[0]
        cp.text = f"{i+1:02d}. {apply_knuth_plass_uz_orphans(item)}"
        cp.font.name = "Arial"
        cp.font.size = Pt(16)
        cp.font.bold = True
        cp.font.color.rgb = palette["primary"]
        cp2 = ctf.add_paragraph()
        cp2.text = apply_knuth_plass_uz_orphans("Asosiy nazariy tamoyillar, tahliliy ko'rsatkichlar va amaliy yechimlar")
        cp2.font.name = "Arial"
        cp2.font.size = Pt(12.5)
        cp2.font.color.rgb = palette["text_muted"]
        cp2.space_before = Pt(4)

    # ── SLAYD 3..N: ASOSIY MAZMUN (ARCHETYPES) ────────────────────────────────
    for idx, slide_info in enumerate(slides_data, 1):
        s = prs.slides.add_slide(blank_layout)
        apply_p16_morph_transition(s)

        bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = palette["light_bg"]
        bg.line.fill.background()

        t_bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
        t_bar.fill.solid()
        t_bar.fill.fore_color.rgb = palette["accent"]
        t_bar.line.fill.background()

        badge = s.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.55), Inches(2.5), Inches(0.35)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = palette["primary"]
        badge.line.fill.background()
        b_p = badge.text_frame.paragraphs[0]
        b_p.text = f"{idx:02d} / {len(slides_data):02d} - BO'LIM"
        b_p.font.name = "Arial"
        b_p.font.size = Pt(9.5)
        b_p.font.bold = True
        b_p.font.color.rgb = palette["text_light"]
        b_p.alignment = PP_ALIGN.CENTER

        s_heading = slide_info.get("heading") or slide_info.get("title") or f"Mavzu {idx}"
        h_tb = s.shapes.add_textbox(Inches(1.0), Inches(1.05), Inches(11.333), Inches(0.95))
        h_tf = h_tb.text_frame
        h_tf.word_wrap = True
        h_p = h_tf.paragraphs[0]
        h_p.text = apply_knuth_plass_uz_orphans(s_heading)
        h_p.font.name = "Arial"
        h_p.font.size = Pt(24)
        h_p.font.bold = True
        h_p.font.color.rgb = palette["text_dark"]
        apply_anti_word_break(h_p)

        raw_points = slide_info.get("points") or slide_info.get("bullets") or ["Asosiy ma'lumot", "Tahlil va natijalar", "Amaliy tatbiq"]
        parsed_pairs = []
        for p_i, p_raw in enumerate(raw_points, 1):
            enriched = TopicLogicValidator.enrich_short_point(str(p_raw), s_heading, topic, p_i)
            if ":" in enriched:
                h_part, d_part = enriched.split(":", 1)
                parsed_pairs.append((h_part.strip(), d_part.strip()))
            else:
                parsed_pairs.append((f"Yo'nalish {p_i}", enriched))

        count = min(4, max(1, len(parsed_pairs)))
        gap = 0.3
        total_w = 11.333
        w_in = (total_w - gap * (count - 1)) / count
        for c_i, (p_head, p_desc) in enumerate(parsed_pairs[:count]):
            c_left = Inches(1.0 + c_i * (w_in + gap))
            c = s.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, c_left, Inches(2.2), Inches(w_in), Inches(4.5)
            )
            c.fill.solid()
            c.fill.fore_color.rgb = palette["card_bg"]
            c.line.color.rgb = palette["card_border"]
            c.line.width = Pt(1.5)
            ctf = c.text_frame
            ctf.word_wrap = True
            ctf.margin_left = Inches(0.18)
            ctf.margin_right = Inches(0.18)
            ctf.margin_top = Inches(0.18)
            cp = ctf.paragraphs[0]
            cp.text = f"0{c_i+1}. {apply_knuth_plass_uz_orphans(p_head)}"
            cp.font.name = "Arial"
            cp.font.size = Pt(15.5)
            cp.font.bold = True
            cp.font.color.rgb = palette["primary"]
            apply_anti_word_break(cp)

            cp2 = ctf.add_paragraph()
            cp2.text = apply_knuth_plass_uz_orphans(p_desc)
            cp2.font.name = "Arial"
            cp2.font.size = Pt(13.5)
            cp2.font.color.rgb = palette["text_dark"]
            cp2.space_before = Pt(10)
            apply_anti_word_break(cp2)

        sp_note = slide_info.get("speaker_note", "")
        s.notes_slide.notes_text_frame.text = sp_note or f"{s_heading} mavzusi bo'yicha ilmiy sharh."

    # ── YAKUNIY SLAYD: OUTRO ──────────────────────────────────────────────────
    s_end = prs.slides.add_slide(blank_layout)
    apply_p16_morph_transition(s_end)

    bg_end = s_end.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg_end.fill.solid()
    bg_end.fill.fore_color.rgb = palette["dark_bg"]
    bg_end.line.fill.background()

    end_tb = s_end.shapes.add_textbox(Inches(1.0), Inches(2.5), Inches(11.333), Inches(2.5))
    end_tf = end_tb.text_frame
    end_tf.word_wrap = True
    end_p = end_tf.paragraphs[0]
    end_p.text = "E'tiboringiz Uchun Rahmat!"
    end_p.font.name = "Arial"
    end_p.font.size = Pt(36)
    end_p.font.bold = True
    end_p.font.color.rgb = palette["text_light"]
    end_p.alignment = PP_ALIGN.CENTER

    end_sub = end_tf.add_paragraph()
    end_sub.text = "Savollar va Munozara Uchun Ochiq"
    end_sub.font.name = "Arial"
    end_sub.font.size = Pt(18)
    end_sub.font.color.rgb = palette["accent"]
    end_sub.alignment = PP_ALIGN.CENTER
    end_sub.space_before = Pt(14)

    prs.save(out_path)
    return out_path


def generate_pptx_file(spec: Dict[str, Any], out_path: str) -> str:
    """
    SlaydHub Ultra v5.4 Taqdimot Generator:
    1. Tanlangan template_id yoki mavzuga mos haqiqiy shablonni oladi.
    2. Har bir slaydga mos arxetip (cover, agenda, cards_grid, timeline_steps, comparison_vs, metrics_stats, conclusion_sources) biriktiradi.
    3. replace_presentation_content orqali takroriy diagrammalarni filtrlab, barcha matnlarni to'liq va yuqori kontrastda joylashtiradi.
    """
    from core.topic_logic_validator import TopicLogicValidator

    topic = spec.get("title") or spec.get("topic") or "Akademik Taqdimot"
    target_count = spec.get("target_slide_count") or (len(spec.get("slides", [])) + 2)
    tpl_id = spec.get("template_id")

    try:
        from core.template_catalog import GLOBAL_TEMPLATE_CATALOG
        from core.pptx_text_replacer import replace_presentation_content

        tpl_path = spec.get("resolved_pptx_path")
        if not tpl_path or not os.path.exists(tpl_path):
            if tpl_id:
                tpl_path = GLOBAL_TEMPLATE_CATALOG.resolve_template_pptx_path(tpl_id)
        if not tpl_path or not os.path.exists(tpl_path):
            tpl_path = find_best_authentic_template(topic)

        if tpl_path and os.path.exists(tpl_path):
            slides_raw = spec.get("slides", [])
            formatted_slides = []

            agenda_titles = [
                str(s.get("heading") or s.get("title") or "").strip()
                for s in slides_raw
                if (s.get("heading") or s.get("title")) and s.get("archetype") != "cover" and s.get("layout_type") != "cover"
            ]

            # 1. Muqova (Cover) slaydi
            has_cover = any((s.get("archetype") == "cover" or s.get("layout_type") == "cover") for s in slides_raw)
            if not has_cover:
                cover_highlights = [
                    TopicLogicValidator.enrich_short_point(t_head, topic, topic, i)
                    for i, t_head in enumerate(agenda_titles[:6], 1)
                ]
                formatted_slides.append({
                    "title": topic,
                    "subtitle": spec.get("description") or spec.get("subtitle") or "Professional tahliliy taqdimot va hisobot",
                    "layout_type": "cover",
                    "points": cover_highlights,
                    "bullets": cover_highlights,
                    "speaker_notes": f"Hurmatli tinglovchilar! Bugungi taqdimotimiz: «{topic}» mavzusiga bag'ishlanadi.",
                })

            # 2. Reja (Mundarija / Agenda) slaydi
            has_agenda = any((s.get("archetype") == "agenda" or s.get("layout_type") == "agenda") for s in slides_raw)
            if not has_agenda and (target_count >= 5 or len(slides_raw) >= 3):
                agenda_rich = [
                    f"{t_head}: Asosiy tushunchalar, tahliliy ko'rsatkichlar va amaliy yechimlar"
                    if ":" not in t_head else t_head
                    for t_head in agenda_titles[:6]
                ]
                formatted_slides.append({
                    "title": "Mundarija",
                    "subtitle": "Taqdimotning asosiy bosqichlari va tuzilishi",
                    "layout_type": "agenda",
                    "points": agenda_rich,
                    "bullets": agenda_rich,
                    "speaker_notes": "Taqdimotimiz quyidagi asosiy reja bandlarini qamrab oladi.",
                })

            # 3. Asosiy mazmunli slaydlar (turli arxetiplar rotatsiyasi bilan)
            archetype_cycle = ["cards_grid", "timeline_steps", "comparison_vs", "metrics_stats", "theory_concept"]
            for s_i, s in enumerate(slides_raw):
                s_atype = s.get("archetype") or s.get("layout_type")
                if s_atype == "cover":
                    continue
                if not s_atype or (s_atype == "cards_grid" and s_i > 0):
                    s_atype = archetype_cycle[s_i % len(archetype_cycle)]

                s_title = s.get("heading") or s.get("title") or topic
                raw_bullets = (
                    s.get("points")
                    or s.get("bullets")
                    or [f"{c.get('title', '')}: {c.get('desc', '')}".strip(": ") for c in s.get("cards", []) if isinstance(c, dict)]
                    or [s.get("content", "")]
                )
                enriched_bullets = [
                    TopicLogicValidator.enrich_short_point(str(b), s_title, topic, b_i)
                    for b_i, b in enumerate(raw_bullets, 1)
                    if b and str(b).strip()
                ]
                slide_dict = {
                    "title": s_title,
                    "subtitle": s.get("subtitle") or "",
                    "layout_type": "table_data" if (s.get("table") or s.get("table_data")) else s_atype,
                    "bullets": enriched_bullets,
                    "points": enriched_bullets,
                    "speaker_notes": s.get("speaker_notes") or s.get("speaker_note") or s.get("notes") or "",
                }
                if s.get("table") or s.get("table_data"):
                    slide_dict["table"] = s.get("table") or s.get("table_data")
                if s.get("metrics"):
                    slide_dict["metrics"] = s.get("metrics")
                if s.get("steps"):
                    slide_dict["steps"] = s.get("steps")
                if s.get("cards"):
                    slide_dict["cards"] = s.get("cards")
                formatted_slides.append(slide_dict)

            # 4. Xulosa (Conclusion / Q&A) slaydi
            has_conclusion = any(
                s.get("archetype") in ("conclusion_sources", "conclusion")
                or s.get("layout_type") in ("conclusion_sources", "conclusion")
                for s in slides_raw
            )
            if not has_conclusion:
                formatted_slides.append({
                    "title": "Xulosa va Tavsiyalar",
                    "subtitle": "Asosiy xulosalar va amaliy takliflar",
                    "layout_type": "conclusion_sources",
                    "points": [
                        f"Ilmiy xulosa: «{topic}» bo'yicha asosiy nazariy va tahliliy natijalar tizimli ravishda umumlashtirildi.",
                        "Amaliy tatbiq: Zamonaviy texnologik yondashuvlarni amaliyotga joriy etishning ustuvor istiqbollari belgilandi.",
                        "Yakuniy qism: E'tiboringiz uchun rahmat! Savollar, fikr-mulohazalar va takliflar uchun tayyormiz.",
                    ],
                    "bullets": [
                        f"Ilmiy xulosa: «{topic}» bo'yicha asosiy nazariy va tahliliy natijalar tizimli ravishda umumlashtirildi.",
                        "Amaliy tatbiq: Zamonaviy texnologik yondashuvlarni amaliyotga joriy etishning ustuvor istiqbollari belgilandi.",
                        "Yakuniy qism: E'tiboringiz uchun rahmat! Savollar, fikr-mulohazalar va takliflar uchun tayyormiz.",
                    ],
                    "speaker_notes": "E'tiboringiz uchun katta rahmat! Mavzu yuzasidan savol va takliflaringizni berishingiz mumkin.",
                })

            content_data = {
                "topic": topic,
                "slides": formatted_slides,
                "donor_pptx_paths": spec.get("donor_pptx_paths") or [],
            }

            res = replace_presentation_content(tpl_path, content_data, out_path, engine_mode="harmonized")

            if os.path.exists(res) and os.path.getsize(res) > 10000:
                logger.info(f"Authentic presentation created from {tpl_path} ({len(formatted_slides)} slides): {res}")
                return res
    except Exception as e:
        logger.warning(f"Authentic template generation failed, falling back to Penpot: {e}")

    return _generate_synthetic_penpot_pptx(spec, out_path)
