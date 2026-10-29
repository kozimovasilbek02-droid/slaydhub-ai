# -*- coding: utf-8 -*-
import re
from typing import Dict, Any
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

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

# ── 2. KNUTH-PLASS O'ZBEKCHA TIPOGRAFIYA HIMOYASI ────────────────────────────
UZ_ORPHAN_PREPOSITIONS_RE = re.compile(
    r"\b(va|ham|bilan|uchun|yoki|ammo|lekin|biroq|to|har|bir|shu|bu|u|o'sha|barcha)\s+",
    re.IGNORECASE,
)


def apply_knuth_plass_uz_orphans(text: str) -> str:
    """
    O'zbekcha bir bo'g'inli bog'lovchi va olmoshlar qator oxirida yetim qolib ketmasligi uchun
    uzilmas bo'shliq (\u00A0) qo'yadi. Python 3.11+ re.sub lambda orqali xatosiz ishlaydi.
    """
    if not text:
        return ""
    return UZ_ORPHAN_PREPOSITIONS_RE.sub(lambda m: f"{m.group(1)}\u00A0", str(text))


def apply_anti_word_break(paragraph) -> None:
    """PowerPoint matn bloklarida o'zbekcha so'zlarning o'rtasidan uzilishini taqiqlaydi."""
    try:
        pPr = paragraph._p.get_or_add_pPr()
        pPr.set("latinLnBrk", "0")
        pPr.set("eaLnBrk", "0")
        pPr.set("hangingPunct", "1")
    except Exception:
        pass


def apply_p16_morph_transition(slide) -> None:
    """
    PowerPoint Office 365 p16:morph silliq o'tish effektini (mc:AlternateContent va Fade fallback bilan)
    lxml ValueError xatosisiz to'g'ridan-to'g'ri OpenXML daraxtiga qo'shadi.
    """
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


def generate_pptx_file(spec: Dict[str, Any], out_path: str) -> str:
    """To'liq multi-mavzuli, multi-arxetipli professional .pptx fayl generatsiyasi."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    theme_name = spec.get("theme", spec.get("recommended_theme", "academic_navy"))
    palette = PENPOT_PALETTES.get(theme_name, PENPOT_PALETTES["academic_navy"])

    slides_data = spec.get("slides", [])
    if not slides_data:
        slides_data = [
            {
                "heading": "Kirish va Dolzarblik",
                "archetype": "cards_grid",
                "points": ["Asosiy maqsad va yo'nalishlar", "Rivojlanish tendensiyalari"],
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

    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(2.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1_t = tf1.paragraphs[0]
    p1_t.text = apply_knuth_plass_uz_orphans(spec.get("title", "Mavzu nomi"))
    p1_t.font.name = "Arial"
    p1_t.font.size = Pt(38 if len(p1_t.text) < 65 else 30)
    p1_t.font.bold = True
    p1_t.font.color.rgb = palette["text_light"]
    apply_anti_word_break(p1_t)

    subtitle_text = spec.get("subtitle", "")
    if subtitle_text:
        p1_sub = tf1.add_paragraph()
        p1_sub.text = apply_knuth_plass_uz_orphans(subtitle_text)
        p1_sub.font.name = "Arial"
        p1_sub.font.size = Pt(18)
        p1_sub.font.color.rgb = palette["accent"]
        p1_sub.space_before = Pt(14)
        apply_anti_word_break(p1_sub)

    s1.notes_slide.notes_text_frame.text = f"Muqova: {spec.get('title')}. Taqdimot rejasiga kirish."

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

    s2_tb = s2.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.0))
    s2_tf = s2_tb.text_frame
    s2_tf.word_wrap = True
    p2_head = s2_tf.paragraphs[0]
    p2_head.text = "Taqdimot Rejasi va Asosiy Yo'nalishlar"
    p2_head.font.name = "Arial"
    p2_head.font.size = Pt(28)
    p2_head.font.bold = True
    p2_head.font.color.rgb = palette["text_dark"]
    apply_anti_word_break(p2_head)

    agenda_items = [slide.get("heading", f"Mavzu {i+1}") for i, slide in enumerate(slides_data)]
    half = (len(agenda_items) + 1) // 2

    for col_idx, col_items in enumerate([agenda_items[:half], agenda_items[half:]]):
        col_left = Inches(1.0 + col_idx * 5.8)
        col_tb = s2.shapes.add_textbox(col_left, Inches(2.2), Inches(5.4), Inches(4.5))
        col_tf = col_tb.text_frame
        col_tf.word_wrap = True

        for i, item in enumerate(col_items):
            item_num = (i + 1) if col_idx == 0 else (half + i + 1)
            p = col_tf.paragraphs[0] if i == 0 else col_tf.add_paragraph()
            p.text = f"{item_num:02d}.  {apply_knuth_plass_uz_orphans(item)}"
            p.font.name = "Arial"
            p.font.size = Pt(14)
            p.font.bold = True
            p.font.color.rgb = palette["text_dark"]
            p.space_after = Pt(14)
            apply_anti_word_break(p)

    s2.notes_slide.notes_text_frame.text = (
        "Taqdimot mundarijasi va ko'riladigan asosiy masalalar sharhi."
    )

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
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.6), Inches(2.5), Inches(0.35)
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

        h_tb = s.shapes.add_textbox(Inches(1.0), Inches(1.1), Inches(11.333), Inches(1.0))
        h_tf = h_tb.text_frame
        h_tf.word_wrap = True
        h_p = h_tf.paragraphs[0]
        h_p.text = apply_knuth_plass_uz_orphans(slide_info.get("heading", f"Mavzu {idx}"))
        h_p.font.name = "Arial"
        h_p.font.size = Pt(24)
        h_p.font.bold = True
        h_p.font.color.rgb = palette["text_dark"]
        apply_anti_word_break(h_p)

        archetype = slide_info.get("archetype", "cards_grid")
        points = slide_info.get("points", [])
        if not points:
            points = ["Asosiy ma'lumot", "Tahlil va natijalar"]

        if archetype == "comparison_vs" and len(points) >= 2:
            card_w = Inches(5.3)
            card_h = Inches(4.5)
            lefts = [Inches(1.0), Inches(6.8)]
            for c_i, p_txt in enumerate(points[:2]):
                c = s.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, lefts[c_i], Inches(2.2), card_w, card_h
                )
                c.fill.solid()
                c.fill.fore_color.rgb = palette["card_bg"]
                c.line.color.rgb = palette["card_border"]
                c.line.width = Pt(1.5)
                ctf = c.text_frame
                ctf.word_wrap = True
                cp = ctf.paragraphs[0]
                cp.text = f"ASPEKT {c_i+1}"
                cp.font.name = "Arial"
                cp.font.size = Pt(12)
                cp.font.bold = True
                cp.font.color.rgb = palette["accent"]
                cp2 = ctf.add_paragraph()
                cp2.text = apply_knuth_plass_uz_orphans(p_txt)
                cp2.font.name = "Arial"
                cp2.font.size = Pt(14)
                cp2.font.color.rgb = palette["text_dark"]
                cp2.space_before = Pt(12)
                apply_anti_word_break(cp2)

        elif archetype == "timeline_steps" and len(points) >= 3:
            step_w = Inches(3.4)
            for c_i, p_txt in enumerate(points[:3]):
                c_left = Inches(1.0 + c_i * 3.9)
                c = s.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, c_left, Inches(2.4), step_w, Inches(4.2)
                )
                c.fill.solid()
                c.fill.fore_color.rgb = palette["card_bg"]
                c.line.color.rgb = palette["card_border"]
                c.line.width = Pt(1.5)
                ctf = c.text_frame
                ctf.word_wrap = True
                cp = ctf.paragraphs[0]
                cp.text = f"BOSQICH {c_i+1}"
                cp.font.name = "Arial"
                cp.font.size = Pt(11)
                cp.font.bold = True
                cp.font.color.rgb = palette["accent"]
                cp2 = ctf.add_paragraph()
                cp2.text = apply_knuth_plass_uz_orphans(p_txt)
                cp2.font.name = "Arial"
                cp2.font.size = Pt(13)
                cp2.font.color.rgb = palette["text_dark"]
                cp2.space_before = Pt(10)
                apply_anti_word_break(cp2)

        else:
            count = min(3, len(points))
            w = Inches(11.0 / count)
            for c_i, p_txt in enumerate(points[:count]):
                c_left = Inches(1.0 + c_i * (11.333 / count))
                c = s.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE, c_left, Inches(2.3), w, Inches(4.3)
                )
                c.fill.solid()
                c.fill.fore_color.rgb = palette["card_bg"]
                c.line.color.rgb = palette["card_border"]
                c.line.width = Pt(1.5)
                ctf = c.text_frame
                ctf.word_wrap = True
                cp = ctf.paragraphs[0]
                cp.text = f"#{c_i+1}"
                cp.font.name = "Arial"
                cp.font.size = Pt(12)
                cp.font.bold = True
                cp.font.color.rgb = palette["accent"]
                cp2 = ctf.add_paragraph()
                cp2.text = apply_knuth_plass_uz_orphans(p_txt)
                cp2.font.name = "Arial"
                cp2.font.size = Pt(13.5)
                cp2.font.color.rgb = palette["text_dark"]
                cp2.space_before = Pt(10)
                apply_anti_word_break(cp2)

        sp_note = slide_info.get("speaker_note", "")
        s.notes_slide.notes_text_frame.text = (
            sp_note or f"{slide_info.get('heading')} mavzusi bo'yicha ilmiy sharh."
        )

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

    s_end.notes_slide.notes_text_frame.text = "Xulosa va savol-javoblar uchun yakuniy slayd."

    prs.save(out_path)
    return out_path
