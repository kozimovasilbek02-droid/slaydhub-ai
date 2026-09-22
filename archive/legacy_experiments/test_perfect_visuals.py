# -*- coding: utf-8 -*-
import os
import sys
import re
import math
from typing import Dict, Any, List, Tuple, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.academic_matcher import get_academic_matcher

# Known dummy placeholder strings commonly found in PowerPoint templates
DUMMY_PATTERNS = [
    r'lorem\s+ipsum',
    r'make\s+a\s+big\s+impact',
    r'content\s+here',
    r'briefly\s+describe',
    r'this\s+is\s+your\s+big\s+image',
    r'this\s+is\s+your\s+chart\s+slide',
    r'table\s+slide',
    r'presentationgo(\.com)?',
    r'slidescarnival(\.com)?',
    r'slidenest(\.com)?',
    r'slidesmania(\.com)?',
    r'pptmon(\.com)?',
    r'thepopp(\.com)?',
    r'allppt(\.com)?',
    r'free\s+powerpoint\s+templates?',
    r'www\.[a-z0-9\-]+\.com',
    r'^(paris|berlin|tokyo|helsinki|bogota|new\s+york|oslo|london|madrid|rome|sydney|chicago|vancouver)(\s*0?[1-9])?$',
    r'^(cfo|coo|ceo|cto|cmo|vp|director|manager|lead)$',
    r'^(karla\s+gard|john\s+doe|jane\s+smith|alex\s+smith|andrew\s+childs|jean\s+marcus|paola\s+frisch|graham\s+gold)$',
    r'the\s+introduction\s+of\s+computers\s+was\s+the\s+basis',
    r'describe\s+your\s+content\s+here',
    r'carrer\s+d.enric\s+granados',
    r'jobs@yourbestbrand\.com',
    r'monday\s+to\s+saturday',
    r'let.s\s+keep\s+in\s+touch',
    r'book\s+now',
    r'write\s+your\s+subtitle\s+here',
    r'three\s+option\s+slide',
    r'four\s+option\s+slide',
    r'six\s+concept\s+slide',
    r'venn\s+diagram',
    r'bullet\s+points',
    r'steps\s+infographic',
    r'timeline\s+infographic',
    r'comparison\s+slide'
]

def is_dummy_text(text: str) -> bool:
    t_clean = text.lower().strip()
    if not t_clean:
        return True
    for pat in DUMMY_PATTERNS:
        if re.search(pat, t_clean):
            return True
    return False

def calculate_safe_font_size(
    text: str,
    width_pt: float,
    height_pt: float,
    max_initial_fs: float = 16.0,
    min_fs: float = 8.5,
    is_title: bool = False
) -> float:
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    if not lines:
        return max_initial_fs

    usable_w = max(15.0, width_pt - 8.0)
    usable_h = max(10.0, height_pt - 6.0)

    font_size = max_initial_fs
    while font_size >= min_fs:
        char_w = font_size * 0.52
        chars_per_line = max(1, int(usable_w / char_w))

        total_lines = 0
        for l in lines:
            words = l.split()
            if not words:
                continue
            cur_line_len = 0
            for w in words:
                w_len = len(w)
                if cur_line_len == 0:
                    cur_line_len = w_len
                elif cur_line_len + 1 + w_len <= chars_per_line:
                    cur_line_len += 1 + w_len
                else:
                    total_lines += 1
                    cur_line_len = w_len
            if cur_line_len > 0:
                total_lines += 1

        line_h = font_size * 1.25
        para_spacing = font_size * 0.35 if len(lines) > 1 else 0
        total_h = (total_lines * line_h) + ((len(lines) - 1) * para_spacing)

        if total_h <= usable_h * 0.94:
            return round(font_size, 1)

        font_size -= 0.5

    return min_fs

def enable_xml_autofit(text_frame) -> None:
    try:
        bodyPr = text_frame._txBody.bodyPr
        for child in list(bodyPr):
            if child.tag.endswith(('noAutofit', 'spAutoFit', 'normAutofit')):
                bodyPr.remove(child)
        normAutofit = OxmlElement('a:normAutofit')
        bodyPr.append(normAutofit)
    except Exception:
        pass

def set_shape_text(
    shape,
    text: str,
    is_title: bool = False,
    is_card_title: bool = False,
    font_size_override: Optional[float] = None,
    color_override: Optional[RGBColor] = None,
    bold_override: Optional[bool] = None
) -> None:
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    enable_xml_autofit(tf)

    orig_font_name = None
    orig_font_size_pt = None
    orig_color = None
    orig_bold = False

    if tf.paragraphs and tf.paragraphs[0].runs:
        r0 = tf.paragraphs[0].runs[0]
        orig_font_name = r0.font.name
        if r0.font.size:
            orig_font_size_pt = r0.font.size.pt
        orig_bold = r0.font.bold
        try:
            if r0.font.color and r0.font.color.type:
                orig_color = r0.font.color.rgb
        except Exception:
            pass

    lines = [line.strip() for line in text.split('\n') if line.strip()]
    if not lines:
        tf.text = ""
        return

    shape_w_pt = shape.width.pt if hasattr(shape.width, 'pt') else float(shape.width) / 12700.0
    shape_h_pt = shape.height.pt if hasattr(shape.height, 'pt') else float(shape.height) / 12700.0

    if is_title:
        max_init_fs = 30.0
        min_fs = 16.0
    elif is_card_title:
        max_init_fs = 17.0
        min_fs = 11.0
    else:
        max_init_fs = 14.0
        min_fs = 9.0

    if font_size_override:
        max_init_fs = font_size_override

    scaled_font_size = calculate_safe_font_size(
        text=text,
        width_pt=shape_w_pt,
        height_pt=shape_h_pt,
        max_initial_fs=max_init_fs,
        min_fs=min_fs,
        is_title=is_title
    )

    tf.text = ""
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = line
        p.line_spacing = 1.15
        p.space_after = Pt(2) if len(lines) > 1 else Pt(0)
        
        for r in p.runs:
            if orig_font_name:
                r.font.name = orig_font_name
            r.font.size = Pt(scaled_font_size)
            if color_override:
                r.font.color.rgb = color_override
            elif orig_color:
                r.font.color.rgb = orig_color
            if bold_override is not None:
                r.font.bold = bold_override
            elif is_title or is_card_title:
                r.font.bold = True
            elif orig_bold:
                r.font.bold = orig_bold

print("Base text tools configured.")

class VisualDeckRenderer:
    @staticmethod
    def render_presentation(template_path: str, content_data: Dict[str, Any], output_path: str) -> str:
        prs = Presentation(template_path)
        sw = prs.slide_width
        sh = prs.slide_height

        # Chosen template slide indices for each blueprint slide
        # 1: Cover (Slide 1, index 0)
        # 2: Agenda (Slide 2, index 1)
        # 3: Theory & Concept (Slide 3, index 2)
        # 4: Cards Grid / Zaifliklar (Slide 9, index 8)
        # 5: Diagram Anatomy / PQC Arxitekturasi (Slide 13, index 12)
        # 6: Comparison VS (Slide 12, index 11)
        # 7: Timeline Steps (Slide 20, index 19)
        # 8: Metrics & Stats (Slide 18, index 17)
        # 9: Q&A Discussion (Slide 14, index 13)
        # 10: Conclusion & Sources (Slide 25, index 24)

        chosen_indices = [0, 1, 2, 8, 12, 11, 19, 17, 13, 24]
        print(f"Selected native template slides: {chosen_indices}")

        # Prune unused slides
        to_delete = [i for i in range(len(prs.slides)) if i not in set(chosen_indices)]
        for del_idx in sorted(to_delete, reverse=True):
            rId = prs.slides._sldIdLst[del_idx].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[del_idx]

        sorted_chosen = sorted(chosen_indices)
        sldIdLst = prs.slides._sldIdLst
        sldId_elements = list(sldIdLst)
        orig_to_elem = {sorted_chosen[i]: sldId_elements[i] for i in range(len(sorted_chosen))}

        for el in sldId_elements:
            sldIdLst.remove(el)
        for orig_idx in chosen_indices:
            sldIdLst.append(orig_to_elem[orig_idx])

        # Render each slide with customized visual perfection
        slides_data = content_data.get("slides", [])
        topic = content_data.get("topic", "Akademik Taqdimot")

        # Slide 1: Cover
        s1 = prs.slides[0]
        s1_data = slides_data[0]
        for shape in s1.shapes:
            if shape.has_text_frame and "quantum" in shape.text_frame.text.lower():
                set_shape_text(shape, s1_data.get("title", topic), is_title=True, font_size_override=42.0)

        # Slide 2: Agenda
        s2 = prs.slides[1]
        s2_data = slides_data[1]
        pts = s2_data.get("points", [])
        # Title: CuadroTexto 24
        for shape in s2.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            t = shape.text_frame.text.strip()
            if name == "CuadroTexto 24" or "table of contents" in t.lower():
                set_shape_text(shape, s2_data.get("title", "Mundarija va Asosiy Yo'nalishlar"), is_title=True, font_size_override=28.0)
            elif name == "CuadroTexto 5": # Card 1 Title
                p_text = pts[0] if len(pts) > 0 else ""
                clean_t = re.sub(r'^\d+[\.\)]\s*', '', p_text).split(":", 1)[0]
                set_shape_text(shape, clean_t, is_card_title=True, font_size_override=16.0)
            elif name == "CuadroTexto 4": # Card 1 Desc
                p_text = pts[0] if len(pts) > 0 else ""
                clean_d = p_text.split(":", 1)[1].strip() if ":" in p_text else p_text
                set_shape_text(shape, clean_d, font_size_override=13.0)
            elif name == "CuadroTexto 11": # Card 2 Title
                p_text = pts[1] if len(pts) > 1 else ""
                clean_t = re.sub(r'^\d+[\.\)]\s*', '', p_text).split(":", 1)[0]
                set_shape_text(shape, clean_t, is_card_title=True, font_size_override=16.0)
            elif name == "CuadroTexto 10": # Card 2 Desc
                p_text = pts[1] if len(pts) > 1 else ""
                clean_d = p_text.split(":", 1)[1].strip() if ":" in p_text else p_text
                set_shape_text(shape, clean_d, font_size_override=13.0)
            elif name == "CuadroTexto 16": # Card 3 Title
                p_text = pts[2] if len(pts) > 2 else ""
                clean_t = re.sub(r'^\d+[\.\)]\s*', '', p_text).split(":", 1)[0]
                set_shape_text(shape, clean_t, is_card_title=True, font_size_override=16.0)
            elif name == "CuadroTexto 15": # Card 3 Desc
                p_text = pts[2] if len(pts) > 2 else ""
                clean_d = p_text.split(":", 1)[1].strip() if ":" in p_text else p_text
                set_shape_text(shape, clean_d, font_size_override=13.0)
            elif name == "CuadroTexto 21": # Card 4 Title
                p_text = pts[3] if len(pts) > 3 else ""
                clean_t = re.sub(r'^\d+[\.\)]\s*', '', p_text).split(":", 1)[0]
                set_shape_text(shape, clean_t, is_card_title=True, font_size_override=16.0)
            elif name == "CuadroTexto 20": # Card 4 Desc
                p_text = pts[3] if len(pts) > 3 else ""
                clean_d = p_text.split(":", 1)[1].strip() if ":" in p_text else p_text
                set_shape_text(shape, clean_d, font_size_override=13.0)

        # Slide 3: Theory & Concept (Circular Photo Slide)
        s3 = prs.slides[2]
        s3_data = slides_data[2]
        c_name = s3_data.get("concept_name", "Kvant Superpozitsiyasi va Kubit")
        definition = s3_data.get("definition", "")
        formula = s3_data.get("key_formula_or_law", "")
        importance = s3_data.get("importance", "")
        for shape in s3.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "TextBox 3": # Upper box -> TITLE
                set_shape_text(shape, s3_data.get("title", "Kvant Hisoblashlari Nazariy Asoslari"), is_title=True, font_size_override=28.0)
            elif name == "TextBox 4": # Lower box -> STRUCTURED BODY
                body_lines = [
                    f"• {c_name}:",
                    f"  {definition}",
                    f"",
                    f"• Asosiy Qonuniyat / Formula:",
                    f"  {formula}",
                    f"",
                    f"• Ilmiy va Amaliy Ahamiyati:",
                    f"  {importance}"
                ]
                set_shape_text(shape, "\n".join(body_lines), font_size_override=13.5)

        # Slide 4: Cards Grid / Zaif Asimmetrik Tizimlar (Slide 9 in template: 6 Cards)
        s4 = prs.slides[3]
        s4_data = slides_data[3]
        cards = s4_data.get("cards", [])
        for shape in s4.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "TITLE":
                set_shape_text(shape, s4_data.get("title", "Zaiflikka Uchrovchi Asimmetrik Tizimlar"), is_title=True, font_size_override=28.0)
            elif name == "TITLE 01" and len(cards) > 0:
                set_shape_text(shape, cards[0].get("title", ""), is_card_title=True, font_size_override=16.0)
            elif name == "BODY 01" and len(cards) > 0:
                set_shape_text(shape, cards[0].get("desc", ""), font_size_override=13.0)
            elif name == "TITLE 02" and len(cards) > 1:
                set_shape_text(shape, cards[1].get("title", ""), is_card_title=True, font_size_override=16.0)
            elif name == "BODY 02" and len(cards) > 1:
                set_shape_text(shape, cards[1].get("desc", ""), font_size_override=13.0)
            elif name == "TITLE 04" and len(cards) > 2:
                set_shape_text(shape, cards[2].get("title", ""), is_card_title=True, font_size_override=16.0)
            elif name == "BODY 04" and len(cards) > 2:
                set_shape_text(shape, cards[2].get("desc", ""), font_size_override=13.0)
            elif name in ["TITLE 05", "BODY 05"]:
                set_shape_text(shape, "") # Clean unused

        # Slide 5: Diagram Anatomy / PQC Arxitekturasi (Slide 13 in template: 3 Option Cards A, B, C)
        s5 = prs.slides[4]
        s5_data = slides_data[4]
        comps = s5_data.get("components", [])
        for shape in s5.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "CuadroTexto 2":
                set_shape_text(shape, s5_data.get("title", "Post-Kvant Kriptografiyasi (PQC) Arxitekturasi"), is_title=True, font_size_override=28.0)
            elif name == "CuadroTexto 3":
                set_shape_text(shape, "Kvant tahdidlariga chidamli xalqaro kriptografik modellar", font_size_override=16.0)
            elif name == "TextBox 1172" and len(comps) > 0: # Option A Title
                set_shape_text(shape, comps[0].get("name", ""), is_card_title=True, font_size_override=16.0)
            elif name == "TextBox 1173" and len(comps) > 0: # Option A Body
                set_shape_text(shape, comps[0].get("function", ""), font_size_override=13.0)
            elif name == "TextBox 1177" and len(comps) > 1: # Option B Title
                set_shape_text(shape, comps[1].get("name", ""), is_card_title=True, font_size_override=16.0)
            elif name == "TextBox 1178" and len(comps) > 1: # Option B Body
                set_shape_text(shape, comps[1].get("function", ""), font_size_override=13.0)
            elif name == "TextBox 1179" and len(comps) > 2: # Option C Title
                set_shape_text(shape, comps[2].get("name", ""), is_card_title=True, font_size_override=16.0)
            elif name == "TextBox 1180" and len(comps) > 2: # Option C Body
                set_shape_text(shape, comps[2].get("function", ""), font_size_override=13.0)

        # Slide 6: Comparison VS (Slide 12 in template: 2 Columns with diamonds and sub-points)
        s6 = prs.slides[5]
        s6_data = slides_data[5]
        left_t = s6_data.get("left_title", "Klassik Asimmetrik Kriptografiya")
        left_pts = s6_data.get("left_points", [])
        right_t = s6_data.get("right_title", "Post-Kvant Kriptografiyasi (PQC)")
        right_pts = s6_data.get("right_points", [])

        for shape in s6.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "TextBox 15": # Top Header
                set_shape_text(shape, s6_data.get("title", "Klassik Kriptografiya vs Post-Kvant Kriptografiyasi"), is_title=True, font_size_override=28.0)
            elif name == "TextBox 30": # Left Diamond Main Header
                set_shape_text(shape, left_t, is_card_title=True, font_size_override=18.0)
            elif name == "TextBox 10": # Left Sub-point 1 Title
                set_shape_text(shape, "Faktorizatsiya & Logarifm", is_card_title=True, font_size_override=14.0)
            elif name == "TextBox 11": # Left Sub-point 1 Desc
                set_shape_text(shape, left_pts[0] if len(left_pts) > 0 else "Faktorizatsiyaga tayanadi", font_size_override=12.5)
            elif name == "TextBox 7": # Left Sub-point 2 Title
                set_shape_text(shape, "Kvant Zaifligi", is_card_title=True, font_size_override=14.0)
            elif name == "TextBox 6": # Left Sub-point 2 Desc
                set_shape_text(shape, left_pts[1] if len(left_pts) > 1 else "Kvant kompyuterlarida buziladi", font_size_override=12.5)
            elif name == "TextBox 19": # Left Sub-point 3 Title
                set_shape_text(shape, "Kalit Parametri", is_card_title=True, font_size_override=14.0)
            elif name == "TextBox 20": # Left Sub-point 3 Desc
                set_shape_text(shape, left_pts[2] if len(left_pts) > 2 else "Kichik kalit va yuqori tezlik", font_size_override=12.5)
            
            elif name == "TextBox 31": # Right Diamond Main Header
                set_shape_text(shape, right_t, is_card_title=True, font_size_override=18.0)
            elif name == "TextBox 23": # Right Sub-point 1 Title
                set_shape_text(shape, "Panjaralar & Xeshlar", is_card_title=True, font_size_override=14.0)
            elif name == "TextBox 24": # Right Sub-point 1 Desc
                set_shape_text(shape, right_pts[0] if len(right_pts) > 0 else "Panjaralarga asoslangan", font_size_override=12.5)
            elif name == "TextBox 26": # Right Sub-point 2 Title
                set_shape_text(shape, "Mutlaq Himoya", is_card_title=True, font_size_override=14.0)
            elif name == "TextBox 18": # Right Sub-point 2 Desc
                set_shape_text(shape, right_pts[1] if len(right_pts) > 1 else "Kvant hujumlariga chidamli", font_size_override=12.5)
            elif name == "TextBox 28": # Right Sub-point 3 Title
                set_shape_text(shape, "Resurs Talabi", is_card_title=True, font_size_override=14.0)
            elif name == "TextBox 29": # Right Sub-point 3 Desc
                set_shape_text(shape, right_pts[2] if len(right_pts) > 2 else "Kattaroq kalit hajmi talab etiladi", font_size_override=12.5)

        # Slide 7: Timeline Steps (Slide 20 in template: Staircase Infographic)
        s7 = prs.slides[6]
        s7_data = slides_data[6]
        steps = s7_data.get("steps", [])
        for shape in s7.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "TextBox 1": # Title
                set_shape_text(shape, s7_data.get("title", "PQC Tizimlariga O'tish Bosqichlari"), is_title=True, font_size_override=28.0)
            elif name == "TextBox 2": # Subtitle
                set_shape_text(shape, "") # Clean
            elif name == "TextBox 49" and len(steps) > 0: # Step 1 Title
                set_shape_text(shape, steps[0].get("title", ""), is_card_title=True, font_size_override=16.0)
            elif name == "TextBox 48" and len(steps) > 0: # Step 1 Desc
                set_shape_text(shape, steps[0].get("desc", ""), font_size_override=13.0)
            elif name == "TextBox 51" and len(steps) > 1: # Step 2 Title
                set_shape_text(shape, steps[1].get("title", ""), is_card_title=True, font_size_override=16.0)
            elif name == "TextBox 50" and len(steps) > 1: # Step 2 Desc
                set_shape_text(shape, steps[1].get("desc", ""), font_size_override=13.0)
            elif name == "TextBox 53" and len(steps) > 2: # Step 3 Title
                set_shape_text(shape, steps[2].get("title", ""), is_card_title=True, font_size_override=16.0)
            elif name == "TextBox 52" and len(steps) > 2: # Step 3 Desc
                set_shape_text(shape, steps[2].get("desc", ""), font_size_override=13.0)
            elif name == "TextBox 55" and len(steps) > 3: # Step 4 Title
                set_shape_text(shape, steps[3].get("title", ""), is_card_title=True, font_size_override=16.0)
            elif name == "TextBox 54" and len(steps) > 3: # Step 4 Desc
                set_shape_text(shape, steps[3].get("desc", ""), font_size_override=13.0)

        # Slide 8: Metrics & Stats (Slide 18 in template: Venn / Stat Diagram)
        s8 = prs.slides[7]
        s8_data = slides_data[7]
        metrics = s8_data.get("metrics", [])
        for shape in s8.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "TextBox 2":
                set_shape_text(shape, s8_data.get("title", "Kvant Xavfsizlik Metriklari va Ko'rsatkichlar"), is_title=True, font_size_override=28.0)
            elif name == "TextBox 27":
                set_shape_text(shape, "")
            elif name == "TextBox 36" and len(metrics) > 0:
                set_shape_text(shape, f"{metrics[0].get('value', '')} - {metrics[0].get('label', '')}", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 37" and len(metrics) > 0:
                set_shape_text(shape, metrics[0].get("desc", ""), font_size_override=13.0)
            elif name == "TextBox 38" and len(metrics) > 1:
                set_shape_text(shape, f"{metrics[1].get('value', '')} - {metrics[1].get('label', '')}", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 39" and len(metrics) > 1:
                set_shape_text(shape, metrics[1].get("desc", ""), font_size_override=13.0)
            elif name == "TextBox 40" and len(metrics) > 2:
                set_shape_text(shape, f"{metrics[2].get('value', '')} - {metrics[2].get('label', '')}", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 41" and len(metrics) > 2:
                set_shape_text(shape, metrics[2].get("desc", ""), font_size_override=13.0)
            elif name == "TextBox 42":
                set_shape_text(shape, "256-bit - Simmetrik Himoya", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 43":
                set_shape_text(shape, "AES-256 algoritmi kvant hujumlariga qarshi 128-bitlik xavfsizlik darajasini saqlab qoladi.", font_size_override=13.0)

        # Slide 9: Q&A Discussion (Slide 14 in template: 4 Cards)
        s9 = prs.slides[8]
        s9_data = slides_data[8]
        qa_list = s9_data.get("qa_list", [])
        for shape in s9.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "TextBox 14":
                set_shape_text(shape, s9_data.get("title", "Ilmiy Munozaralar va Dolzarb Savollar"), is_title=True, font_size_override=28.0)
            elif name == "TextBox 7": # Card 1 Title
                q_text = qa_list[0].get("question", "SNDL Hujumi nima?") if len(qa_list) > 0 else "SNDL Hujumi"
                set_shape_text(shape, f"❓ {q_text}", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 6": # Card 1 Desc
                a_text = qa_list[0].get("answer", "") if len(qa_list) > 0 else ""
                set_shape_text(shape, f"💡 {a_text}", font_size_override=13.0)
            elif name == "TextBox 9": # Card 2 Title
                q_text = qa_list[1].get("question", "Qachon PQC ga o'tish kerak?") if len(qa_list) > 1 else "Migratsiya Vaqti"
                set_shape_text(shape, f"❓ {q_text}", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 8": # Card 2 Desc
                a_text = qa_list[1].get("answer", "") if len(qa_list) > 1 else ""
                set_shape_text(shape, f"💡 {a_text}", font_size_override=13.0)
            elif name == "TextBox 11": # Card 3 Title
                set_shape_text(shape, "❓ Simmetrik Shifrlash Xavfsizmi?", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 16": # Card 3 Desc
                set_shape_text(shape, "💡 AES-256 kvant hisoblashlarida Grover algoritmidan so'ng ham 128-bitlik samarali himoyani to'liq ta'minlaydi.", font_size_override=13.0)
            elif name == "TextBox 13": # Card 4 Title
                set_shape_text(shape, "❓ Asosiy PQC Standartlari Qaysilar?", is_card_title=True, font_size_override=15.0)
            elif name == "TextBox 17": # Card 4 Desc
                set_shape_text(shape, "💡 NIST tasdiqlagan ML-KEM (Kyber) kalit almashish va ML-DSA (Dilithium) raqamli imzo standartlaridir.", font_size_override=13.0)

        # Slide 10: Conclusion & Sources (Slide 25 in template)
        s10 = prs.slides[9]
        s10_data = slides_data[9]
        takeaways = s10_data.get("key_takeaways", [])
        refs = s10_data.get("references", [])
        for shape in s10.shapes:
            if not shape.has_text_frame:
                continue
            name = shape.name
            if name == "TITLE 01":
                set_shape_text(shape, s10_data.get("title", "Ilmiy Xulosa va Foydalanilgan Adabiyotlar"), is_title=True, font_size_override=26.0)
            elif name == "BODY 01":
                set_shape_text(shape, "🎯 Asosiy Ilmiy Xulosalar", is_card_title=True, font_size_override=16.0)
            elif name == "TITLE 02":
                t_lines = "\n".join([f"• {t}" for t in takeaways])
                set_shape_text(shape, t_lines, font_size_override=12.5)
            elif name == "TITLE":
                set_shape_text(shape, "📚 Foydalanilgan Adabiyotlar", is_card_title=True, font_size_override=16.0)
            elif name == "SUBTITLE":
                r_lines = "\n".join([f"📖 {r}" for r in refs])
                set_shape_text(shape, r_lines, font_size_override=11.5)
            elif name in ["BODY 02", "BODY 03", "BODY 04", "TITLE 03", "TITLE 04"]:
                set_shape_text(shape, "") # Clean contact dummy

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        print(f"\n🎉 MASTERPIECE PPTX Generated: {output_path} ({os.path.getsize(output_path):,} bytes)")
        return output_path

matcher = get_academic_matcher()
bp = matcher.build_blueprint("Kvant kompyuterlari va asimmetrik kriptografiya", slide_count=10, language="uz")
template_path = bp["primary_pptx_path"]

sample_content = {
    "topic": "Kvant kompyuterlari va asimmetrik kriptografiya",
    "slides": [
        {
            "slide_number": 1,
            "layout_type": "cover",
            "title": "Kvant kompyuterlari va asimmetrik kriptografiya",
            "subtitle": "Kvant hisoblashlari davrida axborot xavfsizligi arxitekturasi | 2026-yil"
        },
        {
            "slide_number": 2,
            "layout_type": "agenda",
            "title": "Mundarija va Asosiy Yo'nalishlar",
            "points": [
                "1. Kvant hisoblashlarining matematik asoslari: Kubitlar va superpozitsiya tahlili",
                "2. Shor algoritmi va RSA zaifligi: Asimmetrik kalitlarning buzilish mexanizmi",
                "3. Post-kvant kriptografiyasi (PQC): Yangi xavfsiz shifrlash standartlari",
                "4. Xulosalar va amaliy tavsiyalar: Milliy xavfsizlik va migratsiya bosqichlari"
            ]
        },
        {
            "slide_number": 3,
            "layout_type": "theory_concept",
            "title": "Kvant Hisoblashlari Nazariy Asoslari",
            "concept_name": "Kvant Superpozitsiyasi va Kubit",
            "definition": "Klassik bitlardan farqli o'laroq, bir vaqtning o'zida |0> va |1> holatlarining chiziqli kombinatsiyasida bo'la oluvchi axborot birligi.",
            "key_formula_or_law": "|psi> = alpha|0> + beta|1>, |alpha|^2 + |beta|^2 = 1",
            "importance": "Eksponentsial parallel hisoblash imkoniyatini taqdim etadi."
        },
        {
            "slide_number": 4,
            "layout_type": "cards_grid",
            "title": "Zaiflikka Uchrovchi Asimmetrik Tizimlar",
            "cards": [
                {"title": "RSA Shifrlash", "desc": "Katta sonlarni tub ko'paytuvchilarga ajratish Shor algoritmi yordamida polinomial vaqtda yechiladi."},
                {"title": "ECC (Elliptik Egri Chiziqlar)", "desc": "Diskret logarifm masalasi kvant kompyuterlari tomonidan to'liq buziladi."},
                {"title": "Diffie-Hellman Kalit Almashinuvi", "desc": "Klassik kalit almashish protokoli kvant hujumlariga qarshi mutlaqo himoyasiz."}
            ]
        },
        {
            "slide_number": 5,
            "layout_type": "diagram_anatomy",
            "title": "Post-Kvant Kriptografiyasi (PQC) Arxitekturasi",
            "components": [
                {"name": "Lattice-Based Kriptografiya (Kyber/Dilithium)", "function": "Ko'p o'lchovli panjaralardagi eng qisqa vektor masalasiga asoslangan xavfsiz arxitektura."},
                {"name": "Hash-Based Imzolar (SPHINCS+)", "function": "Bir martalik kriptografik xesh funksiyalar daraxtiga tayanuvchi mustahkam himoya."},
                {"name": "Code-Based Tizimlar (Classic McEliece)", "function": "Xatoliklarni tuzatuvchi chiziqli kodlar orqali shifrlash."}
            ]
        },
        {
            "slide_number": 6,
            "layout_type": "comparison_vs",
            "title": "Klassik Kriptografiya vs Post-Kvant Kriptografiyasi",
            "left_title": "Klassik Asimmetrik Kriptografiya",
            "left_points": [
                "Faktorizatsiya va diskret logarifm masalasiga tayanadi.",
                "Shor algoritmi yordamida kvant kompyuterlari tomonidan to'liq buziladi.",
                "Kichik kalit hajmi va yuqori hisoblash tezligiga ega."
            ],
            "right_title": "Post-Kvant Kriptografiyasi (PQC)",
            "right_points": [
                "Ko'p o'lchovli panjaralar va xesh daraxtlariga asoslangan.",
                "Kvant va klassik superkompyuter hujumlariga 100% chidamli.",
                "Kattaroq kalit hajmi talab etiladi, ammo mutlaq xavfsiz."
            ]
        },
        {
            "slide_number": 7,
            "layout_type": "timeline_steps",
            "title": "PQC Tizimlariga O'tish Bosqichlari",
            "steps": [
                {"title": "1. Xatarlarni Baholash", "desc": "Mavjud axborot tizimlaridagi asimmetrik shifrlash nuqtalarini xatlovdan o'tkazish."},
                {"title": "2. Gibrid Kriptografiya", "desc": "Klassik RSA/ECC va PQC algoritmlarini parallel qo'llash."},
                {"title": "3. Standartlashtirish", "desc": "NIST tomonidan tasdiqlangan Kyber va Dilithium standartlariga to'liq o'tish."},
                {"title": "4. To'liq Kvant Himoyasi", "desc": "Infratuzilmani kvantga chidamli arxitektura bilan to'liq yangilash."}
            ]
        },
        {
            "slide_number": 8,
            "layout_type": "metrics_stats",
            "title": "Kvant Xavfsizlik Metriklari va Ko'rsatkichlar",
            "metrics": [
                {"value": "2048-bit", "label": "RSA Zaiflik Chegarasi", "desc": "Taxminan 4000 mantiqiy kubit talab etiladi."},
                {"value": "99.9%", "label": "Lattice PQC Ishonchliligi", "desc": "NIST xalqaro testlarida tasdiqlangan xavfsizlik."},
                {"value": "2030+", "label": "Kvant Taxdidi Davri", "desc": "Kvant ustunligiga erishish kutilayotgan davr."}
            ]
        },
        {
            "slide_number": 9,
            "layout_type": "qa_discussion",
            "title": "Ilmiy Munozaralar va Dolzarb Savollar",
            "qa_list": [
                {"question": "SNDL Hujumi nima?", "answer": "'Hozir saqlab ol, keyin shifrni och' (SNDL) hujumi orqali buzg'unchilar bugun shifrlangan ma'lumotlarni yozib oladi va kelajakda kvant kompyuterida shifrini ochadi."},
                {"question": "Qachon PQC ga o'tish kerak?", "answer": "Zudlik bilan! Global axborot infratuzilmasini yangi standartlarga o'tkazish 5-10 yil vaqt talab qiladi."}
            ]
        },
        {
            "slide_number": 10,
            "layout_type": "conclusion_sources",
            "title": "Ilmiy Xulosalar va Foydalanilgan Adabiyotlar",
            "key_takeaways": [
                "Kvant kompyuterlari asimmetrik kriptografiya xavfsizligiga tubdan xavf soladi.",
                "PQC (Post-Quantum Cryptography) standartlariga o'tish kechiktirib bo'lmas strategik vazifadir.",
                "Gibrid xavfsizlik protokollari orqali xavflarni minimallashtirish zarur."
            ],
            "references": [
                "1. NIST Post-Quantum Cryptography Standardization (2024)",
                "2. Shor P.W. - Polynomial-Time Algorithms for Prime Factorization (SIAM, 1997)"
            ]
        }
    ]
}

out_p = os.path.abspath(r"output\academic_studio\Kvant_MASTERPIECE.pptx")
VisualDeckRenderer.render_presentation(template_path, sample_content, out_p)

# Copy to downloads
import shutil
shutil.copyfile(out_p, r"C:\Users\user\Downloads\Kvant_kompyuterlari_va_asimmetrik_kripto_Customized.pptx")
print("Copied to Downloads!")

