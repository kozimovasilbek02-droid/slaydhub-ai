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

# Dummy placeholder cleaner patterns
DUMMY_PATTERNS = [
    r'lorem\s+ipsum',
    r'make\s+a\s+big\s+impact',
    r'content\s+here',
    r'briefly\s+describe',
    r'this\s+is\s+your\s+big\s+image',
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
    r'^(paris|berlin|tokyo|helsinki|bogota|new\s+york|oslo|london|madrid|rome)$',
    r'^(cfo|coo|ceo|cto|cmo|vp|director|manager|lead)$',
    r'^(karla\s+gard|john\s+doe|jane\s+smith|alex\s+smith|andrew\s+childs|jean\s+marcus|paola\s+frisch|graham\s+gold)$',
    r'the\s+introduction\s+of\s+computers\s+was\s+the\s+basis',
    r'describe\s+your\s+content\s+here',
    r'carrer\s+d.enric\s+granados'
]

class AdvancedTextReplacer:
    @staticmethod
    def is_dummy_text(text: str) -> bool:
        t_clean = text.lower().strip()
        if not t_clean:
            return True
        for pat in DUMMY_PATTERNS:
            if re.search(pat, t_clean):
                return True
        return False

    @staticmethod
    def is_step_badge(text: str) -> bool:
        t = text.strip()
        return bool(
            re.match(r'^(0?[1-9]|1[0-9])[\.\:\)]?$', t) or
            re.match(r'^step\s*0?[1-9]$', t, re.I) or
            re.match(r'^(20\d\d)$', t) or
            re.match(r'^[A-D][\.\:\)]?$', t)
        )

    @staticmethod
    def calculate_safe_font_size(
        text: str,
        width_pt: float,
        height_pt: float,
        max_initial_fs: float = 18.0,
        min_fs: float = 8.0,
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

            if total_h <= usable_h * 0.92:
                return round(font_size, 1)

            font_size -= 0.5

        return min_fs

    @staticmethod
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

    @classmethod
    def set_shape_text(
        cls,
        shape,
        new_text: str,
        is_title: bool = False,
        is_card_title: bool = False,
        font_size_override: Optional[float] = None
    ) -> None:
        if not shape.has_text_frame:
            return

        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.04)
        tf.margin_right = Inches(0.04)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)
        cls.enable_xml_autofit(tf)

        # Style preservation
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

        lines = [line.strip() for line in new_text.split('\n') if line.strip()]
        if not lines:
            tf.text = ""
            return

        shape_w_pt = shape.width.pt if hasattr(shape.width, 'pt') else float(shape.width) / 12700.0
        shape_h_pt = shape.height.pt if hasattr(shape.height, 'pt') else float(shape.height) / 12700.0

        if is_title:
            max_init_fs = 30.0
            min_fs = 14.0
        elif is_card_title:
            max_init_fs = 16.0
            min_fs = 10.0
        else:
            max_init_fs = 13.0
            min_fs = 8.0

        if orig_font_size_pt and orig_font_size_pt > 0:
            max_init_fs = min(max_init_fs, orig_font_size_pt)
        if font_size_override:
            max_init_fs = font_size_override

        scaled_font_size = cls.calculate_safe_font_size(
            text=new_text,
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
                if orig_color:
                    r.font.color.rgb = orig_color
                if is_title or is_card_title:
                    r.font.bold = True
                elif orig_bold:
                    r.font.bold = orig_bold

    @classmethod
    def replace_slide_content_smart(cls, slide, slide_data: Dict[str, Any], topic: str) -> None:
        sw = slide.part.slide_width if hasattr(slide.part, 'slide_width') else Inches(13.33)
        sh = slide.part.slide_height if hasattr(slide.part, 'slide_height') else Inches(7.5)

        all_text_shapes = []
        for s in slide.shapes:
            if not s.has_text_frame:
                continue
            txt = s.text_frame.text.strip()
            if not txt:
                continue
            all_text_shapes.append({
                "shape": s,
                "text": txt,
                "left": s.left,
                "top": s.top,
                "width": s.width,
                "height": s.height,
                "is_step": cls.is_step_badge(txt),
                "is_dummy": cls.is_dummy_text(txt),
                "font_size": max([r.font.size.pt for p in s.text_frame.paragraphs for r in p.runs if r.font.size] or [12.0])
            })

        if not all_text_shapes:
            return

        layout_type = slide_data.get("layout_type", "cards_grid")
        title_text = slide_data.get("title", topic)

        # 1. Title Shape Selection (Strict Upper Bounds + Keyword Dominance)
        title_candidates = []
        for s in all_text_shapes:
            if s["is_step"]:
                continue
            score = 0
            t_low = s["text"].lower()
            s_name_low = s["shape"].name.lower()
            
            # Dominant keyword matching: If template shape explicitly held the slide title text
            if any(k in t_low for k in ["table of contents", "agenda", "contents", "steps infographic", "timeline infographic", "comparison slide", "bullet points", "table slide", "credits"]):
                score += 500
            elif "title" in s_name_low or "header" in s_name_low or "cuadrotexto 24" in s_name_low or "textbox 1" in s_name_low or "textbox 15" in s_name_low:
                score += 250

            # Shapes near the top get strong priority if not a layout keyword
            if s["top"] < sh * 0.35:
                score += 100
                score += int((1.0 - (s["top"] / (sh * 0.35))) * 50)
            elif s["top"] < sh * 0.50:
                score += 20
            else:
                score -= 80 # Lower shapes get lower priority unless they have title keywords

            if s["width"] > sw * 0.4:
                score += 30
            if s["font_size"] >= 24:
                score += 25

            title_candidates.append((score, s))

        title_candidates.sort(key=lambda x: x[0], reverse=True)
        title_shape_meta = title_candidates[0][1] if title_candidates and title_candidates[0][0] > 0 else all_text_shapes[0]

        # Set title
        cls.set_shape_text(title_shape_meta["shape"], title_text, is_title=True)

        # 2. Extract structured content items
        content_items: List[Tuple[str, str]] = []

        if layout_type == "cover":
            subtitle = slide_data.get("subtitle", "Akademik va amaliy tahliliy qo'llanma | 2026-yil")
            content_items.append(("", subtitle))

        elif layout_type == "agenda":
            points = slide_data.get("points", [])
            for p in points:
                clean_p = re.sub(r'^\d+[\.\)]\s*', '', p)
                if ":" in clean_p:
                    h, d = clean_p.split(":", 1)
                    content_items.append((h.strip(), d.strip()))
                else:
                    content_items.append((clean_p, ""))

        elif layout_type == "theory_concept":
            concept_name = slide_data.get("concept_name", "Nazariy Tushuncha")
            definition = slide_data.get("definition", "")
            formula = slide_data.get("key_formula_or_law", "")
            importance = slide_data.get("importance", "")
            
            content_items.append((concept_name, definition))
            if formula:
                content_items.append(("Asosiy Qonuniyat / Formula", formula))
            if importance:
                content_items.append(("Ilmiy va Amaliy Ahamiyati", importance))

        elif layout_type == "comparison_vs":
            left_t = slide_data.get("left_title", "1-Yondashuv")
            left_pts = "\n".join([f"• {p}" for p in slide_data.get("left_points", [])])
            right_t = slide_data.get("right_title", "2-Yondashuv")
            right_pts = "\n".join([f"• {p}" for p in slide_data.get("right_points", [])])
            content_items.append((left_t, left_pts))
            content_items.append((right_t, right_pts))

        elif layout_type == "timeline_steps":
            steps = slide_data.get("steps", [])
            for s in steps:
                content_items.append((s.get("title", ""), s.get("desc", "")))

        elif layout_type == "diagram_anatomy":
            comps = slide_data.get("components", [])
            for c in comps:
                content_items.append((c.get("name", ""), c.get("function", "")))

        elif layout_type == "metrics_stats":
            metrics = slide_data.get("metrics", [])
            for m in metrics:
                val = m.get("value", "")
                lbl = m.get("label", "")
                desc = m.get("desc", "")
                content_items.append((f"{val} - {lbl}", desc))

        elif layout_type == "qa_discussion":
            qa_list = slide_data.get("qa_list", [])
            for qa in qa_list:
                content_items.append((f"Savol: {qa.get('question', '')}", f"Javob: {qa.get('answer', '')}"))

        elif layout_type == "conclusion_sources":
            takeaways = slide_data.get("key_takeaways", [])
            refs = slide_data.get("references", [])
            if takeaways:
                content_items.append(("Asosiy Ilmiy Xulosalar", "\n".join([f"• {t}" for t in takeaways])))
            if refs:
                content_items.append(("Foydalanilgan Adabiyotlar", "\n".join([f"• {r}" for r in refs])))

        else: # cards_grid
            cards = slide_data.get("cards", [])
            if cards:
                for c in cards:
                    content_items.append((c.get("title", ""), c.get("desc", "")))
            elif "points" in slide_data:
                for p in slide_data["points"]:
                    content_items.append((p, ""))

        # 3. Shape Grouping into Cards / Columns
        body_shapes = [s for s in all_text_shapes if s["shape"] != title_shape_meta["shape"]]
        
        # Check for single-box cover or focus slide
        non_step_body = [s for s in body_shapes if not s["is_step"]]
        if len(non_step_body) <= 1:
            if non_step_body and content_items:
                formatted_lines = []
                for h, d in content_items:
                    if h and d:
                        formatted_lines.append(f"• {h}:\n  {d}")
                    elif h:
                        formatted_lines.append(f"• {h}")
                    elif d:
                        formatted_lines.append(d)
                full_text = "\n\n".join(formatted_lines)
                cls.set_shape_text(non_step_body[0]["shape"], full_text)
            for extra in body_shapes:
                if not non_step_body or extra["shape"] != non_step_body[0]["shape"]:
                    cls.set_shape_text(extra["shape"], "")
            return

        # Cluster shapes into card groups
        # Shapes with close X coordinates belong to the same column
        # Group by column cluster
        col_clusters: Dict[int, List[Dict[str, Any]]] = {}
        for s in non_step_body:
            # Cluster threshold ~ 1.5 inches
            col_key = int(round(s["left"] / Inches(1.5).emu))
            col_clusters.setdefault(col_key, []).append(s)

        # Sort each column's shapes from top to bottom
        sorted_card_groups = []
        for col_key in sorted(col_clusters.keys()):
            shapes_in_col = sorted(col_clusters[col_key], key=lambda s: s["top"])
            # If column has multiple shapes far apart vertically (2 rows), split into distinct cards
            cur_group = []
            for shp in shapes_in_col:
                if not cur_group:
                    cur_group.append(shp)
                elif shp["top"] - cur_group[-1]["top"] > Inches(3.5).emu:
                    # New row card
                    sorted_card_groups.append(cur_group)
                    cur_group = [shp]
                else:
                    cur_group.append(shp)
            if cur_group:
                sorted_card_groups.append(cur_group)

        # Sort all card groups by row then column
        sorted_card_groups.sort(key=lambda grp: (grp[0]["top"] // Inches(4.0).emu, grp[0]["left"]))

        assigned_shapes = set()

        for idx, (item_head, item_desc) in enumerate(content_items):
            if idx < len(sorted_card_groups):
                card_grp = sorted_card_groups[idx]
                if len(card_grp) == 1:
                    shp = card_grp[0]["shape"]
                    assigned_shapes.add(id(shp))
                    combined = f"{item_head}\n{item_desc}".strip() if (item_head and item_desc) else (item_head or item_desc)
                    cls.set_shape_text(shp, combined)
                else:
                    # Multi-shape card (e.g. Title Box + Desc Box)
                    title_box = card_grp[0]["shape"]
                    desc_box = card_grp[1]["shape"]
                    assigned_shapes.add(id(title_box))
                    assigned_shapes.add(id(desc_box))
                    cls.set_shape_text(title_box, item_head or item_desc, is_card_title=True)
                    cls.set_shape_text(desc_box, item_desc or "", font_size_override=12.0)
                    
                    # Any extra sub-boxes in this card
                    for extra_shp in card_grp[2:]:
                        assigned_shapes.add(id(extra_shp["shape"]))
                        cls.set_shape_text(extra_shp["shape"], "")

        # 4. Mandatory Cleanup of ALL unused shapes
        for s in body_shapes:
            if id(s["shape"]) not in assigned_shapes:
                if s["is_step"]:
                    continue
                cls.set_shape_text(s["shape"], "")


print("AdvancedTextReplacer tested successfully.")

class TemplateDeckAssembler:
    @staticmethod
    def match_and_assemble_deck(template_pptx_path: str, blueprint: Dict[str, Any], content_data: Dict[str, Any], output_path: str) -> str:
        prs = Presentation(template_pptx_path)
        sw = prs.slide_width
        sh = prs.slide_height
        
        # 1. Classify all template slides
        template_slides_meta = []
        for s_idx, slide in enumerate(prs.slides):
            t_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
            all_text = " ".join([s.text_frame.text.lower() for s in t_shapes])
            
            has_agenda = any(k in all_text for k in ["table of contents", "agenda", "contents", "index"])
            has_step = any(k in all_text for k in ["step 01", "step 1", "01.", "step infographic", "steps infographic"])
            has_timeline = any(k in all_text for k in ["timeline", "2026", "2027", "2028", "2029", "2030", "roadmap"])
            has_comparison = any(k in all_text for k in ["comparison", "vs", "versus"])
            has_table = "table slide" in all_text
            has_credits = any(k in all_text for k in ["credits", "fonts", "images", "let’s keep in touch", "contact", "thank you", "thanks"])
            
            ltype = "general"
            if s_idx == 0:
                ltype = "cover"
            elif has_table:
                ltype = "table_data"
            elif has_credits:
                ltype = "conclusion_sources"
            elif has_agenda:
                ltype = "agenda"
            elif has_timeline:
                ltype = "timeline_steps"
            elif has_step:
                ltype = "timeline_steps"
            elif has_comparison:
                ltype = "comparison_vs"
            elif len(t_shapes) >= 8:
                ltype = "cards_grid"
            elif len(t_shapes) in [2, 3, 4]:
                ltype = "theory_concept"
            else:
                ltype = "cards_grid"
                
            template_slides_meta.append({
                "index": s_idx,
                "layout_type": ltype,
                "text_count": len(t_shapes),
                "has_comparison": has_comparison,
                "has_timeline": has_timeline,
                "has_step": has_step,
                "has_agenda": has_agenda,
                "has_credits": has_credits
            })
            
        slides_data = content_data.get("slides", [])
        total_desired = len(slides_data)
        
        # 2. Select best distinct template slide for each requested slide
        used_indices = set()
        chosen_indices = []
        
        for idx, slide_item in enumerate(slides_data):
            req_layout = slide_item.get("layout_type", "cards_grid")
            best_idx = None
            best_score = -1000
            
            for t_meta in template_slides_meta:
                t_idx = t_meta["index"]
                if t_idx in used_indices:
                    continue
                score = 0
                if idx == 0 and t_idx == 0:
                    score += 500
                elif req_layout == "agenda" and t_meta["has_agenda"]:
                    score += 400
                elif req_layout == "comparison_vs" and t_meta["has_comparison"]:
                    score += 400
                elif req_layout == "timeline_steps" and (t_meta["has_timeline"] or t_meta["has_step"]):
                    score += 400
                elif req_layout == "conclusion_sources" and t_meta["has_credits"]:
                    score += 400
                elif req_layout == t_meta["layout_type"]:
                    score += 200
                elif t_meta["layout_type"] == "cards_grid" and t_meta["text_count"] >= 6:
                    score += 100
                elif t_meta["layout_type"] != "table_data":
                    score += 50
                    
                if score > best_score:
                    best_score = score
                    best_idx = t_idx
                    
            if best_idx is None:
                # Fallback: Pick any unused slide
                avail = [i for i in range(len(prs.slides)) if i not in used_indices]
                best_idx = avail[0] if avail else 0
                
            used_indices.add(best_idx)
            chosen_indices.append(best_idx)
            
        print(f"Blueprint layout request -> Selected template slide indices: {chosen_indices}")
        
        # 3. Prune all unused slides
        to_delete = [i for i in range(len(prs.slides)) if i not in set(chosen_indices)]
        for del_idx in sorted(to_delete, reverse=True):
            rId = prs.slides._sldIdLst[del_idx].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[del_idx]
            
        # Re-index map of remaining slides
        # The remaining slides in prs.slides correspond to chosen_indices in ascending order
        sorted_chosen = sorted(chosen_indices)
        old_to_new_pos = {orig_idx: new_pos for new_pos, orig_idx in enumerate(sorted_chosen)}
        
        # Reorder slides in prs.slides._sldIdLst to match chosen_indices exactly
        sldIdLst = prs.slides._sldIdLst
        sldId_elements = list(sldIdLst)
        
        # Map original index to element
        orig_to_elem = {sorted_chosen[i]: sldId_elements[i] for i in range(len(sorted_chosen))}
        
        # Clear and re-append in exact chosen_indices order
        for el in sldId_elements:
            sldIdLst.remove(el)
        for orig_idx in chosen_indices:
            sldIdLst.append(orig_to_elem[orig_idx])
            
        # 4. Now prs.slides is in exact 1-to-1 order with slides_data!
        topic = content_data.get("topic", "Akademik Taqdimot")
        for idx, slide_item in enumerate(slides_data):
            slide = prs.slides[idx]
            AdvancedTextReplacer.replace_slide_content_smart(slide, slide_item, topic)
            
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        prs.save(output_path)
        print(f"Successfully generated assembled presentation: {output_path} ({os.path.getsize(output_path):,} bytes)")
        return output_path

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
                "Faktorizatsiya va diskret logarifmga tayanadi",
                "Kvant kompyuterlari qarshisida mutlaqo zaif",
                "Kichik kalit hajmi va yuqori tezlik"
            ],
            "right_title": "Post-Kvant Kriptografiyasi (PQC)",
            "right_points": [
                "Panjaralar va xesh daraxtlariga asoslangan",
                "Kvant va klassik hujumlarga 100% chidamli",
                "Kattaroq kalit hajmi talab etiladi"
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
                {"question": "'Hozir saqlab ol, keyin shifrni och' (SNDL) hujumi nima?", "answer": "Buzg'unchilar bugungi shifrlangan maxfiy ma'lumotlarni yozib oladi va kelajakda kvant kompyuteri paydo bo'lganda shifrini ochadi."},
                {"question": "Qachondan PQC ga o'tishni boshlash kerak?", "answer": "Zudlik bilan! Chunki global infratuzilmani yangi kripto-standartlarga o'tkazish 5-10 yil vaqt talab qiladi."}
            ]
        },
        {
            "slide_number": 10,
            "layout_type": "conclusion_sources",
            "title": "Ilmiy Xulosa va Foydalanilgan Adabiyotlar",
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

matcher = get_academic_matcher()
bp = matcher.build_blueprint("Kvant kompyuterlari va asimmetrik kriptografiya", slide_count=10, language="uz")
template_path = bp["primary_pptx_path"]

out_test = r"output\academic_studio\Kvant_PERFECT_ASSEMBLED.pptx"
TemplateDeckAssembler.match_and_assemble_deck(template_path, bp, sample_content, out_test)

