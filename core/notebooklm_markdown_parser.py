# -*- coding: utf-8 -*-
"""
core/notebooklm_markdown_parser.py
Parses raw Markdown outline outputs from Google NotebookLM (and Gemini)
into structured presentation data for SlaydHub AI.

Key Features:
1. Robust Regex extraction for slide blocks (### N-slayd, ### Slide N, etc.).
2. Extracts Action-Oriented Title, Mantiqiy Maqsad (Purpose), and Tarkibiy Tezislar (Bullets).
3. Semantic visual archetype detector that analyzes title, purpose, and bullets to assign
   the optimal archetype (theory_concept, comparison_vs, timeline_steps, process_cycle,
   metrics_stats, diagram_anatomy, cards_grid, conclusion_sources, etc.).
4. Transforms extracted content into archetype-specific fields expected by
   TemplateDeckAssembler and UniversalSlideSplicer.
"""

import re
from typing import Dict, Any, List, Optional, Tuple
from core.slide_archetypes import SLIDE_ARCHETYPES
from core.math_sanitizer import clean_latex_math


class NotebookLMMarkdownParser:
    """
    Parses human-readable NotebookLM Markdown presentations into structured slide data.
    """

    @classmethod
    def parse(
        cls,
        markdown_text: str,
        topic: str = "",
        default_metadata: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Parses full markdown document into a presentation dictionary containing 'slides'.
        """
        if not markdown_text or not markdown_text.strip():
            return {"topic": topic, "slides": []}

        markdown_text = cls._preprocess_notebooklm_text(markdown_text)

        meta = default_metadata or {}
        slides_data: List[Dict[str, Any]] = []

        # Split text into slide blocks
        slide_pattern = re.compile(
            r'(?:^|\n)(?:(?:#{1,4}|\*{2})\s*(?:(\d+)[\.\-\s]*(?:slayd|slide)?|slide\s*(\d+))[\:\.\s\-]*(?:\*{2})?[\:\.\s\-]*([^\n]*))',
            re.IGNORECASE
        )

        matches = list(slide_pattern.finditer(markdown_text))

        if not matches:
            # Fallback: try splitting by horizontal rules `---`
            blocks = re.split(r'\n\s*---\s*\n', markdown_text)
            if len(blocks) > 1:
                for idx, block in enumerate(blocks, 1):
                    if block.strip():
                        slide = cls._parse_slide_block(block, idx, len(blocks), meta)
                        if slide:
                            slides_data.append(slide)
        else:
            total_slides = len(matches)

            # Check if there is valid slide content before the first matched header
            first_num = int(matches[0].group(1) or matches[0].group(2) or 1)
            if first_num > 1 and matches[0].start() > 0:
                pre_text = markdown_text[:matches[0].start()].strip()
                if pre_text and len(pre_text) > 30 and any(k in pre_text.lower() for k in ["action-oriented", "tezis", "maqsad", "sarlavha"]):
                    pre_slide_num = max(1, first_num - 1)
                    pre_slide = cls._parse_slide_block(
                        block_content=pre_text,
                        slide_idx=pre_slide_num,
                        total_slides=total_slides + 1,
                        meta=meta,
                        header_title=f"{pre_slide_num}-Slayd"
                    )
                    if pre_slide:
                        slides_data.append(pre_slide)

            for i, m in enumerate(matches):
                start = m.end()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(markdown_text)
                block_content = markdown_text[start:end]

                slide_num = int(m.group(1) or m.group(2) or (len(slides_data) + 1))
                header_title = (m.group(3) or "").strip()

                slide = cls._parse_slide_block(
                    block_content=block_content,
                    slide_idx=slide_num,
                    total_slides=total_slides,
                    meta=meta,
                    header_title=header_title
                )
                if slide:
                    slides_data.append(slide)

        # Fallback if no slides could be extracted via regex
        if not slides_data:
            slides_data = cls._fallback_parse_raw_text(markdown_text, topic, meta)

        presentation_topic = topic
        if not presentation_topic and slides_data:
            presentation_topic = slides_data[0].get("title", "Akademik Taqdimot")

        return {
            "topic": presentation_topic,
            "total_slides": len(slides_data),
            "slides": slides_data
        }

    @classmethod
    def _parse_slide_block(
        cls,
        block_content: str,
        slide_idx: int,
        total_slides: int,
        meta: Dict[str, Any],
        header_title: str = ""
    ) -> Dict[str, Any]:
        """
        Parses a single slide block text into a structured slide dictionary.
        """
        action_title = ""
        purpose = ""
        explicit_archetype = ""
        theses: List[str] = []

        lines = block_content.strip().split('\n')
        in_theses_section = False

        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            # Check for Action-Oriented Title
            # e.g. * **Action-oriented sarlavha:** Text
            # or **Sarlavha:** Text
            # or Title: Text
            m_title = re.search(
                r'(?:\*|\-)?\s*(?:\*\*)?(?:action[\-\s]*oriented\s*sarlavha|sarlavha|action[\-\s]*oriented\s*title|title)[\:\*\_]*\s*(.+)',
                line_str,
                re.IGNORECASE
            )
            if m_title:
                val = m_title.group(1).strip().strip('*').strip('_').strip()
                if val:
                    action_title = val
                    in_theses_section = False
                    continue

            # Check for Logical Purpose
            # e.g. * **Mantiqiy maqsadi:** Text
            m_purpose = re.search(
                r'(?:\*|\-)?\s*(?:\*\*)?(?:mantiqiy\s*maqsad[ia]?|maqsad|purpose|objective|goal)[\:\*\_]*\s*(.+)',
                line_str,
                re.IGNORECASE
            )
            if m_purpose:
                val = m_purpose.group(1).strip().strip('*').strip('_').strip()
                if val:
                    purpose = val
                    in_theses_section = False
                    continue

            # Check for Explicit Visual Archetype
            # e.g. * **Vizual turi:** comparison_vs
            m_arch = re.search(
                r'(?:\*|\-)?\s*(?:\*\*)?(?:vizual\s*turi|vizual\s*arxetip|archetype|layout(?:_type)?)[\:\*\_]*\s*(.+)',
                line_str,
                re.IGNORECASE
            )
            if m_arch:
                val = m_arch.group(1).strip().strip('*').strip('_').strip().lower()
                for arch_key in SLIDE_ARCHETYPES:
                    if arch_key in val:
                        explicit_archetype = arch_key
                        break
                in_theses_section = False
                continue

            # Check for start of Theses section
            # e.g. * **Tarkibiy tezislar:** or ### Tarkibiy tezislar
            m_theses_head = re.search(
                r'(?:\*|\-|\#)?\s*(?:\*\*)?(?:tarkibiy\s*tezislar|tezislar|asosiy\s*tezislar|key\s*points|points|bullets)[\:\*\_]*',
                line_str,
                re.IGNORECASE
            )
            if m_theses_head and ("tezis" in line_str.lower() or "points" in line_str.lower() or "bullets" in line_str.lower()):
                in_theses_section = True
                # If there's bullet content on the same line after colon:
                parts = line_str.split(':', 1)
                if len(parts) > 1 and parts[1].strip():
                    rem = parts[1].strip()
                    if rem:
                        theses.append(rem)
                continue

            # If in theses section or is a bullet item
            is_bullet = re.match(r'^(?:[\*\-\+]|\d+[\.\)])\s+(.*)', line_str)
            if is_bullet:
                bullet_text = is_bullet.group(1).strip()
                # Exclude lines that were headers
                if not any(k in bullet_text.lower() for k in ["action-oriented", "mantiqiy maqsad", "vizual turi"]):
                    theses.append(bullet_text)
            elif in_theses_section and line_str:
                theses.append(line_str)

        # Clean theses from empty or dangling artifacts (e.g. '---', '5-Bosqich: ---')
        clean_theses = []
        for t in theses:
            t_str = cls._clean_formatting(t).strip()
            if not t_str or t_str in ("---", "***", "___", "-", "*"):
                continue
            if re.match(r'^(?:\d+[\-\s]*bosqich|\d+[\.\)]|step\s*\d+)[\:\.\s\-]*$', t_str, re.I):
                continue
            t_str = re.sub(r'\s*---\s*$', '', t_str).strip()
            if len(t_str) > 2:
                clean_theses.append(t_str)
        theses = clean_theses

        # Fallbacks for title
        final_title = action_title or header_title
        if not final_title:
            if theses:
                final_title = theses.pop(0)
            else:
                final_title = f"{slide_idx}-Slayd"

        # Clean LaTeX / Markdown artifacts from title if overly formatted
        final_title = cls._clean_formatting(final_title)

        # Determine visual archetype
        archetype = cls.infer_archetype(
            slide_idx=slide_idx,
            total_slides=total_slides,
            title=final_title,
            purpose=purpose,
            theses=theses,
            explicit_archetype=explicit_archetype
        )

        # Construct slide structure
        slide_data: Dict[str, Any] = {
            "slide_number": slide_idx,
            "layout_type": archetype,
            "title": final_title,
            "subtitle": purpose or (header_title if header_title != final_title else ""),
            "purpose": purpose,
            "theses": theses,
            "points": theses
        }

        # Populate archetype-specific fields
        cls._populate_archetype_fields(slide_data, archetype, theses, meta, purpose)

        return slide_data

    @classmethod
    def infer_archetype(
        cls,
        slide_idx: int,
        total_slides: int,
        title: str,
        purpose: str,
        theses: List[str],
        explicit_archetype: str = ""
    ) -> str:
        """
        Infers the optimal visual archetype using semantic keyword analysis.
        """
        if explicit_archetype and explicit_archetype in SLIDE_ARCHETYPES:
            return explicit_archetype

        # Slide 1 is always cover
        if slide_idx == 1:
            return "cover"

        combined_text = f"{title} {purpose} {' '.join(theses)}".lower()
        title_purpose_text = f"{title} {purpose}".lower()

        # Last slide or explicit conclusion keywords in title/purpose (avoids false positives on early slides)
        is_deck_end = (slide_idx == total_slides and total_slides >= 4)
        is_explicit_conclusion = any(k in title_purpose_text for k in ["xulosalar", "adabiyotlar", "manbalar", "conclusion", "references", "yakuniy", "bibliografiya"])
        is_late_tavsiya = ("tavsiyalar" in title_purpose_text and slide_idx >= max(1, total_slides - 1))
        if is_deck_end or is_explicit_conclusion or is_late_tavsiya:
            return "conclusion_sources"

        # Agenda / Reja check (typically slide 2 if titled Reja)
        if slide_idx == 2 and any(k in title_purpose_text for k in ["reja", "mundarija", "agenda", "contents", "tuzilishi va mantiqiy"]):
            return "agenda"

        # Theory / Concept check in title / purpose (e.g., Shor algoritmi, Matematik asos, Nazariy, Formula)
        theory_title_triggers = ["matematik asos", "nazariy", "formula", "qonuniyat", "teorema", "shor algoritmi", "faktorizatsiy"]
        if any(k in title_purpose_text for k in theory_title_triggers):
            return "theory_concept"

        # Comparison check in title / purpose
        comparison_triggers = [" vs ", "versus", "taqqoslash", "qiyosiy", "solishtirish", "a vs b", "afzallik va kamchilik", "teng yarimga qisqartiradi", "effektiv kalit"]
        if any(k in title_purpose_text for k in comparison_triggers):
            return "comparison_vs"

        # Timeline / Roadmap check: years, dates, milestones
        timeline_triggers = ["timeline", "roadmap", "2026", "2027", "2028", "2029", "2030", "grafik", "xronologiya", "tarix", "yillar", "o'tish grafik"]
        if any(k in combined_text for k in timeline_triggers) and not any(k in combined_text for k in ["sikl", "cycle"]):
            return "timeline_steps"

        # Process Cycle check: continuous loop, operational cycle
        cycle_triggers = ["sikl", "cycle", "doiraviy", "operatsion sikl", "jarayon", "bosqichli sikl", "loop", "aylanma", "metodologiya"]
        if any(k in combined_text for k in cycle_triggers):
            return "process_cycle"

        # Anatomy / System Breakdown: architecture, components, models
        anatomy_triggers = ["arxitektura", "komponentlar", "tizim tarkib", "diagramma", "gibrid tls", "tuzilma", "anatomy", "arxitekturasi"]
        if any(k in combined_text for k in anatomy_triggers):
            return "diagram_anatomy"

        # Comparison check in bullets if not in title
        if any(k in combined_text for k in comparison_triggers):
            return "comparison_vs"

        # Theory / Concept check in bullets if not in title
        theory_triggers = ["matematik asos", "nazariy", "formula", "qonuniyat", "teorema", "ta'rif", "gipoteza", "shor algoritmi", "grover algoritmi", "faktorizatsiya"]
        if any(k in combined_text for k in theory_triggers):
            return "theory_concept"

        # Pyramid / Hierarchy check
        pyramid_triggers = ["ierarxiya", "piramida", "darajalar", "voronka", "funnel", "pyramid", "hierarchy"]
        if any(k in combined_text for k in pyramid_triggers):
            return "pyramid_funnel"

        # Metrics / KPI check: numbers, %, stats
        if "%" in combined_text or any(k in combined_text for k in ["foiz", "statistika", "kpi", "metrika", "stats"]):
            return "metrics_stats"

        # Table data check
        table_triggers = ["jadval", "parametrlar", "mezonlar", "matritsa", "table"]
        if any(k in combined_text for k in table_triggers):
            return "table_data"

        # Keynote / Quote
        quote_triggers = ["iqtibos", "keynote", "shior", "quote", "aytganidek"]
        if any(k in combined_text for k in quote_triggers):
            return "keynote_quote"

        # Default standard academic layout
        return "cards_grid"

    @classmethod
    def _populate_archetype_fields(
        cls,
        slide_data: Dict[str, Any],
        archetype: str,
        theses: List[str],
        meta: Dict[str, Any],
        purpose: str
    ) -> None:
        """
        Fills layout-specific fields in slide_data matching pptx_text_replacer expectations.
        """
        if archetype == "cover":
            slide_data["author_name"] = meta.get("author_name", "Tadqiqotchi")
            slide_data["organization"] = meta.get("organization", "O'zbekiston Respublikasi Fanlar Akademiyasi")
            slide_data["supervisor"] = meta.get("supervisor", "Ilmiy rahbar: DSc., professor")
            slide_data["email"] = meta.get("email", "tadqiqotchi@edu.uz")
            slide_data["event_name"] = meta.get("event_name", "Akademik Taqdimot | 2026-yil")
            if not slide_data.get("subtitle"):
                slide_data["subtitle"] = purpose or meta.get("subtitle", "Ilmiy va amaliy tadqiqot hisoboti")

        elif archetype == "agenda":
            slide_data["points"] = theses or [
                "1. Muammoning qo'yilishi va dolzarbligi",
                "2. Nazariy asoslar va tahlil",
                "3. Amaliy modellashtirish va yechimlar",
                "4. Xulosa va tavsiyalar"
            ]

        elif archetype == "theory_concept":
            concept_name = ""
            definition = ""
            formula = ""
            importance = ""

            for t in theses:
                cleaned = cls._clean_formatting(t)
                if any(k in cleaned.lower() for k in ["formula", "algoritm", "o((log n)^3)", "2^n", "masalasiga", "faktorizatsiya"]):
                    formula = cleaned
                elif any(k in cleaned.lower() for k in ["maqsad", "ahamiyat", "zarurat", "chekov"]):
                    importance = cleaned
                elif not concept_name:
                    if ":" in cleaned:
                        parts = cleaned.split(":", 1)
                        concept_name = parts[0].strip()
                        definition = parts[1].strip()
                    else:
                        words = cleaned.split()
                        if len(words) <= 6 and len(cleaned) <= 45:
                            concept_name = cleaned
                            definition = ""
                        elif " - " in cleaned:
                            c_parts = cleaned.split(" - ", 1)
                            concept_name = c_parts[0].strip()
                            definition = c_parts[1].strip()
                        else:
                            concept_name = "Asosiy Konseptual Xulosa"
                            definition = cleaned
                elif not definition:
                    definition = cleaned

            slide_data["concept_name"] = concept_name or slide_data.get("title", "Asosiy Nazariya")
            slide_data["definition"] = definition or purpose or "Fundamental ilmiy ta'rif va qonuniyat"
            slide_data["key_formula_or_law"] = formula or ""
            slide_data["importance"] = importance or purpose or ""

        elif archetype == "comparison_vs":
            left_title = ""
            right_title = ""
            left_points = []
            right_points = []

            # If exactly 2 theses, they directly map to Left and Right sides!
            if len(theses) == 2:
                for idx, t in enumerate(theses):
                    cleaned = cls._clean_formatting(t)
                    item_title = ""
                    item_body = cleaned
                    if ":" in cleaned:
                        p_head, p_tail = cleaned.split(":", 1)
                        if len(p_head.split()) <= 6:
                            item_title = p_head.strip()
                            item_body = p_tail.strip()

                    # Expand comma-separated sub-points if present
                    sub_pts = []
                    if "," in item_body:
                        for sub in item_body.split(","):
                            s_clean = sub.strip()
                            if s_clean:
                                sub_pts.append(s_clean)
                    if not sub_pts:
                        sub_pts = [item_body]

                    if idx == 0:
                        left_title = item_title or "1-Yondashuv"
                        left_points = sub_pts
                    else:
                        right_title = item_title or "2-Yondashuv"
                        right_points = sub_pts
            else:
                # Check if theses distinguish between two sides and extract side titles
                for idx, t in enumerate(theses):
                    cleaned = cls._clean_formatting(t)
                    item_title = ""
                    item_body = cleaned
                    if ":" in cleaned:
                        p_head, p_tail = cleaned.split(":", 1)
                        if len(p_head.split()) <= 5:
                            item_title = p_head.strip()
                            item_body = p_tail.strip()

                    t_lower = (item_title + " " + item_body).lower()
                    if any(k in t_lower for k in ["klassik", "rsa", "128", "zaif", "qisqartiradi", "oldin"]):
                        if not left_title and item_title:
                            left_title = item_title
                        # Expand comma-separated points into individual analytical bullet items
                        if "," in item_body and len(theses) <= 2:
                            for sub in item_body.split(","):
                                s_clean = sub.strip()
                                if s_clean:
                                    left_points.append(s_clean)
                        else:
                            left_points.append(item_body)
                    elif any(k in t_lower for k in ["pqc", "post-kvant", "ml-kem", "ml-dsa", "256", "yechim", "barqaror", "himoya", "keyin"]):
                        if not right_title and item_title:
                            right_title = item_title
                        if "," in item_body and len(theses) <= 2:
                            for sub in item_body.split(","):
                                s_clean = sub.strip()
                                if s_clean:
                                    right_points.append(s_clean)
                        else:
                            right_points.append(item_body)
                    else:
                        if idx % 2 == 0:
                            if not left_title and item_title:
                                left_title = item_title
                            left_points.append(item_body)
                        else:
                            if not right_title and item_title:
                                right_title = item_title
                            right_points.append(item_body)

            if not left_title:
                left_title = "Klassik Yondashuv"
            if not right_title:
                right_title = "Post-Kvant Yondashuvi"

            if not left_points and theses:
                left_points = [theses[0]]
            if not right_points and len(theses) > 1:
                right_points = theses[1:]

            slide_data["left_title"] = left_title
            slide_data["left_points"] = left_points
            slide_data["right_title"] = right_title
            slide_data["right_points"] = right_points

        elif archetype in ("timeline_steps", "process_cycle"):
            steps_list = []
            for i, t in enumerate(theses, 1):
                cleaned = cls._clean_formatting(t)
                if ":" in cleaned:
                    parts = cleaned.split(":", 1)
                    title_part = parts[0].strip()
                    desc_part = parts[1].strip()
                else:
                    title_part = f"{i}-Bosqich"
                    desc_part = cleaned
                steps_list.append({
                    "step": f"0{i}" if i < 10 else str(i),
                    "title": title_part,
                    "desc": desc_part
                })

            if archetype == "timeline_steps":
                slide_data["steps"] = steps_list
            else:
                slide_data["cycle_steps"] = steps_list
                slide_data["steps"] = steps_list

        elif archetype == "pyramid_funnel":
            levels_list = []
            for i, t in enumerate(theses, 1):
                cleaned = cls._clean_formatting(t)
                if ":" in cleaned:
                    parts = cleaned.split(":", 1)
                    title_part = parts[0].strip()
                    desc_part = parts[1].strip()
                else:
                    title_part = f"{i}-Daraja"
                    desc_part = cleaned
                levels_list.append({
                    "level": f"0{i}",
                    "title": title_part,
                    "desc": desc_part
                })
            slide_data["levels"] = levels_list

        elif archetype == "metrics_stats":
            metrics_list = []
            for t in theses:
                cleaned = cls._clean_formatting(t)
                # Look for numbers or percentages
                m_stat = re.search(r'(\d+(?:[.,]\d+)?\s*%|\d+(?:[–\-]\d+)?|\b\d{3,}\b)', cleaned)
                if m_stat:
                    val = m_stat.group(1).strip()
                    desc = cleaned.replace(val, "").strip(" :-,()")
                    metrics_list.append({
                        "value": val,
                        "label": cleaned[:30],
                        "desc": desc
                    })
                else:
                    metrics_list.append({
                        "value": "100%",
                        "label": cleaned[:30],
                        "desc": cleaned
                    })
            slide_data["metrics"] = metrics_list

        elif archetype == "diagram_anatomy":
            components_list = []
            for t in theses:
                cleaned = cls._clean_formatting(t)
                if ":" in cleaned:
                    parts = cleaned.split(":", 1)
                    c_name = parts[0].strip()
                    c_fn = parts[1].strip()
                else:
                    c_name = cleaned[:35]
                    c_fn = cleaned
                components_list.append({
                    "name": c_name,
                    "function": c_fn
                })
            slide_data["components"] = components_list

        elif archetype == "conclusion_sources":
            title_p = (slide_data.get("title", "") + " " + slide_data.get("purpose", "")).lower()
            is_pure_biblio = any(k in title_p for k in ["adabiyot", "manba", "bibliograf", "reference", "standartlar ro'yxati"])
            is_recommendations = any(k in title_p for k in ["tavsiya", "istiqbol", "amaliy", "taklif", "rekomendat"])

            if is_pure_biblio:
                standards = []
                academic = []
                for t in theses:
                    cleaned = cls._clean_formatting(t)
                    if any(k in cleaned.lower() for k in ["standart", "fips", "nist", "iso", "rfc", "hujjat", "qonun"]):
                        standards.append(cleaned)
                    else:
                        academic.append(cleaned)
                if not standards and academic:
                    mid = max(1, len(academic) // 2)
                    standards, academic = academic[:mid], academic[mid:]
                elif not academic and standards:
                    mid = max(1, len(standards) // 2)
                    standards, academic = standards[:mid], standards[mid:]
                slide_data["standards_refs"] = standards
                slide_data["academic_refs"] = academic
                slide_data["references"] = theses
            elif is_recommendations:
                takeaways = []
                recs = []
                for t in theses:
                    cleaned = cls._clean_formatting(t)
                    if any(k in cleaned.lower() for k in ["tavsiya", "tatbiq", "joriy", "zarur", "lozim", "maqsadga muvofiq", "kerak", "taklif"]):
                        recs.append(cleaned)
                    else:
                        takeaways.append(cleaned)
                if not takeaways and theses:
                    mid = max(1, len(theses) // 2)
                    takeaways, recs = theses[:mid], theses[mid:]
                elif not recs and theses:
                    mid = max(1, len(theses) // 2)
                    takeaways, recs = theses[:mid], theses[mid:]
                slide_data["key_takeaways"] = takeaways
                slide_data["recommendations"] = recs
            else:
                takeaways = []
                refs = []
                for t in theses:
                    cleaned = cls._clean_formatting(t)
                    if any(k in cleaned.lower() for k in ["manba", "adabiyot", "standart", "fips", "nist", "hujjat", "ieee", "springer", "doi", "19", "20"]):
                        refs.append(cleaned)
                    else:
                        takeaways.append(cleaned)
                if not takeaways and theses:
                    takeaways = theses[:2]
                if not refs and len(theses) > 2:
                    refs = theses[2:]
                slide_data["key_takeaways"] = takeaways
                slide_data["references"] = refs

        else: # cards_grid
            cards_list = []
            for i, t in enumerate(theses, 1):
                cleaned = cls._clean_formatting(t)
                if ":" in cleaned:
                    parts = cleaned.split(":", 1)
                    c_title = parts[0].strip()
                    c_desc = parts[1].strip()
                else:
                    c_title = f"{i}-Aspekt"
                    c_desc = cleaned
                cards_list.append({
                    "title": c_title,
                    "desc": c_desc
                })
            slide_data["cards"] = cards_list

    @classmethod
    def _preprocess_notebooklm_text(cls, text: str) -> str:
        """
        Cleans NotebookLM web UI citation artifacts and separates concatenated headers.
        """
        if not text:
            return ""

        # 1. Clean NotebookLM web UI citation artifacts like "16more_horiz.", "21more_horiz", "more_horiz."
        text = re.sub(r'\d*\s*more_horiz\.?', '', text, flags=re.IGNORECASE)

        # 2. Insert newline before slide headings if glued to previous content
        # e.g., "...xulosalar### 2-slayd" or "...matn\n**2-slayd**"
        text = re.sub(r'([^\n])\s*(#{1,4}\s*(?:\d+[\.\-\s]*(?:slayd|slide)|slide\s*\d+))', r'\1\n\n\2', text, flags=re.IGNORECASE)

        # 3. Insert newline before known field labels if glued to previous line
        # e.g. "TavsiyalarAction-oriented sarlavha:" -> "Tavsiyalar\n* **Action-oriented sarlavha:**"
        labels_regex = r'([^\n])\s*(\*?\s*\*{0,2}(?:action[\-\s]*oriented\s*sarlavha|action[\-\s]*oriented\s*title|mantiqiy\s*maqsad[ia]?|tarkibiy\s*tezislar|asosiy\s*tezislar|vizual\s*turi|vizual\s*arxetip)[\:\*\_]*)'
        text = re.sub(labels_regex, r'\1\n\2', text, flags=re.IGNORECASE)

        # 4. Insert newline before bullet items if glued
        text = re.sub(r'([^\n])(\s{2,}\*\s+)', r'\1\n\2', text)

        return text

    @classmethod
    def _clean_formatting(cls, text: str) -> str:
        """Removes markdown asterisks, backticks, citations, extra brackets, and sanitizes LaTeX math."""
        if not text:
            return ""
        # Clean LaTeX math before stripping other formatting
        text = clean_latex_math(text)
        # Remove web citation artifacts if any remain
        text = re.sub(r'\d*\s*more_horiz\.?', '', text, flags=re.IGNORECASE)
        # Remove footnote citations like [1], [2, 3], [1-4]
        text = re.sub(r'\[\d+(?:[,\s\-–]+\d+)*\]', '', text)
        # Remove markdown bold/italic
        text = re.sub(r'[\*_]{1,3}', '', text)
        # Remove markdown backticks
        text = re.sub(r'`+', '', text)
        # Clean double spaces
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    @classmethod
    def _fallback_parse_raw_text(
        cls,
        text: str,
        topic: str,
        meta: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """Fallback splitter if markdown does not contain standard slide headers."""
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
        slides = []
        for idx, p in enumerate(paragraphs[:10], 1):
            lines = p.split('\n')
            title = lines[0].strip()
            bullets = [l.strip().lstrip('*-• ') for l in lines[1:] if l.strip()]
            archetype = "cover" if idx == 1 else ("conclusion_sources" if idx == len(paragraphs) else "cards_grid")
            slide = {
                "slide_number": idx,
                "layout_type": archetype,
                "title": title or f"{idx}-Slayd",
                "subtitle": "",
                "purpose": "",
                "theses": bullets,
                "points": bullets
            }
            cls._populate_archetype_fields(slide, archetype, bullets, meta, "")
            slides.append(slide)
        return slides
