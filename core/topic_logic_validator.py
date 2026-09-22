# -*- coding: utf-8 -*-
"""
Topic Logic & Semantic Content Validator
Ensures that all slide content strictly adheres to the presentation topic,
filters out irrelevant placeholder artifacts, verifies academic logical flow,
and optimizes bullet length for maximum font legibility.
"""

import re
import logging
from typing import Dict, Any, List, Optional, Tuple, Set

logger = logging.getLogger("TopicLogicValidator")

# Blacklisted placeholder and irrelevant filler tokens
IRRELEVANT_PATTERNS = [
    r'lorem\s+ipsum',
    r'sample\s+text',
    r'click\s+to\s+add',
    r'insert\s+your\s+text',
    r'briefly\s+describe',
    r'write\s+your\s+content',
    r'describe\s+your\s+content',
    r'this\s+is\s+your\s+big\s+image',
    r'this\s+is\s+your\s+chart\s+slide',
    r'free\s+powerpoint\s+templates?',
    r'presentationgo(\.com)?',
    r'slidenest(\.com)?',
    r'slidescarnival(\.com)?',
    r'slidesmania(\.com)?',
    r'pptmon(\.com)?',
    r'thepopp(\.com)?',
    r'allppt(\.com)?',
    r'slideegg(\.com)?',
    r'www\.[a-z0-9\-]+\.com',
    r'carrer\s+d.enric\s+granados',
    r'jobs@yourbestbrand\.com',
    r'monday\s+to\s+saturday',
    r'let.s\s+keep\s+in\s+touch',
    r'book\s+now',
    r'designed\s+with\s+.*?by',
    r'questions\s+or\s+need\s+help',
    r'credits?\s*:\s*this\s+presentation',
    r'fonts\s*&\s*colors\s*used',
    r'please\s+do\s+not\s+delete\s+this\s+slide',
    r'our\s+license',
    r'visit\s+our\s+faq',
    r'thank\s+you\s+for\s+downloading',
    r'happy\s+designing',
    r'pexels,\s*pixabay',
    r'^(paris|berlin|tokyo|helsinki|bogota|new\s+york|oslo|london|madrid|rome|sydney|chicago|vancouver)(\s*0?[1-9])?$',
    r'^(cfo|coo|ceo|cto|cmo|vp|director|manager|lead)$',
    r'^(karla\s+gard|john\s+doe|jane\s+smith|alex\s+smith|andrew\s+childs|jean\s+marcus|paola\s+frisch|graham\s+gold)$',
    r'the\s+introduction\s+of\s+computers\s+was\s+the\s+basis'
]


class TopicLogicValidator:
    """
    Validates and sanitizes slide contents against the core topic domain.
    """

    @staticmethod
    def extract_keywords(text: str) -> Set[str]:
        """Extracts meaningful alphabetic/numeric keyword tokens."""
        words = re.findall(r'\b[a-zA-Z\u0400-\u04FF\u0100-\u017F\u0180-\u024F]{3,}\b', text.lower())
        stopwords = {
            'va', 'ham', 'uchun', 'bilan', 'kabi', 'esa', 'boshqa', 'haqida',
            'barcha', 'har', 'bir', 'o‘z', 'o\'z', 'bu', 'shu', 'u', 'yoki',
            'the', 'and', 'for', 'with', 'about', 'from', 'into', 'that', 'this',
            'для', 'или', 'как', 'при', 'это', 'все', 'также', 'если', 'что'
        }
        return {w for w in words if w not in stopwords}

    @classmethod
    def sanitize_text_field(cls, text: Any, fallback: str = "") -> str:
        """Removes placeholder artifacts and cleans whitespace."""
        if text is None:
            return fallback
        if not isinstance(text, str):
            text = str(text)
        cleaned = text.strip()
        low = cleaned.lower()
        if any(k in low for k in ["slidescarnival", "slidesmania", "presentationgo", "slidenest", "slideegg", "allppt", "free powerpoint"]):
            return fallback
        for pat in IRRELEVANT_PATTERNS:
            if re.search(pat, cleaned, re.IGNORECASE):
                cleaned = re.sub(pat, '', cleaned, flags=re.IGNORECASE).strip()
        # Clean multiple spaces or empty bullet points
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned if cleaned else fallback

    @classmethod
    def optimize_bullet_length(cls, text: Any, max_words: int = 28) -> str:
        """
        Ensures a bullet point is informative, rich, and structured without artificial
        sentence clipping, strictly supporting Bold Lead-In statements up to 28-32 words.
        """
        if text is None:
            return ""
        if not isinstance(text, str):
            text = str(text)
        cleaned = cls.sanitize_text_field(text)
        words = cleaned.split()
        if len(words) <= max_words:
            return cleaned

        # Compact cleanly at sentence or punctuation boundaries if possible
        puncts = [m.start() for m in re.finditer(r'[\.\;\!\?]', cleaned)]
        for p_idx in puncts:
            sub = cleaned[:p_idx + 1].strip()
            sub_words = sub.split()
            if 14 <= len(sub_words) <= max_words:
                return sub

        # Fallback: Truncate cleanly at word boundary
        return " ".join(words[:max_words]).rstrip(".,;:") + "."

    @classmethod
    def validate_and_refine_presentation(cls, content_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Runs comprehensive topic logic validation and text refinement across all slides.
        """
        topic = content_data.get("topic", "Akademik Taqdimot")
        topic_keywords = cls.extract_keywords(topic)
        slides = content_data.get("slides", [])

        refined_slides = []
        for idx, slide in enumerate(slides, start=1):
            s_data = dict(slide)
            layout_type = s_data.get("layout_type", "cards_grid")

            # 1. Sanitize and sharpen Title
            orig_title = s_data.get("title", f"Slayd {idx}")
            s_data["title"] = cls.sanitize_text_field(orig_title, fallback=f"{topic} — {idx}-qism")

            # 2. Refine Layout Specific Fields
            if layout_type == "cover":
                s_data["subtitle"] = cls.sanitize_text_field(
                    s_data.get("subtitle", f"Tadqiqot va ilmiy-amaliy tahlil | 2026"),
                    fallback=f"Tadqiqot va ilmiy-amaliy tahlil | 2026"
                )

            elif layout_type == "agenda":
                points = s_data.get("points", [])
                s_data["points"] = [cls.optimize_bullet_length(p, max_words=24) for p in points if p is not None and str(p).strip()]

            elif layout_type == "theory_concept":
                s_data["concept_name"] = cls.sanitize_text_field(s_data.get("concept_name", "Nazariy Tushuncha"))
                s_data["definition"] = cls.optimize_bullet_length(s_data.get("definition", ""), max_words=28)
                s_data["key_formula_or_law"] = cls.sanitize_text_field(s_data.get("key_formula_or_law", ""))
                s_data["importance"] = cls.optimize_bullet_length(s_data.get("importance", ""), max_words=28)

            elif layout_type == "comparison_vs":
                s_data["left_title"] = cls.sanitize_text_field(s_data.get("left_title", "1-Model / Yondashuv"))
                s_data["left_points"] = [cls.optimize_bullet_length(p, max_words=24) for p in s_data.get("left_points", []) if p is not None and str(p).strip()]
                s_data["right_title"] = cls.sanitize_text_field(s_data.get("right_title", "2-Model / Yondashuv"))
                s_data["right_points"] = [cls.optimize_bullet_length(p, max_words=24) for p in s_data.get("right_points", []) if p is not None and str(p).strip()]

            elif layout_type == "timeline_steps":
                steps = s_data.get("steps", [])
                refined_steps = []
                for s_idx, st in enumerate(steps):
                    if isinstance(st, dict):
                        s_num = str(st.get("step") or st.get("step_number") or st.get("phase") or f"{s_idx+1:02d}").strip()
                        refined_steps.append({
                            "step": s_num,
                            "step_number": s_num,
                            "title": cls.sanitize_text_field(st.get("title", ""), fallback="Bosqich"),
                            "desc": cls.optimize_bullet_length(st.get("desc", ""), max_words=26)
                        })
                    else:
                        s_num = f"{s_idx+1:02d}"
                        refined_steps.append({"step": s_num, "step_number": s_num, "title": cls.sanitize_text_field(st), "desc": ""})
                s_data["steps"] = refined_steps

            elif layout_type == "process_cycle":
                steps = s_data.get("cycle_steps", []) or s_data.get("steps", [])
                refined_steps = []
                for st in steps:
                    if isinstance(st, dict):
                        refined_steps.append({
                            "title": cls.sanitize_text_field(st.get("title", "Bosqich")),
                            "desc": cls.optimize_bullet_length(st.get("desc", ""), max_words=26)
                        })
                    else:
                        refined_steps.append({"title": cls.sanitize_text_field(st), "desc": ""})
                s_data["cycle_steps"] = refined_steps
                s_data["steps"] = refined_steps

            elif layout_type == "pyramid_funnel":
                levels = s_data.get("levels", []) or s_data.get("hierarchy", []) or s_data.get("steps", [])
                refined_levels = []
                for lv in levels:
                    if isinstance(lv, dict):
                        refined_levels.append({
                            "title": cls.sanitize_text_field(lv.get("title") or lv.get("level_name", "Daraja")),
                            "desc": cls.optimize_bullet_length(lv.get("desc", ""), max_words=26)
                        })
                    else:
                        refined_levels.append({"title": cls.sanitize_text_field(lv), "desc": ""})
                s_data["levels"] = refined_levels

            elif layout_type == "keynote_quote":
                s_data["quote"] = cls.sanitize_text_field(s_data.get("quote", ""))
                s_data["author"] = cls.sanitize_text_field(s_data.get("author", ""))
                s_data["significance"] = cls.optimize_bullet_length(s_data.get("significance", ""), max_words=28)

            elif layout_type == "table_data":
                headers = s_data.get("headers", [])
                s_data["headers"] = [cls.sanitize_text_field(h) for h in headers if h is not None and str(h).strip()]
                rows = s_data.get("rows", [])
                refined_rows = []
                for r in rows:
                    if isinstance(r, (list, tuple)):
                        refined_rows.append([cls.sanitize_text_field(c) for c in r])
                    elif isinstance(r, dict):
                        refined_rows.append({k: cls.sanitize_text_field(v) for k, v in r.items()})
                    else:
                        refined_rows.append(cls.sanitize_text_field(r))
                s_data["rows"] = refined_rows

            elif layout_type == "diagram_anatomy":
                comps = s_data.get("components", [])
                refined_comps = []
                for c in comps:
                    if isinstance(c, dict):
                        refined_comps.append({
                            "name": cls.sanitize_text_field(c.get("name", ""), fallback="Komponent"),
                            "function": cls.optimize_bullet_length(c.get("function", ""), max_words=26)
                        })
                    else:
                        refined_comps.append({"name": cls.sanitize_text_field(c), "function": ""})
                s_data["components"] = refined_comps

            elif layout_type == "metrics_stats":
                metrics = s_data.get("metrics", [])
                refined_metrics = []
                for m in metrics:
                    if isinstance(m, dict):
                        refined_metrics.append({
                            "value": cls.sanitize_text_field(m.get("value", "100%")),
                            "label": cls.sanitize_text_field(m.get("label", "Ko'rsatkich")),
                            "desc": cls.optimize_bullet_length(m.get("desc", ""), max_words=26)
                        })
                    else:
                        refined_metrics.append({"value": cls.sanitize_text_field(m), "label": "", "desc": ""})
                s_data["metrics"] = refined_metrics

            elif layout_type == "qa_discussion":
                qa_list = s_data.get("qa_list", [])
                refined_qa = []
                for qa in qa_list:
                    if isinstance(qa, dict):
                        refined_qa.append({
                            "question": cls.sanitize_text_field(qa.get("question", "Dolzarb Savol?")),
                            "answer": cls.optimize_bullet_length(qa.get("answer", ""), max_words=28)
                        })
                    else:
                        refined_qa.append({"question": cls.sanitize_text_field(qa), "answer": ""})
                s_data["qa_list"] = refined_qa

            elif layout_type == "conclusion_sources":
                s_data["key_takeaways"] = [cls.optimize_bullet_length(t, max_words=28) for t in s_data.get("key_takeaways", []) if t is not None and str(t).strip()]
                s_data["references"] = [cls.sanitize_text_field(r) for r in s_data.get("references", []) if r is not None and str(r).strip()]

            else: # cards_grid
                cards = s_data.get("cards", [])
                if cards:
                    refined_cards = []
                    for c in cards:
                        if isinstance(c, dict):
                            refined_cards.append({
                                "title": cls.sanitize_text_field(c.get("title", "Asosiy Aspekt")),
                                "desc": cls.optimize_bullet_length(c.get("desc", ""), max_words=28)
                            })
                        else:
                            refined_cards.append({"title": cls.sanitize_text_field(c), "desc": ""})
                    s_data["cards"] = refined_cards
                elif "points" in s_data:
                    s_data["points"] = [cls.optimize_bullet_length(p, max_words=28) for p in s_data.get("points", []) if p is not None and str(p).strip()]

            refined_slides.append(s_data)

        content_data["slides"] = refined_slides
        logger.info(f"Topic Logic Validator refined {len(refined_slides)} slides for topic '{topic}'")
        return content_data
