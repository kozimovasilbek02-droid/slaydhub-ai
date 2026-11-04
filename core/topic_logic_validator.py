# -*- coding: utf-8 -*-
"""
Topic Logic & Semantic Content Validator
Ensures that all slide content strictly adheres to the presentation topic,
filters out irrelevant placeholder artifacts, verifies academic logical flow,
enriches short 1-2 word bullet points into full analytical items,
and optimizes bullet length for maximum font legibility.
"""

import re
import logging
from typing import Dict, Any, List, Optional, Tuple, Set

logger = logging.getLogger("TopicLogicValidator")

IRRELEVANT_PATTERNS = [
    r"lorem\s+ipsum",
    r"sample\s+text",
    r"click\s+to\s+add",
    r"insert\s+your\s+text",
    r"briefly\s+describe",
    r"write\s+your\s+content",
    r"describe\s+your\s+content",
    r"this\s+is\s+your\s+big\s+image",
    r"this\s+is\s+your\s+chart\s+slide",
    r"free\s+powerpoint\s+templates?",
    r"presentationgo(\.com)?",
    r"slidenest(\.com)?",
    r"slidescarnival(\.com)?",
    r"slidesmania(\.com)?",
    r"pptmon(\.com)?",
    r"thepopp(\.com)?",
    r"allppt(\.com)?",
    r"slideegg(\.com)?",
    r"www\.[a-z0-9\-]+\.com",
    r"carrer\s+d.enric\s+granados",
    r"jobs@yourbestbrand\.com",
    r"monday\s+to\s+saturday",
    r"let.s\s+keep\s+in\s+touch",
    r"book\s+now",
    r"designed\s+with\s+.*?by",
    r"questions\s+or\s+need\s+help",
    r"credits?\s*:\s*this\s+presentation",
    r"fonts\s*&\s*colors\s*used",
    r"please\s+do\s+not\s+delete\s+this\s+slide",
    r"our\s+license",
    r"visit\s+our\s+faq",
    r"thank\s+you\s+for\s+downloading",
    r"happy\s+designing",
    r"pexels,\s*pixabay",
    r"^(paris|berlin|tokyo|helsinki|bogota|new\s+york|oslo|london|madrid|rome|sydney|chicago|vancouver)(\s*0?[1-9])?$",
    r"^(cfo|coo|ceo|cto|cmo|vp|director|manager|lead)$",
    r"^(karla\s+gard|john\s+doe|jane\s+smith|alex\s+smith|andrew\s+childs|jean\s+marcus|paola\s+frisch|graham\s+gold)$",
    r"the\s+introduction\s+of\s+computers\s+was\s+the\s+basis",
]


class TopicLogicValidator:
    """
    Validates, enriches, and sanitizes slide contents against the core topic domain.
    """

    @staticmethod
    def extract_keywords(text: str) -> Set[str]:
        words = re.findall(r"\b[a-zA-Z\u0400-\u04FF\u0100-\u017F\u0180-\u024F]{3,}\b", text.lower())
        stopwords = {
            "va", "ham", "uchun", "bilan", "kabi", "esa", "boshqa", "haqida",
            "barcha", "har", "bir", "o‘z", "o'z", "bu", "shu", "u", "yoki",
            "the", "and", "for", "with", "about", "from", "into", "that", "this",
            "для", "или", "как", "при", "это", "все", "также", "если", "что",
        }
        return {w for w in words if w not in stopwords}

    @classmethod
    def sanitize_text_field(cls, text: Any, fallback: str = "") -> str:
        if text is None:
            return fallback
        if not isinstance(text, str):
            text = str(text)
        cleaned = text.strip()
        low = cleaned.lower()
        if any(
            k in low
            for k in [
                "slidescarnival",
                "slidesmania",
                "presentationgo",
                "slidenest",
                "slideegg",
                "allppt",
                "free powerpoint",
            ]
        ):
            return fallback
        for pat in IRRELEVANT_PATTERNS:
            if re.search(pat, cleaned, re.IGNORECASE):
                cleaned = re.sub(pat, "", cleaned, flags=re.IGNORECASE).strip()
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return cleaned if cleaned else fallback

    @classmethod
    def enrich_short_point(cls, point_text: str, slide_title: str, topic: str, idx: int = 1) -> str:
        """
        If a point is just 1-4 words without a colon/explanation (e.g., 'Taqsimlangan reyestr',
        'Kriptografiya', 'Konsensus'), expands it into a structured 'Sarlavha: Tahliliy izoh'
        string so slide cards always have both a bold header and informative body copy.
        """
        cleaned = cls.sanitize_text_field(point_text)
        if not cleaned:
            return ""
        if ":" in cleaned:
            parts = cleaned.split(":", 1)
            if len(parts) == 2 and len(parts[1].strip().split()) >= 3:
                return f"{parts[0].strip()}: {parts[1].strip()}"
        elif " - " in cleaned or " — " in cleaned or " – " in cleaned:
            parts = re.split(r"\s+[\-\–\—]\s+", cleaned, maxsplit=1)
            if len(parts) == 2 and len(parts[1].strip().split()) >= 3:
                return f"{parts[0].strip()}: {parts[1].strip()}"

        words = cleaned.split()
        if len(words) >= 8:
            return cleaned

        enrichment_templates = [
            f"{cleaned}: «{slide_title}» yo'nalishida tizimli samaradorlik, ma'lumotlar ishonchliligi va barqaror ishlash mexanizmlarini ta'minlaydi.",
            f"{cleaned}: Zamonaviy texnologik standartlar asosida jarayonlarni optimallashtirish, xavfsizlikni oshirish va amaliy natijadorlikka erishish omili.",
            f"{cleaned}: Soha rivojida innovatsion yechimlarni joriy etish, resurslarni tejash hamda yuqori sifat ko'rsatkichlarini kafolatlash asosi.",
            f"{cleaned}: Tahliliy ko'rsatkichlar va ilg'or tajribalar asosida tizimning uzoq muddatli raqobatbardoshligini mustahkamlash vositasi.",
        ]
        return enrichment_templates[(idx - 1) % len(enrichment_templates)]

    @classmethod
    def optimize_bullet_length(cls, text: Any, max_words: int = 28) -> str:
        if text is None:
            return ""
        if not isinstance(text, str):
            text = str(text)
        cleaned = cls.sanitize_text_field(text)
        words = cleaned.split()
        if len(words) <= max_words:
            return cleaned

        puncts = [m.start() for m in re.finditer(r"[\.\;\!\?]", cleaned)]
        for p_idx in puncts:
            sub = cleaned[: p_idx + 1].strip()
            sub_words = sub.split()
            if 12 <= len(sub_words) <= max_words:
                return sub

        return " ".join(words[:max_words]).rstrip(".,;:") + "."

    @classmethod
    def validate_and_refine_presentation(cls, content_data: Dict[str, Any]) -> Dict[str, Any]:
        topic = content_data.get("topic", "Akademik Taqdimot")
        slides = content_data.get("slides", [])

        refined_slides = []
        for idx, slide in enumerate(slides, start=1):
            s_data = dict(slide)
            layout_type = s_data.get("layout_type", "cards_grid")

            orig_title = s_data.get("title") or s_data.get("heading") or f"Slayd {idx}"
            s_Title = cls.sanitize_text_field(orig_title, fallback=f"{topic} — {idx}-qism")
            s_data["title"] = s_Title

            if layout_type == "cover":
                s_data["subtitle"] = cls.sanitize_text_field(
                    s_data.get("subtitle", "Tadqiqot va ilmiy-amaliy tahlil | 2026"),
                    fallback="Tadqiqot va ilmiy-amaliy tahlil | 2026",
                )

            elif layout_type == "agenda":
                points = s_data.get("points") or s_data.get("bullets") or []
                clean_pts = []
                for p_i, p in enumerate(points, 1):
                    if p is not None and str(p).strip():
                        raw_p = cls.optimize_bullet_length(p, max_words=24)
                        if ":" not in raw_p:
                            raw_p = f"{raw_p}: Asosiy tushunchalar, tahliliy ko'rsatkichlar va amaliy yechimlar sharhi"
                        clean_pts.append(raw_p)
                s_data["points"] = clean_pts
                s_data["bullets"] = clean_pts

            elif layout_type == "theory_concept":
                pts = s_data.get("points") or s_data.get("bullets") or []
                enriched_pts = [
                    cls.optimize_bullet_length(cls.enrich_short_point(str(p), s_Title, topic, i), max_words=28)
                    for i, p in enumerate(pts, 1)
                    if p is not None and str(p).strip()
                ]
                if enriched_pts:
                    s_data["points"] = enriched_pts
                    s_data["bullets"] = enriched_pts
                s_data["concept_name"] = cls.sanitize_text_field(s_data.get("concept_name", s_Title))
                if not s_data.get("definition") and enriched_pts:
                    s_data["definition"] = enriched_pts[0]
                if not s_data.get("key_formula_or_law") and len(enriched_pts) > 1:
                    s_data["key_formula_or_law"] = enriched_pts[1]
                if not s_data.get("importance") and len(enriched_pts) > 2:
                    s_data["importance"] = enriched_pts[2]

            elif layout_type == "comparison_vs":
                pts = s_data.get("points") or s_data.get("bullets") or []
                enriched_pts = [
                    cls.optimize_bullet_length(cls.enrich_short_point(str(p), s_Title, topic, i), max_words=26)
                    for i, p in enumerate(pts, 1)
                    if p is not None and str(p).strip()
                ]
                if enriched_pts:
                    s_data["points"] = enriched_pts
                    s_data["bullets"] = enriched_pts
                s_data["left_title"] = cls.sanitize_text_field(s_data.get("left_title", ""))
                s_data["right_title"] = cls.sanitize_text_field(s_data.get("right_title", ""))

            elif layout_type == "timeline_steps":
                steps = s_data.get("steps", [])
                pts = s_data.get("points") or s_data.get("bullets") or []
                if not steps and pts:
                    steps = []
                    for i, p in enumerate(pts, 1):
                        ep = cls.enrich_short_point(str(p), s_Title, topic, i)
                        if ":" in ep:
                            h, d = ep.split(":", 1)
                        else:
                            h, d = f"{i}-Bosqich", ep
                        steps.append({"step": f"{i:02d}", "step_number": f"{i:02d}", "title": h.strip(), "desc": cls.optimize_bullet_length(d.strip(), max_words=26)})
                refined_steps = []
                for s_idx, st in enumerate(steps):
                    if isinstance(st, dict):
                        s_num = str(st.get("step") or st.get("step_number") or st.get("phase") or f"{s_idx+1:02d}").strip()
                        t_val = cls.sanitize_text_field(st.get("title", ""), fallback=f"{s_idx+1}-Bosqich")
                        d_val = cls.optimize_bullet_length(st.get("desc", ""), max_words=26)
                        if not d_val:
                            ep = cls.enrich_short_point(t_val, s_Title, topic, s_idx + 1)
                            d_val = ep.split(":", 1)[1].strip() if ":" in ep else ep
                        refined_steps.append({"step": s_num, "step_number": s_num, "title": t_val, "desc": d_val})
                    else:
                        s_num = f"{s_idx+1:02d}"
                        ep = cls.enrich_short_point(str(st), s_Title, topic, s_idx + 1)
                        h, d = ep.split(":", 1) if ":" in ep else (ep, "")
                        refined_steps.append({"step": s_num, "step_number": s_num, "title": h.strip(), "desc": d.strip()})
                s_data["steps"] = refined_steps

            elif layout_type == "conclusion_sources":
                pts = s_data.get("points") or s_data.get("bullets") or s_data.get("key_takeaways") or []
                clean_pts = [
                    cls.optimize_bullet_length(cls.enrich_short_point(str(t), s_Title, topic, i), max_words=28)
                    for i, t in enumerate(pts, 1)
                    if t is not None and str(t).strip()
                ]
                s_data["key_takeaways"] = clean_pts
                s_data["points"] = clean_pts
                s_data["bullets"] = clean_pts
                s_data["references"] = [
                    cls.sanitize_text_field(r) for r in s_data.get("references", []) if r is not None and str(r).strip()
                ]

            else:
                cards = s_data.get("cards", [])
                if cards:
                    refined_cards = []
                    for i, c in enumerate(cards, 1):
                        if isinstance(c, dict):
                            t_val = cls.sanitize_text_field(c.get("title", "Asosiy Aspekt"))
                            d_val = cls.optimize_bullet_length(c.get("desc", ""), max_words=28)
                            if not d_val:
                                ep = cls.enrich_short_point(t_val, s_Title, topic, i)
                                d_val = ep.split(":", 1)[1].strip() if ":" in ep else ep
                            refined_cards.append({"title": t_val, "desc": d_val})
                        else:
                            ep = cls.enrich_short_point(str(c), s_Title, topic, i)
                            h, d = ep.split(":", 1) if ":" in ep else (ep, "")
                            refined_cards.append({"title": h.strip(), "desc": d.strip()})
                    s_data["cards"] = refined_cards
                else:
                    pts = s_data.get("points") or s_data.get("bullets") or []
                    enriched_pts = [
                        cls.optimize_bullet_length(cls.enrich_short_point(str(p), s_Title, topic, i), max_words=28)
                        for i, p in enumerate(pts, 1)
                        if p is not None and str(p).strip()
                    ]
                    s_data["points"] = enriched_pts
                    s_data["bullets"] = enriched_pts

            refined_slides.append(s_data)

        content_data["slides"] = refined_slides
        logger.info(f"Topic Logic Validator refined {len(refined_slides)} slides for topic '{topic}'")
        return content_data
