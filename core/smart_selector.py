# -*- coding: utf-8 -*-
"""
core/smart_selector.py
AI-Powered Smart Template Selector & Search Engine for SlaydHub Ultra.
Searches and ranks across all 17,250+ verified templates using multilingual
semantic keyword expansion, category relevance weighting, and quality scoring.
"""

import re
import logging
from typing import List, Dict, Any, Optional, Set

logger = logging.getLogger(__name__)

# Multilingual (UZ / RU / EN) domain concepts map
MULTILINGUAL_SEMANTIC_MAP: Dict[str, List[str]] = {
    # Nature, Ecology, Clean Energy
    "quyosh": ["solar", "sun", "renewable", "clean energy"],
    "energiya": ["energy", "power", "electricity", "solar", "battery"],
    "tabiat": ["nature", "ecology", "green", "forest", "earth", "plant", "environmental"],
    "ekologiya": ["ecology", "environment", "eco", "green", "recycling", "climate"],
    "suv": ["water", "aqua", "ocean", "river"],
    "qishloq": ["agriculture", "farming", "harvest", "rural"],
    "turizm": ["tourism", "travel", "journey", "vacation", "trip"],

    # Medicine, Healthcare, Biology
    "yurak": ["heart", "cardio", "cardiology", "cardiovascular", "cardiogram"],
    "kardio": ["cardio", "cardiology", "heart", "pulse"],
    "tibbiyot": ["medical", "medicine", "healthcare", "health", "clinic", "hospital", "doctor"],
    "tish": ["dental", "dentist", "teeth", "oral", "stomatology"],
    "dori": ["pharmacy", "pharma", "medication", "capsule", "drug", "prescription"],
    "kasallik": ["disease", "pathology", "syndrome", "infection", "virus"],
    "jarrohlik": ["surgery", "surgical", "operation", "surgeon"],
    "anatomiya": ["anatomy", "biological", "human body", "spine"],
    "psixologiya": ["psychology", "mental", "brain", "behavior"],
    "pediatriya": ["pediatric", "child health", "baby", "infant"],

    # Technology, AI, IT, Engineering
    "ai": ["ai", "artificial intelligence", "robot", "machine learning", "neural"],
    "sun'iy intellekt": ["ai", "artificial intelligence", "robotics", "automation", "deep learning"],
    "robot": ["robot", "robotics", "automation", "android"],
    "dasturlash": ["coding", "programming", "software", "developer", "code", "python", "javascript"],
    "kiber": ["cyber", "security", "hacker", "network", "firewall"],
    "bulut": ["cloud", "server", "data center", "devops"],
    "blokcheyn": ["blockchain", "crypto", "bitcoin", "web3"],
    "3d": ["3d printing", "cad", "modeling", "industrial design"],
    "muhandislik": ["engineering", "mechanical", "industry", "technology"],

    # Business, Finance, Startups, Strategy
    "biznes": ["business", "company", "corporate", "strategy", "enterprise"],
    "moliya": ["finance", "financial", "money", "banking", "fintech", "accounting", "budget"],
    "invest": ["invest", "investment", "investor", "pitch deck", "venture", "funding"],
    "startap": ["startup", "pitch deck", "entrepreneur", "innovation"],
    "reja": ["plan", "business plan", "roadmap", "strategy", "timeline"],
    "marketing": ["marketing", "market", "advertising", "brand", "sales", "promotion"],
    "savdo": ["sales", "commerce", "ecommerce", "retail"],
    "kpi": ["kpi", "dashboard", "metric", "report", "analytics"],

    # Education, Academic, Science, History
    "maktab": ["school", "education", "classroom", "student", "teacher", "elementary"],
    "universitet": ["university", "college", "academic", "campus", "thesis", "defense"],
    "dars": ["lesson", "lecture", "study", "teaching", "course"],
    "tarix": ["history", "historical", "ancient", "vintage", "heritage", "medieval"],
    "adabiyot": ["literature", "poetry", "books", "reading", "library"],
    "til": ["language", "english", "linguistics", "grammar"],
    "matematika": ["math", "mathematics", "algebra", "geometry", "calculus"],
    "fizika": ["physics", "quantum", "optics", "mechanics"],
    "kimyo": ["chemistry", "chemical", "laboratory", "molecule"],

    # General / Lifestyle / Media
    "kosmos": ["space", "galaxy", "planet", "astronomy", "rocket", "universe"],
    "avtomobil": ["car", "automotive", "vehicle", "transport", "racing"],
    "sport": ["sport", "fitness", "workout", "gym", "football", "soccer"],
    "musiqa": ["music", "concert", "sound", "audio", "instrument"],
    "oshxona": ["food", "cooking", "restaurant", "recipe", "culinary"],
    "bolalar": ["kids", "children", "cute", "cartoon"],
}


class SmartTemplateSelector:
    """
    Ranks templates from the 17,250 catalog against a user's prompt
    using multilingual semantic expansion and weighted relevance scoring.
    """

    @classmethod
    def expand_query_terms(cls, query: str) -> Set[str]:
        """Extracts and semantically expands query words into multi-language keywords."""
        q_clean = (query or "").lower().strip()
        terms = set(re.findall(r"\w+", q_clean))

        # Check multi-word keys first, then single words
        for phrase, expansions in MULTILINGUAL_SEMANTIC_MAP.items():
            if phrase in q_clean:
                terms.update(expansions)
            else:
                for word in phrase.split():
                    if word in terms:
                        terms.update(expansions)

        # Filter out trivial stop words
        stop_words = {"va", "bilan", "uchun", "haqida", "dagi", "and", "the", "for", "with", "about"}
        return {t for t in terms if len(t) >= 2 and t not in stop_words}

    @classmethod
    def search(
        cls,
        catalog: Dict[str, Dict[str, Any]],
        query: str,
        top_k: int = 6,
        target_category: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        if not catalog:
            return []

        expanded_terms = cls.expand_query_terms(query)
        if not expanded_terms:
            return []

        # Regex'ni 17,250 marta emas, qidiruv boshida FAQAT 1 MARTA kompilyatsiya qilamiz
        escaped = [re.escape(t) for t in expanded_terms if len(t) >= 2]
        word_boundary_re = re.compile(r"\b(?:" + "|".join(escaped) + r")\b") if escaped else None
        substr_terms = [t for t in expanded_terms if len(t) >= 3]

        scored: List[tuple[float, Dict[str, Any]]] = []

        for item in catalog.values():
            title = item.get("_title_lc") or item.get("title", "").lower()
            filename = item.get("_fn_lc") or item.get("filename", "").lower()
            tags = item.get("tags") or ()
            cat = item.get("category", "")

            term_score = 0.0
            if word_boundary_re:
                exact_hits = set(word_boundary_re.findall(title))
                term_score += len(exact_hits) * 40.0
            else:
                exact_hits = set()

            for term in substr_terms:
                if term not in exact_hits and term in title:
                    term_score += 20.0
                if term in filename:
                    term_score += 15.0

            if tags:
                tag_hits = sum(1 for t in expanded_terms if t in tags)
                term_score += tag_hits * 30.0

            # MUHIM: Toifa va HD muqova bonuslari faqat kamida 1 ta so'z mos kelgandagina beriladi!
            if term_score > 0:
                score = term_score
                if target_category and cat == target_category:
                    score += 35.0
                if item.get("thumb_file_id"):
                    score += 10.0
                if item.get("pptx_file_id"):
                    score += 10.0
                scored.append((score, item))

        scored.sort(key=lambda x: -x[0])
        return [item for _, item in scored[:top_k]]
