# -*- coding: utf-8 -*-
"""
core/smart_selector.py
AI-Powered Smart Template Selector & Search Engine for SlaydHub Ultra.
Searches and ranks across all 17,250+ verified templates using multilingual
semantic keyword expansion, category relevance weighting, and quality scoring.
"""

import heapq
import re
import logging
from typing import List, Dict, Any, Optional, Set

logger = logging.getLogger(__name__)

# Comprehensive Multilingual (UZ / RU / EN) domain concepts map
MULTILINGUAL_SEMANTIC_MAP: Dict[str, List[str]] = {
    # Time, Philosophy, Physics, Science, Cosmology
    "vaqt": ["time", "timeline", "chronology", "clock", "history", "roadmap", "evolution", "process", "future"],
    "zamon": ["time", "modern", "era", "timeline", "future", "chronology"],
    "falsafa": ["philosophy", "theory", "concept", "logic", "science", "mind", "academic", "education"],
    "falsafiy": ["philosophy", "theory", "concept", "science", "academic", "logic"],
    "ilmiy": ["science", "scientific", "research", "academic", "analysis", "education", "innovation"],
    "ilm": ["science", "knowledge", "education", "research", "academic", "study"],
    "fan": ["science", "scientific", "academic", "research", "education", "technology"],
    "kvant": ["quantum", "physics", "science", "atom", "energy", "technology", "future"],
    "nisbiylik": ["physics", "science", "space", "universe", "time", "theory"],
    "termodinamika": ["physics", "energy", "science", "process", "heat", "mechanics"],
    "atom": ["atom", "physics", "science", "nuclear", "energy", "chemistry"],
    "koinot": ["space", "universe", "galaxy", "astronomy", "planet", "cosmos", "science"],
    "kosmos": ["space", "galaxy", "planet", "astronomy", "rocket", "universe"],
    "evolyutsiya": ["evolution", "timeline", "growth", "history", "development", "process"],
    "nazariya": ["theory", "concept", "science", "research", "analysis", "education"],
    "mantiq": ["logic", "mind", "brain", "strategy", "process", "analysis"],

    # Nature, Ecology, Clean Energy
    "quyosh": ["solar", "sun", "renewable", "clean energy"],
    "energiya": ["energy", "power", "electricity", "solar", "battery"],
    "ekologiya": ["ecology", "environment", "eco", "green", "recycling", "climate", "nature"],
    "atrof-muhit": ["environment", "ecology", "green", "nature", "climate"],
    "iqlim": ["climate", "weather", "environment", "ecology", "earth"],
    "o'simlik": ["plant", "botany", "green", "leaf", "nature", "agriculture"],
    "hayvon": ["animal", "wildlife", "biology", "nature", "veterinary"],
    "suv": ["water", "aqua", "ocean", "river", "environment"],
    "qishloq": ["agriculture", "farming", "harvest", "rural"],
    "turizm": ["tourism", "travel", "journey", "vacation", "trip"],

    # Medicine, Healthcare, Biology
    "yurak": ["heart", "cardio", "cardiology", "cardiovascular", "cardiogram", "medical"],
    "kardio": ["cardio", "cardiology", "heart", "pulse", "medical"],
    "tibbiyot": ["medical", "medicine", "healthcare", "health", "clinic", "hospital", "doctor"],
    "salomatlik": ["health", "healthcare", "medical", "wellness", "medicine"],
    "tish": ["dental", "dentist", "teeth", "oral", "stomatology"],
    "dori": ["pharmacy", "pharma", "medication", "capsule", "drug", "prescription"],
    "farmakologiya": ["pharmacy", "pharma", "medication", "capsule", "medical"],
    "kasallik": ["disease", "pathology", "syndrome", "infection", "virus", "medical"],
    "jarrohlik": ["surgery", "surgical", "operation", "surgeon", "hospital"],
    "anatomiya": ["anatomy", "biological", "human body", "spine", "medical"],
    "biologiya": ["biology", "biological", "dna", "genetics", "cell", "science"],
    "genetika": ["dna", "gene", "genetics", "biotechnology", "biology"],
    "psixologiya": ["psychology", "mental", "brain", "behavior", "mind"],
    "pediatriya": ["pediatric", "child health", "baby", "infant", "medical"],

    # Technology, AI, IT, Engineering
    "ai": ["ai", "artificial intelligence", "robot", "machine learning", "neural", "technology"],
    "sun'iy intellekt": ["ai", "artificial intelligence", "robotics", "automation", "deep learning", "technology"],
    "suniy intellekt": ["ai", "artificial intelligence", "robotics", "automation", "deep learning", "technology"],
    "sun'iy": ["ai", "artificial intelligence", "robotics", "neural", "technology"],
    "suniy": ["ai", "artificial intelligence", "robotics", "neural", "technology"],
    "intellekt": ["intelligence", "ai", "artificial intelligence", "smart", "brain", "technology"],
    "robot": ["robot", "robotics", "automation", "android", "technology"],
    "robototexnika": ["robotics", "robot", "automation", "engineering", "technology"],
    "texnologiya": ["technology", "tech", "digital", "innovation", "modern", "future"],
    "raqamli": ["digital", "technology", "tech", "data", "network", "innovation"],
    "dasturlash": ["coding", "programming", "software", "developer", "code", "computer"],
    "dastur": ["software", "app", "application", "program", "coding", "digital"],
    "kiber": ["cyber", "security", "network", "firewall", "technology"],
    "axborot": ["information", "data", "digital", "technology", "network", "it"],
    "bulut": ["cloud", "server", "data center", "network"],
    "blokcheyn": ["blockchain", "crypto", "bitcoin", "web3", "fintech", "digital"],
    "kripto": ["crypto", "cryptocurrency", "bitcoin", "blockchain", "finance"],
    "3d": ["3d printing", "3d", "modeling", "industrial design", "technology"],
    "muhandislik": ["engineering", "mechanical", "industry", "technology"],

    # Business, Economics, Finance, Startups, Strategy
    "biznes": ["business", "company", "corporate", "strategy", "enterprise", "management"],
    "iqtisodiyot": ["economy", "economics", "finance", "market", "business", "investment", "growth"],
    "iqtisod": ["economy", "economics", "finance", "market", "business", "analytics"],
    "moliya": ["finance", "financial", "money", "banking", "fintech", "accounting", "budget"],
    "bank": ["bank", "banking", "finance", "money", "investment", "credit"],
    "invest": ["invest", "investment", "investor", "pitch deck", "venture", "funding"],
    "investitsiya": ["investment", "investor", "finance", "business", "funding", "growth"],
    "startap": ["startup", "pitch deck", "entrepreneur", "innovation", "business"],
    "reja": ["plan", "business plan", "roadmap", "strategy", "timeline"],
    "strategiya": ["strategy", "strategic", "roadmap", "plan", "business", "management"],
    "marketing": ["marketing", "market", "advertising", "brand", "sales", "promotion"],
    "savdo": ["sales", "commerce", "ecommerce", "retail", "market"],
    "boshqaruv": ["management", "leadership", "corporate", "strategy", "business"],
    "menejment": ["management", "business", "strategy", "organization", "corporate"],
    "kpi": ["kpi", "dashboard", "metric", "report", "analytics", "chart"],

    # Education, Academic, History, Law, Literature
    "ta'lim": ["education", "school", "university", "learning", "academic", "teaching"],
    "talim": ["education", "school", "university", "learning", "academic", "teaching"],
    "maktab": ["school", "education", "classroom", "student", "teacher"],
    "universitet": ["university", "college", "academic", "campus", "education"],
    "dars": ["lesson", "lecture", "study", "teaching", "course", "education"],
    "pedagogika": ["education", "teaching", "learning", "school", "academic"],
    "tarix": ["history", "historical", "timeline", "ancient", "heritage", "culture"],
    "tarixiy": ["history", "historical", "timeline", "heritage"],
    "huquq": ["law", "legal", "justice", "governance", "policy"],
    "qonun": ["law", "legal", "justice", "policy", "regulation"],
    "adabiyot": ["literature", "poetry", "books", "reading", "library", "education"],
    "til": ["language", "english", "linguistics", "grammar", "education"],
    "matematika": ["math", "mathematics", "algebra", "geometry", "calculus", "science"],
    "fizika": ["physics", "quantum", "optics", "mechanics", "science", "energy"],
    "kimyo": ["chemistry", "chemical", "laboratory", "molecule", "science"],
    "geografiya": ["geography", "map", "earth", "world", "global"],
    "jamiyat": ["society", "social", "community", "culture", "people"],
    "madaniyat": ["culture", "art", "heritage", "history", "society"],
    "sport": ["sport", "fitness", "workout", "gym", "athletic"],
    "musiqa": ["music", "concert", "sound", "audio", "art"],
}

UZBEK_STOP_WORDS: Set[str] = {
    "va", "bilan", "uchun", "haqida", "dagi", "ning", "dan", "ga", "da", "ni",
    "bu", "shu", "u", "ular", "biz", "siz", "men", "sen",
    "nima", "qanday", "nega", "qachon", "qayerda", "kim", "qaysi",
    "tabiati", "tabiatini", "tushunish", "tushunchasi", "mohiyati", "ahamiyati",
    "asosi", "asoslari", "turlari", "xususiyatlari", "muammolari", "yechimlari",
    "rivojlanishi", "rivoji", "tahlili", "jarayoni", "bosqichlari", "usullari",
    "orqali", "bo'yicha", "boyicha", "asosida", "doirasida", "hamda", "yoki",
    "and", "the", "for", "with", "about", "from", "into", "that", "this", "what", "how",
}


class SmartTemplateSelector:
    """
    Ranks templates from the 17,250 catalog against a user's prompt
    using multilingual semantic expansion and weighted relevance scoring.
    """

    @classmethod
    def expand_query_terms(cls, query: str) -> tuple[Set[str], Set[str]]:
        """
        Returns (all_expanded_terms, english_semantic_terms).
        Normalizes apostrophes and filters out Uzbek stop words so words like 'nima'
        never match 'minimal' as a substring!
        """
        q_clean = re.sub(r"[ʻʼ`’‘]", "'", (query or "").lower().strip())
        raw_tokens = [w for w in re.findall(r"\w+(?:'\w+)?", q_clean) if len(w) >= 2 and w not in UZBEK_STOP_WORDS]

        # Normalize Uzbek suffixes (-da, -ga, -ni, -ning, -dagi, -lar, -lari, -iy, -viy) for dictionary lookup
        normalized_tokens = set(raw_tokens)
        for tok in raw_tokens:
            for suf in ("ning", "dagi", "lari", "lar", "ni", "da", "ga", "dan", "da", "si", "i"):
                if len(tok) > len(suf) + 2 and tok.endswith(suf):
                    stem = tok[:-len(suf)]
                    if stem not in UZBEK_STOP_WORDS:
                        normalized_tokens.add(stem)

        semantic_terms: Set[str] = set()
        for phrase, expansions in MULTILINGUAL_SEMANTIC_MAP.items():
            if " " in phrase:
                if phrase in q_clean:
                    semantic_terms.update(expansions)
            elif phrase in normalized_tokens:
                semantic_terms.update(expansions)

        # Only keep raw tokens >= 4 chars that aren't stop words
        valid_raw = {t for t in normalized_tokens if len(t) >= 4 and t not in UZBEK_STOP_WORDS}
        all_terms = semantic_terms | valid_raw
        return all_terms, semantic_terms

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

        expanded_terms, semantic_terms = cls.expand_query_terms(query)
        escaped = [re.escape(t) for t in expanded_terms if len(t) >= 2]
        word_boundary_re = re.compile(r"(?:" + "|".join(escaped) + r")") if escaped else None
        # Only allow prefix/substring matching for long semantic terms (>= 5 chars)
        long_substr_terms = [t for t in semantic_terms if len(t) >= 5]

        scored: List[tuple[float, Dict[str, Any]]] = []

        for item in catalog.values():
            title = item.get("_title_lc") or item.get("title", "").lower()
            filename = item.get("_fn_lc") or item.get("filename", "").lower()
            tags = item.get("tags") or ()
            cat = item.get("category", "")

            # Skip calendar/holiday/certificate templates unless explicitly requested
            if any(bad in title for bad in ("calendar", "holiday", "christmas", "valentine", "birthday", "certificate", "invitation", "resume", "cv ")):
                if not any(bad in (query or "").lower() for bad in ("kalendar", "calendar", "bayram", "sertifikat", "rezyume")):
                    continue

            term_score = 0.0
            generic_words = {"business", "modern", "digital", "science", "education", "technology", "tech", "process", "plan", "company", "corporate", "data", "market"}
            if word_boundary_re:
                exact_hits = set(word_boundary_re.findall(title))
                for eh in exact_hits:
                    term_score += 30.0 if eh in generic_words else 75.0
            else:
                exact_hits = set()

            for term in long_substr_terms:
                if term not in exact_hits and term in title:
                    term_score += 20.0
                if term in filename:
                    term_score += 12.0

            tag_set = item.get("_tags_set") or tags
            if tag_set:
                tag_hits = sum(1 for t in expanded_terms if t in tag_set)
                term_score += tag_hits * 35.0

            if term_score > 0:
                score = term_score
                if target_category and cat == target_category:
                    score += 35.0
                # Reward multi-slide presentations & rich infographic diagrams in the 17,250 catalog
                if any(w in title for w in ("presentation", "powerpoint templates", "pitch deck", "overview", "infographic", "timeline", "process", "diagram", "steps", "matrix", "roadmap")):
                    score += 25.0
                if item.get("total_slides") and int(item.get("total_slides", 0)) >= 15:
                    score += 20.0
                if item.get("thumb_file_id"):
                    score += 10.0
                if item.get("pptx_file_id"):
                    score += 10.0
                scored.append((score, item))

        # Fallback: if query had fewer than top_k keyword matches, fill with high-structure templates from target_category
        if len(scored) < top_k:
            seen_ids = {it["id"] for _, it in scored}
            for item in catalog.values():
                if item["id"] in seen_ids or not item.get("pptx_file_id"):
                    continue
                title = item.get("_title_lc") or item.get("title", "").lower()
                if any(bad in title for bad in ("calendar", "holiday", "christmas", "birthday", "certificate")):
                    continue
                cat = item.get("category", "")
                if target_category and cat != target_category and cat not in ("minimal", "biz", "edu", "tech"):
                    continue
                fb_score = 15.0 if cat == target_category else 5.0
                if any(w in title for w in ("infographic", "presentation", "timeline", "process", "steps", "diagram", "powerpoint templates")):
                    fb_score += 20.0
                if item.get("total_slides") and int(item.get("total_slides", 0)) >= 20:
                    fb_score += 25.0
                if fb_score >= 25.0:
                    scored.append((fb_score, item))
                    if len(scored) >= top_k * 4:
                        break

        top_items = heapq.nlargest(top_k, scored, key=lambda x: x[0])
        return [item for _, item in top_items]
