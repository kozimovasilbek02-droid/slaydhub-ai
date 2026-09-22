# -*- coding: utf-8 -*-
"""
Academic Template Catalog & Indexing Engine (Native PPTX Edition)
Indexes and manages 16,000+ authentic, editable PPTX template files from:
- Slidenest (2,741 PPTX)
- PresentationGO (3,314 PPTX)
- PPTMON (2,168 PPTX)
- SlidesCarnival (1,162 PPTX)
- AllPPT (3,441 PPTX)
- Showeet (903 PPTX)
- SlideHunter (648 PPTX)
- PresentationMagazine (1,225 PPTX)
- SlidesMania (370 PPTX)
"""

import os
import glob
import json
import re
import logging
from typing import List, Dict, Any, Optional

from pptx import Presentation
from core.slide_archetypes import SlideArchetypeDetector

from core.config import config

logger = logging.getLogger("AcademicCatalog")

# Base library roots resolved dynamically through config
PPTX_LIBRARY_ROOTS = config.TEMPLATE_DIRS
INDEX_CACHE_PATH = str(config.INDEX_CACHE_PATH)

# Multi-lingual synonyms dictionary for query expansion (UZ/RU -> EN)
MULTILINGUAL_SYNONYMS = {
    # Exact Sciences
    "fizika": ["physics", "quantum", "optics", "atom", "mechanics", "gravity", "thermodynamics", "energy", "force"],
    "kvant": ["quantum", "quantum computing", "quantum ai", "qubit"],
    "matematika": ["mathematics", "math", "algebra", "geometry", "calculus", "matrix", "statistics", "data"],
    "kimyo": ["chemistry", "chemical", "molecule", "reaction", "formula", "organic", "compound"],
    "astronomiya": ["astronomy", "space", "planet", "galaxy", "solar", "universe", "cosmic"],
    
    # Natural Sciences
    "biologiya": ["biology", "biological", "cell", "dna", "genetics", "organism", "evolution", "flora", "fauna"],
    "genetika": ["genetics", "dna", "genome", "gene", "mutation"],
    "ekologiya": ["ecology", "environment", "nature", "climate", "green", "sustainability", "eco"],
    "anatomiya": ["anatomy", "body", "human body", "organ", "physiology", "skeleton", "brain"],
    
    # Medicine & Healthcare
    "tibbiyot": ["medical", "medicine", "health", "healthcare", "clinical", "hospital", "doctor", "clinic"],
    "farmatsevtika": ["pharmacy", "pharmaceutical", "drug", "medicine", "capsule"],
    "kasallik": ["disease", "pathology", "syndrome", "virus", "infection", "disorder"],
    "neyron": ["neuron", "neuroscience", "brain", "nervous system"],
    "psixologiya": ["psychology", "mental health", "mind", "behavior", "psychological"],

    # Engineering & IT
    "dasturlash": ["programming", "coding", "software", "developer", "development", "code", "python", "java"],
    "sun'iy intellekt": ["artificial intelligence", "ai", "machine learning", "deep learning", "neural network"],
    "ai": ["artificial intelligence", "machine learning", "ai tech", "neural network", "deep learning"],
    "kiberxavfsizlik": ["cybersecurity", "cyber", "security", "encryption", "infosec", "firewall", "hacking"],
    "kriptografiya": ["cryptography", "encryption", "blockchain", "security", "cipher"],
    "ma'lumotlar": ["data", "big data", "analytics", "database", "data science"],
    "tarmoq": ["network", "cloud", "telecom", "internet", "iot"],
    "robototexnika": ["robotics", "robot", "automation", "mechatronics"],

    # Humanities & Arts
    "tarix": ["history", "historical", "ancient", "civilization", "archaeology", "heritage", "empire", "era"],
    "falsafa": ["philosophy", "philosophical", "ethics", "thinker", "epistemology", "wisdom", "logic"],
    "adabiyot": ["literature", "poetry", "books", "author", "prose", "writer", "reading"],
    "tilshunoslik": ["linguistics", "language", "grammar", "phonetics", "communication"],
    "madaniyat": ["culture", "cultural", "art", "tradition", "heritage", "folklore"],

    # Social Sciences & Economics
    "iqtisodiyot": ["economics", "economy", "finance", "financial", "market", "money", "banking", "trade"],
    "moliya": ["finance", "financial", "investment", "banking", "budget", "accounting"],
    "boshqaruv": ["management", "leadership", "organization", "administration", "strategy"],
    "marketing": ["marketing", "market", "advertising", "sales", "brand", "promotion"],
    "huquq": ["law", "legal", "justice", "court", "legislation", "constitution"],
    "talim": ["education", "academic", "learning", "student", "university", "course"],
    "tadqiqot": ["research", "methodology", "study", "analysis", "scientific", "investigation"]
}

# Academic Categories definition
ACADEMIC_CATEGORIES = {
    "exact_sciences": {
        "name_uz": "Aniq fanlar (Fizika, Matematika, Kimyo)",
        "name_ru": "Точные науки (Физика, Математика, Химия)",
        "name_en": "Exact Sciences (Physics, Mathematics, Chemistry)",
        "keywords": [
            "physics", "math", "mathematics", "quantum", "chemistry", "formula", "calculus",
            "equation", "atom", "mechanics", "optics", "thermodynamics", "algebra", "geometry",
            "fizika", "matematika", "kimyo", "mexanika", "kvant", "formula", "nuclear"
        ]
    },
    "natural_sciences": {
        "name_uz": "Tabiiy fanlar (Biologiya, Ekologiya, Genetika)",
        "name_ru": "Естественные науки (Биология, Экология, Генетика)",
        "name_en": "Natural Sciences (Biology, Ecology, Genetics)",
        "keywords": [
            "biology", "genetics", "dna", "ecology", "botany", "zoology", "evolution",
            "nature", "cell", "organism", "environment", "biotechnology", "bio",
            "biologiya", "genetika", "dnk", "ekologiya", "tabiat", "hujayra"
        ]
    },
    "medicine_health": {
        "name_uz": "Tibbiyot va Salomatlik (Anatomiya, Farmatsevtika)",
        "name_ru": "Медицина и Здоровье (Анатомия, Фармацевтика)",
        "name_en": "Medicine & Healthcare (Anatomy, Pharmacology)",
        "keywords": [
            "medical", "medicine", "health", "healthcare", "anatomy", "clinical", "hospital",
            "pharmacy", "doctor", "disease", "treatment", "virus", "vaccine", "surgery", "neuron",
            "tibbiyot", "salomatlik", "anatomiya", "klinik", "kasallik", "dori", "neyron"
        ]
    },
    "engineering_it": {
        "name_uz": "Muhandislik va IT (Dasturlash, AI, Kiberxavfsizlik)",
        "name_ru": "Инженерия и IT (Программирование, ИИ, Кибербезопасность)",
        "name_en": "Engineering & IT (Programming, AI, Cybersecurity)",
        "keywords": [
            "ai", "artificial intelligence", "computer", "data", "big data", "cyber", "security",
            "software", "programming", "algorithm", "network", "robotics", "tech", "technology",
            "cloud", "blockchain", "information", "it", "dasturlash", "kiber", "digital"
        ]
    },
    "humanities_arts": {
        "name_uz": "Gumanitar fanlar (Tarix, Falsafa, Tilshunoslik, Adabiyot)",
        "name_ru": "Гуманитарные науки (История, Философия, Литература)",
        "name_en": "Humanities & Arts (History, Philosophy, Linguistics)",
        "keywords": [
            "history", "philosophy", "literature", "language", "linguistics", "culture", "art",
            "ancient", "civilization", "education", "heritage", "theology", "ethics",
            "tarix", "falsafa", "adabiyot", "tilshunoslik", "madaniyat", "sanat"
        ]
    },
    "social_sciences_economy": {
        "name_uz": "Ijtimoiy-iqtisodiy fanlar (Iqtisodiyot, Huquq, Psixologiya)",
        "name_ru": "Социально-экономические науки (Экономика, Право, Психология)",
        "name_en": "Social & Economic Sciences (Economics, Law, Psychology)",
        "keywords": [
            "economy", "economics", "finance", "business", "management", "psychology", "law",
            "sociology", "market", "marketing", "strategy", "analysis", "leadership", "policy",
            "iqtisod", "moliya", "boshqaruv", "psixologiya", "huquq", "strategiya"
        ]
    },
    "general_academic": {
        "name_uz": "Umumiy Ilmiy / Tadqiqot (Metodologiya, Dissertatsiya)",
        "name_ru": "Общенаучные / Исследования (Методология, Диссертация)",
        "name_en": "General Academic / Research (Methodology, Thesis)",
        "keywords": [
            "academic", "research", "science", "thesis", "methodology", "study", "report",
            "process", "action plan", "timeline", "infographic", "overview", "evaluation",
            "ilmiy", "tadqiqot", "metodologiya", "dissertatsiya", "hisobot"
        ]
    }
}

# Slide Layout Classifications
SLIDE_LAYOUT_TYPES = {
    "cover": {"name_uz": "Muqova / Kirish Sarlavhasi", "name_ru": "Титульный лист / Введение", "name_en": "Title Cover / Intro"},
    "agenda": {"name_uz": "Mundarija / Reja", "name_ru": "Оглавление / План", "name_en": "Agenda / Table of Contents"},
    "theory_concept": {"name_uz": "Nazariy Asoslar va Ta'riflar", "name_ru": "Теоретические основы", "name_en": "Theoretical Foundations"},
    "cards_grid": {"name_uz": "3/4 Talik Konseptual Kartalar", "name_ru": "Карточки классификации", "name_en": "3/4 Column Concept Cards"},
    "comparison_vs": {"name_uz": "Qiyosiy Tahlil (A vs B)", "name_ru": "Сравнительный анализ", "name_en": "Comparative Analysis (A vs B)"},
    "timeline_steps": {"name_uz": "Bosqichma-bosqich Jarayon / Xronologiya", "name_ru": "Пошаговый процесс", "name_en": "Step-by-Step Process"},
    "diagram_anatomy": {"name_uz": "Tizim, Anatomiya va Diagramma", "name_ru": "Система и Диаграмма", "name_en": "System Architecture / Diagram"},
    "metrics_stats": {"name_uz": "Empirik Ko'rsatkichlar va Metrikalar", "name_ru": "Эмпирические данные", "name_en": "Empirical Metrics & Statistics"},
    "case_study": {"name_uz": "Amaliy Tadqiqot va Keyslar", "name_ru": "Практические кейсы", "name_en": "Case Study & Practical Application"},
    "qa_discussion": {"name_uz": "Munozarali Savol-Javoblar (Q&A)", "name_ru": "Вопросы и Дискуссии", "name_en": "Scientific Questions & Discussion"},
    "conclusion_sources": {"name_uz": "Ilmiy Xulosalar va Adabiyotlar", "name_ru": "Научные выводы", "name_en": "Academic Conclusions & References"}
}


class AcademicCatalog:
    """
    Catalog of authentic native PPTX slide templates.
    """
    def __init__(self, index_file: str = INDEX_CACHE_PATH):
        self.index_file = index_file
        self.templates: List[Dict[str, Any]] = []
        self.category_index: Dict[str, List[int]] = {}
        self._slide_archetype_cache: Dict[str, List[Dict[str, Any]]] = {}
        self.is_loaded = False
        self.load_or_build_index()

    def expand_query_terms(self, query: str) -> List[str]:
        """Expands Uzbek/Russian/English terms into English search keywords."""
        q_clean = query.lower().strip()
        q_tokens = re.findall(r'\b\w+\b', q_clean)
        expanded = set(q_tokens)
        
        if q_clean in MULTILINGUAL_SYNONYMS:
            expanded.update(MULTILINGUAL_SYNONYMS[q_clean])

        for token in q_tokens:
            if token in MULTILINGUAL_SYNONYMS:
                expanded.update(MULTILINGUAL_SYNONYMS[token])
            for uz_key, en_synonyms in MULTILINGUAL_SYNONYMS.items():
                if len(token) >= 4 and (token.startswith(uz_key) or uz_key.startswith(token)):
                    expanded.update(en_synonyms)

        return list(expanded)

    def classify_template_name(self, name: str) -> List[str]:
        """Classifies template into academic categories based on title/path."""
        name_lower = name.lower().replace("_", " ").replace("-", " ")
        matched_cats = []
        
        for cat_key, cat_data in ACADEMIC_CATEGORIES.items():
            for kw in cat_data["keywords"]:
                if re.search(r'\b' + re.escape(kw) + r'\b', name_lower):
                    matched_cats.append(cat_key)
                    break

        if not matched_cats:
            matched_cats.append("general_academic")
        return matched_cats

    def classify_topic(self, topic: str) -> str:
        """Classifies a topic into an academic category key."""
        from core.academic_matcher import get_academic_matcher
        return get_academic_matcher().detect_category(topic)

    @staticmethod
    def is_valid_base_template(path: str) -> bool:
        """Strictly validates that file is an authentic PPTX base template, not a test or output artifact."""
        fname = os.path.basename(path).lower()
        if fname.startswith("~$") or fname.startswith("."):
            return False
        if not fname.endswith(".pptx"):
            return False
        # Exclude generated/test output files
        ignore_tokens = [
            "customized", "flawless", "masterpiece", "assembled", "purged",
            "taqdimot", "test_", "temp_", "scratch_", "test-", "render_test",
            "precision", "session", "auto", "torn", "clean.pptx"
        ]
        if any(tok in fname for tok in ignore_tokens):
            return False
        # Exclude if path is inside an output or temp directory
        norm_p = path.lower().replace("\\", "/")
        if "/output/" in norm_p or "/temp_" in norm_p or "/scratch" in norm_p:
            return False
        # Strictly exclude low-quality folders blacklisted by user
        if config.is_folder_blacklisted(path):
            return False
        return True

    def build_index(self, max_templates: int = 18000) -> None:
        """Scans PPTX libraries and indexes all authentic PPTX files."""
        logger.info("Scanning all authentic PPTX template libraries...")
        self.templates = []
        seen_paths = set()

        for lib_root in PPTX_LIBRARY_ROOTS:
            if not os.path.exists(lib_root):
                continue
            
            # Check direct PPTX files in directory
            direct_pptxs = glob.glob(os.path.join(lib_root, "*.pptx"))
            for p_path in direct_pptxs:
                if p_path in seen_paths or not self.is_valid_base_template(p_path):
                    continue
                seen_paths.add(p_path)

                base_name = os.path.splitext(os.path.basename(p_path))[0]
                clean_title = base_name.replace("_", " ").title()
                categories = self.classify_template_name(base_name)

                self.templates.append({
                    "id": f"tmpl_{len(self.templates):05d}",
                    "title": clean_title,
                    "folder_name": base_name,
                    "pptx_path": p_path,
                    "has_pptx": True,
                    "total_slides": 10, # default estimation, refined on demand
                    "categories": categories,
                    "cover_image": "",
                    "slides": []
                })

            # Check subdirectories containing authentic PPTX files
            for root_dir, dirs, files in os.walk(lib_root):
                # Prune blacklisted low-quality folders so they are never traversed
                dirs[:] = [d for d in dirs if not config.is_folder_blacklisted(d)]
                pptx_files = [f for f in files if self.is_valid_base_template(os.path.join(root_dir, f))]
                if not pptx_files:
                    continue

                for pf in pptx_files:
                    full_p = os.path.join(root_dir, pf)
                    if full_p in seen_paths:
                        continue
                    seen_paths.add(full_p)

                    base_name = os.path.splitext(pf)[0]
                    # Only accept preview images that strictly match the PPTX base name or are in dedicated slide folders
                    matched_slide_imgs = sorted(glob.glob(os.path.join(root_dir, f"{glob.escape(base_name)}_slide_*.jpg")) + 
                                               glob.glob(os.path.join(root_dir, f"{glob.escape(base_name)}_slide_*.png")) +
                                               glob.glob(os.path.join(root_dir, "slide_*.jpg")) +
                                               glob.glob(os.path.join(root_dir, "Slide_*.png")))
                    
                    cov_img = matched_slide_imgs[0] if matched_slide_imgs else ""
                    clean_title = base_name.replace("_", " ").title()
                    categories = self.classify_template_name(base_name + " " + os.path.basename(root_dir))

                    self.templates.append({
                        "id": f"tmpl_{len(self.templates):05d}",
                        "title": clean_title,
                        "folder_name": os.path.basename(root_dir),
                        "pptx_path": full_p,
                        "has_pptx": True,
                        "total_slides": len(matched_slide_imgs) if matched_slide_imgs else 10,
                        "categories": categories,
                        "cover_image": cov_img,
                        "slides": [{"slide_number": i+1, "image_path": img} for i, img in enumerate(matched_slide_imgs)]
                    })

                if len(self.templates) >= max_templates:
                    break

        logger.info(f"Indexed {len(self.templates)} authentic PPTX templates successfully.")
        self.save_index()
        self.build_lookups()

    def build_lookups(self) -> None:
        """Builds lookup index for instant retrieval."""
        self.category_index = {cat: [] for cat in ACADEMIC_CATEGORIES.keys()}
        for t_idx, tmpl in enumerate(self.templates):
            for cat in tmpl.get("categories", []):
                if cat in self.category_index:
                    self.category_index[cat].append(t_idx)

    def save_index(self) -> None:
        """Saves index to JSON cache file."""
        try:
            with open(self.index_file, "w", encoding="utf-8") as f:
                json.dump(self.templates, f, ensure_ascii=False, indent=2)
            logger.info(f"Saved native PPTX template index to {self.index_file}")
        except Exception as e:
            logger.error(f"Failed to save template index: {e}")

    def load_or_build_index(self) -> None:
        """Loads index from file or builds if not existing."""
        if os.path.exists(self.index_file):
            try:
                with open(self.index_file, "r", encoding="utf-8") as f:
                    cached = json.load(f)
                valid = [t for t in cached if t.get("pptx_path") and os.path.exists(t["pptx_path"]) and self.is_valid_base_template(t["pptx_path"])]
                if len(valid) >= 5:
                    self.templates = valid
                    self.build_lookups()
                    self.is_loaded = True
                    logger.info(f"Loaded {len(self.templates)} native PPTX templates from cache.")
                    return
            except Exception as e:
                logger.warning(f"Error loading cache ({e}), trying fallback...")
        
        fallback_file = str(config.FALLBACK_INDEX_PATH)
        if os.path.exists(fallback_file):
            try:
                with open(fallback_file, "r", encoding="utf-8") as f:
                    cached = json.load(f)
                valid = [t for t in cached if t.get("pptx_path") and os.path.exists(t["pptx_path"]) and self.is_valid_base_template(t["pptx_path"])]
                if len(valid) >= 5:
                    self.templates = valid
                    self.build_lookups()
                    self.is_loaded = True
                    logger.info(f"Loaded {len(self.templates)} native PPTX templates from fallback cache.")
                    return
            except Exception as e:
                logger.warning(f"Error loading fallback cache ({e}), rebuilding...")
        
        self.build_index()
        self.is_loaded = True

    def search_templates(
        self,
        query: str = "",
        category: Optional[str] = None,
        min_slides: int = 1,
        limit: int = 24
    ) -> List[Dict[str, Any]]:
        """
        Searches authentic PPTX templates with multilingual keywords and relevance ranking.
        """
        if not self.templates:
            return []

        expanded_terms = self.expand_query_terms(query) if query else []
        results = []

        candidate_indices = range(len(self.templates))
        if category and category in self.category_index:
            candidate_indices = self.category_index[category]

        for idx in candidate_indices:
            tmpl = self.templates[idx]
            if tmpl.get("total_slides", 10) < min_slides:
                continue
            score = 0
            if not expanded_terms:
                score = 10
            else:
                title_lower = tmpl["title"].lower()
                folder_lower = tmpl.get("folder_name", "").lower()
                
                for term in expanded_terms:
                    term_l = term.lower()
                    if term_l in title_lower:
                        score += 10
                    elif term_l in folder_lower:
                        score += 6
                    elif any(term_l in cat for cat in tmpl.get("categories", [])):
                        score += 4

            if score > 0:
                results.append((score, tmpl))

        if not results:
            fallback_indices = candidate_indices if candidate_indices else range(len(self.templates))
            for idx in fallback_indices:
                results.append((1, self.templates[idx]))
                if len(results) >= limit:
                    break

        results.sort(key=lambda x: x[0], reverse=True)
        return [r[1] for r in results[:limit]]

    def get_template_by_id(self, template_id: str) -> Optional[Dict[str, Any]]:
        """Returns template metadata by ID."""
        for t in self.templates:
            if t["id"] == template_id:
                return t
        return None

    def inspect_template_archetypes(self, tmpl: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        Inspects all slides in a template PPTX, classifies each slide's visual archetype,
        and caches the result.
        """
        pptx_p = tmpl.get("pptx_path", "")
        if not pptx_p or not os.path.exists(pptx_p):
            return []

        if pptx_p in self._slide_archetype_cache:
            return self._slide_archetype_cache[pptx_p]

        slide_records = []
        try:
            prs = Presentation(pptx_p)
            total = len(prs.slides)
            preview_slides = tmpl.get("slides", [])

            for s_idx, slide in enumerate(prs.slides):
                archetype = SlideArchetypeDetector.detect_archetype(slide, s_idx, total)
                img_p = ""
                if s_idx < len(preview_slides):
                    img_p = preview_slides[s_idx].get("image_path", "")
                elif tmpl.get("cover_image"):
                    img_p = tmpl["cover_image"]

                slide_records.append({
                    "pptx_path": pptx_p,
                    "slide_index": s_idx,
                    "archetype": archetype,
                    "template_id": tmpl.get("id", ""),
                    "template_title": tmpl.get("title", ""),
                    "image_path": img_p
                })
        except Exception as e:
            logger.debug(f"Could not inspect archetypes for {pptx_p}: {e}")

        self._slide_archetype_cache[pptx_p] = slide_records
        return slide_records

    def search_slides_by_archetype(
        self,
        archetype: str,
        category: Optional[str] = None,
        topic: str = "",
        limit: int = 5
    ) -> List[Dict[str, Any]]:
        """
        Searches authentic templates and returns candidate slides matching the requested archetype.
        Prioritizes templates matching the topic and academic category.
        """
        candidate_decks = self.search_templates(query=topic, category=category, limit=12)
        if not candidate_decks:
            candidate_decks = self.search_templates(query="", category=category, limit=12)
        if not candidate_decks:
            candidate_decks = self.templates[:12]

        matches: List[Dict[str, Any]] = []

        # 1. Search candidate decks matching the topic/category
        for tmpl in candidate_decks:
            slides = self.inspect_template_archetypes(tmpl)
            for s in slides:
                if s["archetype"] == archetype:
                    matches.append(s)
                    if len(matches) >= limit:
                        return matches

        # 2. If not enough matches, search broader category pool
        if len(matches) < limit and category and category in self.category_index:
            broad_indices = self.category_index[category][:25]
            for idx in broad_indices:
                tmpl = self.templates[idx]
                if tmpl in candidate_decks:
                    continue
                slides = self.inspect_template_archetypes(tmpl)
                for s in slides:
                    if s["archetype"] == archetype:
                        matches.append(s)
                        if len(matches) >= limit:
                            return matches

        # 3. Fallback: if still no match found, provide best available non-cover slide from top deck
        if not matches and candidate_decks:
            top_deck_slides = self.inspect_template_archetypes(candidate_decks[0])
            non_cover = [s for s in top_deck_slides if s["slide_index"] > 0]
            if non_cover:
                matches.append(non_cover[0])
            elif top_deck_slides:
                matches.append(top_deck_slides[0])

        return matches[:limit]


# Singleton instance
_catalog_instance: Optional[AcademicCatalog] = None

def get_academic_catalog() -> AcademicCatalog:
    global _catalog_instance
    if _catalog_instance is None:
        _catalog_instance = AcademicCatalog()
    return _catalog_instance
