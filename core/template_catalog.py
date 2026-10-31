# -*- coding: utf-8 -*-
"""
core/template_catalog.py
Professional Template Catalog & Cloud Delivery Engine for SlaydHub Ultra.
- Manages 1,000+ authentic presentation templates categorized by sector.
- Automatically extracts embedded high-res cover thumbnails (docProps/thumbnail.jpeg).
- Works transparently across local disk storage and Telegram Cloud (TELEGRAM_TEMPLATES_CHANNEL_ID).
- Built-in metadata seed for instant 24/7 Render deployment.
"""

import os
import io
import json
import zipfile
import logging
import tempfile
from pathlib import Path
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)

# Toifalar ro'yxati
CATEGORIES = {
    "med": {"name": "Tibbiyot va Salomatlik", "emoji": "🩺", "badge": "Tibbiyot"},
    "biz": {"name": "Biznes, Startap va Moliya", "emoji": "💼", "badge": "Biznes"},
    "tech": {"name": "IT, AI va Texnologiya", "emoji": "💻", "badge": "Texnologiya"},
    "edu": {"name": "Ta'lim, Fan va Akademik", "emoji": "🎓", "badge": "Ta'lim"},
    "nature": {"name": "Tabiat, Ekologiya va Turizm", "emoji": "🌿", "badge": "Ekologiya"},
    "minimal": {"name": "Zamonaviy Minimalizm", "emoji": "🎨", "badge": "Minimalist"},
}

# Dastlabki 8 ta asosiy shablon (har qanday muhitda darhol ishlashi uchun)
BUILTIN_SEED_TEMPLATES: List[Dict[str, Any]] = [
    {
        "id": "tpl_med_heartbeat",
        "title": "Humanity Heartbeat & Cardiology",
        "category": "med",
        "filename": "11_Humanity_Heartbeat.pptx",
        "total_slides": 48,
        "tags": ["kardio", "yurak", "tibbiyot", "salomatlik", "doktor", "infarkt", "qon"],
    },
    {
        "id": "tpl_med_healthcare",
        "title": "Modern Healthcare Solutions",
        "category": "med",
        "filename": "07_Modern_Healthcare_Solutions.pptx",
        "total_slides": 35,
        "tags": ["tibbiyot", "shifoxona", "terapiya", "kasallik", "bemor", "tashxis"],
    },
    {
        "id": "tpl_med_dental",
        "title": "Dental Health & Oral Care",
        "category": "med",
        "filename": "16_Dental_Health_Care.pptx",
        "total_slides": 32,
        "tags": ["tish", "stomatolog", "dental", "implant", "og'iz"],
    },
    {
        "id": "tpl_med_pharma",
        "title": "Capsules Medication & Balancing",
        "category": "med",
        "filename": "20_Capsules_Medication_Balancing.pptx",
        "total_slides": 30,
        "tags": ["dori", "farmakologiya", "kapsula", "tabletka", "retsept"],
    },
    {
        "id": "tpl_med_virology",
        "title": "Virology & Epidemic Prevention",
        "category": "med",
        "filename": "01_Coronavirus.pptx",
        "total_slides": 28,
        "tags": ["virus", "epidemiya", "infeksiya", "profilaktika", "vaksina"],
    },
    {
        "id": "tpl_tech_digital_health",
        "title": "Digital Healthcare & AI Tech",
        "category": "tech",
        "filename": "09_Digital_Healthcare.pptx",
        "total_slides": 36,
        "tags": ["ai", "sun'iy intellekt", "raqamli", "texnologiya", "dastur"],
    },
    {
        "id": "tpl_tech_3d_print",
        "title": "3D Printing & Future Engineering",
        "category": "tech",
        "filename": "3D Printing Technology PowerPoint Templates.pptx",
        "total_slides": 27,
        "tags": ["muhandislik", "3d", "texnika", "kompyuter", "innovatsiya", "robot"],
    },
    {
        "id": "tpl_biz_plan_pro",
        "title": "Corporate Business Plan & Strategy",
        "category": "biz",
        "filename": "2019 Business Plan PowerPoint Templates.pptx",
        "total_slides": 42,
        "tags": ["biznes", "moliya", "reja", "marketing", "investitsiya", "daromad", "startap"],
    },
]


class TemplateCatalog:
    """Shablonlar katalogi va bulut yetkazib berish markazi."""

    def __init__(self, catalog_json_path: Optional[str] = None):
        self.catalog_json_path = catalog_json_path or os.getenv(
            "SLAYDHUB_CATALOG_PATH", "templates_catalog.json"
        )
        self.templates: Dict[str, Dict[str, Any]] = {}
        self._load_catalog()

    def _load_catalog(self) -> None:
        """Katalogni yuklaydi: avval builtin shablonlar, so'ngra json fayl ma'lumotlari."""
        for tpl in BUILTIN_SEED_TEMPLATES:
            self.templates[tpl["id"]] = dict(tpl)

        # Agar templates_catalog.json fayli mavjud bo'lsa, ustiga birlashtiramiz
        if os.path.exists(self.catalog_json_path):
            try:
                with open(self.catalog_json_path, "r", encoding="utf-8") as f:
                    extra = json.load(f)
                    if isinstance(extra, list):
                        for item in extra:
                            if "id" in item:
                                self.templates[item["id"]] = item
                    elif isinstance(extra, dict):
                        self.templates.update(extra)
                logger.info(f"Templates catalog yuklandi: {len(self.templates)} ta shablon.")
            except Exception as e:
                logger.warning(f"templates_catalog.json o'qishda xatolik: {e}")

    def detect_category(self, topic: str) -> str:
        """Mavzu yoki fayl nomi bo'yicha eng mos keladigan toifani aniqlaydi (UZ / EN)."""
        t = (topic or "").lower()
        if any(w in t for w in [
            "yurak", "tibbiyot", "salomat", "kasallik", "dori", "tish", "shifoxona", "bemor", "doktor",
            "heart", "cardio", "dental", "pharma", "medic", "health", "hospital", "doctor", "clinic",
            "virus", "disease", "coron", "capsule", "spine", "pediatric", "surgical", "anatomy"
        ]):
            return "med"
        if any(w in t for w in [
            "biznes", "moliya", "reja", "marketing", "invest", "pul", "iqtisod", "startap", "daromad",
            "business", "market", "finance", "plan", "startup", "invest", "company", "corporate", "strategy"
        ]):
            return "biz"
        if any(w in t for w in [
            "ai", "sun'iy", "dastur", "kod", "it", "robot", "kompyuter", "3d", "texnolog",
            "tech", "software", "comput", "digital", "data", "cyber", "cloud", "network", "analytics", "coding"
        ]):
            return "tech"
        if any(w in t for w in [
            "maktab", "universitet", "dars", "talaba", "tarix", "pedagog", "adabiyot", "fan", "til",
            "education", "school", "universit", "student", "study", "history", "lesson", "academic", "learn", "teach"
        ]):
            return "edu"
        if any(w in t for w in [
            "tabiat", "ekolog", "o'simlik", "hayvon", "iqlim", "suv", "geograf",
            "nature", "green", "ecolog", "plant", "earth", "animal", "flower", "water", "tree", "forest", "leaf", "leaves"
        ]):
            return "nature"
        return "minimal"

    def get_templates_by_category(self, cat_key: str) -> List[Dict[str, Any]]:
        """Toifa bo'yicha shablonlar ro'yxatini qaytaradi."""
        results = [t for t in self.templates.values() if t.get("category") == cat_key]
        if not results:
            # Fallback to general templates if category has no specific templates yet
            results = [t for t in self.templates.values() if t.get("category") in ("biz", "tech")]
        return results

    def get_template(self, tpl_id: str) -> Optional[Dict[str, Any]]:
        return self.templates.get(tpl_id)

    def resolve_template_pptx_path(self, tpl_id: str) -> Optional[str]:
        """
        Shablonning lokal PPTX fayl yo'lini topadi.
        Render (templates/ papkasi) yoki lokal mashina diskidan qidiradi.
        """
        tpl = self.get_template(tpl_id)
        if not tpl:
            return None

        # 1. To'g'ridan-to'g'ri ko'rsatilgan local_path
        if "local_path" in tpl and os.path.exists(tpl["local_path"]):
            return tpl["local_path"]

        filename = tpl.get("filename", "")
        if not filename:
            return None

        # 2. telegram_bot/templates papkasi (Render va lokalda)
        curr_dir = Path(__file__).resolve().parent.parent
        bundled = curr_dir / "templates" / filename
        if bundled.exists():
            return str(bundled)

        # 3. Medical_Presentations_English_50 papkasi
        med_dir = Path(r"C:\Users\user\Desktop\Institut\Suniy intelekt\05_Taqdimotlar_va_Maqolalar\Medical_Presentations_English_50") / filename
        if med_dir.exists():
            return str(med_dir)

        # 4. Desktop slayd katalogi
        slayd_dir = Path(r"C:\Users\user\Desktop\slayd")
        if slayd_dir.exists():
            matched = list(slayd_dir.glob(f"**/{filename}"))
            if matched:
                return str(matched[0])

        return None

    async def ensure_template_pptx_available(self, tpl_id: str, bot=None) -> Optional[str]:
        """
        Shablon faylini doimiy mavjud bo'lishini ta'minlaydi:
        1. Agar shablon lokal diskda bo'lsa (templates/ papkasida), shu zahoti qaytaradi.
        2. Agar kompyuter o'chiq bo'lsa va bot Render bulutida ishlayotgan bo'lsa,
           shablonni Telegram yopiq kanalidan (pptx_file_id) tezkor CDN orqali /tmp ga yuklab oladi!
        """
        local_path = self.resolve_template_pptx_path(tpl_id)
        if local_path and os.path.exists(local_path):
            return local_path

        tpl = self.get_template(tpl_id)
        if not tpl:
            return None

        # Telegram Cloud CDN yuklab olish
        pptx_file_id = tpl.get("pptx_file_id")
        if pptx_file_id and bot:
            try:
                dest_dir = Path(tempfile.gettempdir()) / "soff_ready" / "cloud_templates"
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest_path = dest_dir / f"{tpl_id}.pptx"
                if dest_path.exists() and dest_path.stat().st_size > 5000:
                    return str(dest_path)

                file_info = await bot.get_file(pptx_file_id)
                await bot.download_file(file_info.file_path, destination=dest_path)
                if dest_path.exists() and dest_path.stat().st_size > 5000:
                    logger.info(f"Telegram Cloud'dan shablon muvaffaqiyatli yuklab olindi: {tpl_id}")
                    return str(dest_path)
            except Exception as e:
                logger.error(f"Telegram Cloud'dan shablon yuklashda xatolik ({tpl_id}): {e}")

        return None

    def extract_cover_thumbnail(self, pptx_path: str, out_thumb_path: str) -> bool:
        """
        PPTX ichidagi docProps/thumbnail.jpeg ni 0.002 soniyada ajratib oladi.
        PowerPoint yoki og'ir kutubxonalarsiz, sof OpenXML orqali ishlaydi.
        """
        if not os.path.exists(pptx_path):
            return False
        try:
            with zipfile.ZipFile(pptx_path, "r") as z:
                # 1. Rasmiy thumbnail rasmi
                for candidate in ["docProps/thumbnail.jpeg", "docProps/thumbnail.jpg", "docProps/thumbnail.png"]:
                    if candidate in z.namelist():
                        with open(out_thumb_path, "wb") as f:
                            f.write(z.read(candidate))
                        return True
                # 2. Fallback: 1-slaydning asosiy fotosurati
                for candidate in z.namelist():
                    if candidate.startswith("ppt/media/image1.") and candidate.endswith((".jpg", ".jpeg", ".png")):
                        with open(out_thumb_path, "wb") as f:
                            f.write(z.read(candidate))
                        return True
        except Exception as e:
            logger.warning(f"Thumbnail ajratishda xatolik ({pptx_path}): {e}")
        return False


# Global singleton instance
GLOBAL_TEMPLATE_CATALOG = TemplateCatalog()
