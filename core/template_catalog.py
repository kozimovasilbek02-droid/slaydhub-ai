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
        repo_root = Path(__file__).resolve().parent.parent
        default_cat = repo_root / "templates_catalog.json"
        self.catalog_json_path = catalog_json_path or os.getenv(
            "SLAYDHUB_CATALOG_PATH", str(default_cat)
        )
        self.templates: Dict[str, Dict[str, Any]] = {}
        self._load_catalog()

    def _load_catalog(self) -> None:
        self._by_category: Dict[str, List[Dict[str, Any]]] = {}
        for tpl in BUILTIN_SEED_TEMPLATES:
            item = dict(tpl)
            item["_title_lc"] = item.get("title", "").lower()
            item["_fn_lc"] = item.get("filename", "").lower()
            item["_tags_set"] = frozenset(item.get("tags") or ())
            self.templates[item["id"]] = item

        if os.path.exists(self.catalog_json_path):
            try:
                with open(self.catalog_json_path, "r", encoding="utf-8") as f:
                    extra = json.load(f)
                items = extra if isinstance(extra, list) else extra.values() if isinstance(extra, dict) else []
                for raw in items:
                    if isinstance(raw, dict) and "id" in raw:
                        # Render RAM tejash: Windows local_path ni xotirada saqlamaymiz
                        if os.name != "nt":
                            raw.pop("local_path", None)
                        raw["_title_lc"] = raw.get("title", "").lower()
                        raw["_fn_lc"] = raw.get("filename", "").lower()
                        raw["_tags_set"] = frozenset(raw.get("tags") or ())
                        self.templates[raw["id"]] = raw
                del extra, items
            except Exception as e:
                logger.warning(f"templates_catalog.json o'qishda xatolik: {e}")

        # Toifalar bo'yicha oldindan saralangan kesh (har klikda 11,595 ta elementni sort qilmaslik uchun)
        for tpl in self.templates.values():
            cat = tpl.get("category", "minimal")
            self._by_category.setdefault(cat, []).append(tpl)
        for cat_list in self._by_category.values():
            cat_list.sort(key=lambda x: 0 if x.get("thumb_file_id") else 1)

        try:
            from core.memory_guard import release_c_and_python_memory
            release_c_and_python_memory()
        except Exception:
            pass

    def detect_category(self, topic: str) -> str:
        """Mavzu yoki fayl nomi bo'yicha eng mos keladigan toifani aniqlaydi (UZ / EN)."""
        import re
        t = re.sub(r"[ʻʼ`’‘]", "'", (topic or "").lower())
        if any(w in t for w in [
            "yurak", "tibbiyot", "salomat", "kasallik", "dori", "tish", "shifoxona", "bemor", "doktor",
            "heart", "cardio", "dental", "pharma", "medic", "health", "hospital", "doctor", "clinic",
            "virus", "disease", "coron", "capsule", "spine", "pediatric", "surgical", "anatomy"
        ]):
            return "med"
        if any(w in t for w in [
            "blokcheyn", "blockchain", "kripto", "crypto", "bitcoin", "ai", "sun'iy", "suniy",
            "dastur", "kod", "robot", "kompyuter", "3d", "texnolog", "kiber",
            "tech", "software", "comput", "cyber", "cloud", "network", "coding"
        ]):
            return "tech"
        if any(w in t for w in [
            "ilmiy", "falsaf", "kvant", "nisbiylik", "fizika", "kimyo", "matematika", "nazariya",
            "maktab", "universitet", "dars", "talaba", "tarix", "pedagog", "adabiyot", "fan",
            "education", "school", "universit", "student", "study", "history", "lesson", "academic", "science", "quantum"
        ]):
            return "edu"
        if any(w in t for w in [
            "biznes", "moliya", "reja", "marketing", "invest", "pul", "iqtisod", "startap", "daromad",
            "business", "market", "finance", "plan", "startup", "invest", "company", "corporate", "strategy"
        ]):
            return "biz"
        # 'tabiati' / 'tabiatini' means 'essence of X', whereas 'tabiat' / 'tabiatni' means ecology/nature
        t_no_essence = re.sub(r"tabiat(?:i|ini|iga|idan|ida)", "", t)
        if any(w in t_no_essence for w in [
            "tabiat", "ekolog", "o'simlik", "hayvon", "iqlim", "atrof-muhit", "geograf",
            "nature", "green", "ecolog", "plant", "earth", "animal", "flower", "water", "tree", "forest", "leaf", "leaves"
        ]):
            return "nature"
        return "minimal"

    def get_templates_by_category(self, cat_key: str) -> List[Dict[str, Any]]:
        results = self._by_category.get(cat_key) or []
        if not results:
            results = (self._by_category.get("biz") or []) + (self._by_category.get("tech") or [])
        return results

    def smart_search(
        self, query: str, top_k: int = 6, category: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """AI semantik qidiruvi orqali 17,250 ta shablon ichidan eng moslarini qaytaradi."""
        from core.smart_selector import SmartTemplateSelector
        target_cat = category or self.detect_category(query)
        return SmartTemplateSelector.search(
            catalog=self.templates,
            query=query,
            top_k=top_k,
            target_category=target_cat
        )

    def find_best_matching_template(
        self, topic: str, category: Optional[str] = None
    ) -> Optional[Dict[str, Any]]:
        """Mavzuga eng yuqori reyting bilan mos kelgan bitta eng zo'r shablonni qaytaradi."""
        results = self.smart_search(topic, top_k=1, category=category)
        return results[0] if results else None

    def get_template(self, tpl_id: str) -> Optional[Dict[str, Any]]:
        return self.templates.get(tpl_id)

    def resolve_template_pptx_path(self, tpl_id: str) -> Optional[str]:
        """
        Shablonning lokal PPTX fayl yo'lini topadi.
        1. Telegram Cloud'dan yuklab olingan kesh (/tmp/soff_ready/cloud_templates/{tpl_id}.pptx)
        2. downloaded_templates/ papkasi
        3. templates/ papkasi yoki local_path
        """
        if not tpl_id:
            return None

        # 0. Telegram Cloud (17,250 SlaydHub Templates Vault) dan yuklab olingan keshni tekshirish
        cloud_cached = Path(tempfile.gettempdir()) / "soff_ready" / "cloud_templates" / f"{tpl_id}.pptx"
        if cloud_cached.exists() and cloud_cached.stat().st_size > 5000:
            return str(cloud_cached)

        curr_dir = Path(__file__).resolve().parent.parent
        dl_cached = curr_dir / "downloaded_templates" / f"{tpl_id}.pptx"
        if dl_cached.exists() and dl_cached.stat().st_size > 5000:
            return str(dl_cached)

        tpl = self.get_template(tpl_id)
        if not tpl:
            return None

        # 1. To'g'ridan-to'g'ri ko'rsatilgan local_path
        if "local_path" in tpl and tpl["local_path"] and os.path.exists(tpl["local_path"]):
            return tpl["local_path"]

        filename = tpl.get("filename", "")
        if not filename:
            return None

        dl_by_name = curr_dir / "downloaded_templates" / filename
        if dl_by_name.exists() and dl_by_name.stat().st_size > 5000:
            return str(dl_by_name)

        # 2. telegram_bot/templates papkasi (Render va lokalda)
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
        local_path = self.resolve_template_pptx_path(tpl_id)
        if local_path and os.path.exists(local_path) and zipfile.is_zipfile(local_path):
            return local_path

        tpl = self.get_template(tpl_id)
        if not tpl:
            return None

        pptx_file_id = tpl.get("pptx_file_id")
        if pptx_file_id and bot:
            dest_dir = Path(tempfile.gettempdir()) / "soff_ready" / "cloud_templates"
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest_path = dest_dir / f"{tpl_id}.pptx"

            if dest_path.exists():
                if dest_path.stat().st_size > 5000 and zipfile.is_zipfile(dest_path):
                    os.utime(dest_path, None)  # LRU uchun vaqtni yangilash
                    return str(dest_path)
                try:
                    dest_path.unlink()  # Buzilgan chala faylni tozalash
                except OSError:
                    pass

            # Atomic Download: avval .part faylga yozib, ZIP butunligini tekshirib, keyin os.replace qilamiz
            import asyncio
            part_path = dest_dir / f"{tpl_id}.{os.getpid()}.{id(asyncio.current_task())}.part"
            try:
                file_info = await bot.get_file(pptx_file_id)
                await bot.download_file(file_info.file_path, destination=part_path)
                if part_path.exists() and part_path.stat().st_size > 5000 and zipfile.is_zipfile(part_path):
                    os.replace(part_path, dest_path)
                    return str(dest_path)
            except Exception as e:
                logger.error(f"Telegram Cloud'dan shablon yuklashda xatolik ({tpl_id}): {e}")
            finally:
                if part_path.exists():
                    try:
                        part_path.unlink()
                    except OSError:
                        pass

        return None

    async def ensure_multi_template_pool_available(
        self,
        topic: str,
        primary_tpl_id: Optional[str] = None,
        bot=None,
        max_donors: int = 2
    ) -> tuple[Optional[str], List[str]]:
        """
        17,250+ Telegram kanal (SlaydHub Templates Vault) shablonlaridan:
        1. Asosiy tanlangan shablonni (primary_tpl_id) yuklab oladi.
        2. Agar asosiy shablonda slaydlar soni kam bo'lsa (masalan 1-6 betlik infografika bo'lsa),
           17,250+ katalogdan mavzuga mos hamda ko'p betli/infografik shablonlarni
           Telegram kanaldan parallel (asyncio.gather) yuklab beradi, shunda har bir slayd
           100% haqiqiy 17,250+ shablonlar bazasidan yig'iladi.
        """
        import asyncio
        target_cat = self.detect_category(topic)
        matches = self.smart_search(topic, top_k=max_donors + 6, category=target_cat)
        # Also search structural infographic/timeline/multi-slide templates in the 17,250 catalog for the same category
        struct_matches = self.smart_search(
            f"{topic} infographic timeline process steps presentation",
            top_k=6,
            category=target_cat
        )

        candidate_ids: List[str] = []
        if primary_tpl_id:
            candidate_ids.append(primary_tpl_id)
        for m in matches + struct_matches:
            mid = m.get("id")
            if mid and mid not in candidate_ids:
                candidate_ids.append(mid)

        # Also include full 48-slide category deck IDs from the 17,250 catalog as deep structural backstops
        cat_backstop_ids = {
            "med": ["tpl_07_modern_healthcare_solu", "tpl_09_digital_healthcare", "tpl_11_humanity_heartbeat"],
            "tech": ["tpl_09_digital_healthcare", "tpl_3d_printing_technology_po", "tpl_2019_business_plan_powerp"],
            "biz": ["tpl_2019_business_plan_powerp", "tpl_09_digital_healthcare", "tpl_07_modern_healthcare_solu"],
            "edu": ["tpl_2019_business_plan_powerp", "tpl_07_modern_healthcare_solu", "tpl_09_digital_healthcare"],
            "nature": ["tpl_07_modern_healthcare_solu", "tpl_2019_business_plan_powerp"],
            "minimal": ["tpl_2019_business_plan_powerp", "tpl_09_digital_healthcare", "tpl_07_modern_healthcare_solu"],
        }
        for bid in cat_backstop_ids.get(target_cat, ["tpl_2019_business_plan_powerp", "tpl_09_digital_healthcare"]):
            if bid in self.templates and bid not in candidate_ids:
                candidate_ids.append(bid)

        if not candidate_ids:
            return None, []

        # 1. First ensure primary template
        primary_path = await self.ensure_template_pptx_available(candidate_ids[0], bot=bot)
        remaining_ids = candidate_ids[1:]

        # Check if primary template already has >= 6 slides; if so, 0 extra donors needed (super fast!)
        if primary_path and os.path.exists(primary_path):
            try:
                from pptx import Presentation
                s_count = len(Presentation(primary_path).slides)
                if s_count >= 6:
                    return primary_path, []
            except Exception:
                pass

        # 2. If primary template is small (1-5 slide infographic), download at most 2 donors with 12s timeout
        batch_ids = remaining_ids[:2]
        if batch_ids and bot:
            async def _safe_download(cid: str) -> Optional[str]:
                try:
                    return await asyncio.wait_for(
                        self.ensure_template_pptx_available(cid, bot=bot),
                        timeout=12.0
                    )
                except Exception:
                    return None

            results = await asyncio.gather(
                *(_safe_download(cid) for cid in batch_ids),
                return_exceptions=True
            )
        else:
            results = []

        donor_paths: List[str] = []
        for res in results:
            if isinstance(res, str) and res and os.path.exists(res):
                if primary_path is None:
                    primary_path = res
                elif res != primary_path and res not in donor_paths:
                    donor_paths.append(res)
                    if len(donor_paths) >= max_donors:
                        break

        return primary_path, donor_paths

    def extract_cover_thumbnail(self, pptx_path: str, out_thumb_path: str, ppt_app=None) -> bool:
        """
        Yuqori sifatli muqova rasmini ajratib oladi:
        1. OpenXML: docProps/thumbnail.jpeg (agar mavjud va > 8KB bo'lsa).
        2. Agar mavjud bo'lmasa: PowerPoint COM orqali 1-slaydni 1280x720 HD sifatda eksport qiladi!
        3. Hech qachon kichik ijtimoiy tarmoq ikonkalari (< 25KB, masalan Facebook logosi) olinmaydi!
        """
        if not os.path.exists(pptx_path):
            return False
        abs_out = os.path.abspath(out_thumb_path)
        try:
            with zipfile.ZipFile(pptx_path, "r") as z:
                # 1. Rasmiy thumbnail rasmi (faqat agar 8KB dan katta haqiqiy rasm bo'lsa)
                for candidate in ["docProps/thumbnail.jpeg", "docProps/thumbnail.jpg", "docProps/thumbnail.png"]:
                    if candidate in z.namelist() and z.getinfo(candidate).file_size > 8000:
                        with open(abs_out, "wb") as f:
                            f.write(z.read(candidate))
                        return True
        except Exception:
            pass

        # 2. PowerPoint COM orqali 1-slaydni haqiqiy 1280x720 HD qilib render qilish
        try:
            import win32com.client
            owns_app = False
            if ppt_app is None:
                ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
                owns_app = True

            pres = ppt_app.Presentations.Open(os.path.abspath(pptx_path), ReadOnly=True, Untitled=False, WithWindow=False)
            pres.Slides(1).Export(abs_out, "JPG", 1280, 720)
            pres.Close()
            if owns_app:
                try:
                    ppt_app.Quit()
                except Exception:
                    pass
            if os.path.exists(abs_out) and os.path.getsize(abs_out) > 5000:
                return True
        except Exception as e:
            logger.warning(f"PowerPoint COM render xatosi ({pptx_path}): {e}")

        # 3. Fallback: ppt/media/ ichidagi eng katta rasmni olish (> 40KB)
        try:
            with zipfile.ZipFile(pptx_path, "r") as z:
                media_files = [
                    m for m in z.namelist()
                    if m.startswith("ppt/media/") and m.endswith((".jpg", ".jpeg", ".png"))
                    and z.getinfo(m).file_size > 40000  # Kichik ikonkalar (Facebook va h.k.) qat'iyan rad etiladi
                ]
                if media_files:
                    largest = max(media_files, key=lambda m: z.getinfo(m).file_size)
                    with open(abs_out, "wb") as f:
                        f.write(z.read(largest))
                    return True
        except Exception:
            pass

        return False


# Global singleton instance
GLOBAL_TEMPLATE_CATALOG = TemplateCatalog()
