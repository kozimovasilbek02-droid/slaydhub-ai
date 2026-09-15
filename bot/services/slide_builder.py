# -*- coding: utf-8 -*-
import os
import sys
import glob
import random
import re
import json
import urllib.request
import urllib.parse
from PIL import Image
from pathlib import Path
from typing import Dict, Any, List, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = Path(r"c:\Users\user\Desktop\Antigravity\Power Point")
sys.path.insert(0, str(BASE_DIR))

from bot.config import TEMPLATES_DIRS, OUTPUT_DIR, TEMP_DIR
from backend.core.pptx_processor import safe_load_presentation


class SlideBuilderService:
    """
    Mukammal A'lo 5 darajasidagi taqdimot generatori:
    - Ko'p tilli semantik shablon tanlash (tabiat, tibbiyot, falsafa, tarix, iqtisod...)
    - Mavzuga moslashuvchan ranglar palitrasi (Domain Color Themes)
    - To'liq responsive o'lchamlar (16:9 widescreen va 4:3 formatga 100% moslashish, chetga chiqib ketmaslik)
    - Har qanday shablonda (hatto bo'sh yoki rasmli bo'lsa ham) 100% o'qiluvchan Muqova (Hero Card) va Sarlavha
    - Turli xil zamonaviy maketlar: 2 ustunli Qiyosiy tahlil, Bosqichlar (Steps), Savol-Javob (Q&A), Mundarija va Kartalar.
    """

    SEMANTIC_ONTOLOGY = {
        'tabiat': ['nature', 'biology', 'science', 'green', 'environment', 'plant', 'ecology', 'earth'],
        'biolog': ['biology', 'science', 'nature', 'microbiology', 'life'],
        'tibbiyot': ['medical', 'doctor', 'health', 'hospital', 'clinic', 'first aid', 'ambulance'],
        'bemor': ['medical', 'hospital', 'patient', 'health', 'doctor', 'first aid'],
        'transport': ['medical', 'transport', 'ambulance', 'hospital'],
        'санпин': ['medical', 'hospital', 'clean', 'hygiene', 'health', 'doctor'],
        'инфекция': ['medical', 'microbiology', 'health', 'hospital', 'virus'],
        'xulq': ['psychology', 'behavior', 'brain', 'mind', 'science', 'cognitive', 'neural', 'human'],
        'psixolog': ['psychology', 'brain', 'mind', 'neuron', 'neuroscience', 'thinking', 'mental', 'cognitive'],
        'fiziolog': ['physiology', 'biology', 'brain', 'science', 'nerve', 'body', 'neural'],
        'neyro': ['neuro', 'neuron', 'neuroscience', 'brain', 'science', 'technology'],
        'ong': ['mind', 'philosophy', 'psychology', 'brain', 'thinking'],
        'asab': ['neural', 'nerve', 'brain', 'biology', 'science'],
        'tarix': ['history', 'ancient', 'education', 'book', 'culture', 'architecture'],
        'din': ['religion', 'philosophy', 'history', 'ancient', 'culture', 'book'],
        'hurfikr': ['philosophy', 'mind', 'idea', 'thinking', 'debate', 'book'],
        'berdyaev': ['philosophy', 'theology', 'book', 'history', 'culture'],
        'gazzoliy': ['philosophy', 'islam', 'oriental', 'culture', 'book', 'history'],
        'zardusht': ['ancient', 'religion', 'history', 'culture', 'architecture'],
        'buddaviy': ['religion', 'ancient', 'history', 'culture', 'asia', 'monument'],
        'tafakkur': ['psychology', 'mind', 'brain', 'thinking', 'intellect', 'education', 'strategy'],
        'islohot': ['reform', 'government', 'social', 'strategy', 'business', 'development', 'policy'],
        'ijtimoiy': ['social', 'community', 'people', 'society', 'business'],
        'savol': ['quiz', 'education', 'school', 'exam', 'book', 'study', 'class'],
        'maktab': ['school', 'education', 'class', 'student', 'book', 'learning'],
        'sinf': ['school', 'education', 'class', 'study', 'book'],
        'mutafakkir': ['philosophy', 'thinker', 'wisdom', 'book', 'history', 'culture'],
        'kamolot': ['education', 'growth', 'mind', 'wisdom', 'success'],
        'kvant': ['quantum', 'computer', 'technology', 'digital', 'physics', 'ai', 'cyber']
    }

    THEMES = {
        'nature': {
            'name': 'nature',
            'primary': RGBColor(22, 163, 74),       # Forest / Emerald Green
            'primary_dark': RGBColor(20, 83, 45),   # Deep Dark Green
            'card_bg': RGBColor(240, 253, 244),     # Light Emerald
            'border': RGBColor(187, 247, 208),      # Crisp Mint Border
            'title_color': RGBColor(6, 78, 59),     # Dark Green Title
            'badge_bg': RGBColor(22, 163, 74),      # Emerald Badge
            'badge_text': RGBColor(255, 255, 255),
            'tag': 'BIOLOGIYA VA EKOLOGIYA'
        },
        'medical': {
            'name': 'medical',
            'primary': RGBColor(2, 132, 199),       # Clinical Sky/Teal
            'primary_dark': RGBColor(12, 74, 110),  # Deep Medical Navy
            'card_bg': RGBColor(240, 249, 255),     # Ice Blue
            'border': RGBColor(186, 230, 253),      # Sky Blue Border
            'title_color': RGBColor(15, 23, 42),    # Deep Navy
            'badge_bg': RGBColor(2, 132, 199),
            'badge_text': RGBColor(255, 255, 255),
            'tag': 'TIBBIYOT VA SOG\'LIQNI SAQLASH'
        },
        'psychology': {
            'name': 'psychology',
            'primary': RGBColor(99, 102, 241),       # Indigo Violet
            'primary_dark': RGBColor(49, 46, 129),   # Deep Midnight Violet
            'card_bg': RGBColor(245, 243, 255),     # Soft Lavender/Lilac
            'border': RGBColor(221, 214, 254),      # Pastel Violet Border
            'title_color': RGBColor(30, 27, 75),    # Deep Slate Indigo
            'badge_bg': RGBColor(99, 102, 241),
            'badge_text': RGBColor(255, 255, 255),
            'tag': 'PSIXOLOGIYA VA NEYROBIOLOGIYA'
        },
        'philosophy': {
            'name': 'philosophy',
            'primary': RGBColor(217, 119, 6),       # Warm Amber / Bronze
            'primary_dark': RGBColor(120, 53, 15),  # Deep Terracotta
            'card_bg': RGBColor(255, 251, 235),     # Warm Parchment
            'border': RGBColor(253, 230, 138),      # Golden Amber Border
            'title_color': RGBColor(69, 26, 3),     # Rich Brown
            'badge_bg': RGBColor(217, 119, 6),
            'badge_text': RGBColor(255, 255, 255),
            'tag': 'FALSAFA VA DINSHUNOSLIK'
        },
        'education': {
            'name': 'education',
            'primary': RGBColor(79, 70, 229),       # Academic Indigo
            'primary_dark': RGBColor(49, 46, 129),  # Deep Indigo
            'card_bg': RGBColor(238, 242, 255),     # Soft Lavender
            'border': RGBColor(199, 210, 254),      # Indigo Border
            'title_color': RGBColor(30, 27, 75),    # Dark Indigo
            'badge_bg': RGBColor(79, 70, 229),
            'badge_text': RGBColor(255, 255, 255),
            'tag': 'TA\'LIM VA PEDAGOGIKA'
        },
        'reform': {
            'name': 'reform',
            'primary': RGBColor(29, 78, 216),       # Executive Royal Blue
            'primary_dark': RGBColor(30, 58, 138),  # Deep Navy
            'card_bg': RGBColor(239, 246, 255),     # Soft Blue
            'border': RGBColor(191, 219, 254),      # Light Blue Border
            'title_color': RGBColor(15, 23, 42),    # Slate Dark
            'badge_bg': RGBColor(29, 78, 216),
            'badge_text': RGBColor(255, 255, 255),
            'tag': 'DAVLAT VA IJTIMOIY ISLOHOTLAR'
        },
        'default': {
            'name': 'default',
            'primary': RGBColor(218, 165, 32),      # Golden Amber
            'primary_dark': RGBColor(15, 23, 42),   # Navy Slate
            'card_bg': RGBColor(250, 252, 255),     # Clean Crisp White-Blue
            'border': RGBColor(210, 220, 230),      # Neutral Silver Border
            'title_color': RGBColor(15, 23, 42),    # Deep Slate
            'badge_bg': RGBColor(218, 165, 32),
            'badge_text': RGBColor(255, 255, 255),
            'tag': 'AKADEMIK VA ILMIY TADQIQOT'
        }
    }

    TOPIC_ICONS = {
        'psychology': ["🧠", "⚡", "🎯", "🔬", "🛡️", "⚖️", "💡", "🔄"],
        'medical': ["🩺", "💊", "🏥", "🧬", "🛡️", "🔬", "📋", "✅"],
        'philosophy': ["🏛️", "📜", "💡", "⚖️", "📖", "🔍", "🕊️", "✨"],
        'education': ["🎓", "📚", "✍️", "🎯", "🏆", "💡", "📊", "✅"],
        'reform': ["🏢", "📈", "🌐", "🤝", "⚖️", "💼", "🚀", "🎯"],
        'nature': ["🌱", "🌿", "🌍", "🐾", "☀️", "💧", "🌳", "🍃"],
        'default': ["📌", "🎯", "💡", "📊", "⚡", "🛡️", "✅", "🚀"]
    }

    JUNK_LAYOUT_KEYWORDS = [
        "thank", "closing", "color", "palette", "typography", "font",
        "credit", "designed", "resource", "team", "contact", "q&a",
        "question", "blank", "end", "section", "divider", "chapter",
        "vertical", "caption", "two content", "comparison"
    ]

    DUMMY_TEXT_PATTERNS = [
        r"your footer here", r"\bfooter\b", r"\bdate\b", r"allppt",
        r"presentationgo", r"slidescarnival", r"slidesmania",
        r"designed with", r"mobile / email", r"\bcompany\b", r"\bname\b"
    ]

    DANGEROUS_TEMPLATE_KEYWORDS = [
        'gun', 'bullet', 'handgun', 'pistol', 'rifle', 'shotgun', 'weapon', 'ammo',
        'military', 'army', 'war', 'soldier', 'tank', 'grenade', 'missile', 'bomb',
        'crime', 'gang', 'police', 'handcuff', 'knife', 'sword', 'skull', 'death',
        'funeral', 'coffin', 'cemetery', 'tombstone', 'blood', 'kill', 'killer',
        'doctor', 'nurse', 'surgeon', 'patient', 'baby', 'kid', 'child', 'children',
        'cartoon', 'character', 'doll', 'wedding', 'party', 'birthday', 'cake', 'cute',
        'animal', 'dog', 'cat', 'pet', 'cooking', 'chef', 'food', 'pizza', 'sushi',
        'burger', 'wine', 'beer', 'team workers', 'television man', 'wanted', 'pirate',
        'isolated on sand'
    ]

    @staticmethod
    def _apply_anti_break(paragraph):
        """PowerPoint-da so'zlarni o'rtasidan bo'lib tashlashni (word break) to'xtatadi."""
        pPr = paragraph._p.get_or_add_pPr()
        pPr.set('latinLnBrk', '0')
        pPr.set('eaLnBrk', '0')
        pPr.set('hangingPunct', '1')

    DIAGRAM_QUERY_MAP = {
        'xulq': 'human brain nervous system neuron diagram',
        'psixolog': 'cognitive psychology human brain diagram',
        'fiziolog': 'human neuron synapse action potential diagram',
        'neyro': 'neuron synapse neural network diagram',
        'asab': 'central nervous system brain diagram',
        'miya': 'human brain lobes structure diagram',
        'ong': 'cognitive psychology brain mind diagram',
        'tabiat': 'ecosystem food web plant photosynthesis diagram',
        'biolog': 'biological cell structure diagram organelle',
        'tibbiyot': 'human internal organs anatomy diagram',
        'bemor': 'patient hospital care first aid diagram',
        'transport': 'medical patient transport ambulance equipment',
        'санпин': 'hygiene medical sterilization disinfection diagram',
        'инфекция': 'virus bacteria transmission infection diagram',
        'kvant': 'quantum computer circuit architecture qubit',
        'din': 'world religions symbols history',
        'zardusht': 'zoroastrianism faravahar ancient history',
        'buddaviy': 'buddhism dharmachakra stupa history',
        'falsafa': 'philosophy thinking wisdom symbol',
        'berdyaev': 'philosophy existentialism book history',
        'gazzoliy': 'islamic philosophy scholars history',
        'tafakkur': 'thinking cognitive process intellect diagram',
        'tarix': 'ancient history architecture manuscript',
        'islohot': 'economic growth social development chart',
        'savol': 'education learning knowledge concept diagram',
        'maktab': 'education school classroom learning concept'
    }

    @staticmethod
    def fetch_diagram_image(query: str, cache_dir: str = "output/diagram_cache") -> Optional[str]:
        """
        Wikimedia Commons API orqali autentik ilmiy diagramma/rasmlarni topadi va keshlaydi.
        """
        try:
            os.makedirs(cache_dir, exist_ok=True)
            q_lower = query.lower()

            # 1. Mavzuga mos inglizcha ilmiy kalit so'zlarni tanlash
            search_query = None
            for k, val in SlideBuilderService.DIAGRAM_QUERY_MAP.items():
                if k in q_lower:
                    search_query = val
                    break

            if not search_query:
                clean_q = re.sub(r'[^a-zA-Z0-9\s]', ' ', query)
                words = [w for w in clean_q.split() if len(w) > 3][:3]
                search_query = f"{' '.join(words)} diagram"

            # 2. Fayl nomi bo'yicha keshni tekshirish
            safe_slug = re.sub(r'[^a-z0-9]', '_', search_query.lower())[:35]
            cached_files = glob.glob(os.path.join(cache_dir, f"{safe_slug}*"))
            if cached_files and os.path.exists(cached_files[0]) and os.path.getsize(cached_files[0]) > 5000:
                return cached_files[0]

            # 3. Wikimedia Commons API orqali qidirish (Faqat File namespace - 6)
            url = (
                f"https://commons.wikimedia.org/w/api.php?action=query"
                f"&generator=search&gsrnamespace=6&gsrsearch={urllib.parse.quote(search_query)}"
                f"&gsrlimit=8&prop=imageinfo&iiprop=url|thumburl&iiurlwidth=900&format=json"
            )
            req = urllib.request.Request(url, headers={'User-Agent': 'SlaydHubBot/2.0 (slide_generator@academic.uz)'})
            with urllib.request.urlopen(req, timeout=12) as response:
                data = json.loads(response.read().decode('utf-8'))

            pages = data.get('query', {}).get('pages', {})
            best_img_url = None

            for pid, p in pages.items():
                t = p.get('title', '').lower()
                if any(bad in t for bad in ['.ogg', '.ogv', '.pdf', '.tif', 'flag', 'coat of arms', 'map', 'icon', 'logo']):
                    continue
                ii_list = p.get('imageinfo', [])
                if ii_list:
                    ii = ii_list[0]
                    t_url = ii.get('thumburl') or ii.get('url')
                    if t_url and (t_url.endswith('.png') or t_url.endswith('.jpg') or t_url.endswith('.jpeg') or 'thumb' in t_url):
                        best_img_url = t_url
                        if 'diagram' in t or 'scheme' in t or 'en' in t or 'structure' in t:
                            break

            if not best_img_url:
                return None

            # 4. Rasmni yuklab olish va saqlash
            ext = ".png" if (".png" in best_img_url.lower() or ".svg" in best_img_url.lower()) else ".jpg"
            dest_path = os.path.join(cache_dir, f"{safe_slug}{ext}")
            img_req = urllib.request.Request(best_img_url, headers={'User-Agent': 'SlaydHubBot/2.0'})
            with urllib.request.urlopen(img_req, timeout=15) as img_resp, open(dest_path, 'wb') as out_f:
                out_f.write(img_resp.read())

            im = Image.open(dest_path)
            im.verify()
            return dest_path
        except Exception:
            return None

    @staticmethod
    def get_theme_for_topic(topic: str) -> Dict[str, Any]:
        t = topic.lower()
        if any(k in t for k in ['xulq', 'psixolog', 'fiziolog', 'neyro', 'asab', 'ong', 'tafakkur', 'kognitiv']):
            return SlideBuilderService.THEMES['psychology']
        if any(k in t for k in ['tabiat', 'biolog', 'ekolog', 'o\'simlik', 'hayvon', 'yashil']):
            return SlideBuilderService.THEMES['nature']
        if any(k in t for k in ['bemor', 'tibbiyot', 'transport', 'санпин', 'инфекция', 'vrach', 'shifoxona']):
            return SlideBuilderService.THEMES['medical']
        if any(k in t for k in ['din', 'hurfikr', 'berdyaev', 'gazzoliy', 'zardusht', 'buddaviy', 'falsafa', 'tarix', 'mutafakkir']):
            return SlideBuilderService.THEMES['philosophy']
        if any(k in t for k in ['sinf', 'maktab', 'savol', 'talaba', 'dars', 'ta\'lim', 'pedagog']):
            return SlideBuilderService.THEMES['education']
        if any(k in t for k in ['islohot', 'ijtimoiy', 'iqtisod', 'davlat', 'boshqaruv', 'taraqqiyot']):
            return SlideBuilderService.THEMES['reform']
        return SlideBuilderService.THEMES['default']

    @staticmethod
    def find_best_template(topic: str) -> Optional[str]:
        t_lower = topic.lower()
        is_military_topic = any(m in t_lower for m in ['qurol', 'armiya', 'harbiy', 'urush', 'jang', 'military', 'war', 'weapon'])
        is_medical_doctor_topic = any(m in t_lower for m in ['vrach', 'shifoxona', 'doktor', 'bemor', 'hamshira', 'operatsiya'])

        candidates = []
        for tdir in TEMPLATES_DIRS:
            if tdir.exists():
                for root, _, files in os.walk(str(tdir)):
                    # Qolgan_slaydlar (FreePPT7 ning sifatsiz 2010 cliparti) ni butunlay o'tkazib yuboramiz
                    if 'qolgan_slaydlar' in root.lower():
                        continue
                    for f in files:
                        fl = f.lower()
                        if (fl.endswith(".pptx") or fl.endswith(".potx")) and not f.startswith("~$"):
                            if any(bad in fl for bad in ['icon collection', 'clipart', 'gantt chart', 'funnel diagram', 'venn diagram']):
                                continue
                            if not is_military_topic and any(bad in fl for bad in ['gun', 'bullet', 'handgun', 'pistol', 'rifle', 'weapon', 'military', 'army', 'war']):
                                continue
                            if not is_medical_doctor_topic and any(bad in fl for bad in ['doctor', 'nurse', 'surgeon', 'patient', 'baby']):
                                continue
                            if any(bad in fl for bad in ['cartoon', 'doll', 'wedding', 'party', 'birthday', 'cake', 'cute', 'animal', 'dog', 'cat', 'pet', 'cooking', 'chef', 'food', 'wine', 'beer', 'team workers', 'television man', 'wanted', 'pirate']):
                                continue
                            candidates.append(os.path.join(root, f))
        
        if not candidates:
            return None

        target_keywords = set()
        for root_key, kw_list in SlideBuilderService.SEMANTIC_ONTOLOGY.items():
            if root_key in t_lower:
                target_keywords.update(kw_list)

        matched = []
        for c in candidates:
            c_name = os.path.basename(c).lower()
            score = sum(1 for kw in target_keywords if kw in c_name)
            if score > 0:
                if 'PresentationGO' in c:
                    score += 5
                elif 'PowerPointSchool' in c:
                    score += 4
                elif 'SlidesMania' in c:
                    score += 4
                elif 'PPTTemplate' in c:
                    score += 3
                elif 'FreePPT7' in c:
                    score -= 3
                matched.append((score, c))

        matched.sort(key=lambda x: x[0], reverse=True)
        if matched:
            pool = [m[1] for m in matched[:25]]
        else:
            pool = [
                c for c in candidates
                if any(k in os.path.basename(c).lower() for k in ['education', 'academic', 'clean', 'simple', 'minimal', 'business', 'modern', 'blue', 'tech', 'presentation', 'corporate', 'science'])
            ] or candidates

        # Eng yuqori ball olgan 5 ta ichidan tanlash
        top_tier = pool[:5] if len(pool) >= 5 else pool
        random.shuffle(top_tier)

        for cand in top_tier + pool[5:]:
            try:
                prs = safe_load_presentation(cand)
                if len(prs.slide_layouts) >= 2:
                    return cand
            except Exception:
                continue

        return None

    @staticmethod
    def get_cover_and_content_layouts(prs: Presentation):
        cover_layout = None
        content_layouts = []

        for idx, layout in enumerate(prs.slide_layouts):
            lname = layout.name.lower()
            if any(junk in lname for junk in SlideBuilderService.JUNK_LAYOUT_KEYWORDS):
                continue
            if not cover_layout and any(k in lname for k in ["cover", "title slide", "main", "start"]):
                cover_layout = layout
                continue
            if any(k in lname for k in ["title and content", "content 1", "content", "body"]):
                content_layouts.append(layout)

        if not cover_layout:
            cover_layout = prs.slide_layouts[0]
        if not content_layouts:
            # Clean fallback - avoid vertical or caption
            valid_layouts = [l for l in prs.slide_layouts[1:] if not any(j in l.name.lower() for j in SlideBuilderService.JUNK_LAYOUT_KEYWORDS)]
            content_layouts = valid_layouts if valid_layouts else [prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]]

        return cover_layout, content_layouts

    @staticmethod
    def purge_slide_junk(slide):
        """
        Shablon layoutidan meros bo'lib o'tgan barcha keraksiz placeholder shakllarini
        (Click to edit master title, ruscha vertikal matnlar, insert ikonkalari)
        jismonan XML daraxtidan to'liq o'chirib tashlaydi.
        Faqat slide master foni saqlanib qoladi.
        """
        for shape in list(slide.shapes):
            try:
                sp = shape._element
                sp.getparent().remove(sp)
            except Exception:
                pass

    @staticmethod
    def purge_presentation_watermarks_and_junk(prs):
        """
        Shablon masterlari va layoutlaridagi barcha reklamalar, watermarklar,
        sayt nomlari (presentationgo, allppt, powerpointschool, by:, .com) va
        keraksiz footer/date/slide number shakllarini XML daraxtidan butunlay tozalaydi.
        """
        bad_words = [
            'presentationgo', 'allppt', 'powerpointschool', 'slidescarnival',
            'slidesmania', 'by:', '.com', 'designed with', 'your footer here',
            'freeppt7', 'ppttemplate', 'copyright'
        ]

        # 1. Slide Masters tozalash
        for master in prs.slide_masters:
            for shape in list(master.shapes):
                txt = ""
                if shape.has_text_frame:
                    txt = shape.text_frame.text.lower()
                elif shape.shape_type == 6:  # GROUP shape
                    try:
                        for sub in shape.shapes:
                            if hasattr(sub, 'text_frame'):
                                txt += " " + sub.text_frame.text.lower()
                            if hasattr(sub, 'shapes'):
                                for sub2 in sub.shapes:
                                    if hasattr(sub2, 'text_frame'):
                                        txt += " " + sub2.text_frame.text.lower()
                    except Exception:
                        pass

                is_ad = any(w in txt for w in bad_words)
                is_footer_junk = shape.name.lower().startswith(('date', 'footer', 'slide number')) or 'date' in txt or 'footer' in txt

                if is_ad or is_footer_junk:
                    try:
                        sp = shape._element
                        sp.getparent().remove(sp)
                    except Exception:
                        pass

        # 2. Slide Layouts tozalash
        for layout in prs.slide_layouts:
            for shape in list(layout.shapes):
                txt = ""
                if shape.has_text_frame:
                    txt = shape.text_frame.text.lower()
                is_ad = any(w in txt for w in bad_words)
                is_footer_junk = shape.name.lower().startswith(('date', 'footer', 'slide number'))
                if is_ad or is_footer_junk:
                    try:
                        sp = shape._element
                        sp.getparent().remove(sp)
                    except Exception:
                        pass

    @staticmethod
    def build_presentation(spec: Dict[str, Any], template_path: Optional[str] = None, output_path: Optional[str] = None) -> str:
        topic = spec.get("topic", "Taqdimot")
        main_title = spec.get("title", topic)
        slides_spec = spec.get("slides", [])
        if not slides_spec:
            slides_spec = [{"title": topic, "points": ["Asosiy ma'lumotlar"]}]

        theme = SlideBuilderService.get_theme_for_topic(topic)

        prs = None
        if not template_path or not os.path.exists(template_path):
            template_path = SlideBuilderService.find_best_template(topic)

        if template_path and os.path.exists(template_path):
            try:
                prs = safe_load_presentation(template_path)
                SlideBuilderService.purge_presentation_watermarks_and_junk(prs)
            except Exception:
                prs = None

        if prs is None:
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

        # Shablon masteridagi barcha latinLnBrk="1" larni 0 ga aylantirish (so'zlarning bo'linishiga qarshi)
        try:
            for master in prs.slide_masters:
                for elem in master._element.iter():
                    if elem.tag.endswith('lvl1pPr') or elem.tag.endswith('pPr'):
                        elem.set('latinLnBrk', '0')
                        elem.set('eaLnBrk', '0')
                        elem.set('hangingPunct', '1')
        except Exception:
            pass

        slide_w = prs.slide_width
        slide_h = prs.slide_height

        # Responsive o'lchamlar (16:9 widescreen va 4:3 formatga to'liq mos)
        margin_x = Inches(0.8) if slide_w.inches > 11.0 else Inches(0.5)
        content_w = slide_w - (margin_x * 2)

        cover_layout, content_layouts = SlideBuilderService.get_cover_and_content_layouts(prs)

        # Eski shablon slaydlarini tozalash
        while len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[0]

        # Slaydlarni generatsiya qilish
        for idx, s_data in enumerate(slides_spec):
            layout_type = s_data.get("layout_type", "content").lower()
            title_text = s_data.get("title", f"{idx+1}-qism")
            is_last_slide = (idx == len(slides_spec) - 1)
            is_conclusion = any(k in title_text.lower() for k in ["xulosa", "tavsiya", "yakun", "natija", "istiqbol"])

            # 1. MUQOVA (COVER HERO SLIDE)
            if idx == 0 or layout_type == "cover":
                slide = prs.slides.add_slide(cover_layout)
                SlideBuilderService.purge_slide_junk(slide)
                SlideBuilderService._render_cover_hero(slide, main_title, spec, theme, margin_x, content_w, slide_w, slide_h)

            # 2. BOSHQA KONTENT SLAYDLARI
            else:
                current_layout = content_layouts[0]
                slide = prs.slides.add_slide(current_layout)
                SlideBuilderService.purge_slide_junk(slide)

                # Sarlavhani joylashtirish (Shablonning o'z Title shakli yoki maxsus Header bilan)
                SlideBuilderService._render_slide_header(slide, title_text, theme, margin_x, content_w)

                top_start = Inches(1.85) if slide_h.inches > 7.0 else Inches(1.55)

                # A) QIYOSIY TAHLIL (COMPARISON / VS) + Markaziy VS aylanasi
                if not is_last_slide and not is_conclusion and (
                    layout_type == "comparison" or "vs" in title_text.lower() or "qiyos" in title_text.lower() or "farq" in title_text.lower()
                ):
                    SlideBuilderService._render_comparison(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

                # B) ILMIY DIAGRAMMA / RASMLI SLAYD (DIAGRAM + CARDS SPLIT)
                elif not is_last_slide and not is_conclusion and (
                    layout_type in ["diagram", "illustration", "image"] or (idx in [2, 6] and any(k in title_text.lower() for k in ["tuzilma", "mexanizm", "markaz", "anatomiya", "struktura"]))
                ):
                    diag_img = SlideBuilderService.fetch_diagram_image(f"{topic} {title_text}")
                    if diag_img:
                        SlideBuilderService._render_diagram_split(slide, s_data, theme, margin_x, top_start, content_w, slide_h, diag_img)
                    else:
                        SlideBuilderService._render_grid(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

                # C) BOSQICHLAR / JARAYON OQIM DIAGRAMMASI (STEPS / PROCESS FLOW + ARROWS)
                elif not is_last_slide and not is_conclusion and (
                    layout_type in ["steps", "process"] or "bosqich" in title_text.lower() or "tartib" in title_text.lower() or "algoritm" in title_text.lower()
                ):
                    SlideBuilderService._render_steps(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

                # D) METRIKA VA PROGRESS-BAR DIAGRAMMASI (METRICS / DASHBOARD)
                elif not is_last_slide and not is_conclusion and (
                    layout_type in ["metrics", "dashboard", "indicators"] or "metod" in title_text.lower() or "ko'rsatkich" in title_text.lower()
                ):
                    SlideBuilderService._render_metric_dashboard(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

                # E) SAVOL-JAVOB (Q&A / QUIZ)
                elif not is_last_slide and not is_conclusion and (
                    layout_type in ["qa", "quiz"] or "savol" in title_text.lower()
                ):
                    SlideBuilderService._render_qa(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

                # F) MUNDARIJA (AGENDA)
                elif layout_type == "agenda" or idx == 1 or "mundarija" in title_text.lower() or "reja" in title_text.lower():
                    SlideBuilderService._render_agenda(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

                # G) USTUNLI GRID KARTALAR (GRID / 3-PILLARS)
                elif not is_last_slide and not is_conclusion and (layout_type == "grid" or (idx % 2 == 1 and len(s_data.get("points", [])) == 3)):
                    SlideBuilderService._render_grid(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

                # H) STANDART AKADEMIK KARTALAR (CARDS) - Xulosa va asosiy slaydlar uchun eng toza format
                else:
                    SlideBuilderService._render_cards(slide, s_data, theme, margin_x, top_start, content_w, slide_h)

        if not output_path:
            clean_name = "".join(c for c in topic if c.isalnum() or c in (" ", "_", "-")).strip()[:50]
            output_path = os.path.join(OUTPUT_DIR, f"{clean_name}_Taqdimot.pptx")

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        prs.save(output_path)
        return output_path

    @staticmethod
    def _render_cover_hero(slide, main_title, spec, theme, margin_x, content_w, slide_w, slide_h):
        """
        Har qanday shablonda (hatto bo'sh yoki rasm bo'lsa ham) 100% o'qiluvchan,
        markazlashgan premium Hero Card muqovasi.
        """
        subtitle_text = spec.get("subtitle", "Akademik va amaliy tahliliy qo'llanma / 2026-yil")
        if not subtitle_text or "taqdimot" in subtitle_text.lower():
            subtitle_text = "Akademik va amaliy tahliliy qo'llanma / 2026-yil"

        # Markaziy premium Hero karta
        hero_w = min(content_w, Inches(10.5))
        hero_left = (slide_w - hero_w) / 2
        hero_h = Inches(3.8) if slide_h.inches > 7.0 else Inches(3.2)
        hero_top = (slide_h - hero_h) / 2

        # 1. Shaffof-oq / engil fonli zamonaviy karta
        hero_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, hero_left, hero_top, hero_w, hero_h)
        hero_card.fill.solid()
        hero_card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        hero_card.line.color.rgb = theme['primary']
        hero_card.line.width = Pt(2.0)

        # 2. Mavzu Kategoriya Teg-Nishoni
        tag_w = Inches(3.2)
        tag_h = Inches(0.4)
        tag_left = hero_left + (hero_w - tag_w) / 2
        tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tag_left, hero_top + Inches(0.25), tag_w, tag_h)
        tag.fill.solid()
        tag.fill.fore_color.rgb = theme['primary']
        tag.line.fill.background()
        tag.text_frame.text = theme['tag']
        for tp in tag.text_frame.paragraphs:
            SlideBuilderService._apply_anti_break(tp)
            tp.alignment = PP_ALIGN.CENTER
            tp.font.size = Pt(10)
            tp.font.bold = True
            tp.font.color.rgb = RGBColor(255, 255, 255)

        # 3. Asosiy Sarlavha (Keng text box va so'zlar bo'linishiga qarshi nozik o'lchamlar)
        t_box = slide.shapes.add_textbox(hero_left + Inches(0.15), hero_top + Inches(0.72), hero_w - Inches(0.3), Inches(1.85))
        tf = t_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.01)
        tf.margin_right = Inches(0.01)
        p1 = tf.paragraphs[0]
        SlideBuilderService._apply_anti_break(p1)
        p1.text = main_title
        p1.alignment = PP_ALIGN.CENTER
        p1.font.bold = True
        if len(main_title) > 65:
            p1.font.size = Pt(20)
        elif len(main_title) > 45:
            p1.font.size = Pt(22)
        elif len(main_title) > 30:
            p1.font.size = Pt(25)
        else:
            p1.font.size = Pt(28)
        p1.font.color.rgb = theme['title_color']

        # 4. Kichik Sarlavha (Subtitle)
        s_box = slide.shapes.add_textbox(hero_left + Inches(0.5), hero_top + hero_h - Inches(0.9), hero_w - Inches(1.0), Inches(0.7))
        sf = s_box.text_frame
        sf.word_wrap = True
        sf.margin_left = Inches(0.05)
        sf.margin_right = Inches(0.05)
        p2 = sf.paragraphs[0]
        SlideBuilderService._apply_anti_break(p2)
        p2.text = subtitle_text
        p2.alignment = PP_ALIGN.CENTER
        p2.font.size = Pt(13)
        p2.font.color.rgb = RGBColor(71, 85, 105)

    @staticmethod
    def _render_slide_header(slide, title_text, theme, margin_x, content_w):
        """
        Slayd sarlavhasini shablon fonidan qat'i nazar 100% o'qiluvchan qiluvchi
        zamonaviy Header formati. Shablonning o'z Title shakli bo'lsa undan foydalanadi,
        aks holda maxsus Header Ribbon chizadi.
        """
        # Duplicate so'zlarni (masalan '3 omil 3 omil') tozalash
        title_text = re.sub(r'\b(\w+(?:\s+\w+)?)\s+\1\b', r'\1', title_text.strip(), flags=re.IGNORECASE)

        title_shape = None
        for sh in list(slide.shapes):
            sh_l = sh.name.lower()
            if "title" in sh_l and sh.has_text_frame and not title_shape and sh.top < Inches(2.2):
                title_shape = sh
            elif ("content" in sh_l or "placeholder" in sh_l or "body" in sh_l or "text" in sh_l) and sh.has_text_frame:
                sh.text_frame.text = ""

        if title_shape:
            tf = title_shape.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0] if len(tf.paragraphs) > 0 else tf.add_paragraph()
            p.text = title_text
            SlideBuilderService._apply_anti_break(p)
            p.font.bold = True
            p.font.size = Pt(22)
            p.font.color.rgb = theme['title_color']
        else:
            # Aksent pill va maxsus text box
            pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, margin_x, Inches(0.55), Inches(0.12), Inches(0.75))
            pill.fill.solid()
            pill.fill.fore_color.rgb = theme['primary']
            pill.line.fill.background()

            tb = slide.shapes.add_textbox(margin_x + Inches(0.25), Inches(0.55), content_w - Inches(0.3), Inches(0.8))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.02)
            tf.margin_right = Inches(0.02)
            p = tf.paragraphs[0]
            p.text = title_text
            SlideBuilderService._apply_anti_break(p)
            p.font.bold = True
            p.font.size = Pt(22)
            p.font.color.rgb = theme['title_color']

    @staticmethod
    def _render_cards(slide, s_data, theme, margin_x, top_start, content_w, slide_h):
        """
        Standart akademik kontent: Har bir matn ostiga ortiqcha qutilar chizmasdan,
        toza, o'qiluvchan va zamonaviy tipografik ro'yxat shaklida chiqaradi.
        """
        points = s_data.get("points", [])
        num_pts = len(points)
        if num_pts == 0:
            return

        available_h = slide_h - top_start - Inches(0.6)
        item_h = available_h / num_pts

        for p_i, pt in enumerate(points):
            c_top = top_start + p_i * item_h

            tb = slide.shapes.add_textbox(margin_x, c_top, content_w, item_h - Inches(0.1))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.0)
            tf.margin_top = Inches(0.0)
            tf.margin_right = Inches(0.0)
            tf.margin_bottom = Inches(0.0)

            clean_pt = pt.lstrip("•-* \t\u2022\u25cf\ufffd")
            clean_pt = re.sub(r'^\d+[\.\)]\s*', '', clean_pt)

            if ":" in clean_pt:
                parts = clean_pt.split(":", 1)
                t_p = tf.paragraphs[0]
                SlideBuilderService._apply_anti_break(t_p)
                t_p.text = f"{p_i+1}. {parts[0].strip()}"
                t_p.font.size = Pt(17)
                t_p.font.bold = True
                t_p.font.color.rgb = theme['title_color']
                t_p.space_after = Pt(4)

                d_p = tf.add_paragraph()
                SlideBuilderService._apply_anti_break(d_p)
                d_p.text = parts[1].strip()
                d_p.font.size = Pt(13)
                d_p.font.color.rgb = RGBColor(51, 65, 85)
            else:
                t_p = tf.paragraphs[0]
                SlideBuilderService._apply_anti_break(t_p)
                t_p.text = f"• {clean_pt}"
                t_p.font.size = Pt(15)
                t_p.font.bold = True
                t_p.font.color.rgb = theme['title_color']

    @staticmethod
    def _render_comparison(slide, s_data, theme, margin_x, top_start, content_w, slide_h):
        """2 ustunli yonma-yon qiyosiy tahlil maketi + Markaziy 'VS' aylanasi."""
        col_w = (content_w - Inches(0.5)) / 2

        comp_data = s_data.get("comparison_data", {})
        left_title = comp_data.get("left_title")
        right_title = comp_data.get("right_title")
        left_pts = comp_data.get("left_points", [])
        right_pts = comp_data.get("right_points", [])

        if not left_title or not right_title:
            t_title = s_data.get("title", "")
            if " vs " in t_title.lower():
                parts = re.split(r'\s+vs\s+', t_title, flags=re.IGNORECASE)
                if len(parts) >= 2:
                    left_title = left_title or parts[0].strip()
                    right_title = right_title or parts[1].strip()
            elif " va " in t_title.lower():
                parts = re.split(r'\s+va\s+', t_title, flags=re.IGNORECASE)
                if len(parts) >= 2:
                    left_title = left_title or parts[0].strip()
                    right_title = right_title or parts[1].strip()

        if not left_pts or not right_pts:
            pts = s_data.get("points", [])
            mid = max(1, len(pts) // 2)
            left_pts = pts[:mid]
            right_pts = pts[mid:]

        left_title = left_title or "1-Yo'nalish Tahlili"
        right_title = right_title or "2-Yo'nalish Tahlili"

        max_pts = max(len(left_pts), len(right_pts), 1)
        target_h = Inches(1.3) + Inches(0.8) * max_pts
        col_h = min(slide_h - top_start - Inches(0.6), max(Inches(3.5), target_h))

        # 1-USTUN (CHAP)
        c1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, margin_x, top_start, col_w, col_h)
        c1.fill.solid()
        c1.fill.fore_color.rgb = theme['card_bg']
        c1.line.color.rgb = theme['primary']
        c1.line.width = Pt(1.5)

        h1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, margin_x + Inches(0.15), top_start + Inches(0.15), col_w - Inches(0.3), Inches(0.55))
        h1.fill.solid()
        h1.fill.fore_color.rgb = theme['primary']
        h1.line.fill.background()
        h1.text_frame.text = left_title
        for p in h1.text_frame.paragraphs:
            SlideBuilderService._apply_anti_break(p)
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(255, 255, 255)

        tb1 = slide.shapes.add_textbox(margin_x + Inches(0.2), top_start + Inches(0.8), col_w - Inches(0.4), col_h - Inches(0.95))
        tf1 = tb1.text_frame
        tf1.word_wrap = True
        tf1.margin_left = Inches(0.05)
        tf1.margin_right = Inches(0.05)
        for i, pt in enumerate(left_pts):
            p = tf1.paragraphs[0] if i == 0 else tf1.add_paragraph()
            clean_pt = pt.lstrip("•-* \t\u2022\u25cf\ufffd")
            p.text = f"• {clean_pt}"
            SlideBuilderService._apply_anti_break(p)
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(30, 41, 59)
            p.space_after = Pt(12)

        # 2-USTUN (O'NG)
        x2 = margin_x + col_w + Inches(0.5)
        c2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x2, top_start, col_w, col_h)
        c2.fill.solid()
        c2.fill.fore_color.rgb = RGBColor(255, 255, 255)
        c2.line.color.rgb = theme['border']
        c2.line.width = Pt(1.5)

        h2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x2 + Inches(0.15), top_start + Inches(0.15), col_w - Inches(0.3), Inches(0.55))
        h2.fill.solid()
        h2.fill.fore_color.rgb = theme['primary_dark']
        h2.line.fill.background()
        h2.text_frame.text = right_title
        for p in h2.text_frame.paragraphs:
            SlideBuilderService._apply_anti_break(p)
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(255, 255, 255)

        tb2 = slide.shapes.add_textbox(x2 + Inches(0.2), top_start + Inches(0.8), col_w - Inches(0.4), col_h - Inches(0.95))
        tf2 = tb2.text_frame
        tf2.word_wrap = True
        tf2.margin_left = Inches(0.05)
        tf2.margin_right = Inches(0.05)
        for i, pt in enumerate(right_pts):
            p = tf2.paragraphs[0] if i == 0 else tf2.add_paragraph()
            clean_pt = pt.lstrip("•-* \t\u2022\u25cf\ufffd")
            p.text = f"• {clean_pt}"
            SlideBuilderService._apply_anti_break(p)
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(30, 41, 59)
            p.space_after = Pt(12)

        # MARKAZIY DUMALOQ "VS" NISHONI
        vs_size = Inches(0.7)
        vs_x = margin_x + col_w + (Inches(0.5) - vs_size) / 2
        vs_y = top_start + (col_h - vs_size) / 2
        vs_badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, vs_x, vs_y, vs_size, vs_size)
        vs_badge.fill.solid()
        vs_badge.fill.fore_color.rgb = theme['primary']
        vs_badge.line.color.rgb = RGBColor(255, 255, 255)
        vs_badge.line.width = Pt(2.0)
        vs_badge.text_frame.text = "VS"
        for p in vs_badge.text_frame.paragraphs:
            SlideBuilderService._apply_anti_break(p)
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(255, 255, 255)

    @staticmethod
    def _render_diagram_split(slide, s_data, theme, margin_x, top_start, content_w, slide_h, diag_img: str):
        """
        Chap tomonda ilmiy diagramma / sxema (ortiqcha quti/ramkasiz, toza rasm),
        o'ng tomonda toza, keng matnli tahlil (har bir matn ostiga quti chizmasdan!).
        """
        gap = Inches(0.4)
        col_w = (content_w - gap) / 2
        col_h = slide_h - top_start - Inches(0.55)

        # 1. CHAP TOMON: DIAGRAMMA (Ortiqcha qutisiz, toza ilmiy rasm!)
        try:
            with Image.open(diag_img) as im:
                orig_w, orig_h = im.size

            avail_w = col_w
            avail_h = col_h
            ratio = min(avail_w / Inches(orig_w / 96.0), avail_h / Inches(orig_h / 96.0))
            draw_w = Inches(orig_w / 96.0) * ratio
            draw_h = Inches(orig_h / 96.0) * ratio

            pic_left = margin_x + (avail_w - draw_w) / 2
            pic_top = top_start + (avail_h - draw_h) / 2

            slide.shapes.add_picture(diag_img, pic_left, pic_top, width=draw_w, height=draw_h)
        except Exception:
            pass

        # 2. O'NG TOMON: TOZA MATNLI TAHLIL (Ortiqcha qutisiz, chiroyli tipografiya!)
        right_x = margin_x + col_w + gap
        points = s_data.get("points", [])
        if not points:
            points = ["Komponent 1: Strukturaviy xususiyati", "Komponent 2: Faoliyat mexanizmi", "Komponent 3: Integratsiya va natija"]

        num_pts = min(4, len(points))
        item_h = col_h / num_pts

        for i, pt in enumerate(points[:num_pts]):
            i_top = top_start + i * item_h

            tb = slide.shapes.add_textbox(right_x, i_top, col_w, item_h - Inches(0.1))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.0)
            tf.margin_top = Inches(0.0)
            tf.margin_right = Inches(0.0)
            tf.margin_bottom = Inches(0.0)

            clean_pt = pt.lstrip("•-* \t\u2022\u25cf\ufffd")
            clean_pt = re.sub(r'^\d+[\.\)]\s*', '', clean_pt)

            if ":" in clean_pt:
                parts = clean_pt.split(":", 1)
                p1 = tf.paragraphs[0]
                SlideBuilderService._apply_anti_break(p1)
                p1.text = f"{i+1}. {parts[0].strip()}"
                p1.font.bold = True
                p1.font.size = Pt(17)
                p1.font.color.rgb = theme['title_color']
                p1.space_after = Pt(4)

                p2 = tf.add_paragraph()
                SlideBuilderService._apply_anti_break(p2)
                p2.text = parts[1].strip()
                p2.font.size = Pt(13)
                p2.font.color.rgb = RGBColor(51, 65, 85)
            else:
                p1 = tf.paragraphs[0]
                SlideBuilderService._apply_anti_break(p1)
                p1.text = f"• {clean_pt}"
                p1.font.size = Pt(15)
                p1.font.bold = True
                p1.font.color.rgb = theme['title_color']

    @staticmethod
    def _render_steps(slide, s_data, theme, margin_x, top_start, content_w, slide_h):
        """3 ta yonma-yon jarayon / bosqich maketi + Bog'lovchi strelkalar va dumaloq tugunlar."""
        steps = s_data.get("steps_data", [])
        if not steps:
            pts = s_data.get("points", [])
            steps = [{"step": f"0{i+1}", "title": p.split(":")[0] if ":" in p else f"{i+1}-Bosqich", "desc": p.split(":")[1] if ":" in p else p} for i, p in enumerate(pts[:3])]

        num_steps = min(3, len(steps))
        gap = Inches(0.35)
        col_w = (content_w - gap * (num_steps - 1)) / num_steps

        max_desc_len = max([len(st.get('desc', '')) for st in steps[:num_steps]] or [80])
        if max_desc_len < 120:
            col_h = Inches(3.5)
        elif max_desc_len < 190:
            col_h = Inches(4.2)
        else:
            col_h = min(Inches(4.8), slide_h - top_start - Inches(0.6))

        # Ustunlar orasiga bog'lovchi strelkalar (RIGHT_ARROW)
        arrow_w = Inches(0.24)
        arrow_h = Inches(0.28)
        for i in range(num_steps - 1):
            arr_x = margin_x + (i + 1) * col_w + i * gap + (gap - arrow_w) / 2
            arr_y = top_start + (col_h - arrow_h) / 2
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, arr_x, arr_y, arrow_w, arrow_h)
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = theme['primary']
            arrow.line.fill.background()

        for i, st in enumerate(steps[:num_steps]):
            col_x = margin_x + i * (col_w + gap)

            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x, top_start, col_w, col_h)
            card.fill.solid()
            card.fill.fore_color.rgb = theme['card_bg']
            card.line.color.rgb = theme['primary'] if i == 0 else theme['border']
            card.line.width = Pt(1.5 if i == 0 else 1.2)

            # Dumaloq raqam tuguni (Circle Node)
            node_size = Inches(0.65)
            node = slide.shapes.add_shape(MSO_SHAPE.OVAL, col_x + (col_w - node_size) / 2, top_start + Inches(0.18), node_size, node_size)
            node.fill.solid()
            node.fill.fore_color.rgb = theme['primary']
            node.line.color.rgb = RGBColor(255, 255, 255)
            node.line.width = Pt(1.5)
            node.text_frame.text = f"{i+1}"
            for p in node.text_frame.paragraphs:
                SlideBuilderService._apply_anti_break(p)
                p.alignment = PP_ALIGN.CENTER
                p.font.bold = True
                p.font.size = Pt(14)
                p.font.color.rgb = RGBColor(255, 255, 255)

            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x + Inches(0.15), top_start + Inches(0.92), col_w - Inches(0.3), Inches(0.42))
            badge.fill.solid()
            badge.fill.fore_color.rgb = theme['primary']
            badge.line.fill.background()
            badge.text_frame.text = f"BOSQICH {st.get('step', f'0{i+1}')}"
            for bp in badge.text_frame.paragraphs:
                SlideBuilderService._apply_anti_break(bp)
                bp.alignment = PP_ALIGN.CENTER
                bp.font.bold = True
                bp.font.size = Pt(11)
                bp.font.color.rgb = RGBColor(255, 255, 255)

            tb = slide.shapes.add_textbox(col_x + Inches(0.15), top_start + Inches(1.42), col_w - Inches(0.3), col_h - Inches(1.55))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.05)
            tf.margin_right = Inches(0.05)

            p1 = tf.paragraphs[0]
            SlideBuilderService._apply_anti_break(p1)
            p1.text = st.get('title', f'{i+1}-Qism').strip()
            p1.font.bold = True
            p1.font.size = Pt(14)
            p1.font.color.rgb = theme['title_color']
            p1.space_after = Pt(6)

            p2 = tf.add_paragraph()
            SlideBuilderService._apply_anti_break(p2)
            p2.text = st.get('desc', '').strip()
            p2.font.size = Pt(12)
            p2.font.color.rgb = RGBColor(51, 65, 85)

    @staticmethod
    def _render_metric_dashboard(slide, s_data, theme, margin_x, top_start, content_w, slide_h):
        """
        Metrika, tahliliy ko'rsatkichlar va progress barlar bilan jihozlangan tahlil paneli.
        """
        metrics = s_data.get("metrics_data", [])
        if not metrics:
            pts = s_data.get("points", [])
            defaults = [
                {"label": "Neyron Faolligi", "value": "88%", "desc": "Qo'zg'alish va tormozlanish jarayonlarining muvozanati"},
                {"label": "Reflektor Aniqlik", "value": "94%", "desc": "Signallarning sinaptik o'tkazuvchanlik ko'rsatkichi"},
                {"label": "Kognitiv Barqarorlik", "value": "82%", "desc": "Adaptatsiya va stressga nisbatan fiziologik bardoshlilik"}
            ]
            metrics = []
            for i, p in enumerate(pts[:3]):
                clean_p = p.lstrip("•-* \t\u2022\u25cf\ufffd")
                parts = clean_p.split(":", 1) if ":" in clean_p else (clean_p, "Ko'rsatkich tahlili")
                pct = [92, 85, 78][i % 3]
                metrics.append({
                    "label": parts[0].strip()[:35],
                    "value": f"{pct}%",
                    "desc": parts[1].strip() if len(parts) > 1 else defaults[i % len(defaults)]["desc"]
                })
            if not metrics:
                metrics = defaults

        num_m = min(3, len(metrics))
        gap = Inches(0.28)
        col_w = (content_w - gap * (num_m - 1)) / num_m
        col_h = min(Inches(4.4), slide_h - top_start - Inches(0.6))

        theme_name = theme.get('name', 'default')
        icons = SlideBuilderService.TOPIC_ICONS.get(theme_name, SlideBuilderService.TOPIC_ICONS['default'])

        for i, m in enumerate(metrics[:num_m]):
            col_x = margin_x + i * (col_w + gap)

            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x, top_start, col_w, col_h)
            card.fill.solid()
            card.fill.fore_color.rgb = theme['card_bg']
            card.line.color.rgb = theme['primary'] if i == 0 else theme['border']
            card.line.width = Pt(1.5 if i == 0 else 1.2)

            # Icon badge
            icon = icons[i % len(icons)]
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x + Inches(0.2), top_start + Inches(0.2), Inches(0.7), Inches(0.45))
            badge.fill.solid()
            badge.fill.fore_color.rgb = theme['primary']
            badge.line.fill.background()
            badge.text_frame.text = icon
            for p in badge.text_frame.paragraphs:
                SlideBuilderService._apply_anti_break(p)
                p.alignment = PP_ALIGN.CENTER
                p.font.size = Pt(14)

            # Katta Metrika qiymati (masalan: 94% yoki 88%)
            val_tb = slide.shapes.add_textbox(col_x + Inches(0.18), top_start + Inches(0.75), col_w - Inches(0.36), Inches(0.9))
            v_tf = val_tb.text_frame
            v_tf.word_wrap = True
            v_p = v_tf.paragraphs[0]
            SlideBuilderService._apply_anti_break(v_p)
            v_p.text = str(m.get("value", "85%"))
            v_p.font.bold = True
            v_p.font.size = Pt(38)
            v_p.font.color.rgb = theme['primary']

            # Label
            lbl_tb = slide.shapes.add_textbox(col_x + Inches(0.18), top_start + Inches(1.75), col_w - Inches(0.36), Inches(0.7))
            l_tf = lbl_tb.text_frame
            l_tf.word_wrap = True
            l_p = l_tf.paragraphs[0]
            SlideBuilderService._apply_anti_break(l_p)
            l_p.text = m.get("label", "Ko'rsatkich")
            l_p.font.bold = True
            l_p.font.size = Pt(15)
            l_p.font.color.rgb = theme['title_color']

            # Progress bar track & fill
            pb_w = col_w - Inches(0.4)
            pb_h = Inches(0.18)
            pb_left = col_x + Inches(0.2)
            pb_top = top_start + Inches(2.55)

            track = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, pb_left, pb_top, pb_w, pb_h)
            track.fill.solid()
            track.fill.fore_color.rgb = RGBColor(226, 232, 240)
            track.line.fill.background()

            # Fill bar
            val_str = str(m.get("value", "80%")).replace("%", "").strip()
            try:
                pct_float = float(val_str) / 100.0
                pct_float = max(0.1, min(1.0, pct_float))
            except Exception:
                pct_float = 0.85

            fill_w = pb_w * pct_float
            fill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, pb_left, pb_top, fill_w, pb_h)
            fill.fill.solid()
            fill.fill.fore_color.rgb = theme['primary']
            fill.line.fill.background()

            # Description
            desc_tb = slide.shapes.add_textbox(col_x + Inches(0.18), top_start + Inches(2.85), col_w - Inches(0.36), col_h - Inches(3.0))
            d_tf = desc_tb.text_frame
            d_tf.word_wrap = True
            d_p = d_tf.paragraphs[0]
            SlideBuilderService._apply_anti_break(d_p)
            d_p.text = m.get("desc", "")
            d_p.font.size = Pt(12)
            d_p.font.color.rgb = RGBColor(51, 65, 85)

    @staticmethod
    def _render_grid(slide, s_data, theme, margin_x, top_start, content_w, slide_h):
        """3 yoki 4 ustunli vertikal kartalar to'plami (Pillars / Components / Grid) + Mavzu ikonkasi."""
        points = s_data.get("points", [])
        num_cols = min(3, len(points)) if len(points) >= 3 else len(points)
        if num_cols == 0:
            return

        gap = Inches(0.25)
        col_w = (content_w - gap * (num_cols - 1)) / num_cols

        max_chars = max([len(pt) for pt in points[:num_cols]] or [60])
        if max_chars > 130:
            col_h = Inches(3.2)
        elif max_chars > 70:
            col_h = Inches(2.6)
        else:
            col_h = Inches(2.2)

        title_l = s_data.get("title", "").lower()
        tag_prefix = "KOMPONENT" if "komponent" in title_l else "OMIL" if "omil" in title_l else "TUZILMA" if "tuzilma" in title_l else "ASPEKT"
        theme_name = theme.get('name', 'default')
        icons = SlideBuilderService.TOPIC_ICONS.get(theme_name, SlideBuilderService.TOPIC_ICONS['default'])

        for i, pt in enumerate(points[:num_cols]):
            col_x = margin_x + i * (col_w + gap)

            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x, top_start, col_w, col_h)
            card.fill.solid()
            card.fill.fore_color.rgb = theme['card_bg']
            card.line.color.rgb = theme['border']
            card.line.width = Pt(1.3)

            # Top Accent Badge with Icon
            icon = icons[i % len(icons)]
            acc = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col_x + Inches(0.2), top_start + Inches(0.2), Inches(1.5), Inches(0.35))
            acc.fill.solid()
            acc.fill.fore_color.rgb = theme['primary']
            acc.line.fill.background()
            acc.text_frame.text = f"{icon} {tag_prefix} 0{i+1}"
            for p in acc.text_frame.paragraphs:
                SlideBuilderService._apply_anti_break(p)
                p.alignment = PP_ALIGN.CENTER
                p.font.bold = True
                p.font.size = Pt(9)
                p.font.color.rgb = RGBColor(255, 255, 255)

            tb = slide.shapes.add_textbox(col_x + Inches(0.18), top_start + Inches(0.68), col_w - Inches(0.36), col_h - Inches(0.75))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.04)
            tf.margin_right = Inches(0.04)

            clean_pt = pt.lstrip("•-* \t\u2022\u25cf\ufffd")
            clean_pt = re.sub(r'^\d+[\.\)]\s*', '', clean_pt)

            if ":" in clean_pt:
                parts = clean_pt.split(":", 1)
                p1 = tf.paragraphs[0]
                SlideBuilderService._apply_anti_break(p1)
                p1.text = parts[0].strip()
                p1.font.bold = True
                p1.font.size = Pt(15)
                p1.font.color.rgb = theme['title_color']
                p1.space_after = Pt(6)

                p2 = tf.add_paragraph()
                SlideBuilderService._apply_anti_break(p2)
                p2.text = parts[1].strip()
                p2.font.size = Pt(12)
                p2.font.color.rgb = RGBColor(51, 65, 85)
            else:
                p1 = tf.paragraphs[0]
                SlideBuilderService._apply_anti_break(p1)
                p1.text = clean_pt
                p1.font.bold = True
                p1.font.size = Pt(14)
                p1.font.color.rgb = theme['title_color']

    @staticmethod
    def _render_qa(slide, s_data, theme, margin_x, top_start, content_w, slide_h):
        """Savol-Javob (Q&A) ta'limiy maketi (mutanosib ixcham bloklar)."""
        qa_items = s_data.get("qa_data", [])
        if not qa_items:
            pts = s_data.get("points", [])
            qa_items = [{"q": p.split(":")[0] if ":" in p else f"Savol {i+1}", "a": p.split(":")[1] if ":" in p else p} for i, p in enumerate(pts[:2])]

        num_items = min(2, len(qa_items))
        if num_items == 0:
            return

        card_h = Inches(1.8)
        gap = Inches(0.35)

        for i, qa in enumerate(qa_items[:num_items]):
            c_top = top_start + i * (card_h + gap)

            # Yagona ixcham va mukammal blok
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, margin_x, c_top, content_w, card_h)
            card.fill.solid()
            card.fill.fore_color.rgb = theme['card_bg']
            card.line.color.rgb = theme['border']
            card.line.width = Pt(1.2)

            # Chap aksent chiziq (Accent pill)
            accent = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, margin_x + Inches(0.12), c_top + Inches(0.15), Inches(0.08), card_h - Inches(0.3))
            accent.fill.solid()
            accent.fill.fore_color.rgb = theme['primary']
            accent.line.fill.background()

            # Matn
            tb = slide.shapes.add_textbox(margin_x + Inches(0.35), c_top + Inches(0.15), content_w - Inches(0.55), card_h - Inches(0.3))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.0)
            tf.margin_top = Inches(0.0)

            p1 = tf.paragraphs[0]
            SlideBuilderService._apply_anti_break(p1)
            p1.text = f"SAVOL 0{i+1}: {qa.get('q', '').strip()}"
            p1.font.bold = True
            p1.font.size = Pt(15)
            p1.font.color.rgb = theme['primary']
            p1.space_after = Pt(6)

            p2 = tf.add_paragraph()
            SlideBuilderService._apply_anti_break(p2)
            p2.text = f"Tahliliy Javob: {qa.get('a', '').strip()}"
            p2.font.size = Pt(13)
            p2.font.color.rgb = RGBColor(51, 65, 85)

    @staticmethod
    def _render_agenda(slide, s_data, theme, margin_x, top_start, content_w, slide_h):
        """Mundarija (Ixcham va nafis gorizontal qatorlar shaklida)."""
        pts = s_data.get("points", [])
        num_pts = min(4, len(pts))
        if num_pts == 0:
            return

        row_h = Inches(0.85)
        gap = Inches(0.22)
        total_w = min(content_w, Inches(10.5))

        for i, pt in enumerate(pts[:num_pts]):
            c_top = top_start + i * (row_h + gap)

            # Shaffof-oq / engil kartochka (qalin emas, ixcham va nafis)
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, margin_x, c_top, total_w, row_h)
            card.fill.solid()
            card.fill.fore_color.rgb = theme['card_bg']
            card.line.color.rgb = theme['border']
            card.line.width = Pt(1.0)

            # Ixcham kvadrat nishon (Number badge)
            badge_size = Inches(0.55)
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, margin_x + Inches(0.18), c_top + (row_h - badge_size) / 2, badge_size, badge_size)
            badge.fill.solid()
            badge.fill.fore_color.rgb = theme['primary']
            badge.line.fill.background()
            badge.text_frame.text = f"0{i+1}"
            for bp in badge.text_frame.paragraphs:
                SlideBuilderService._apply_anti_break(bp)
                bp.alignment = PP_ALIGN.CENTER
                bp.font.size = Pt(13)
                bp.font.bold = True
                bp.font.color.rgb = RGBColor(255, 255, 255)

            # Matn
            tb = slide.shapes.add_textbox(margin_x + Inches(0.9), c_top + Inches(0.12), total_w - Inches(1.1), row_h - Inches(0.24))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.0)
            tf.margin_top = Inches(0.0)
            tf.margin_right = Inches(0.0)
            tf.margin_bottom = Inches(0.0)
            p = tf.paragraphs[0]
            SlideBuilderService._apply_anti_break(p)
            clean_pt = pt.lstrip("•-* \t\u2022\u25cf\ufffd")
            clean_pt = re.sub(r'^\d+[\.\)]\s*', '', clean_pt)
            p.text = clean_pt
            p.font.size = Pt(15)
            p.font.bold = True
            p.font.color.rgb = theme['title_color']
