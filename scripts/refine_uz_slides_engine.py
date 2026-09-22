# -*- coding: utf-8 -*-
"""
scripts/refine_uz_slides_engine.py
-----------------------------------
01_Tayyor_Sara_Slaydlar_UZ papkasidagi taqdimotlarni to'liq audit qilib, barcha kamchiliklarni
tuzatuvchi, ruscha/kirillcha qoldiqlarni professional o'zbek tiliga o'giruvchi,
suvbelgi va reklama mualliflarini tozalovchi, nomlarni standartlashtiruvchi master skript.
"""

import os
import glob
import re
import sys
import json
import logging
from typing import List, Dict, Any, Tuple
from pptx import Presentation
from pptx.util import Pt, Inches

sys.stdout.reconfigure(encoding='utf-8')
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger("RefineEngine")

from backend.core.gemini_translator import GeminiTranslator, sanitize_text
from backend.core.pptx_processor import PPTXProcessor, safe_load_presentation
from backend.core.image_translator import ImageTranslator
from backend.core.transliteration import ensure_script

CYRILLIC_PATTERN = re.compile(r'[\u0400-\u04FF]')
LATIN_RU_WORDS = [
    r'\bchto\b', r'\bkak\b', r'\besli\b', r'\bkogda\b', r'\bpotomu\b', r'\btak\b',
    r'\bdlya\b', r'\bcherez\b', r'\bposle\b', r'\bmejdu\b', r'\bvsego\b', r'\bkotory\w*',
    r'\buchenik\w*', r'\bklass\b', r'\burok\b', r'\bzadach\w*', r'\bzadanie\b',
    r'\btema\b', r'\bopredelenie\b', r'\bponyati\w*', r'\bsvoystv\w*', r'\bsilы\b',
    r'\bvremya\b', r'\bveshestv\w*', r'\bprotsess\w*', r'\brabota\b', r'\bistochnik\w*',
    r'\bavtor\b', r'\buchitel\b', r'\bpovtorenie\b', r'\bprimer\w*', r'\bvyvod\w*'
]
LATIN_RU_REGEX = re.compile('|'.join(LATIN_RU_WORDS), re.IGNORECASE)

WATERMARK_KEYWORDS = [
    r"antonenkova", r"allppt", r"slidesgo", r"presentationgo", r"shkola", 
    r"o['’`]?qituvchi", r"учитель", r"maktab", r"pedsovet", r"infourok", r"myshared",
    r"prezentacii", r"videouroki", r"viki\.rdf\.ru", r"uroki\.net", r"uchportal",
    r"kopilkaurokov", r"mbou\s*sosh", r"sovg['’`]?a\s*uchun", r"скачано\s*с",
    r"муниципальное\s*казенное", r"uchitel\s+kononov", r"fominova\s+elena"
]
WATERMARK_REGEX = re.compile('|'.join(WATERMARK_KEYWORDS), re.IGNORECASE)

FILENAME_RENAMES = {
    "Fizika - Sila Ampera.pptx": "Fizika - Amper kuchi va uning amaliy qo'llanilishi.pptx",
    "Fizika - Zarozhdenie f.pptx": "Fizika - Elementar zarralar fizikasining shakllanishi.pptx",
    "Fizika - УРОК ФИЗИКИ В 10 КЛАССЕ.pptx": "Fizika - Elektr maydoni ishi va potensiali (10-sinf).pptx",
    "Fizika - УРОК ФИЗИКИ В 9 КЛАССЕ.pptx": "Fizika - Moddiy nuqta va sanoq sistemasi (9-sinf).pptx",
    "Geografiya - klass atmosfernoe davlenie.pptx": "Geografiya - Atmosfera bosimi va barometrlar (7-sinf).pptx"
}

def is_text_needing_translation(text: str) -> bool:
    """Tekshiradi: Ushbu matn ruscha (kirill yoki lotin transkripsiyasida) ekanmi?"""
    if not text or len(text.strip()) < 2:
        return False
    # 1. Kirill harflari bormi?
    cyr = len(CYRILLIC_PATTERN.findall(text))
    if cyr >= 4:
        return True
    # 2. Lotin yozuvidagi ruscha so'zlar bormi?
    ru_words = LATIN_RU_REGEX.findall(text)
    if len(ru_words) >= 2 or (len(ru_words) >= 1 and len(text.split()) <= 4):
        return True
    return False

def clean_watermarks_and_authors(prs: Presentation) -> int:
    """Slaydlardagi maktab, o'qituvchi, sayt va reklama qoldiqlarini butunlay tozalaydi."""
    cleaned_count = 0
    for s_idx, slide in enumerate(prs.slides):
        shapes_to_remove = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if WATERMARK_REGEX.search(txt):
                    # Agar matn faqat muallif / reklama / maktab ma'lumotlaridan iborat bo'lsa
                    if len(txt) < 180 or "uchitel" in txt.lower() or "uchportal" in txt.lower() or "mbou" in txt.lower():
                        shapes_to_remove.append(shape)
                        cleaned_count += 1
                    else:
                        # Agar ichida muhim dars matni bo'lsa, faqat suvbelgi qatorlarini o'chiramiz
                        new_paras = []
                        for p in shape.text_frame.paragraphs:
                            if not WATERMARK_REGEX.search(p.text):
                                new_paras.append(p.text)
                        shape.text_frame.text = "\n".join(new_paras)
                        cleaned_count += 1
        for sh in shapes_to_remove:
            try:
                sp = sh._element
                sp.getparent().remove(sp)
            except Exception:
                pass
    return cleaned_count

def harmonize_slide_typography(prs: Presentation):
    """Barcha slaydlardagi shrift o'lchamlari va joylashuvini bir xillashtiradi."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                tf = shape.text_frame
                tf.word_wrap = True
                
                # Sarlavha yoki asosiy matn ekanligini aniqlash
                paras = [p for p in tf.paragraphs if p.text.strip()]
                if not paras:
                    continue
                
                is_title = (shape.top or 0) < Inches(1.8) and len(paras) <= 2 and len(tf.text) < 120
                
                # Har bir xatboshidagi font size larni muvozanatlash
                for p in paras:
                    for r in p.runs:
                        if is_title:
                            if not r.font.size or r.font.size.pt < 20:
                                r.font.size = Pt(28)
                            elif r.font.size.pt > 44:
                                r.font.size = Pt(36)
                        else:
                            if not r.font.size:
                                r.font.size = Pt(16)
                            elif r.font.size.pt > 28:
                                r.font.size = Pt(22)
                            elif r.font.size.pt < 12:
                                r.font.size = Pt(14)
                                
                        # Default shriftni toza va zamonaviy saqlash
                        if not r.font.name or r.font.name.lower() in ["times new roman", "arial cyr"]:
                            r.font.name = "Calibri"

def process_single_presentation(
    file_path: str, 
    translator: GeminiTranslator,
    img_translator: ImageTranslator
) -> Dict[str, Any]:
    """Bitta taqdimotni to'liq audit qilib, kamchiliklarni bartaraf etadi."""
    fname = os.path.basename(file_path)
    logger.info(f"Tekshirilmoqda: {fname}")
    
    prs = safe_load_presentation(file_path)
    
    # 1. Suvbelgilarni tozalash
    wm_cleaned = clean_watermarks_and_authors(prs)
    
    # 2. Tarjima talab qiluvchi barcha matnlarni yig'ish
    items_to_translate = []
    item_map = {}
    
    for s_idx, slide in enumerate(prs.slides, start=1):
        for sh_idx, shape in enumerate(slide.shapes):
            if shape.has_text_frame:
                for p_idx, p in enumerate(shape.text_frame.paragraphs):
                    for r_idx, r in enumerate(p.runs):
                        text = r.text
                        if is_text_needing_translation(text):
                            item_id = f"s{s_idx}_sh{sh_idx}_p{p_idx}_r{r_idx}"
                            items_to_translate.append({
                                "id": item_id,
                                "text": text,
                                "original_text": text
                            })
                            item_map[item_id] = r
            if shape.has_table:
                for row_idx, row in enumerate(shape.table.rows):
                    for col_idx, cell in enumerate(row.cells):
                        text = cell.text
                        if is_text_needing_translation(text):
                            item_id = f"s{s_idx}_tbl_{row_idx}_{col_idx}"
                            items_to_translate.append({
                                "id": item_id,
                                "text": text,
                                "original_text": text
                            })
                            item_map[item_id] = cell
                            
    translated_count = 0
    if items_to_translate:
        logger.info(f"  -> {fname}: {len(items_to_translate)} ta ruscha matn bloki topildi. Tarjima qilinmoqda...")
        # Gemini orqali professional o'zbek tiliga o'girish
        tr_results = translator.translate_items_batch(items_to_translate, target_script="latin", batch_size=30)
        for res in tr_results:
            rid = res["id"]
            tr_text = res.get("translated_text", "")
            if tr_text and rid in item_map:
                obj = item_map[rid]
                if hasattr(obj, 'text'):
                    obj.text = tr_text
                    translated_count += 1
                    
    # 3. Tipografiya va shriftlarni muvozanatlash
    harmonize_slide_typography(prs)
    
    # 4. Agar fayl nomi ruscha bo'lsa, yangilash
    new_fname = FILENAME_RENAMES.get(fname, fname)
    new_fpath = os.path.join(os.path.dirname(file_path), new_fname)
    
    # 5. Saqlash
    prs.save(new_fpath)
    if new_fname != fname and os.path.exists(file_path):
        os.remove(file_path)
        logger.info(f"  -> Fayl nomi o'zbekchalashtirildi: {fname} => {new_fname}")
        
    return {
        "original_name": fname,
        "final_name": new_fname,
        "translated_items": translated_count,
        "watermarks_cleaned": wm_cleaned
    }

if __name__ == "__main__":
    folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"
    files = sorted(glob.glob(os.path.join(folder, "*.pptx")))
    print(f"Jami {len(files)} ta taqdimotni tekshirish va kamchiliklarni to'g'rilash boshlandi...")
    
    translator = GeminiTranslator()
    img_translator = ImageTranslator()
    
    summary = []
    for f in files:
        res = process_single_presentation(f, translator, img_translator)
        summary.append(res)
        
    print("\n" + "="*50)
    print("AUDIT VA TAHRIRLASH YAKUNLANDI")
    print(f"Jami qayta ishlangan taqdimotlar: {len(summary)}")
    print(f"Tarjima qilingan ruscha matn bloklari: {sum(s['translated_items'] for s in summary)}")
    print(f"Tozalangan suvbelgilar / muallif reklamalari: {sum(s['watermarks_cleaned'] for s in summary)}")
    renamed = [s for s in summary if s['original_name'] != s['final_name']]
    print(f"Nomlari o'zbekchaga o'zgartirilgan fayllar: {len(renamed)}")
    for r in renamed:
        print(f"  - {r['original_name']} -> {r['final_name']}")
    print("="*50)
