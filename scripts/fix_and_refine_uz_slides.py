# -*- coding: utf-8 -*-
"""
scripts/fix_and_refine_uz_slides.py
------------------------------------
01_Tayyor_Sara_Slaydlar_UZ papkasidagi barcha taqdimotlarni tahrirlash:
1. Suvbelgilar, o'qituvchi/maktab ma'lumotlari va reklama havolalarini tozalash
2. Tipografiya, shrift o'lchamlari va matn bloklarini tekislash (anti-overflow)
3. Ruscha/kirillcha qolgan slayd va matnlarni to'liq professional o'zbek tiliga o'girish
4. Barcha fayl nomlarini toza va tushunarli o'zbekcha nomlarga standartlashtirish
"""

import os
import glob
import re
import sys
import json
import time
import logging
import shutil
from typing import List, Dict, Any, Tuple
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger("SlideRefiner")

from backend.core.gemini_translator import GeminiTranslator, sanitize_text
from backend.core.pptx_processor import PPTXProcessor, safe_load_presentation

CYRILLIC_PATTERN = re.compile(r'[\u0400-\u04FF]')
LATIN_RU_WORDS = [
    r'\bchto\b', r'\bkak\b', r'\besli\b', r'\bkogda\b', r'\bgde\b', r'\bpochemu\b', r'\bpotomu\b', 
    r'\bdlya\b', r'\bposle\b', r'\bcherez\b', r'\bmejdu\b', r'\bpered\b', r'\bokolo\b', r'\bpri\b', 
    r'\bbez\b', r'\biz\b', r'\bot\b', r'\bdo\b', r'\bpod\b', r'\bnad\b', r'\bvsego\b',
    r'\bkotory\w*', r'\buchenik\w*', r'\bklass\w*', r'\burok\w*', r'\bzadach\w*', r'\bzadanie\w*',
    r'\btema\b', r'\bopredelenie\w*', r'\bponyati\w*', r'\bsvoystv\w*', r'\byavleni\w*', r'\bvremya\b',
    r'\bveshestv\w*', r'\bprotsess\w*', r'\brabot\w*', r'\bistochnik\w*', r'\bavtor\b', r'\buchitel\w*',
    r'\bpovtorenie\b', r'\bprimer\w*', r'\bvyvod\w*', r'\bzaryad\w*', r'\bnapryajen\w*', r'\bsoprotivlen\w*',
    r'\btraektor\w*', r'\bopыt\w*', r'\bshkol\w*', r'\bpodgotov\w*', r'\bvopros\w*', r'\botvet\w*',
    r'\bresheni\w*', r'\bdann\w*', r'\btsel\w*', r'\bnazvanie\w*', r'\bvid\w*', r'\bstrukturn\w*',
    r'\bznacheni\w*', r'\bprimeneni\w*', r'\bedinits\w*', r'\bizmereni\w*', r'\bosnovn\w*', r'\bpar\w*',
    r'\bparoobrazovan\w*', r'\bisparen\w*', r'\bkondensats\w*', r'\bkipen\w*', r'\bdavlen\w*',
    r'\batmosfern\w*', r'\bdeystv\w*', r'\bnapravlen\w*', r'\bstrelk\w*', r'\bmagnitn\w*', r'\bmagnitnogo\b',
    r'\bpolya\b', r'\bmagnitnuyu\b', r'\bodnorodn\w*', r'\bvektor\w*', r'\bindukts\w*', r'\bprovodnik\w*',
    r'\bdvijen\w*', r'\bskorost\w*', r'\buskoren\w*', r'\bmassa\w*', r'\bplotnost\w*', r'\benergi\w*',
    r'\btemperat\w*', r'\bteplot\w*', r'\bsostavlyat\b', r'\bsostavlyaet\b', r'\bnaxodit\b', r'\bnaydite\b',
    r'\bopredelit\b', r'\bvыchislit\b', r'\breshite\b', r'\botvette\b', r'\bukajite\b', r'\bvыberite\b',
    r'\bposmotrite\b', r'\bvpervыe\b', r'\bvesomost\b', r'\bvozduxa\b', r'\bprivela\b', r'\blyudey\b',
    r'\bgodu\b', r'\bpersonal\w*', r'\bmotivat\w*', r'\bpobujdeni\w*', r'\bstimul\w*', r'\borganizats\w*',
    r'\bupravleni\w*', r'\brukovodstv\w*', r'\bnizol\w*', r'\bkonflikt\w*', r'\bstress\w*', r'\beffektivn\w*',
    r'\bchitayte\b', r'\bzapomnite\b', r'\bpravilo\b', r'\buprajnenie\w*'
]
LATIN_RU_REGEX = re.compile('|'.join(LATIN_RU_WORDS), re.IGNORECASE)

UZ_WORDS = [
    r'\bva\b', r'\bbu\b', r'\bbilan\b', r'\buchun\b', r'\bham\b', r'\bbo\'lgan\b', r'\bbo\'lib\b',
    r'\borqali\b', r'\bhaqida\b', r'\bto\'g\'risida\b', r'\bmavzu\b', r'\bdars\b', r'\bsinf\b',
    r'\bqonun\b', r'\bqonuni\b', r'\bamalga\b', r'\boshirish\b', r'\bmisol\b', r'\bmasala\b',
    r'\byechish\b', r'\bjavob\b', r'\bsavol\b', r'\bjadval\b', r'\bchizma\b', r'\bo\'rganish\b',
    r'\btushuncha\b', r'\bxususiyat\b', r'\bturlari\b', r'\btuzilishi\b', r'\btarixi\b', r'\busullari\b',
    r'\bkuch\b', r'\bbosim\b', r'\btezlik\b', r'\bharakat\b', r'\bmassa\b', r'\bzichlik\b',
    r'\bharorat\b', r'\bissiqlik\b', r'\belektr\b', r'\bmagnit\b', r'\bzaryad\b', r'\bkuchlanish\b',
    r'\btok\b', r'\bqarshilik\b', r'\bguruh\b', r'\bbosqich\b', r'\bnatija\b', r'\byo\'nalish\b'
]
UZ_REGEX = re.compile('|'.join(UZ_WORDS), re.IGNORECASE)

FILENAME_MAP = {
    "Fizika - Sila Ampera.pptx": "Fizika - Amper kuchi va uning amaliy qo'llanilishi (11-sinf).pptx",
    "Fizika - Zarozhdenie f.pptx": "Fizika - Elementar zarralar fizikasining paydo bo'lishi (11-sinf).pptx",
    "Fizika - УРОК ФИЗИКИ В 10 КЛАССЕ.pptx": "Fizika - Elektr maydoni ishi va potensiali (10-sinf).pptx",
    "Fizika - УРОК ФИЗИКИ В 9 КЛАССЕ.pptx": "Fizika - Moddiy nuqta va sanoq sistemasi (9-sinf).pptx",
    "Geografiya - klass atmosfernoe davlenie.pptx": "Geografiya - Atmosfera bosimi va barometrlar (7-sinf).pptx",
    "Iqtisodiyot - Taqdimot pt.pptx": "Iqtisodiyot - Korxona xodimlari va ish haqi tizimi.pptx",
    "Kimyo - 8 sinf Oksidlar.pptx": "Kimyo - Oksidlar (8-sinf).pptx"
}

def is_text_russian(text: str) -> bool:
    """Tekshiradi: Matnda ruscha (kirill yoki lotin transkripsiyasida) so'zlar bormi?"""
    if not text or len(text.strip()) < 2:
        return False
    cyr = len(CYRILLIC_PATTERN.findall(text))
    if cyr >= 2:
        return True
    ru_hits = len(LATIN_RU_REGEX.findall(text))
    uz_hits = len(UZ_REGEX.findall(text))
    if ru_hits >= 1 and uz_hits == 0:
        return True
    if ru_hits >= 2 and ru_hits >= uz_hits:
        return True
    return False

def refine_single_file(file_path: str, translator: GeminiTranslator) -> Dict[str, Any]:
    fname = os.path.basename(file_path)
    
    # 1. Matnlarni ajratib olish
    data = PPTXProcessor.extract_presentation_data(file_path)
    items = []
    for s in data.get("slides", []):
        items.extend(s.get("items", []))
        
    if not items:
        return {"original_name": fname, "final_name": fname, "total_items": 0, "translated_items": 0, "status": "EMPTY"}

    # 2. Butun fayldagi ruscha matnlar sonini hisoblash
    ru_count = sum(1 for it in items if is_text_russian(it.get("original_text", "")))
    is_heavily_russian = ru_count >= 5
    
    untranslated_items = []
    translations_map = {}
    
    for it in items:
        orig = it.get("original_text", "")
        if is_heavily_russian:
            # Agar butun taqdimot ruscha bo'lsa, barcha elementlarni to'liq tarjima qilish
            untranslated_items.append(it)
        else:
            if is_text_russian(orig):
                untranslated_items.append(it)
            else:
                translations_map[it["id"]] = orig
            
    translated_count = 0
    if untranslated_items:
        logger.info(f"  [{fname}] {len(untranslated_items)} / {len(items)} ta matn bloki tarjima qilinmoqda...")
        tr_results = translator.translate_items_batch(
            untranslated_items, 
            target_script="latin", 
            batch_size=30
        )
        for res in tr_results:
            translations_map[res["id"]] = res.get("translated_text", "")
            translated_count += 1
            
    # 3. Vaqtinchalik faylga eksport qilish
    temp_out = file_path + ".refined.tmp"
    PPTXProcessor.apply_translations_and_export(
        original_pptx_path=file_path,
        translations_map=translations_map,
        output_pptx_path=temp_out,
        auto_fit=True,
        target_script="latin",
        clean_watermarks=True,
        translate_images=False
    )
    
    # 4. Yangi nom (standartlashtirish)
    new_fname = FILENAME_MAP.get(fname, fname)
    final_path = os.path.join(os.path.dirname(file_path), new_fname)
    
    if os.path.exists(temp_out):
        if os.path.exists(file_path):
            os.remove(file_path)
        shutil.move(temp_out, final_path)
        
    return {
        "original_name": fname,
        "final_name": new_fname,
        "total_items": len(items),
        "translated_items": translated_count,
        "status": "REFINED"
    }

def main():
    folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"
    files = sorted(glob.glob(os.path.join(folder, "*.pptx")))
    total_files = len(files)
    print(f"Boshlandi: {total_files} ta taqdimot tekshiriladi va kamchiliklari tuzatiladi.")
    
    translator = GeminiTranslator()
    
    # Faqat ruscha yoki tuzatish talab qiladigan fayllarni tanlash
    target_files = []
    for f in files:
        fname = os.path.basename(f)
        try:
            prs = Presentation(f)
            has_ru = False
            for s in prs.slides:
                for sp in s.shapes:
                    if sp.has_text_frame:
                        if is_text_russian(sp.text_frame.text):
                            has_ru = True
                            break
                if has_ru:
                    break
            if has_ru or fname in FILENAME_MAP:
                target_files.append(f)
        except Exception:
            pass

    print(f"Tahrir va to'liq tarjima talab qilinadigan taqdimotlar soni: {len(target_files)} / {total_files}")
    
    results = []
    for idx, f in enumerate(target_files, 1):
        fname = os.path.basename(f)
        try:
            res = refine_single_file(f, translator)
            results.append(res)
            tr_info = f" (Tarjima: {res['translated_items']} blok)" if res['translated_items'] > 0 else ""
            print(f"[{idx}/{len(target_files)}] Yakunlandi: {res['final_name']}{tr_info}")
        except Exception as e:
            logger.error(f"Xatolik yuz berdi ({fname}): {e}")
            results.append({
                "original_name": fname,
                "status": f"ERROR: {e}"
            })
            
    print("\n" + "="*60)
    print("01_Tayyor_Sara_Slaydlar_UZ REFINEMENT VA AUDIT YAKUNLANDI")
    print(f"Jami yangilangan fayllar: {len(results)}")
    print(f"Muvaffaqiyatli yakunlangan: {len([r for r in results if r.get('status') == 'REFINED'])}")
    print(f"Jami tarjima qilingan ruscha bloklar: {sum(r.get('translated_items', 0) for r in results)}")
    print("="*60)

if __name__ == "__main__":
    main()

