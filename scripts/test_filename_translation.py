import re
import os
import sys
from pathlib import Path

# Add project root to sys.path
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from backend.core.gemini_translator import GeminiTranslator
from backend.core.transliteration import ensure_script

PREFIX_MAP = {
    r"^Pravo_VUZ_?": "Huquq_OTM_",
    r"^Psikh_VUZ_?": "Psixologiya_OTM_",
    r"^Sotsiologiya_VUZ_?": "Sotsiologiya_OTM_",
    r"^Tekhnika_VUZ_?": "Texnika_OTM_",
    r"^Pedagogika_VUZ_?": "Pedagogika_OTM_",
    r"^Ekonomika_VUZ_?": "Iqtisodiyot_OTM_",
    r"^Meditsina_VUZ_?": "Tibbiyot_OTM_",
    r"^Biologiya_VUZ_?": "Biologiya_OTM_",
    r"^Fizika_VUZ_?": "Fizika_OTM_",
    r"^Khimiya_VUZ_?": "Kimyo_OTM_",
    r"^Istoriya_VUZ_?": "Tarix_OTM_",
    r"^Filosofiya_VUZ_?": "Falsafa_OTM_",
    r"^Ekologiya_VUZ_?": "Ekologiya_OTM_",
    r"^Geografiya_VUZ_?": "Geografiya_OTM_",
    r"^Informatika_VUZ_?": "Informatika_OTM_",
    r"^Matematika_VUZ_?": "Matematika_OTM_",
    r"^Dizayn_VUZ_?": "Dizayn_OTM_",
    r"^Arxitektura_VUZ_?": "Arxitektura_OTM_",
}

def clean_and_translate_filename(raw_filename: str, translator: GeminiTranslator, cache: dict) -> str:
    if raw_filename in cache:
        return cache[raw_filename]
        
    stem = Path(raw_filename).stem
    ext = ".pptx"  # Standardize all outputs to .pptx
    
    # 1. Check known prefixes
    prefix_uz = ""
    rest_stem = stem
    for pat, uz_p in PREFIX_MAP.items():
        if re.search(pat, stem, re.IGNORECASE):
            prefix_uz = uz_p
            rest_stem = re.sub(pat, "", stem, flags=re.IGNORECASE)
            break
            
    # Clean up junk suffixes
    rest_stem = re.sub(r'_wecompress\.com_', '', rest_stem, flags=re.IGNORECASE)
    rest_stem = re.sub(r'[\(\[\{]\d+[\)\]\}]', '', rest_stem)
    rest_stem = re.sub(r'[-_]?(kopi[jy]a|copy|szhatyy|compressed)', '', rest_stem, flags=re.IGNORECASE)
    
    # Replace underscores with spaces for translation
    words_to_translate = rest_stem.replace('_', ' ').strip()
    
    if not words_to_translate:
        translated_stem = "Taqdimot"
    else:
        try:
            prompt = f"Ushbu ruscha yoki translit qilingan mavzu nomini O'zbek tili (Lotin yozuvi)ga lo'nda va toza tarjima qiling (faqat nomni qaytaring, ortiqcha so'zsiz):\n\n\"{words_to_translate}\""
            res = translator.client.models.generate_content(
                model=translator.model_name,
                contents=prompt
            )
            trans = res.text.strip().strip('"\'').strip()
            trans = re.sub(r'[/\\:*?"<>|]', '', trans)
            # Replace spaces with underscores
            words = [w.strip() for w in trans.split() if w.strip()]
            translated_stem = "_".join(words)
        except Exception as e:
            # Fallback transliteration/cleanup
            translated_stem = rest_stem.strip('_')
            
    final_name = f"{prefix_uz}{translated_stem}{ext}"
    final_name = re.sub(r'_+', '_', final_name).strip('_')
    if not final_name.endswith('.pptx'):
        final_name += '.pptx'
        
    cache[raw_filename] = final_name
    return final_name

if __name__ == '__main__':
    t = GeminiTranslator()
    c = {}
    test_files = [
        "Pravo_VUZ_stadii_ugolovnogo_protsessa.pptx",
        "Tekhnika_VUZ_alternativnye_istochniki_energii.pptx",
        "Sotsiologiya_VUZ_vrednye_privychki_i_ih_profilaktika.pptx",
        "Psikh_VUZ_kak_nauchit_rebenka_druzhit.pptx",
        "Pravo_VUZ_tema_6.2-6.3_osobennosti_meny_zemelnyh_uchastkov_wecompress.com_.pptx"
    ]
    for tf in test_files:
        res = clean_and_translate_filename(tf, t, c)
        print(f"{tf} \n  -> {res}\n")
