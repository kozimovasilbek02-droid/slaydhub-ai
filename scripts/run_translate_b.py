"""
Loyiha B — AI Image-to-PPTX Studio
Standalone skript: PPTX faylni tarjima qiladi.

Foydalanish:
  python run_translate_b.py <input.pptx> <output.pptx> [uz-Latn|uz-Cyrl]

API kalit GEMINI_API_KEY muhit o'zgaruvchisidan olinadi.
"""
import sys
import os
import time
import pathlib

# --- modul yo'llarini sozlash ---
PROJECT_ROOT = pathlib.Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

# Loyiha B modullarini import qilish
from backend.core.pptx_processor import PPTXProcessor
from backend.core.gemini_translator import GeminiTranslator

def main():
    if len(sys.argv) < 3:
        print("Foydalanish: python run_translate_b.py <input.pptx> <output.pptx> [uz-Latn|uz-Cyrl]")
        sys.exit(1)

    input_path  = sys.argv[1]
    output_path = sys.argv[2]
    target_lang = sys.argv[3] if len(sys.argv) > 3 else "uz-Latn"
    target_script = "latin" if target_lang == "uz-Latn" else "cyrillic"

    print(f"\n{'='*60}")
    print(f"  Loyiha B — AI Image-to-PPTX Studio")
    print(f"  Fayl: {input_path}")
    print(f"  Til:  {target_lang}")
    print(f"{'='*60}\n")

    t0 = time.perf_counter()

    # Matnlarni chiqarish
    print("📂 PPTX tahlil qilinmoqda...")
    extracted = PPTXProcessor.extract_presentation_data(input_path)
    print(f"✅ PPTX yuklandi: {extracted['slides_count']} ta slayd, {extracted['total_items']} ta element\n")

    all_items = []
    for s in extracted['slides']:
        for item in s['items']:
            all_items.append(item)

    if not all_items:
        print("⚠️  Tarjima qilinadigan matn topilmadi.")
        sys.exit(0)

    # Dastlabki matnlar
    print("📋 Dastlabki matnlar (birinchi 5 ta):")
    for item in all_items[:5]:
        txt = item.get('original_text') or item.get('text', '')
        print(f"  [{item['id']}] {txt[:80]!r}")
    print()

    # Tarjima
    translator = GeminiTranslator()  # API kalit muhit o'zgaruvchisidan
    print("🔄 Gemini tarjima boshlandi...")
    t1 = time.perf_counter()
    translated_results = translator.translate_items_batch(
        items=all_items,
        target_script=target_script,
        domain="Tarix, san'at va biografiya taqdimotlari",
    )
    t2 = time.perf_counter()

    failed = sum(1 for r in translated_results
                 if r.get('translated_text') == (r.get('original_text') or r.get('text', '')))
    print(f"✅ Tarjima tugadi: {len(translated_results)} ta element — {t2-t1:.1f}s\n")

    # Natijani ko'rsatish (birinchi 5 ta)
    trans_map = {r['id']: r['translated_text'] for r in translated_results}
    print("📋 Tarjima qilingan matnlar (birinchi 5 ta):")
    for item in all_items[:5]:
        iid = item['id']
        print(f"  [{iid}] {trans_map.get(iid, '???')[:80]!r}")
    print()

    # Tarjimani saqlab chiqarish
    PPTXProcessor.apply_translations_and_export(
        original_pptx_path=input_path,
        translations_map=trans_map,
        output_pptx_path=output_path,
        auto_fit=True,
        target_script=target_script
    )

    t3 = time.perf_counter()
    print(f"✅ Saqlandi: {output_path}")
    print(f"\n⏱  Jami vaqt: {t3-t0:.1f}s")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
