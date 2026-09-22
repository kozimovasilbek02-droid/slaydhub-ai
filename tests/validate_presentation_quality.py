# -*- coding: utf-8 -*-
from typing import Dict, Any
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def inspect_deck(pptx_path: str) -> Dict[str, Any]:
    """
    Audits an assembled PPTX deck for:
    - Empty placeholders (which trigger PowerPoint's default ghost prompt text)
    - Ghost prompt strings ("Click to edit title", "Текст слайда", etc.)
    - Promotional watermarks/logos (Logotype, site names)
    - Pictures count
    """
    prs = Presentation(pptx_path)
    print(f"=== Presentation Quality Audit: {pptx_path} ===")
    print(f"Total Slides: {len(prs.slides)}")

    empty_placeholders = 0
    ghost_prompts = 0
    watermarks_found = 0
    pictures_count = 0

    for idx, s in enumerate(prs.slides):
        print(f"\n--- Slide #{idx+1} (Layout: '{s.slide_layout.name}') ---")
        
        # Check placeholders
        ph_list = list(s.placeholders)
        for p in ph_list:
            txt = p.text_frame.text.strip() if p.has_text_frame else ""
            if not txt:
                empty_placeholders += 1
                print(f"  [WARN] Empty Placeholder: '{p.name}'")

        # Check all shapes
        for shp in s.shapes:
            if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures_count += 1
                print(f"  [IMG] Picture Shape: '{shp.name}' ({shp.width.pt:.1f}pt x {shp.height.pt:.1f}pt)")
            if shp.has_text_frame:
                txt = shp.text_frame.text.strip()
                low = txt.lower()
                if any(k in low for k in ["click to edit", "click to add", "текст слайда", "заголовок слайда"]):
                    ghost_prompts += 1
                    print(f"  [FAIL] Ghost Prompt detected: '{txt}'")
                if any(k in low for k in ["logotype", "presentationgo", "slidesmania", "slideegg"]):
                    watermarks_found += 1
                    print(f"  [FAIL] Watermark detected: '{txt}'")
            if txt:
                safe_txt = txt[:60].replace("\n", " ").encode("ascii", "replace").decode("ascii")
                print(f"  [TEXT] '{shp.name}': {safe_txt}...")

    print("\n=== AUDIT SUMMARY ===")
    print(f"Empty Placeholders: {empty_placeholders} (Should be 0)")
    print(f"Ghost Prompts: {ghost_prompts} (Should be 0)")
    print(f"Watermarks: {watermarks_found} (Should be 0)")
    print(f"Scientific Pictures: {pictures_count}")

    return {
        "slides_count": len(prs.slides),
        "empty_placeholders": empty_placeholders,
        "ghost_prompts": ghost_prompts,
        "watermarks_found": watermarks_found,
        "pictures_count": pictures_count
    }


if __name__ == '__main__':
    default_path = 'output/academic_studio/Kvant_kompyuterlari_va_asimmetrik_kripto_Customized.pptx'
    import sys
    p = sys.argv[1] if len(sys.argv) > 1 else default_path
    inspect_deck(p)
