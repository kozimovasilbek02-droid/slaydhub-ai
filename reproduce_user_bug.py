import os
from pptx import Presentation
from core.pptx_text_replacer import PPTXTextReplacer
from core.slide_preview import export_pptx_to_pdf, get_presentation_previews

tmpl_path = r"C:\Users\user\Desktop\slayd\PresentationGO\Templates\Atom Diagram (3 Orbits) for PowerPoint and Google Slides.pptx"
prs = Presentation(tmpl_path)
s2 = prs.slides[0]

slide_data = {
    "title": "Kvant Superpozitsiya va Entanglement Prinsiplari Klassik Algoritmlarning Hisoblash Murakkabligini Keskin Oshiradi",
    "points": [
        "Qubitning Holat Vektori: Bloch sferasi va $\\alpha|0\\rangle + \\beta|1\\rangle$ superpozitsiya holati.",
        "Kvant Entanglement: Zarrachalar orasidagi kvant bog'liqlik axborotni masofaga uzatish tezligini oshiradi.",
        "Klassik Turing: Klassik bitlar faqat 0 yoki 1 qiymatini qabul qilsa, kvant parallelizmi barcha holatlarni bir vaqtda qayta ishlaydi."
    ],
    "layout_type": "cards_grid"
}

topic = "Kvant kompyuterlari"
print("Before replacement shapes:", len(s2.shapes))
for shp in s2.shapes:
    print(f"  {shp.name} type={shp.shape_type} pos=({shp.left.pt:.1f}, {shp.top.pt:.1f}) size=({shp.width.pt:.1f}, {shp.height.pt:.1f})")

PPTXTextReplacer.replace_slide_content(s2, slide_data, topic)

print("\nAfter replacement shapes:", len(s2.shapes))
for shp in s2.shapes:
    txt = repr(shp.text_frame.text[:40].replace('\n', ' ')) if shp.has_text_frame else 'NO TEXT'
    print(f"  {shp.name} type={shp.shape_type} pos=({shp.left.pt:.1f}, {shp.top.pt:.1f}) size=({shp.width.pt:.1f}, {shp.height.pt:.1f}) txt={txt}")

out_test = r"output\test_reproduce_atom.pptx"
prs.save(out_test)
print(f"Saved to {out_test}")
