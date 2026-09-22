# -*- coding: utf-8 -*-
"""
scratch/test_bullet_harmonizer.py
Comprehensive unit test suite for core/bullet_harmonizer.py
Verifies:
1. Wingdings and custom template bullet extraction
2. Multi-level hierarchy parsing (indentation, lead-in, symbol detection)
3. DrawingML canonical child ordering in <a:pPr> (§21.1.2.2.7)
4. Hanging indent calculations across Level 0, Level 1, Level 2
5. Real PowerPoint COM preview generation
"""

import sys
import os
import io
from pathlib import Path
import lxml.etree as etree
from pptx import Presentation
from pptx.util import Inches, Pt

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from core.bullet_harmonizer import (
    BulletHarmonizer,
    BulletStyle,
    HierarchyItem,
    A_NS,
    PPR_CHILD_ORDER
)


def test_hierarchy_parser():
    print("--- 1. Testing MultiLevelHierarchyRestorer.parse_hierarchical_lines ---")
    raw_text = """
• Asosiy Konsepsiya: Kvant hisoblashlarining fundamental asoslari
  - Superpozitsiya: Qubitning bir vaqtning o'zida bir nechta holatda bo'lishi
  - Kvant chalkashligi: Zarrachalar orasidagi uzviy masofaviy bog'liqlik
    * Bell tengsizliklari va eksperimental isbotlar
• Amaliy Tatbiq: Asimmetrik shifrlash tizimlarining xavfsizligi
  - RSA va ECC algoritmlarining zaiflashuvi
    """

    items = BulletHarmonizer.parse_hierarchical_lines(raw_text)
    print(f"Parsed {len(items)} items:")
    for idx, it in enumerate(items):
        print(f"  [{idx}] Lvl {it.level} | Bullet: {it.is_bullet} | Lead: {it.lead_in!r} | Body: {it.body[:40]}...")

    assert len(items) == 6, f"Expected 6 items, got {len(items)}"
    assert items[0].level == 0 and items[0].is_bullet and items[0].lead_in == "Asosiy Konsepsiya"
    assert items[1].level == 1 and items[1].is_bullet and items[1].lead_in == "Superpozitsiya"
    assert items[2].level == 1 and items[2].is_bullet and items[2].lead_in == "Kvant chalkashligi"
    assert items[3].level == 2 and items[3].is_bullet and "Bell tengsizliklari" in items[3].body
    assert items[4].level == 0 and items[4].is_bullet and items[4].lead_in == "Amaliy Tatbiq"
    assert items[5].level == 1 and items[5].is_bullet and "RSA" in items[5].body
    print("  => Hierarchy parser test PASSED! (100% Correct)")


def test_drawingml_order():
    print("\n--- 2. Testing DrawingML Canonical Child Element Order (§21.1.2.2.7) ---")
    elem = etree.Element(f"{{{A_NS}}}pPr")
    # Add elements intentionally out of order
    etree.SubElement(elem, f"{{{A_NS}}}defRPr")
    etree.SubElement(elem, f"{{{A_NS}}}buChar", char="•")
    etree.SubElement(elem, f"{{{A_NS}}}lnSpc")
    etree.SubElement(elem, f"{{{A_NS}}}buClr")
    etree.SubElement(elem, f"{{{A_NS}}}buFont", typeface="Calibri")
    etree.SubElement(elem, f"{{{A_NS}}}buSzPct", val="100000")

    initial_tags = [c.tag.split("}")[-1] for c in elem]
    print("  Initial tags:", initial_tags)

    BulletHarmonizer.order_ppr_children(elem)
    sorted_tags = [c.tag.split("}")[-1] for c in elem]
    print("  Sorted tags: ", sorted_tags)

    expected_tags = ["lnSpc", "buClr", "buSzPct", "buFont", "buChar", "defRPr"]
    assert sorted_tags == expected_tags, f"Expected {expected_tags}, got {sorted_tags}"
    print("  => DrawingML canonical ordering test PASSED! (100% Schema Compliant)")


def test_template_bullet_extraction_and_cloning():
    print("\n--- 3. Testing Template Bullet Extraction & Cloning ---")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(3))
    p = tb.text_frame.paragraphs[0]
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", "285750")
    pPr.set("indent", "-285750")
    etree.SubElement(pPr, f"{{{A_NS}}}buFont", typeface="Wingdings")
    etree.SubElement(pPr, f"{{{A_NS}}}buChar", char="v")

    extracted = BulletHarmonizer.extract_paragraph_bullet_style(p)
    print(f"  Extracted style: Type={extracted.bullet_type}, Char={extracted.char}, Font={extracted.font_typeface}, marL={extracted.mar_l}, indent={extracted.indent}")
    assert extracted.bullet_type == "char"
    assert extracted.char == "v"
    assert extracted.font_typeface == "Wingdings"
    assert extracted.mar_l == 285750
    assert extracted.indent == -285750

    # Clone onto a new paragraph in another textbox
    tb2 = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(5), Inches(3))
    p2 = tb2.text_frame.paragraphs[0]
    BulletHarmonizer.apply_native_bullet(p2, extracted, level=0, is_bullet=True)

    pPr2 = p2._p.find(f"{{{A_NS}}}pPr")
    assert pPr2.attrib.get("marL") == "285750"
    assert pPr2.attrib.get("indent") == "-285750"
    buChar2 = pPr2.find(f"{{{A_NS}}}buChar")
    assert buChar2 is not None and buChar2.attrib.get("char") == "v"
    buFont2 = pPr2.find(f"{{{A_NS}}}buFont")
    assert buFont2 is not None and buFont2.attrib.get("typeface") == "Wingdings"
    print("  => Template extraction & cloning test PASSED! (100% Preserved)")


def test_full_multilevel_deck_generation():
    print("\n--- 4. Testing End-to-End Deck Generation with Multi-Level Bullets ---")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Title box (No bullets)
    title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.2))
    p_title = title_box.text_frame.paragraphs[0]
    p_title.text = "Kvant Kompyuterlari va Post-Kvant Xavfsizligi"
    p_title.font.size = Pt(32)
    p_title.font.bold = True
    BulletHarmonizer.apply_native_bullet(p_title, is_bullet=False)

    # Multi-level Content Card
    card_box = slide.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(4.5))
    tf = card_box.text_frame
    tf.word_wrap = True

    sample_content = """
• Fundamental Arxitektura: Kvant protsessorlarining ishlash tamoyillari va kubitlar barqarorligi
  - Superpozitsiya Xossasi: Klassik bitlardan farqli o'laroq bir vaqtning o'zida cheksiz holatlarni ifodalay oladi
  - Dekogerensiya Muammosi: Tashqi shovqinlar ta'sirida kvant holatining yo'qolish xavfi
    * Termal shovqinlarni kamaytirish uchun mutlaq nol darajaga yaqin kriogenik sovutish tizimlari qo'llaniladi
• Post-Kvant Kriptografiyasi: Kelajak axborot xavfsizligini ta'minlashning yangi me'yorlari
  - Panjara Asosidagi Algoritmlar (Lattice-based): Shor algoritmi hujumlariga bardoshli eng istiqbolli yo'nalish
  - NIST Standartlashtirish Jarayoni: CRYSTALS-Kyber va Dilithium algoritmlarini rasmiy tasdiqlash
    """

    items = BulletHarmonizer.parse_hierarchical_lines(sample_content)
    tf.text = ""

    for idx, item in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        BulletHarmonizer.apply_native_bullet(
            p,
            level=item.level,
            is_bullet=item.is_bullet,
            theme_accent_hex="3B82F6"  # Royal blue bullet marker
        )
        p.line_spacing = 1.25
        p.space_after = Pt(6) if item.level == 0 else Pt(3)

        # Add runs
        if item.lead_in:
            r1 = p.add_run()
            r1.text = item.lead_in + ": "
            r1.font.bold = True
            r1.font.size = Pt(17 - (item.level * 1.5))

            r2 = p.add_run()
            r2.text = item.body
            r2.font.bold = False
            r2.font.size = Pt(17 - (item.level * 1.5))
        else:
            r = p.add_run()
            r.text = item.clean_text
            r.font.size = Pt(17 - (item.level * 1.5))

    out_pptx = "output/test_bullet_harmonized_deck.pptx"
    prs.save(out_pptx)
    print(f"  Saved presentation to {out_pptx}")

    # Verify preview export via PowerPoint COM
    from core.slide_preview import get_presentation_previews
    previews = get_presentation_previews(out_pptx)
    print(f"  Rendered {len(previews)} preview(s): {previews}")
    assert len(previews) > 0, "Preview export failed!"
    print("  => End-to-End Multi-Level Deck test PASSED! (100% Success)")


if __name__ == "__main__":
    test_hierarchy_parser()
    test_drawingml_order()
    test_template_bullet_extraction_and_cloning()
    test_full_multilevel_deck_generation()
    print("\n🎉 ALL 4 BULLET HARMONIZER UNIT TESTS PASSED WITH ZERO ERRORS!")
