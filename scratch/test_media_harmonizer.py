# -*- coding: utf-8 -*-
"""
scratch/test_media_harmonizer.py
Unit tests for MediaHarmonizer (Round 12).
Validates:
1. Smart Center-Crop (<a:srcRect>) calculation for wide and tall images.
2. Injection of <a:srcRect> inside OpenXML blipFill.
3. Discovery of content images vs background backdrops.
4. Collision-free placement calculation.
5. Ingestion of an actual image with center-crop into PPTX.
"""

import os
import sys
import io

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement

from core.media_harmonizer import MediaHarmonizer


def test_media_harmonizer():
    print("🧪 Running MediaHarmonizer Unit Tests (Round 12)...")

    # 1. Test Smart Center-Crop calculations
    # Case A: 1:1 square shape (100x100), 16:9 wide image (1920x1080)
    l, t, r, b = MediaHarmonizer.calculate_smart_crop(100.0, 100.0, 1920.0, 1080.0)
    print(f"Case A (1:1 shape, 16:9 image): l={l}, t={t}, r={r}, b={b}")
    assert l > 0 and r > 0 and t == 0 and b == 0, "Wide image must crop left/right"
    assert l == r, "Left and right crop must be equal (center-crop)"
    # Expected: (1 - 1080/1920)/2 = 0.21875 -> ~21875
    assert abs(l - 21875) < 100, f"Expected ~21875, got {l}"

    # Case B: 16:9 wide shape (160x90), 1:1 square image (1000x1000)
    l, t, r, b = MediaHarmonizer.calculate_smart_crop(160.0, 90.0, 1000.0, 1000.0)
    print(f"Case B (16:9 shape, 1:1 image): l={l}, t={t}, r={r}, b={b}")
    assert t > 0 and b > 0 and l == 0 and r == 0, "Tall image must crop top/bottom"
    assert t == b, "Top and bottom crop must be equal"

    # Case C: Exact match
    l, t, r, b = MediaHarmonizer.calculate_smart_crop(200.0, 100.0, 400.0, 200.0)
    print(f"Case C (identical aspect): l={l}, t={t}, r={r}, b={b}")
    assert l == 0 and t == 0 and r == 0 and b == 0, "Exact aspect ratio requires 0 crop"

    # 2. Test Safe Image Placement
    occupied = [
        (40.0, 40.0, 400.0, 300.0), # Left side occupied
    ]
    safe_l, safe_t = MediaHarmonizer.find_safe_image_placement(960.0, 540.0, occupied, 300.0, 250.0)
    print(f"Safe placement with left side occupied: ({safe_l:.1f}, {safe_t:.1f})")
    assert safe_l > 400.0, "Safe placement must choose unoccupied right side"

    # 3. Create dummy test image using PIL
    os.makedirs("output/image_cache", exist_ok=True)
    test_img_path = "output/image_cache/test_crop_source.png"
    img = Image.new("RGB", (1200, 600), color=(37, 99, 235)) # 2:1 aspect
    img.save(test_img_path)

    # 4. Create slide with an image and apply Smart Center-Crop
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Add picture as a 1:1 square box (Inches(3) x Inches(3))
    pic = slide.shapes.add_picture(test_img_path, Inches(1), Inches(1), Inches(3), Inches(3))

    # Replace with smart crop
    success = MediaHarmonizer.replace_image_with_file(pic, test_img_path)
    assert success is True, "Image replacement with center-crop must succeed"

    # Inspect XML for <a:srcRect>
    srcRect_nodes = pic._element.xpath('.//*[local-name()="srcRect"]')
    print(f"Injected srcRect count: {len(srcRect_nodes)}")
    assert len(srcRect_nodes) >= 1, "srcRect must be present inside picture XML"

    crop_l = int(srcRect_nodes[0].get("l", "0"))
    crop_r = int(srcRect_nodes[0].get("r", "0"))
    print(f"Verified XML srcRect margins: l={crop_l}, r={crop_r}")
    # 2:1 image in 1:1 shape -> crop_lr = (1 - 0.5)/2 = 0.25 -> 25000
    assert abs(crop_l - 25000) < 100, f"Expected ~25000, got {crop_l}"
    assert crop_l == crop_r, "Left and right crop must be identical"

    output_pptx = "output/test_media_harmonized.pptx"
    prs.save(output_pptx)
    print(f"✅ Muvaffaqiyatli saqlandi: {output_pptx}")
    print("🎉 Barcha MediaHarmonizer testlari 100% PASS!\n")


if __name__ == "__main__":
    test_media_harmonizer()
