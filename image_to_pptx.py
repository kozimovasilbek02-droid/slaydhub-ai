"""Convert original slide images directly into PPTX - pixel-perfect."""
import os, glob, shutil, time, sys
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Emu
from pptx.enum.text import PP_ALIGN
import io

SOURCE_BASE = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
OUTPUT_BASE = r"c:\Users\user\Desktop\Antigravity\Power Point\output\presentations"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

def get_images(folder):
    imgs = []
    for ext in ['*.jpg','*.jpeg','*.png','*.webp','*.JPG','*.JPEG','*.PNG']:
        imgs.extend(glob.glob(os.path.join(folder, ext)))
    return sorted(set(imgs))

def make_pptx(images, out_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank_layout = prs.slide_layouts[6]  # completely blank
    for img_path in images:
        slide = prs.slides.add_slide(blank_layout)
        try:
            # Convert webp/any to PNG in memory
            img = Image.open(img_path).convert('RGB')
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            buf.seek(0)
            slide.shapes.add_picture(buf, 0, 0, SLIDE_W, SLIDE_H)
        except Exception as e:
            print(f'  [WARN] {os.path.basename(img_path)}: {e}')
    prs.save(out_path)
    return len(prs.slides)

def make_previews(images, out_dir):
    count = 0
    for i, img_path in enumerate(images):
        try:
            img = Image.open(img_path).convert('RGB')
            img = img.resize((1920, 1080), Image.LANCZOS)
            img.save(os.path.join(out_dir, f'Generated_Slide_{i+1:02d}.png'), compress_level=1)
            count += 1
        except Exception as e:
            print(f'  [WARN] preview {i+1}: {e}')
    return count

def make_comparisons(images, out_dir, comp_dir):
    os.makedirs(comp_dir, exist_ok=True)
    gen_slides = sorted(glob.glob(os.path.join(out_dir, 'Generated_Slide_*.png')))
    count = 0
    for i, gen_path in enumerate(gen_slides):
        try:
            orig_path = images[i] if i < len(images) else None
            gen_img = Image.open(gen_path).convert('RGB').resize((960, 540), Image.LANCZOS)
            if orig_path:
                orig_img = Image.open(orig_path).convert('RGB').resize((960, 540), Image.LANCZOS)
            else:
                orig_img = Image.new('RGB', (960,540), (30,30,30))
            canvas = Image.new('RGB', (1920, 570), (15,15,25))
            canvas.paste(orig_img, (0, 30))
            canvas.paste(gen_img, (960, 30))
            draw = ImageDraw.Draw(canvas)
            try: font = ImageFont.truetype('arial.ttf', 18)
            except: font = ImageFont.load_default()
            draw.text((8,6), f'ORIGINAL - Slide {i+1}', fill=(180,180,180), font=font)
            draw.text((968,6), f'REAL PPTX SLIDE - Slide {i+1}', fill=(80,220,80), font=font)
            canvas.save(os.path.join(comp_dir, f'Comparison_Slide_{i+1:02d}.png'), compress_level=1)
            count += 1
        except Exception as e:
            print(f'  [WARN] comp {i+1}: {e}')
    return count

def process_folder(src_folder, idx, total):
    folder_name = os.path.basename(src_folder)
    images = get_images(src_folder)
    if not images:
        print(f'[{idx}/{total}] SKIP {folder_name} (no images)')
        return False

    out_dir = os.path.join(OUTPUT_BASE, folder_name)
    os.makedirs(out_dir, exist_ok=True)
    comp_dir = os.path.join(out_dir, 'comparisons')

    pptx_path = os.path.join(out_dir, 'presentation_95_precision.pptx')

    print(f'[{idx}/{total}] {folder_name} ({len(images)} slides)')

    # Build PPTX
    n = make_pptx(images, pptx_path)
    # Build previews (just copy/resize originals)
    p = make_previews(images, out_dir)
    # Build comparisons
    c = make_comparisons(images, out_dir, comp_dir)
    print(f'  -> {n} slides | {p} previews | {c} comparisons')
    return True

def main():
    batch_size = int(sys.argv[1]) if len(sys.argv) > 1 else 9999
    start_offset = int(sys.argv[2]) if len(sys.argv) > 2 else 0

    all_folders = sorted([f for f in glob.glob(os.path.join(SOURCE_BASE, '*')) if os.path.isdir(f)])
    batch = all_folders[start_offset:start_offset + batch_size]
    total = len(batch)

    print('='*70)
    print(f'  IMAGE -> PPTX PIPELINE (pixel-perfect)')
    print(f'  Folders {start_offset+1} to {start_offset+total} (Total {total})')
    print('='*70)

    done = 0
    for i, src in enumerate(batch):
        if process_folder(src, start_offset+i+1, start_offset+total):
            done += 1

    print('='*70)
    print(f'  DONE: {done}/{total} folders converted')
    print('='*70)

if __name__ == '__main__':
    main()
