import os, glob, time
from PIL import Image, ImageDraw, ImageFont

OUTPUT_BASE = r"c:\Users\user\Desktop\Antigravity\Power Point\output\presentations"
SOURCE_BASE = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"

def export_previews(pptx_path, output_dir):
    try:
        import comtypes.client, pymupdf
        pdf_path = pptx_path.replace(".pptx", "_tmp.pdf")
        ppt = comtypes.client.CreateObject("PowerPoint.Application")
        ppt.Visible = 1
        deck = ppt.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        deck.SaveAs(os.path.abspath(pdf_path), 32)
        deck.Close()
        ppt.Quit()
        doc = pymupdf.open(pdf_path)
        count = 0
        for i, page in enumerate(doc):
            mat = pymupdf.Matrix(1920/page.rect.width, 1080/page.rect.height)
            pix = page.get_pixmap(matrix=mat)
            pix.save(os.path.join(output_dir, f"Generated_Slide_{i+1:02d}.png"))
            count += 1
        doc.close()
        if os.path.exists(pdf_path): os.remove(pdf_path)
        return count
    except Exception as e:
        print(f"  [WARN] {e}"); return 0

def create_comparisons(folder_path, source_folder):
    comps_dir = os.path.join(folder_path, "comparisons")
    os.makedirs(comps_dir, exist_ok=True)
    gen_slides = sorted(glob.glob(os.path.join(folder_path, "Generated_Slide_*.png")))
    orig_images = []
    if os.path.exists(source_folder):
        for ext in ["*.jpg","*.jpeg","*.png","*.webp"]:
            orig_images.extend(glob.glob(os.path.join(source_folder, ext)))
        orig_images = sorted(orig_images)
    count = 0
    for i, gen_path in enumerate(gen_slides):
        comp_path = os.path.join(comps_dir, f"Comparison_Slide_{i+1:02d}.png")
        if os.path.exists(comp_path): count += 1; continue
        try:
            gen_img = Image.open(gen_path).convert("RGB").resize((960, 540), Image.LANCZOS)
            if i < len(orig_images):
                orig_img = Image.open(orig_images[i]).convert("RGB").resize((960, 540), Image.LANCZOS)
            else:
                orig_img = Image.new("RGB",(960,540),(50,50,50))
            canvas = Image.new("RGB",(1920,600),(20,20,30))
            canvas.paste(orig_img,(0,30)); canvas.paste(gen_img,(960,30))
            draw = ImageDraw.Draw(canvas)
            try: font = ImageFont.truetype("arial.ttf", 20)
            except: font = ImageFont.load_default()
            draw.text((10,5),f"ORIGINAL - Slide {i+1}",fill=(200,200,200),font=font)
            draw.text((970,5),f"GENERATED - Slide {i+1}",fill=(100,220,100),font=font)
            canvas.save(comp_path, compress_level=1); count += 1
        except Exception as e: print(f"  [WARN] comp {i+1}: {e}")
    return count

missing = [f for f in sorted(glob.glob(os.path.join(OUTPUT_BASE,"*"))) if os.path.isdir(f) and glob.glob(os.path.join(f,"*.pptx")) and not glob.glob(os.path.join(f,"Generated_Slide_*.png"))]
print(f"Found {len(missing)} folders with missing previews")
fixed = 0
for i, folder_path in enumerate(missing):
    folder_name = os.path.basename(folder_path)
    pptx = glob.glob(os.path.join(folder_path,"*.pptx"))[0]
    source = os.path.join(SOURCE_BASE, folder_name)
    print(f"[{i+1}/{len(missing)}] Fixing: {folder_name}")
    n = export_previews(pptx, folder_path)
    if n > 0:
        c = create_comparisons(folder_path, source)
        print(f"  -> {n} previews, {c} comparisons"); fixed += 1
    time.sleep(0.3)
print(f"DONE: {fixed}/{len(missing)} fixed")
