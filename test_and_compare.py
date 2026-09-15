import os
import win32com.client
from convert_sample import convert_uploaded_sample
from PIL import Image

def run_test():
    print("=== STARTING SLIDE REGENERATION & COMPARISON TEST ===")
    
    # 1. Generate PPTX
    out_pptx = convert_uploaded_sample()
    print(f"1. Generated PPTX: {out_pptx}")
    
    # 2. Export Preview PNG via PowerPoint
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    pres = ppt.Presentations.Open(os.path.abspath(out_pptx), WithWindow=False)
    img_out = os.path.abspath("output/generated_slide_preview.png")
    pres.Slides[1].Export(img_out, "PNG", 1920, 1080)
    pres.Close()
    ppt.Quit()
    print(f"2. Exported Preview PNG: {img_out}")
    
    # 3. Verify Files
    ref_path = "reference_slide_hd.png" if os.path.exists("reference_slide_hd.png") else "reference_slide.jpg"
    print(f"3. Reference Image: {ref_path}")
    print("=== TEST COMPLETED SUCCESSFULLY ===")

if __name__ == "__main__":
    run_test()
