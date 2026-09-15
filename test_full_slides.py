import os, glob
import slide_recreator_engine
import slide_manager

CURRENT_DIR = r"c:\Users\user\Desktop\Antigravity\Power Point"
STORAGE_DIR = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
OUTPUT_DIR = os.path.join(CURRENT_DIR, "output", "presentations")

# Let's test on 3 folders with different slide counts:
test_folders = [
    "100_Day_New_Leadership_Position_Plan",  # 11 slides
    "10_Types_Of_Innovation_Ppt",            # 1 slide
    "10_Step_Business_Process_Ppt_Editable_Free_Download", # 11 slides
    "10Vs_Of_Big_Data",                      # 13 slides
]

for f in test_folders:
    src = os.path.join(STORAGE_DIR, f)
    out = os.path.join(OUTPUT_DIR, f)
    print(f"\n--- Testing {f} ---")
    pptx, imgs = slide_recreator_engine.build_full_presentation(src, out)
    if pptx:
        previews = slide_manager.export_pptx_previews(pptx, f)
        print(f"Exported {len(previews)} previews")
        comps = slide_recreator_engine.generate_comparisons(src, out, imgs, previews)
        print(f"Generated {len(comps)} comparisons")
