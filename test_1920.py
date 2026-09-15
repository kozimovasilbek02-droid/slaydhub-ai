import os, glob
import slide_recreator_engine
import slide_manager

src = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar\1920S_Theme"
out = r"c:\Users\user\Desktop\Antigravity\Power Point\output\presentations\1920S_Theme"

pptx, imgs = slide_recreator_engine.build_full_presentation(src, out)
previews = slide_manager.export_pptx_previews(pptx, "1920S_Theme")
slide_recreator_engine.generate_comparisons(src, out, imgs, previews)
