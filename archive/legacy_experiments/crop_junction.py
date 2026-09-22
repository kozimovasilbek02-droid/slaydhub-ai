import os
from PIL import Image

ref_path = r'C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar\100_Days_Action_Plan\Slide_02.png'
im = Image.open(ref_path)
w, h = im.size

# Crop junction between Circle 1 and Circle 2
crop_junction = im.crop((int(0.20*w), int(0.35*h), int(0.45*w), int(0.85*h)))
os.makedirs('output/inspection', exist_ok=True)
junc_path = os.path.abspath('output/inspection/original_junction.png')
crop_junction.save(junc_path)
print('Saved original junction to:', junc_path)
