import os, glob

base_dir = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
folders = sorted([f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))])
batch2 = folders[20:40]

for name in batch2:
    p = os.path.join(base_dir, name)
    slides = glob.glob(os.path.join(p, "Slide_*.png")) + glob.glob(os.path.join(p, "Slide_*.jpg"))
    print(f"{name}: {len(slides)} slides")
