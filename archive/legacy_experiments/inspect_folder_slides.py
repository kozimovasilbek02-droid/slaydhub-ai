import os, glob

base_dir = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
folders = sorted([f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))])

print(f"Total folders: {len(folders)}")
for i, f in enumerate(folders[:25], 1):
    fp = os.path.join(base_dir, f)
    imgs = sorted(glob.glob(os.path.join(fp, "*.png")) + glob.glob(os.path.join(fp, "*.jpg")))
    print(f"[{i}] {f} -> {len(imgs)} images: {[os.path.basename(x) for x in imgs[:5]]}")
