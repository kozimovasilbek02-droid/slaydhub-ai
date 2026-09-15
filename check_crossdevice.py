import os

base_dir = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
if os.path.exists(base_dir):
    folders = sorted([f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))])
    print(f"Total available folders: {len(folders)}")
    for i, f in enumerate(folders[:25], 1):
        print(f"[{i}] {f}")
else:
    print(f"Base dir not found: {base_dir}")
