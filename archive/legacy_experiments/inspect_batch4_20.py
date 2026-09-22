import os

base_dir = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
folders = sorted([f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))])

batch4 = folders[90:110]
print(f"Total Folders: {len(folders)}")
print(f"Batch 4 (91 to 110) - {len(batch4)} presentations:")
for i, f in enumerate(batch4, 91):
    print(f"[{i}] {f}")
