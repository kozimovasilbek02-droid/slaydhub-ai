import os

base_dir = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
folders = sorted([f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))])

batch3 = folders[40:90]
print(f"Total Folders: {len(folders)}")
print(f"Batch 3 (41 to 90) - {len(batch3)} presentations:")
for i, f in enumerate(batch3, 41):
    print(f"[{i}] {f}")
