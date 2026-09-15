import os

base_dir = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
folders = sorted([f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))])

batch2 = folders[20:40]
print(f"Total Folders: {len(folders)}")
print(f"Batch 2 (21 to 40): {len(batch2)} presentations:")
for i, f in enumerate(batch2, 21):
    print(f"[{i}] {f}")
