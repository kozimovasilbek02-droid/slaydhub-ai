import os

base_dir = r"C:\Users\user\CrossDevice\HONOR X9c (1)\storage\SlideEgg_Downloads\Prezentatsiyalar"
folders = sorted([f for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))])

batch5 = folders[110:160]
print(f"Total Folders: {len(folders)}")
print(f"Batch 5 (111 to 160) - {len(batch5)} presentations:")
for i, f in enumerate(batch5, 111):
    print(f"[{i}] {f}")
