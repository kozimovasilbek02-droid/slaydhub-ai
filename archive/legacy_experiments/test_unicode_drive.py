import os, sys

# Use explicit Unicode escape for Russian 'Мой диск'
# М = \u041c, о = \u043e, й = \u0439, д = \u0434, и = \u0438, с = \u0441, к = \u043a
gdrive_base = "G:\\\u041c\u043e\u0439 \u0434\u0438\u0441\u043a\\SlideEgg_Downloads\\Prezentatsiyalar"

print("Checking Unicode G-Drive path:", os.path.exists(gdrive_base))
if os.path.exists(gdrive_base):
    folders = sorted(os.listdir(gdrive_base))
    print(f"Total Folders: {len(folders)}")
    for i, f in enumerate(folders[:20], 1):
        print(f"[{i}] {f}")
