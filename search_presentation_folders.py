import os, glob

print("Searching for presentation folders with Slide_01.png...")

search_roots = [
    r"C:\Users\user\Desktop",
    r"C:\Users\user\Downloads",
    r"C:\Users\user\Documents",
    r"C:\Users\user",
]

found_folders = []
for sroot in search_roots:
    if os.path.exists(sroot):
        for root, dirs, files in os.walk(sroot):
            if "Slide_01.png" in files or "Slide_01.jpg" in files:
                found_folders.append(root)
                print(f"  FOUND: {root}")
            # Prune deep node_modules, appdata, virtualenvs
            if any(p in root.lower() for p in ['appdata', 'node_modules', '.git', 'site-packages']):
                dirs.clear()

print(f"\nTotal presentation folders found: {len(found_folders)}")
