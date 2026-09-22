import os

desktop_base = os.path.join(r"C:\Users\user\Desktop", "SlideEgg_Downloads", "Prezentatsiyalar")
if os.path.exists(desktop_base):
    folders = sorted([f for f in os.listdir(desktop_base) if os.path.isdir(os.path.join(desktop_base, f))])
    print(f"Found {len(folders)} presentations in Desktop SlideEgg_Downloads!")
    for i, f in enumerate(folders, 1):
        print(f"[{i}] {f}")
else:
    print(f"Desktop path not found: {desktop_base}")
