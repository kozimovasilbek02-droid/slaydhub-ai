import os, glob

# Check all possible locations
candidates = [
    r"G:\Мой диск\SlideEgg_Downloads\Prezentatsiyalar",
    r"G:\My Drive\SlideEgg_Downloads\Prezentatsiyalar",
    r"C:\Users\user\Desktop\SlideEgg_Downloads\Prezentatsiyalar",
    r"C:\Users\user\Desktop\5M_Of_Advertising",
]

for c in candidates:
    print(c, "->", os.path.exists(c))

# Search Desktop subdirectories for any Presentation folders
desktop = r"C:\Users\user\Desktop"
for item in os.listdir(desktop):
    full = os.path.join(desktop, item)
    if os.path.isdir(full):
        slides = glob.glob(os.path.join(full, "Slide_*.png"))
        if slides:
            print(f"Presentation on Desktop: {item} ({len(slides)} slides)")
