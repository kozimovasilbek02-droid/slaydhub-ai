import shutil, os

src = r"c:\Users\user\Desktop\Antigravity\Power Point\skills\slide-layout-intelligence\SKILL.md"
dst = r"c:\Users\user\Desktop\Antigravity\Power Point\skills\clean-canvas-slide-engineering\SKILL.md"
shutil.copyfile(src, dst)
print("Skills synchronized successfully.")
