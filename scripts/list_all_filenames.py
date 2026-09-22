import os
import glob
import sys

sys.stdout.reconfigure(encoding='utf-8')

folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"
files = sorted([os.path.basename(f) for f in glob.glob(os.path.join(folder, "*.pptx"))])

for idx, f in enumerate(files, 1):
    print(f"{idx}. {f}")
