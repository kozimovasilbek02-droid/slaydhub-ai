import os
import glob
import sys
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')

folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"
files = sorted([f for f in glob.glob(os.path.join(folder, "*.pptx")) if not os.path.basename(f).startswith("~$")])

bug_count = 0
for f in files:
    prs = Presentation(f)
    for s_idx, s in enumerate(prs.slides):
        for shape in s.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text
                if len(txt) >= 5 and len(txt) % 2 == 1 and all(c == "'" for c in txt[::2]):
                    print(f"Bug found in {os.path.basename(f)} Slide {s_idx+1}: {repr(txt[:60])}")
                    bug_count += 1
                lines = txt.split('\n')
                for line in lines:
                    if len(line) >= 5 and len(line) % 2 == 1 and all(c == "'" for c in line[::2]):
                        print(f"Bug line in {os.path.basename(f)} Slide {s_idx+1}: {repr(line[:60])}")
                        bug_count += 1

print(f"\nAUDIT VERIFICATION RESULT:")
print(f"Total files verified: {len(files)}")
print(f"Total remaining bugged paragraphs: {bug_count}")
