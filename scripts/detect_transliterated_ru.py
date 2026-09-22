import glob
import os
import re
import sys
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')

folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"
files = sorted(glob.glob(os.path.join(folder, "*.pptx")))

# Common Russian words transliterated into Latin (or in Cyrillic)
ru_latin_keywords = [
    r'\bchto\b', r'\bkak\b', r'\besli\b', r'\bkogda\b', r'\bpotomu\b', r'\btak\b',
    r'\bdlya\b', r'\bcherez\b', r'\bposle\b', r'\bmejdu\b', r'\bvsego\b', r'\bkotory\w*',
    r'\bochen\b', r'\bmojno\b', r'\bnelzya\b', r'\bvы\b', r'\bmы\b', r'\boni\b', r'\bona\b',
    r'\byavlyaetsya\b', r'\byavlyayutsya\b', r'\btakje\b', r'\btolko\b', r'\bdrug\w*',
    r'\bklass\b', r'\burok\b', r'\bzaryad\w*', r'\bopredelenie\b', r'\bsvoystv\w*',
    r'\bzadach\w*', r'\bsilы\b', r'\benergiy\w*', r'\bvremya\b', r'\btemperatur\w*'
]

ru_latin_regex = re.compile('|'.join(ru_latin_keywords), re.IGNORECASE)

print(f"Checking {len(files)} files for Russian words (both Cyrillic and Latin transliteration)...")

suspicious_files = []

for f in files:
    fname = os.path.basename(f)
    try:
        prs = Presentation(f)
        ru_matches = []
        for s_idx, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    matches = ru_latin_regex.findall(text)
                    # Filter out false positives in Uzbek (e.g. 'vaqt', etc.)
                    if len(matches) >= 3:
                        ru_matches.append((s_idx+1, matches[:5], text[:80]))
        if len(ru_matches) >= 2:
            suspicious_files.append((fname, len(ru_matches), ru_matches))
    except Exception as e:
        print(f"Error reading {fname}: {e}")

print(f"\nFound {len(suspicious_files)} files with significant Russian transliteration:")
for fname, count, matches in suspicious_files:
    print(f"\n- {fname} ({count} suspicious text blocks):")
    for s_idx, m, snippet in matches[:3]:
        snippet_clean = snippet.replace('\n', ' ')
        print(f"    Slide {s_idx} [matched: {m}]: {snippet_clean}")
