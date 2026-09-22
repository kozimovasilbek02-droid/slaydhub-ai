import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r'C:\Users\user\Desktop\Antigravity\Soff.uz\audit_01_tayyor_sara.json', encoding='utf-8') as f:
    data = json.load(f)

print('=== 13 FILES WITH CYRILLIC LEAKS ===')
for d in data:
    if d['status'] == 'CYRILLIC_LEAK':
        print(f"{d['filename']} (Cyr chars: {d['cyrillic_chars']})")
        for s in d['cyr_samples'][:3]:
            print('    [Leak]:', s)

print('\n=== WATERMARKS FOUND ===')
for d in data:
    if d['watermarks']:
        print(f"\n{d['filename']}:")
        for wm in d['watermarks']:
            print(f"    Slide {wm['slide']} [{wm['keyword']}]: {wm['snippet']}")
