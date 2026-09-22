import sys
import re
import zipfile
from pathlib import Path
from pptx import Presentation

sys.stdout.reconfigure(encoding='utf-8')

import xml.etree.ElementTree as ET

def sample_file(pptx_path):
    print("=" * 60)
    print("FILE:", pptx_path.name)
    try:
        with zipfile.ZipFile(pptx_path, 'r') as z:
            slide_names = sorted([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')],
                                 key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 0)
            for sname in slide_names[:4]:
                print(f"  [{sname}]")
                xml_data = z.read(sname)
                root = ET.fromstring(xml_data)
                texts = [elem.text for elem in root.iter() if elem.tag.endswith('}t') and elem.text and elem.text.strip()]
                for t in texts[:10]:
                    print(f"    - {t}")
    except Exception as e:
        print("ERROR:", e)

en_dir = Path(r"C:\Users\user\Desktop\Antigravity\Soff.uz\Tarjima_qilingan_slaydlar\Slaydlar_EN")
files_to_check = [
    "Simsiz aloqa va ko'p antennali tizimlar (1-ma'ruza).pptx",
    "Kompilyatorlar texnologiyasi va LLVM (1-qism Intro).pptx",
    "Amaliy va tabiiy fanlar - avtomobil kak istochnik zagryazneniya atmosfery.pptx",
    "Kompyuter ko'rishi va tasvirlarni tahlil qilish (1-ma'ruza).pptx",
    "Kiberxavfsizlik asoslari (1-ma'ruza).pptx"
]

for fname in files_to_check:
    fpath = en_dir / fname
    if fpath.exists():
        sample_file(fpath)
    else:
        print(f"Not found: {fname}")
