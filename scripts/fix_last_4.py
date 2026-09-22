import os, sys, json, requests, re
from pptx import Presentation
import json_repair

sys.stdout.reconfigure(encoding='utf-8')

folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"
target_4 = [
    "Fizika - Ichki energiya va uni o'zgartirish usullari (10-sinf).pptx",
    "Fizika - Moddiy nuqta va sanoq sistemasi (9-sinf).pptx",
    "Geografiya - Atmosfera bosimi va barometrlar (7-sinf).pptx",
    "Menejment - Nizolar va stresslarni boshqarish.pptx"
]

api_key = os.environ.get('GROQ_API_KEY')
url = 'https://api.groq.com/openai/v1/chat/completions'
headers = {
    'Authorization': f'Bearer {api_key}',
    'Content-Type': 'application/json'
}

def translate_paragraphs_list(p_list):
    prompt = """Siz professional akademik tarjimonsiz. Ushbu taqdimot matnlarini toza, tabiiy va ravon O'ZBEK tiliga (Lotin alifbosida) tarjima qiling.
Har qanday ruscha yoki lotin-ruscha transliteratsiyani to'liq o'zbek tiliga o'giring (masalan: 'ITOGI UROKA' -> 'Dars yakunlari', 'DOMAShNEE ZADANIE' -> 'Uyga vazifa', 'Chto my izuchili v 8-om klasse?' -> '8-sinfda nimani o\\'rgangan edik?', 'Vpervye vesomost vozduxa' -> 'Havoning og\\'irligi birinchi marta...', 'Faza trevogi' -> 'Xavotir fazasi', 'Kakie byvayut fazy stressa?' -> 'Stressning qanday fazalari bor?', 'Snijaetsya kontsentratsiya vnimaniya' -> 'Diqqat jamlanishi pasayadi').
Formulalar va raqamlarni aynan saqlang.

Natijani FAQATGINA JSON massiv ko'rinishida bering:
[
  {"id": 0, "uzbek": "Tarjima"}
]"""
    payload = {
        'model': 'qwen/qwen3.8-27b',
        'messages': [
            {'role': 'system', 'content': prompt},
            {'role': 'user', 'content': json.dumps([{"id": i, "text": t} for i, t in enumerate(p_list)], ensure_ascii=False)}
        ],
        'temperature': 0.1
    }
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=20)
        if r.status_code == 200:
            raw_text = r.json()['choices'][0]['message']['content'].strip()
            raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
            raw_text = re.sub(r"\s*```$", "", raw_text)
            parsed = json_repair.repair_json(raw_text, return_objects=True)
            res = {}
            if isinstance(parsed, list):
                for row in parsed:
                    if isinstance(row, dict) and 'id' in row:
                        val = row.get('uzbek') or row.get('translated') or row.get('text') or ''
                        res[int(row['id'])] = val
            return res
    except Exception as e:
        print("Error:", e)
    return {}

for fname in target_4:
    fpath = os.path.join(folder, fname)
    prs = Presentation(fpath)
    
    p_elements = []
    
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    t = p.text.strip()
                    if t:
                        p_elements.append((p, t))
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for p in cell.text_frame.paragraphs:
                                t = p.text.strip()
                                if t:
                                    p_elements.append((p, t))
                                    
    print(f"\nProcessing {fname} ({len(p_elements)} total paragraphs)...")
    chunk_size = 20
    for i in range(0, len(p_elements), chunk_size):
        chunk = p_elements[i:i+chunk_size]
        text_list = [item[1] for item in chunk]
        tr_map = translate_paragraphs_list(text_list)
        for idx, (p, orig) in enumerate(chunk):
            tr = tr_map.get(idx)
            if tr and tr.strip():
                clean_tr = tr.replace('З', 'Z').replace('з', 'z').replace('ы', 'y')
                p.text = clean_tr
                
    prs.save(fpath)
    print(f"Saved {fname} 100% clean!")

print("\nDone!")
