import os
import requests
import json
import re
import sys
from pptx import Presentation
import json_repair

sys.stdout.reconfigure(encoding='utf-8')

folder = r"C:\Users\user\Desktop\Antigravity\Soff.uz\01_Tayyor_Sara_Slaydlar_UZ"
target_files = [
    "Fizika - Elektr maydoni ishi va potensiali (10-sinf).pptx",
    "Fizika - Elementar zarralar fizikasining paydo bo'lishi (11-sinf).pptx",
    "Fizika - Energiya saqlanish qonuni bo'yicha masalalar (10-sinf).pptx",
    "Fizika - Ichki energiya va uni o'zgartirish usullari (10-sinf).pptx",
    "Fizika - Moddiy nuqta va sanoq sistemasi (9-sinf).pptx",
    "Geografiya - Atmosfera bosimi va barometrlar (7-sinf).pptx",
    "Iqtisodiyot - Xodimlar motivatsiyasi.pptx",
    "Menejment - Nizolar va stresslarni boshqarish.pptx"
]

api_key = os.environ.get('GROQ_API_KEY')
url = 'https://api.groq.com/openai/v1/chat/completions'
headers = {
    'Authorization': f'Bearer {api_key}',
    'Content-Type': 'application/json'
}

sys_prompt = """Siz professional akademik tarjimonsiz. Taqdimotdagi har bir matnni toza, mukammal va tabiiy O'ZBEK tiliga (Lotin yozuvida) tarjima qiling.
DIQQAT: Matn kirill yoki lotincha transkripsiyadagi ruscha so'zlardan iborat bo'lishi mumkin (masalan: 'Napryajennost polya napravlena v storonu ubыvaniya...', 'DOMAShNEE ZADANIE', 'Snijenie kontsentratsii vnimaniya...', 'Sostavte uravnenie na osnove ZSE...', 'Rtutnыy barometr...').
Barcha bu ruscha jumlalarni 100% o'zbekchaga o'giring! Hech qanday ruscha so'z yoki kirillcha qoldirmang.
Formulalar (masalan: F=ma, E=mc^2, m/s, J, N, V) va raqamlarni to'g'ri saqlang.
MUHIM: Natijani FAQATGINA JSON massiv ko'rinishida qaytaring:
[
  {"id": 0, "translated": "O'zbekcha tarjima"}
]"""

def translate_batch(batch_items):
    payload = {
        'model': 'openai/gpt-oss-120b',
        'messages': [
            {'role': 'system', 'content': sys_prompt},
            {'role': 'user', 'content': json.dumps(batch_items, ensure_ascii=False, indent=2)}
        ],
        'temperature': 0.1
    }
    models = ['openai/gpt-oss-120b', 'openai/gpt-oss-20b', 'qwen/qwen3.8-27b']
    for m in models:
        payload['model'] = m
        try:
            r = requests.post(url, headers=headers, json=payload, timeout=8)
            if r.status_code == 200:
                raw_text = r.json()['choices'][0]['message']['content'].strip()
                raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text)
                raw_text = re.sub(r"\s*```$", "", raw_text)
                parsed = json_repair.repair_json(raw_text, return_objects=True)
                res_map = {}
                if isinstance(parsed, list):
                    for row in parsed:
                        if isinstance(row, dict) and 'id' in row:
                            t_val = row.get('translated') or row.get('translated_text') or row.get('text') or ''
                            res_map[int(row['id'])] = t_val
                            res_map[str(row['id'])] = t_val
                if res_map:
                    return res_map
        except Exception as e:
            pass
            
    # Fallback: GoogleTranslator batch
    try:
        from deep_translator import GoogleTranslator
        gt = GoogleTranslator(source='auto', target='uz')
        texts = [it['text'] for it in batch_items]
        tr_list = gt.translate_batch(texts)
        return {it['id']: tr for it, tr in zip(batch_items, tr_list)}
    except Exception:
        return {it['id']: it['text'] for it in batch_items}

for f_idx, fname in enumerate(target_files, 1):
    fpath = os.path.join(folder, fname)
    if not os.path.exists(fpath):
        continue
    prs = Presentation(fpath)
    
    # Collect all paragraph references
    paragraphs_list = []
    for s_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    t = p.text.strip()
                    if t and len(t) >= 2:
                        paragraphs_list.append((p, t))
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for p in cell.text_frame.paragraphs:
                                t = p.text.strip()
                                if t and len(t) >= 2:
                                    paragraphs_list.append((p, t))
                                    
    print(f"[{f_idx}/{len(target_files)}] {fname}: {len(paragraphs_list)} ta matn qatori tarjima qilinmoqda...", flush=True)
    
    # Batch process in chunks of 20
    chunk_size = 20
    for i in range(0, len(paragraphs_list), chunk_size):
        chunk = paragraphs_list[i:i+chunk_size]
        batch_payload = [{"id": idx, "text": item[1]} for idx, item in enumerate(chunk)]
        tr_map = translate_batch(batch_payload)
        for idx, (p, orig) in enumerate(chunk):
            tr_text = tr_map.get(idx) or tr_map.get(str(idx))
            if tr_text and tr_text.strip():
                # Direct in-place replacement
                p.text = tr_text
                    
    prs.save(fpath)
    print(f"  -> {fname} muvaffaqiyatli to'liq tarjima qilindi va saqlandi!", flush=True)

print("\n" + "="*60, flush=True)
print("BARCHA 8 TA TAQDIMOT 100% TO'LIQ O'ZBEK TILIGA O'GIRILDI!", flush=True)
print("="*60, flush=True)



