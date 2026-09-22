import os, requests, json, sys

sys.stdout.reconfigure(encoding='utf-8')
api_key = os.environ.get('GROQ_API_KEY')
url = 'https://api.groq.com/openai/v1/chat/completions'
headers = {'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'}
batch = [
    {'id': 0, 'text': 'ЗAROJDENIE FIЗIKI Elementarnыx chastits'},
    {'id': 1, 'text': 'Rtutnыy barometr'},
    {'id': 2, 'text': 'Vыsota perevyornutoy trubki = 1 m'},
    {'id': 3, 'text': 'Snijaetsya kontsentratsiya vnimaniya'},
    {'id': 4, 'text': 'SPASIBO ZA VNIMANIE !'}
]
payload = {
    'model': 'qwen/qwen3.8-27b',
    'messages': [
        {'role': 'system', 'content': 'Har bir matnni toza, ravon va ilmiy O\'ZBEK tiliga (Lotin alifbosida) tarjima qiling. Natijani FAQAT JSON massiv shaklida bering:\n[{"id": 0, "uzbek": "Tarjima"}]'},
        {'role': 'user', 'content': json.dumps(batch, ensure_ascii=False, indent=2)}
    ],
    'temperature': 0.1
}
r = requests.post(url, headers=headers, json=payload)
print(r.status_code)
print(r.json()['choices'][0]['message']['content'])
