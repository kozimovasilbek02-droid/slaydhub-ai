import os
import requests
import json
import time

key = os.environ.get("GROQ_API_KEY")
url = "https://api.groq.com/openai/v1/chat/completions"
headers = {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}

items = [
    {"id": f"it_{i}", "text": f"Slide {i}: Comprehensive analysis of distributed computing systems and fault tolerance in cloud networks."}
    for i in range(25)
]

payload = {
    "model": "qwen/qwen3.8-27b",
    "messages": [
        {"role": "system", "content": 'Professional tarjimonsiz. O\'zbek tili (Lotin)ga aniq tarjima qiling. Qat\'iy JSON massiv qaytaring: [{"id": "...", "translated": "..."}]'},
        {"role": "user", "content": json.dumps(items, ensure_ascii=False, indent=2)}
    ],
    "temperature": 0.1
}

t0 = time.time()
r = requests.post(url, headers=headers, json=payload, timeout=30)
t1 = time.time()
print("Groq Status:", r.status_code, f"in {t1 - t0:.2f}s")
if r.status_code == 200:
    res = r.json()
    content = res["choices"][0]["message"]["content"]
    print("Groq content preview:", content[:150])

from google import genai
from google.genai import types
client = genai.Client()
t0 = time.time()
res = client.models.generate_content(
    model="gemini-3.6-flash",
    contents=json.dumps(items, ensure_ascii=False),
    config=types.GenerateContentConfig(
        system_instruction='Professional tarjimonsiz. O\'zbek tili (Lotin)ga aniq tarjima qiling. Qat\'iy JSON massiv qaytaring: [{"id": "...", "translated": "..."}]',
        response_mime_type="application/json"
    )
)
print("Gemini 3.6 flash:", f"{time.time() - t0:.2f}s")
print("Gemini content preview:", res.text[:150])
