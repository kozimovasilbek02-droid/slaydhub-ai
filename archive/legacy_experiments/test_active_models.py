# -*- coding: utf-8 -*-
import os
import sys
from google import genai
from google.genai import types

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from bot.config import GEMINI_API_KEY
client = genai.Client(api_key=GEMINI_API_KEY)

models_to_test = [
    "gemini-3.1-flash-lite",
    "gemini-3.5-flash-lite",
    "gemini-3.7-flash",
    "gemini-3.8-flash",
    "gemini-3.5-flash",
    "gemini-flash-lite-latest",
    "gemini-flash-latest",
    "gemini-2.5-flash-lite"
]

print("--- Testing Active Responsive Models ---")
for m in models_to_test:
    try:
        resp = client.models.generate_content(
            model=m,
            contents='Return a valid JSON object: {"status": "ready"}',
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        print(f"  [ACTIVE] {m}: {resp.text.strip()}")
    except Exception as e:
        print(f"  [FAILED] {m}: {e}")
