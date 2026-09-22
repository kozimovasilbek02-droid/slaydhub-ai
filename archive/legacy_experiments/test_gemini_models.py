# -*- coding: utf-8 -*-
import os
import sys
from google import genai
from google.genai import types

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from bot.config import GEMINI_API_KEY

api_key = GEMINI_API_KEY or os.environ.get("GEMINI_API_KEY")
print(f"Using API Key: {api_key[:8]}...{api_key[-4:] if api_key else 'None'}")

client = genai.Client(api_key=api_key)

candidate_models = [
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-2.5-pro",
    "gemini-1.5-pro",
    "gemini-2.0-flash-exp",
    "gemini-3.6-flash",
    "gemini-3.1-pro-preview",
    "gemini-3.1-flash-lite"
]

print("\n--- Testing Models ---")
for m in candidate_models:
    try:
        resp = client.models.generate_content(
            model=m,
            contents='Return a valid JSON object: {"status": "ok"}',
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        print(f"  [SUCCESS] Model '{m}': {resp.text.strip()}")
    except Exception as e:
        print(f"  [FAILED] Model '{m}': {e}")
