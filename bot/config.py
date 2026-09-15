# -*- coding: utf-8 -*-
import os
from pathlib import Path
from dotenv import load_dotenv

# Loyiha asosiy katalogi
BASE_DIR = Path(__file__).resolve().parent.parent

# .env faylini yuklash
load_dotenv(BASE_DIR / ".env")

# Bot sozlamalari
BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# Vaqtinchalik va saqlash kataloglari
TEMP_DIR = BASE_DIR / "temp_sessions" / "telegram_bot"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_DIR = BASE_DIR / "output" / "bot_generated"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Shablonlar joylashuvi
TEMPLATES_DIRS = [
    BASE_DIR / "downloaded_templates",
    BASE_DIR / "output" / "presentations",
    Path(r"C:\Users\user\Desktop\slayd"),
]

# Soff.uz sessiya fayli
SOFF_SESSION_FILE = Path(r"C:\Users\user\.gemini\antigravity\scratch\soff_seller_agent\soff_session.json")
