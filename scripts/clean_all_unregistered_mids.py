# -*- coding: utf-8 -*-
import asyncio
import os
import sys
import json
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

from aiogram import Bot

REPO_ROOT = Path(r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot")
sys.path.insert(0, str(REPO_ROOT))

from core.config import config

async def clean_unregistered():
    bot = Bot(token=config.bot_token)
    ch = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "-1004476836723")

    with open(REPO_ROOT / "scripts" / "valid_message_ids.json", "r") as f:
        valid_ids = set(json.load(f))

    print(f"Loaded {len(valid_ids)} protected message IDs.", flush=True)

    probe = await bot.send_message(ch, "probe")
    latest_id = probe.message_id
    await bot.delete_message(ch, latest_id)

    print(f"Scanning from {latest_id - 1} down to 25000...", flush=True)
    deleted = 0
    checked = 0

    # Scan in chunks
    for mid in range(latest_id - 1, 25000, -1):
        if mid in valid_ids:
            continue
        try:
            await bot.delete_message(ch, mid)
            deleted += 1
            print(f"🧹 Deleted unregistered/stray message ID: {mid}", flush=True)
            await asyncio.sleep(0.04)
        except Exception:
            # Message already doesn't exist
            pass
        checked += 1

    print(f"\n🎉 COMPLETED! Scanned unregistered IDs. Deleted: {deleted} stray messages.", flush=True)
    await bot.session.close()

if __name__ == "__main__":
    asyncio.run(clean_unregistered())
