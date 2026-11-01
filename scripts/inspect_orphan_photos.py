import asyncio
import os
import sys
from pathlib import Path

REPO_ROOT = Path(r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot")
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from aiogram import Bot
from core.config import config

async def inspect():
    bot = Bot(token=config.bot_token)
    channel_id = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "-1004476836723")
    
    probe = await bot.send_message(channel_id, "🔍 Inspecting...")
    latest_id = probe.message_id
    await bot.delete_message(channel_id, latest_id)
    print(f"Latest ID: {latest_id}")
    
    # We want to scan back from latest_id - 1 down to 9250
    messages = []
    
    for mid in range(latest_id - 1, 9250, -1):
        try:
            fwd = await bot.forward_message(chat_id=channel_id, from_chat_id=channel_id, message_id=mid)
            is_photo = bool(fwd.photo)
            is_doc = bool(fwd.document)
            cap = fwd.caption or ""
            text = fwd.text or ""
            # Delete forwarded copy immediately
            await bot.delete_message(channel_id, fwd.message_id)
            
            messages.append({
                "id": mid,
                "is_photo": is_photo,
                "is_doc": is_doc,
                "caption": cap,
                "text": text
            })
            print(f"[{mid}] Photo: {is_photo}, Doc: {is_doc} | Cap: {cap[:50]}")
            await asyncio.sleep(0.08)
        except Exception as e:
            # message doesn't exist
            pass

    print(f"\nTotal messages analyzed: {len(messages)}")
    import json
    with open(REPO_ROOT / "scripts" / "channel_tail_scan.json", "w", encoding="utf-8") as f:
        json.dump(messages, f, indent=2, ensure_ascii=False)

    await bot.session.close()

if __name__ == "__main__":
    asyncio.run(inspect())
