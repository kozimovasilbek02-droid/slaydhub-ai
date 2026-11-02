# -*- coding: utf-8 -*-
"""
scripts/fix_facebook_thumbnails.py
Detects all messages in Telegram channel with the bad Facebook logo thumbnail
and replaces them in-place with true 1280x720 HD Slide 1 exports using PowerPoint COM.
"""

import os
import sys
import json
import asyncio
import tempfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import win32com.client
from aiogram import Bot
from aiogram.types import FSInputFile, InputMediaPhoto

REPO_ROOT = Path(r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot")
sys.path.insert(0, str(REPO_ROOT))

from core.config import config

BAD_THUMB_FILE_ID = "AgACAgIAAyEGAAMBCtcbcwACY55q6KahopJk5SaG8h-MmmOXkeB9XQACVyFrG3xSQEskBIOLQeK6rAEAAwIAA20AAz0E"

async def fix_thumbnails():
    bot = Bot(token=config.bot_token)
    channel_id = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "-1004476836723")

    catalog_path = REPO_ROOT / "templates_catalog.json"
    with open(catalog_path, "r", encoding="utf-8") as f:
        catalog = json.load(f)

    # Map of template titles/stems to catalog item
    title_map = {}
    for item in catalog.values():
        if item.get("filename"):
            stem = Path(item["filename"]).stem.strip().lower()
            title_map[stem] = item

    # Initialize PowerPoint COM
    ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
    print("PowerPoint COM initialized.")

    probe = await bot.send_message(channel_id, "🔍 Scanning...")
    latest_id = probe.message_id
    await bot.delete_message(channel_id, latest_id)

    print(f"Scanning messages from {latest_id - 1} down to 24000...")
    fixed_count = 0

    with tempfile.TemporaryDirectory() as td:
        for mid in range(latest_id - 1, 24000, -1):
            try:
                # Forward to inspect message
                fwd = await bot.forward_message(chat_id=channel_id, from_chat_id=channel_id, message_id=mid)
                is_photo = bool(fwd.photo)
                caption = fwd.caption or ""
                photo_file_id = fwd.photo[-1].file_id if fwd.photo else None
                # Immediately delete forward copy
                await bot.delete_message(channel_id, fwd.message_id)

                if is_photo and photo_file_id == BAD_THUMB_FILE_ID:
                    print(f"\n🚨 Found Facebook logo thumbnail at Msg ID: {mid} | Caption: {caption}")
                    
                    # Extract template name from caption: "Muqova: <name>\nToifa: ..."
                    first_line = caption.split("\n")[0].replace("🖼️", "").replace("Muqova:", "").strip()
                    first_line_clean = first_line.lower().rstrip(".")
                    
                    # Find matching item in catalog
                    matched_item = None
                    for stem, item in title_map.items():
                        if stem.startswith(first_line_clean[:20]) or first_line_clean[:20] in stem:
                            matched_item = item
                            break

                    if not matched_item:
                        print(f"⚠️ Could not match item for: '{first_line}'")
                        continue

                    local_pptx = matched_item.get("local_path")
                    if not local_pptx or not os.path.exists(local_pptx):
                        print(f"⚠️ Local PPTX not found: {local_pptx}")
                        continue

                    # Render true Slide 1 in 1280x720 HD via PowerPoint COM
                    out_jpg = os.path.abspath(os.path.join(td, f"slide1_{mid}.jpg"))
                    pres = ppt_app.Presentations.Open(str(local_pptx), ReadOnly=True, Untitled=False, WithWindow=False)
                    pres.Slides(1).Export(out_jpg, "JPG", 1280, 720)
                    pres.Close()

                    print(f"🎨 Rendered true HD Slide 1: {out_jpg} ({os.path.getsize(out_jpg)} bytes)")

                    # Edit Telegram photo message in-place!
                    new_media = InputMediaPhoto(media=FSInputFile(out_jpg), caption=caption)
                    res = await bot.edit_message_media(chat_id=channel_id, message_id=mid, media=new_media)
                    new_thumb_id = res.photo[-1].file_id if res.photo else None
                    print(f"✅ Replaced Telegram photo at Msg ID {mid} in-place! New file_id: {new_thumb_id[:20]}...")

                    # Update catalog
                    if new_thumb_id:
                        matched_item["thumb_file_id"] = new_thumb_id
                        with open(catalog_path, "w", encoding="utf-8") as f:
                            json.dump(catalog, f, indent=2, ensure_ascii=False)

                    fixed_count += 1
                    await asyncio.sleep(0.5)

            except Exception as e:
                # ignore deleted or missing message errors
                if "message to forward not found" not in str(e).lower():
                    pass

    ppt_app.Quit()
    await bot.session.close()
    print(f"\n🎉 Finished! Total bad Facebook thumbnails replaced with HD covers: {fixed_count}")

if __name__ == "__main__":
    asyncio.run(fix_thumbnails())
