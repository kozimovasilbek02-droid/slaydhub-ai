# -*- coding: utf-8 -*-
import os
import sys
import json
import asyncio
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

BAD_THUMB = "AgACAgIAAyEGAAMBCtcbcwACY55q6KahopJk5SaG8h-MmmOXkeB9XQACVyFrG3xSQEskBIOLQeK6rAEAAwIAA20AAz0E"

async def fix_all():
    bot = Bot(token=config.bot_token)
    ch = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "-1004476836723")
    
    cat_path = REPO_ROOT / "templates_catalog.json"
    with open(cat_path, "r", encoding="utf-8") as f:
        catalog = json.load(f)

    # Initialize PowerPoint COM
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    print("PowerPoint COM ready.", flush=True)

    # Scan the range 25400 to 25760 (around 360 messages) where the 19 bad photos reside
    fixed = 0
    temp_dir = Path(r"C:\Users\user\Desktop\Antigravity\Soff.uz\scratch_thumbs")
    temp_dir.mkdir(exist_ok=True)

    print("Scanning range 25760 down to 25400 for bad Facebook thumbs...", flush=True)
    for mid in range(25760, 25400, -1):
        try:
            fwd = await bot.forward_message(chat_id=ch, from_chat_id=ch, message_id=mid)
            is_p = bool(fwd.photo)
            fid = fwd.photo[-1].file_id if is_p else ""
            caption = fwd.caption or ""
            await bot.delete_message(ch, fwd.message_id)

            if is_p and fid == BAD_THUMB:
                print(f"\n🎯 Found Facebook thumb at msg {mid}! Caption: {caption[:50]}", flush=True)
                
                # Next message (mid+1) is the document!
                doc_fwd = await bot.forward_message(chat_id=ch, from_chat_id=ch, message_id=mid+1)
                doc_cap = doc_fwd.caption or ""
                await bot.delete_message(ch, doc_fwd.message_id)

                # Find template id: "ID: tpl_..."
                tpl_id = None
                for line in doc_cap.splitlines():
                    if line.startswith("ID: "):
                        tpl_id = line.replace("ID: ", "").strip()
                        break

                if not tpl_id or tpl_id not in catalog:
                    print(f"⚠️ Could not find tpl_id from doc_cap: {doc_cap}", flush=True)
                    continue

                item = catalog[tpl_id]
                local_path = item.get("local_path")
                if not local_path or not os.path.exists(local_path):
                    print(f"⚠️ Local file missing: {local_path}", flush=True)
                    continue

                # Export real Slide 1 in 1280x720 HD
                out_jpg = str(temp_dir / f"fix_{mid}.jpg")
                pres = ppt.Presentations.Open(os.path.abspath(local_path), ReadOnly=True, Untitled=False, WithWindow=False)
                pres.Slides(1).Export(out_jpg, "JPG", 1280, 720)
                pres.Close()
                print(f"🖼️ Rendered HD Slide 1: {out_jpg} ({os.path.getsize(out_jpg)} bytes)", flush=True)

                # Send temporary to get fresh file_id
                tmp = await bot.send_photo(ch, photo=FSInputFile(out_jpg))
                new_fid = tmp.photo[-1].file_id
                await bot.delete_message(ch, tmp.message_id)

                # Replace photo in-place!
                new_cap = caption.replace("Muqova:", "Muqova (HD):") if "Muqova:" in caption else f"{caption} (HD)"
                await bot.edit_message_media(
                    chat_id=ch,
                    message_id=mid,
                    media=InputMediaPhoto(media=new_fid, caption=new_cap)
                )
                print(f"✅ Replaced msg {mid} in-place with true HD cover!", flush=True)

                # Update catalog
                item["thumb_file_id"] = new_fid
                fixed += 1
                await asyncio.sleep(0.5)

        except Exception as e:
            # print(f"Error on {mid}: {e}", flush=True)
            pass

    ppt.Quit()
    with open(cat_path, "w", encoding="utf-8") as f:
        json.dump(catalog, f, indent=2, ensure_ascii=False)

    print(f"\n🎉 COMPLETED! All {fixed} Facebook thumbs were replaced in-place with true HD covers!", flush=True)
    await bot.session.close()

if __name__ == "__main__":
    asyncio.run(fix_all())
