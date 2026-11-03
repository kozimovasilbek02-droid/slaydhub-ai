# -*- coding: utf-8 -*-
"""
scripts/batch_fix_hd_covers.py
Directly replaces all low-quality social media icons (Facebook, Twitter, etc.)
in Telegram messages in-place with true 1280x720 HD Slide 1 covers
generated via Microsoft PowerPoint COM.

ZERO temporary messages sent to channel!
Direct multipart upload via InputMediaPhoto(media=FSInputFile(...)).
"""

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
from aiogram.exceptions import TelegramBadRequest, TelegramRetryAfter

REPO_ROOT = Path(r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot")
sys.path.insert(0, str(REPO_ROOT))

from core.config import config

async def run_hd_replacements():
    bot = Bot(token=config.bot_token)
    ch = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "-1004476836723")

    targets_file = REPO_ROOT / "scripts" / "true_bad_icon_targets.json"
    if not targets_file.exists():
        print("targets_file not found!")
        return

    with open(targets_file, "r", encoding="utf-8") as f:
        targets = json.load(f)

    cat_path = REPO_ROOT / "templates_catalog.json"
    with open(cat_path, "r", encoding="utf-8") as f:
        catalog = json.load(f)

    print(f"Total target templates with valid message IDs: {len(targets)}", flush=True)

    temp_dir = Path(r"C:\Users\user\Desktop\Antigravity\Soff.uz\hd_covers")
    temp_dir.mkdir(parents=True, exist_ok=True)

    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    print("PowerPoint COM initialized successfully.", flush=True)

    fixed_count = 0
    skipped_count = 0
    already_good = 0

    for i, it in enumerate(targets):
        tpl_id = it["id"]
        local_path = it.get("local_path")
        mid = it["photo_mid"]
        title = it.get("title", "")

        if not local_path or not os.path.exists(local_path):
            skipped_count += 1
            continue

        out_jpg = str(temp_dir / f"hd_{mid}.jpg")

        # 1. Export Slide 1 in HD (1280x720)
        try:
            pres = ppt.Presentations.Open(os.path.abspath(local_path), ReadOnly=True, Untitled=False, WithWindow=False)
            pres.Slides(1).Export(out_jpg, "JPG", 1280, 720)
            pres.Close()
        except Exception as e:
            # If PowerPoint cannot open this specific file
            skipped_count += 1
            continue

        # 2. Edit message in Telegram in-place
        clean_title = Path(local_path).stem.rstrip(".")
        caption = f"🖼️ Muqova (HD): {clean_title}\nToifa: minimal"

        for attempt in range(3):
            try:
                res = await bot.edit_message_media(
                    chat_id=ch,
                    message_id=mid,
                    media=InputMediaPhoto(media=FSInputFile(out_jpg), caption=caption)
                )
                new_fid = res.photo[-1].file_id
                if tpl_id in catalog:
                    catalog[tpl_id]["thumb_file_id"] = new_fid
                fixed_count += 1
                print(f"[{i+1}/{len(targets)}] ✅ Fixed {tpl_id} (Msg {mid}) -> HD Cover ({clean_title[:30]})", flush=True)
                break
            except TelegramBadRequest as e:
                err = str(e).lower()
                if "not modified" in err:
                    already_good += 1
                    # Already updated earlier
                    break
                elif "message to edit not found" in err:
                    # Message was deleted or not accessible
                    break
                else:
                    print(f"[{i+1}/{len(targets)}] ⚠️ BadRequest for Msg {mid}: {e}", flush=True)
                    break
            except TelegramRetryAfter as e:
                print(f"Rate limited, sleeping {e.retry_after}s...", flush=True)
                await asyncio.sleep(e.retry_after + 1)
            except Exception as e:
                print(f"[{i+1}/{len(targets)}] ⚠️ Error for Msg {mid}: {e}", flush=True)
                await asyncio.sleep(0.5)

        # Periodic catalog save
        if fixed_count > 0 and fixed_count % 20 == 0:
            with open(cat_path, "w", encoding="utf-8") as f:
                json.dump(catalog, f, indent=2, ensure_ascii=False)
            print(f"💾 Synced catalog checkpoint ({fixed_count} fixed so far)...", flush=True)

        await asyncio.sleep(0.05)

    try:
        ppt.Quit()
    except Exception:
        pass

    # Final catalog save
    with open(cat_path, "w", encoding="utf-8") as f:
        json.dump(catalog, f, indent=2, ensure_ascii=False)

    print(f"\n🎉 ALL DONE! Fixed: {fixed_count}, Already good: {already_good}, Skipped: {skipped_count}", flush=True)
    await bot.session.close()

if __name__ == "__main__":
    asyncio.run(run_hd_replacements())
