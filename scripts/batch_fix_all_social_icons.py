# -*- coding: utf-8 -*-
"""
scripts/batch_fix_all_social_icons.py
Automated batch fixer for all SlidesMania and web-based templates:
Replaces any social media icon (Facebook, Twitter bird, Instagram, etc.)
with the TRUE 1280x720 HD Slide 1 cover generated directly via PowerPoint COM!
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

async def batch_fix():
    bot = Bot(token=config.bot_token)
    ch = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "-1004476836723")

    bad_json = REPO_ROOT / "scripts" / "all_bad_icon_templates.json"
    if not bad_json.exists():
        print("all_bad_icon_templates.json not found yet, checking slidesmania_decoded_msgs.json...")
        bad_json = REPO_ROOT / "scripts" / "slidesmania_decoded_msgs.json"

    with open(bad_json, "r", encoding="utf-8") as f:
        targets = json.load(f)

    cat_path = REPO_ROOT / "templates_catalog.json"
    with open(cat_path, "r", encoding="utf-8") as f:
        catalog = json.load(f)

    print(f"Loaded {len(targets)} candidate templates to check/fix.", flush=True)

    # PowerPoint COM initialization
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    print("PowerPoint COM initialized successfully.", flush=True)

    temp_dir = Path(r"C:\Users\user\Desktop\Antigravity\Soff.uz\hd_covers")
    temp_dir.mkdir(exist_ok=True)

    fixed_count = 0
    skipped_count = 0

    for i, it in enumerate(targets):
        tpl_id = it["id"]
        local_path = it.get("local_path")
        photo_mid = it.get("photo_mid")
        filename = it.get("filename", "")

        if not local_path or not os.path.exists(local_path):
            continue

        out_jpg = str(temp_dir / f"hd_{photo_mid}.jpg")

        try:
            # 1. Export Slide 1 in HD using PowerPoint COM
            pres = ppt.Presentations.Open(os.path.abspath(local_path), ReadOnly=True, Untitled=False, WithWindow=False)
            pres.Slides(1).Export(out_jpg, "JPG", 1280, 720)
            pres.Close()

            # 2. Upload temporary to get fresh file_id
            tmp = await bot.send_photo(ch, photo=FSInputFile(out_jpg))
            new_fid = tmp.photo[-1].file_id
            await bot.delete_message(ch, tmp.message_id)

            # 3. Edit Telegram message in-place
            clean_stem = Path(filename).stem.rstrip(".")
            caption = f"🖼️ Muqova (HD): {clean_stem}\nToifa: minimal"

            await bot.edit_message_media(
                chat_id=ch,
                message_id=photo_mid,
                media=InputMediaPhoto(media=new_fid, caption=caption)
            )

            # 4. Update catalog
            if tpl_id in catalog:
                catalog[tpl_id]["thumb_file_id"] = new_fid

            fixed_count += 1
            print(f"[{fixed_count}/{len(targets)}] ✅ Replaced Msg {photo_mid} ({clean_stem[:35]}) with true HD cover!", flush=True)

            if fixed_count % 15 == 0:
                with open(cat_path, "w", encoding="utf-8") as f:
                    json.dump(catalog, f, indent=2, ensure_ascii=False)

            await asyncio.sleep(0.4)

        except Exception as e:
            # print(f"Skipped {photo_mid}: {e}", flush=True)
            skipped_count += 1

    ppt.Quit()
    with open(cat_path, "w", encoding="utf-8") as f:
        json.dump(catalog, f, indent=2, ensure_ascii=False)

    print(f"\n🎉 ALL COMPLETED! Fixed: {fixed_count}, Skipped: {skipped_count}")
    await bot.session.close()

if __name__ == "__main__":
    asyncio.run(batch_fix())
