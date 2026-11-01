# -*- coding: utf-8 -*-
import asyncio
import os
import sys
import json
from pathlib import Path

REPO_ROOT = Path(r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot")
sys.path.insert(0, str(REPO_ROOT))

from aiogram import Bot
from core.config import config

async def clean_orphans():
    bot = Bot(token=config.bot_token)
    channel_id = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "-1004476836723")

    scan_file = REPO_ROOT / "scripts" / "channel_tail_scan.json"
    if not scan_file.exists():
        print("Scan file not found.")
        await bot.session.close()
        return

    with open(scan_file, "r", encoding="utf-8") as f:
        msgs = json.load(f)

    # Sort messages ascending by ID
    msgs = sorted(msgs, key=lambda m: m["id"])

    # Load verified catalog
    catalog_path = REPO_ROOT / "templates_catalog.json"
    catalog = {}
    if catalog_path.exists():
        with open(catalog_path, "r", encoding="utf-8") as f:
            catalog = json.load(f)

    # All verified pptx_file_id and thumb_file_id
    verified_thumbs = {item.get("thumb_file_id") for item in catalog.values() if item.get("thumb_file_id")}

    # Find orphan photos:
    # A message is orphan if it's a photo and its caption is "Muqova: ..." but the next message is NOT a document for this template,
    # or if its caption matches files that failed document upload.
    # In msgs list (sorted ascending by id):
    orphan_ids = []
    for i, m in enumerate(msgs):
        if m.get("is_photo"):
            # Check if immediately followed by document or preceding document
            has_paired_doc = False
            # Check next message
            if i + 1 < len(msgs) and msgs[i+1].get("is_doc"):
                has_paired_doc = True
            # Check previous message
            if i - 1 >= 0 and msgs[i-1].get("is_doc"):
                has_paired_doc = True
                
            if not has_paired_doc:
                orphan_ids.append(m["id"])

    print(f"Total orphan photo messages identified: {len(orphan_ids)}")
    print(f"Orphan IDs: {orphan_ids}")

    deleted_count = 0
    for mid in orphan_ids:
        try:
            await bot.delete_message(chat_id=channel_id, message_id=mid)
            deleted_count += 1
            print(f"Deleted orphan photo message ID: {mid}")
            await asyncio.sleep(0.1)
        except Exception as e:
            print(f"Failed to delete {mid}: {e}")

    print(f"Successfully deleted {deleted_count} orphan photos from channel.")
    await bot.session.close()

if __name__ == "__main__":
    asyncio.run(clean_orphans())
