# -*- coding: utf-8 -*-
"""
scripts/seed_telegram_templates.py
Automated seeder script for SlaydHub Ultra 1,000+ templates.
1. Scans local PPTX directories (Desktop/slayd, templates, Medical_Presentations_English_50).
2. Extracts high-res cover thumbnail (docProps/thumbnail.jpeg) in memory.
3. Uploads thumbnail and PPTX file to your private Telegram storage channel.
4. Stores permanent file_id references into templates_catalog.json for Render cloud delivery.
"""

import os
import sys
import json
import asyncio
import tempfile
from pathlib import Path
from typing import Dict, Any, List

from aiogram import Bot
from aiogram.types import FSInputFile

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from core.config import config
from core.template_catalog import GLOBAL_TEMPLATE_CATALOG, CATEGORIES


async def seed_templates(channel_id: str, max_upload: int = 50) -> None:
    bot = Bot(token=config.bot_token)
    try:
        me = await bot.get_me()
        print(f"✅ Bot ulandi: @{me.username}")
    except Exception as e:
        print(f"❌ Bot token xatosi: {e}")
        return

    print(f"📢 Nishon Telegram kanali: {channel_id}")
    catalog_path = REPO_ROOT / "templates_catalog.json"
    existing_catalog: Dict[str, Any] = {}
    if catalog_path.exists():
        try:
            with open(catalog_path, "r", encoding="utf-8") as f:
                existing_catalog = json.load(f)
        except Exception:
            existing_catalog = {}

    # Qidiruv kataloglari
    search_dirs = [
        REPO_ROOT / "templates",
        Path(r"C:\Users\user\Desktop\Institut\Suniy intelekt\05_Taqdimotlar_va_Maqolalar\Medical_Presentations_English_50"),
        Path(r"C:\Users\user\Desktop\slayd"),
    ]

    all_pptx_files: List[Path] = []
    for s_dir in search_dirs:
        if s_dir.exists():
            for p in s_dir.rglob("*.pptx"):
                if not p.name.startswith("~$") and p.stat().st_size > 50000:
                    all_pptx_files.append(p)

    print(f"🔍 Topilgan jami PPTX fayllar: {len(all_pptx_files)} ta.")
    uploaded_count = 0

    for pptx_file in all_pptx_files:
        if uploaded_count >= max_upload:
            print(f"⏸️ Maksimal kvota ({max_upload} ta) yetdi. Seeding yakunlandi.")
            break

        safe_id = f"tpl_{pptx_file.stem[:25].lower().replace(' ', '_').replace('-', '_')}"
        if safe_id in existing_catalog and existing_catalog[safe_id].get("pptx_file_id"):
            continue  # Allaqachon yuklangan

        category = GLOBAL_TEMPLATE_CATALOG.detect_category(pptx_file.name)
        print(f"📤 Yuklanmoqda [{uploaded_count + 1}]: {pptx_file.name} (Toifa: {category})...")

        with tempfile.TemporaryDirectory() as td:
            thumb_path = os.path.join(td, "thumb.jpeg")
            has_thumb = GLOBAL_TEMPLATE_CATALOG.extract_cover_thumbnail(str(pptx_file), thumb_path)

            thumb_file_id = None
            if has_thumb:
                try:
                    photo_msg = await bot.send_photo(
                        chat_id=channel_id,
                        photo=FSInputFile(thumb_path),
                        caption=f"🖼️ Muqova: {pptx_file.stem}\nToifa: {category}",
                    )
                    if photo_msg.photo:
                        thumb_file_id = photo_msg.photo[-1].file_id
                except Exception as ex:
                    print(f"⚠️ Rasm yuklashda ogohlantirish: {ex}")

            try:
                doc_msg = await bot.send_document(
                    chat_id=channel_id,
                    document=FSInputFile(str(pptx_file)),
                    caption=f"📁 Slayd: {pptx_file.name}\nID: {safe_id}",
                )
                pptx_file_id = doc_msg.document.file_id

                existing_catalog[safe_id] = {
                    "id": safe_id,
                    "title": pptx_file.stem.replace("_", " ").title(),
                    "category": category,
                    "filename": pptx_file.name,
                    "pptx_file_id": pptx_file_id,
                    "thumb_file_id": thumb_file_id,
                    "local_path": str(pptx_file),
                }
                uploaded_count += 1
                await asyncio.sleep(1.5)  # Telegram rate limit himoyasi
            except Exception as e:
                print(f"❌ Fayl yuklashda xatolik ({pptx_file.name}): {e}")

    with open(catalog_path, "w", encoding="utf-8") as f:
        json.dump(existing_catalog, f, indent=2, ensure_ascii=False)

    print(f"🎉 Muvaffaqiyatli yakunlandi! {uploaded_count} ta yangi shablon katalogga saqlandi.")
    await bot.session.close()


if __name__ == "__main__":
    ch_id = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "")
    if not ch_id and len(sys.argv) > 1:
        ch_id = sys.argv[1]
    if not ch_id:
        print("Iltimos, TELEGRAM_TEMPLATES_CHANNEL_ID parametrini .env ga yoki argument sifatida kiriting!")
        print("Masalan: python scripts/seed_telegram_templates.py -1001234567890")
    else:
        asyncio.run(seed_templates(ch_id))
