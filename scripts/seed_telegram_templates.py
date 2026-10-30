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

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

from aiogram import Bot
from aiogram.types import FSInputFile
from aiogram.exceptions import TelegramRetryAfter

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
        Path(r"C:\Users\user\Desktop\Institut\Suniy intelekt\05_Taqdimotlar_va_Maqolalar\Medical_Presentations_English_50"),
        REPO_ROOT / "templates",
        Path(r"C:\Users\user\Desktop\slayd\Yuklab_Olingan_Top50"),
        Path(r"C:\Users\user\Desktop\slayd\PowerPointSchool"),
        Path(r"C:\Users\user\Desktop\slayd\PresentationGO"),
        Path(r"C:\Users\user\Desktop\slayd\FreePPT7"),
        Path(r"C:\Users\user\Desktop\slayd\PPTTemplate"),
    ]

    # Toifalar bo'yicha saralash (Category Balancer)
    category_buckets: Dict[str, List[Path]] = {
        "med": [],
        "biz": [],
        "tech": [],
        "edu": [],
        "nature": [],
        "minimal": [],
    }

    seen_names = set()
    for s_dir in search_dirs:
        if s_dir.exists():
            for p in s_dir.rglob("*.pptx"):
                if p.name.startswith("~$") or p.stat().st_size < 50000 or p.name in seen_names:
                    continue
                seen_names.add(p.name)
                cat = GLOBAL_TEMPLATE_CATALOG.detect_category(p.name)
                category_buckets[cat].append(p)

    # Mavjud toifalar sonini hisoblash
    cat_counts = {c: 0 for c in category_buckets}
    for item in existing_catalog.values():
        c = item.get("category", "minimal")
        if c in cat_counts:
            cat_counts[c] += 1
    print(f"📊 Hozirgi katalogdagi toifalar: {cat_counts}")

    target_files: List[Path] = []
    # Kamroq yuklangan toifalardan boshlab to'ldirish (tech, edu, nature, med, biz, minimal)
    for cat in ["tech", "edu", "nature", "med", "biz", "minimal"]:
        flist = category_buckets.get(cat, [])
        needed = max(5, 12 - cat_counts.get(cat, 0))
        added = 0
        for p in flist:
            safe_id = f"tpl_{p.stem[:25].lower().replace(' ', '_').replace('-', '_')}"
            if safe_id not in existing_catalog:
                target_files.append(p)
                added += 1
                if added >= needed:
                    break

    print(f"🎯 Saralangan maqsadli yangi shablonlar: {len(target_files)} ta.")
    uploaded_count = 0

    for pptx_file in target_files:
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
                for _ in range(3):
                    try:
                        photo_msg = await bot.send_photo(
                            chat_id=channel_id,
                            photo=FSInputFile(thumb_path),
                            caption=f"🖼️ Muqova: {pptx_file.stem}\nToifa: {category}",
                        )
                        if photo_msg.photo:
                            thumb_file_id = photo_msg.photo[-1].file_id
                        break
                    except TelegramRetryAfter as tra:
                        print(f"⏳ FloodWait SendPhoto: {tra.retry_after}s kutilmoqda...", flush=True)
                        await asyncio.sleep(tra.retry_after + 2)
                    except Exception as ex:
                        print(f"⚠️ Rasm yuklashda ogohlantirish: {ex}", flush=True)
                        break

            doc_msg = None
            for _ in range(3):
                try:
                    doc_msg = await bot.send_document(
                        chat_id=channel_id,
                        document=FSInputFile(str(pptx_file)),
                        caption=f"📁 Slayd: {pptx_file.name}\nID: {safe_id}",
                    )
                    break
                except TelegramRetryAfter as tra:
                    print(f"⏳ FloodWait SendDocument: {tra.retry_after}s kutilmoqda...", flush=True)
                    await asyncio.sleep(tra.retry_after + 2)
                except Exception as e:
                    print(f"❌ Fayl yuklashda xatolik ({pptx_file.name}): {e}", flush=True)
                    break

            if doc_msg and doc_msg.document:
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
                print(f"✅ Saqlandi [{uploaded_count}/{max_upload}]: {safe_id} (Toifa: {category})", flush=True)

                # Har bir yuklashdan keyin oraliq holatni saqlab borish
                with open(catalog_path, "w", encoding="utf-8") as f:
                    json.dump(existing_catalog, f, indent=2, ensure_ascii=False)

                await asyncio.sleep(2.5)  # Barqaror Telegram intervali

    with open(catalog_path, "w", encoding="utf-8") as f:
        json.dump(existing_catalog, f, indent=2, ensure_ascii=False)

    print(f"🎉 Muvaffaqiyatli yakunlandi! Jami {len(existing_catalog)} ta shablon katalogda.", flush=True)
    await bot.session.close()


if __name__ == "__main__":
    ch_id = os.getenv("TELEGRAM_TEMPLATES_CHANNEL_ID", "")
    if len(sys.argv) > 1 and sys.argv[1].startswith("-100"):
        ch_id = sys.argv[1]

    max_up = 60
    if len(sys.argv) > 2 and sys.argv[2].isdigit():
        max_up = int(sys.argv[2])
    elif os.getenv("SEED_MAX_UPLOAD", "").isdigit():
        max_up = int(os.getenv("SEED_MAX_UPLOAD"))

    if not ch_id:
        print("Iltimos, TELEGRAM_TEMPLATES_CHANNEL_ID parametrini .env ga yoki argument sifatida kiriting!")
        print("Masalan: python scripts/seed_telegram_templates.py -1001234567890 50")
    else:
        print(f"🚀 Telegram Vault Seeding boshlanmoqda (Maks: {max_up} ta shablon)...")
        asyncio.run(seed_templates(ch_id, max_upload=max_up))
