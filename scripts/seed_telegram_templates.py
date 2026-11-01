# -*- coding: utf-8 -*-
"""
scripts/seed_telegram_templates.py
Automated seeder script for SlaydHub Ultra 1,000+ templates.
1. Scans local PPTX directories (Desktop/slayd, templates, Medical_Presentations_English_50, Powepoint).
2. Extracts high-res cover thumbnail (docProps/thumbnail.jpeg) in memory.
3. Uploads thumbnail and PPTX file to your private Telegram storage channel with ATOMIC ROLLBACK.
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

import aiohttp
from aiogram import Bot
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.types import FSInputFile
from aiogram.exceptions import TelegramRetryAfter

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from core.config import config
from core.template_catalog import GLOBAL_TEMPLATE_CATALOG, CATEGORIES


async def seed_templates(channel_id: str, max_upload: int = 50) -> None:
    # 300 soniyalik (5 daqiqa) kengaytirilgan timeout - katta PPTX fayllar uzilib qolmasligi uchun
    custom_timeout = aiohttp.ClientTimeout(total=300, connect=60, sock_read=300)
    session = AiohttpSession(timeout=custom_timeout)
    bot = Bot(token=config.bot_token, session=session)

    try:
        me = await bot.get_me()
        print(f"✅ Bot ulandi: @{me.username}", flush=True)
    except Exception as e:
        print(f"❌ Bot token xatosi: {e}", flush=True)
        await bot.session.close()
        return

    print(f"📢 Nishon Telegram kanali: {channel_id}", flush=True)
    catalog_path = REPO_ROOT / "templates_catalog.json"
    existing_catalog: Dict[str, Any] = {}
    if catalog_path.exists():
        try:
            with open(catalog_path, "r", encoding="utf-8") as f:
                existing_catalog = json.load(f)
        except Exception:
            existing_catalog = {}

    # Qidiruv kataloglari (Kompyuteringizdagi barcha 16,000+ ta shablonlar ombori)
    search_dirs = [
        Path(r"C:\Users\user\Desktop\Institut\Suniy intelekt\05_Taqdimotlar_va_Maqolalar\Medical_Presentations_English_50"),
        REPO_ROOT / "templates",
        Path(r"C:\Users\user\Desktop\slayd"),
        Path(r"C:\Users\user\Desktop\Powepoint"),
        Path(r"G:\Мой диск\Powepoint"),
        Path(r"G:\Мой диск\Slaydlar"),
        Path(r"G:\Мой диск\SlideEgg_Downloads"),
    ]

    # Allaqachon to'liq (fayli bilan) katalogda mavjud bo'lgan fayllar
    existing_filenames = {
        item.get("filename")
        for item in existing_catalog.values()
        if item.get("filename") and item.get("pptx_file_id")
    }
    print(f"📦 Avval to'liq yuklangan tasdiqlangan fayllar: {len(existing_filenames)} ta.", flush=True)

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
                # Vaqtinchalik fayllar, 50KB dan kichik yoki 48MB dan katta (Telegram bot limiti 50MB) fayllarni o'tkazib yuborish
                if p.name.startswith("~$") or p.stat().st_size < 50000 or p.stat().st_size > 48 * 1024 * 1024:
                    continue
                if p.name in seen_names or p.name in existing_filenames:
                    continue
                seen_names.add(p.name)
                cat = GLOBAL_TEMPLATE_CATALOG.detect_category(p.name)
                category_buckets[cat].append(p)

    target_files: List[Path] = []
    if max_upload > 200:
        # Barcha arxiv rejimida: barcha yangi yuklanadigan shablonlarni navbat bilan qo'shish
        for cat, flist in category_buckets.items():
            print(f"📁 Toifa [{cat}]: {len(flist)} ta yangi yuklanadigan shablon topildi", flush=True)
            target_files.extend(flist)
    else:
        # Balanslangan kvota rejimi
        per_cat = max(5, max_upload // 6)
        for cat, flist in category_buckets.items():
            target_files.extend(flist[:per_cat])

    print(f"🎯 Jami yuklanishi kerak bo'lgan yangi shablonlar: {len(target_files)} ta.", flush=True)
    uploaded_count = 0

    for pptx_file in target_files:
        if uploaded_count >= max_upload:
            print(f"⏸️ Maksimal kvota ({max_upload} ta) yetdi. Seeding yakunlandi.", flush=True)
            break

        # Unikal xeshli ID yaratish (nomlar to'qnashmasligi uchun)
        safe_name = "".join(c if c.isalnum() else "_" for c in pptx_file.stem.lower())[:30].strip("_")
        name_hash = hex(abs(hash(pptx_file.name)))[2:8]
        safe_id = f"tpl_{safe_name}_{name_hash}"

        category = GLOBAL_TEMPLATE_CATALOG.detect_category(pptx_file.name)
        print(f"📤 Yuklanmoqda [{uploaded_count + 1}/{min(len(target_files), max_upload)}]: {pptx_file.name} (Toifa: {category}, Hajm: {round(pptx_file.stat().st_size / (1024*1024), 2)} MB)...", flush=True)

        with tempfile.TemporaryDirectory() as td:
            thumb_path = os.path.join(td, "thumb.jpeg")
            has_thumb = GLOBAL_TEMPLATE_CATALOG.extract_cover_thumbnail(str(pptx_file), thumb_path)

            photo_msg = None
            thumb_file_id = None
            if has_thumb:
                for attempt in range(3):
                    try:
                        photo_msg = await bot.send_photo(
                            chat_id=channel_id,
                            photo=FSInputFile(thumb_path),
                            caption=f"🖼️ Muqova: {pptx_file.stem}\nToifa: {category}",
                        )
                        if photo_msg and photo_msg.photo:
                            thumb_file_id = photo_msg.photo[-1].file_id
                        break
                    except TelegramRetryAfter as tra:
                        print(f"⏳ FloodWait SendPhoto: {tra.retry_after}s kutilmoqda...", flush=True)
                        await asyncio.sleep(tra.retry_after + 2)
                    except Exception as ex:
                        print(f"⚠️ Rasm yuklashda ogohlantirish (urinish {attempt+1}): {ex}", flush=True)
                        await asyncio.sleep(2)

            doc_msg = None
            for attempt in range(4):
                try:
                    doc_msg = await bot.send_document(
                        chat_id=channel_id,
                        document=FSInputFile(str(pptx_file)),
                        caption=f"📁 Slayd: {pptx_file.name}\nID: {safe_id}",
                    )
                    if doc_msg and doc_msg.document:
                        break
                except TelegramRetryAfter as tra:
                    print(f"⏳ FloodWait SendDocument: {tra.retry_after}s kutilmoqda...", flush=True)
                    await asyncio.sleep(tra.retry_after + 2)
                except Exception as e:
                    print(f"❌ Fayl yuklashda xatolik (urinish {attempt+1}, {pptx_file.name}): {e}", flush=True)
                    await asyncio.sleep(3)

            # ATOMIK ROLLBACK KAFOLATI: Agar fayl yuklanmasa, yetim qolgan muqova darhol kanaldan o'chiriladi!
            if not doc_msg or not doc_msg.document:
                if photo_msg:
                    try:
                        print(f"🧹 ROLLBACK: Fayl yuklanmadi! Yetim muqova kanaldan o'chirilmoqda (msg_id: {photo_msg.message_id})...", flush=True)
                        await bot.delete_message(chat_id=channel_id, message_id=photo_msg.message_id)
                    except Exception as dex:
                        print(f"⚠️ Muqovani o'chirishda xatolik: {dex}", flush=True)
                print(f"⏭️ {pptx_file.name} o'tkazib yuborildi (yetim muqova qoldirilmadi).", flush=True)
                continue

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
            print(f"✅ Muvaffaqiyatli saqlandi [{uploaded_count}/{max_upload}]: {safe_id} (Toifa: {category})", flush=True)

            # Har bir yuklashdan keyin darhol katalog faylini yangilash
            with open(catalog_path, "w", encoding="utf-8") as f:
                json.dump(existing_catalog, f, indent=2, ensure_ascii=False)

            await asyncio.sleep(2.0)  # Telegram API limitlariga tushmaslik uchun xavfsiz interval

    with open(catalog_path, "w", encoding="utf-8") as f:
        json.dump(existing_catalog, f, indent=2, ensure_ascii=False)

    print(f"🎉 Yakunlandi! Jami {len(existing_catalog)} ta shablon (fayli va muqovasi bilan) katalogda.", flush=True)
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
