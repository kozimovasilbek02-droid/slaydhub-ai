# -*- coding: utf-8 -*-
import os
import sys
import asyncio
import logging
from pathlib import Path
from aiogram import Bot, Dispatcher
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.fsm.storage.memory import MemoryStorage

# Loyiha ildiz yo'lini qo'shish
current_dir = Path(__file__).resolve().parent
project_root = current_dir.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

from bot.config import BOT_TOKEN, OUTPUT_DIR, TEMP_DIR
from bot.handlers import start, translation, generator

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("TaqdimotUzBot")


async def main():
    if not BOT_TOKEN:
        logger.error("BOT_TOKEN aniqlanmadi! Iltimos, config.py yoki muhit o'zgaruvchilarini tekshiring.")
        return

    logger.info("Bot ishga tushmoqda...")
    session = AiohttpSession(timeout=300.0)
    bot = Bot(token=BOT_TOKEN, session=session)
    dp = Dispatcher(storage=MemoryStorage())

    # Handler routerlarini ulash
    dp.include_router(start.router)
    dp.include_router(translation.router)
    dp.include_router(generator.router)

    # Kerakli kataloglar mavjudligini ta'minlash
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    TEMP_DIR.mkdir(parents=True, exist_ok=True)

    bot_info = await bot.get_me()
    logger.info(f"✅ Bot muvaffaqiyatli ishga tushdi: @{bot_info.username} ({bot_info.first_name})")
    print(f"\n{'='*60}", flush=True)
    print(f"🚀 TAQDIMOTUZ AI BOTI ISHLAMOQDA!", flush=True)
    print(f"🤖 Bot Username: @{bot_info.username}", flush=True)
    print(f"🔗 Havola: https://t.me/{bot_info.username}", flush=True)
    print(f"{'='*60}\n", flush=True)

    # Polling boshlash
    await dp.start_polling(bot)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logger.info("Bot to'xtatildi.")
