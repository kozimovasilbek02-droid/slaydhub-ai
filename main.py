# -*- coding: utf-8 -*-
"""
SlaydHub Ultra (v4.0) — All-In-One Academic & Presentation Ecosystem
Audited & Approved by Gemini 4 Argon (Lead Architect & QA Auditor)
Target Infrastructure: Render.com 512MB RAM Web Service, Single PORT=10000
"""

import os
import sys
import time
import asyncio
import tempfile
from aiohttp import web
from aiogram import Bot, Dispatcher
from aiogram.enums import ParseMode
from aiogram.client.default import DefaultBotProperties
from aiogram.fsm.storage.memory import MemoryStorage

# Unicode konsol chiqishini sozlash (Windows va Linux uchun)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Linux glibc arenalarini cheklash (512MB RAM himoyasi)
os.environ.setdefault("MALLOC_ARENA_MAX", "2")
os.environ.setdefault("MALLOC_TRIM_THRESHOLD_", "65536")

from core.database import init_db
from core.memory_guard import SlaydHubTaskScheduler, enforce_tmp_quota, PENDING_UPLOADS
from billing.webhooks import handle_click_prepare, handle_click_complete, handle_payme_rpc
from handlers.start_cabinet import router as start_cabinet_router
from handlers.agenda_flow import router as agenda_flow_router
from handlers.document_router import router as document_router

# ── KONFIGURATSIYA VA ENV ─────────────────────────────────────────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN") or os.getenv("TELEGRAM_BOT_TOKEN")
PORT = int(os.getenv("PORT", 10000))
START_TIME = time.time()

# Local .env fallback
if not BOT_TOKEN:
    for env_path in [
        os.path.join(os.path.dirname(__file__), ".env"),
        os.path.join(os.path.dirname(__file__), "..", ".env"),
    ]:
        if os.path.exists(env_path):
            with open(env_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith(("BOT_TOKEN=", "TELEGRAM_BOT_TOKEN=")) and not BOT_TOKEN:
                        BOT_TOKEN = line.split("=", 1)[1].strip().strip("\"'")
                    elif line.startswith("GEMINI_API_KEY=") and not os.getenv("GEMINI_API_KEY"):
                        os.environ["GEMINI_API_KEY"] = line.split("=", 1)[1].strip().strip("\"'")
                    elif line.startswith("SELLER_TOKEN=") and not os.getenv("SELLER_TOKEN"):
                        os.environ["SELLER_TOKEN"] = line.split("=", 1)[1].strip().strip("\"'")


# ── YAGONA VEB-SERVER: KEEPALIVE + BILLING (PORT 10000) ───────────────────────
async def handle_ping(request: web.Request) -> web.Response:
    """Render.com 24/7 Keepalive sog'liq tekshiruvi (< 2ms, zero-lock)."""
    uptime_sec = int(time.time() - START_TIME)
    return web.json_response({
        "status": "ok",
        "service": "SlaydHub Ultra v4.0 (Argon Approved)",
        "uptime": f"{uptime_sec // 3600}h {(uptime_sec % 3600) // 60}m {uptime_sec % 60}s",
        "pending_uploads": len(PENDING_UPLOADS),
    })


async def start_unified_web_server(port: int = 10000) -> web.AppRunner:
    app = web.Application(client_max_size=5 * 1024 * 1024)
    # 1. 24/7 Healthcheck Keepalive
    app.router.add_get("/", handle_ping)
    app.router.add_get("/health", handle_ping)
    # 2. Click & Payme Webhook shlyuzlari
    app.router.add_post("/api/pay/click/prepare", handle_click_prepare)
    app.router.add_post("/api/pay/click/complete", handle_click_complete)
    app.router.add_post("/api/pay/payme", handle_payme_rpc)

    runner = web.AppRunner(app, access_log=None)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    print(f"[WEB] Yagona aiohttp veb-server ishga tushdi: 0.0.0.0:{port}")
    return runner


# ── BACKGROUND /TMP TOZALAGICH ────────────────────────────────────────────────
async def background_cleanup_loop(perm_dir: str):
    while True:
        try:
            await asyncio.sleep(300)  # Har 5 daqiqada
            enforce_tmp_quota(perm_dir)
        except asyncio.CancelledError:
            break
        except Exception as e:
            print(f"[CLEANUP ERROR] {e}")


# ── ASOSIY BOT ISHGA TUSHIRISH ────────────────────────────────────────────────
async def main():
    print("=" * 65)
    print("  🚀 SLAYDHUB ULTRA (v4.0) — PRODUCTION ENGINE")
    print("  Arxitektor: Gemini 4 Argon | Infratuzilma: Render 512MB RAM")
    print("=" * 65)

    if not BOT_TOKEN:
        print("[CRITICAL] BOT_TOKEN topilmadi! Render Environment Variables ga kiriting.")
        sys.exit(1)

    # 1. SQLite WAL ma'lumotlar bazasini ishga tushirish
    db_path = os.getenv("SLAYDHUB_DB_PATH", "slaydhub_ultra.db")
    await init_db(db_path)
    print(f"[DB] SQLite WAL bazasi tayyor: {db_path}")

    # 2. PriorityQueue vazifalar rejalashtiruvchisini boshlash
    scheduler = SlaydHubTaskScheduler(max_queue_size=50, concurrency=2)
    await scheduler.start()
    print("[SCHEDULER] SlaydHub Task Scheduler (2 Workers) faollashtirildi")

    # 3. Yagona aiohttp veb-serverni ishga tushirish (PORT 10000)
    server_runner = await start_unified_web_server(port=PORT)

    # 4. /tmp/soff_ready tozalash foni
    perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
    os.makedirs(perm_dir, exist_ok=True)
    cleanup_task = asyncio.create_task(background_cleanup_loop(perm_dir))

    # 5. Aiogram 3.x bot va routerlarni ulash
    bot = Bot(
        token=BOT_TOKEN,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML)
    )
    dp = Dispatcher(storage=MemoryStorage())

    dp.include_router(start_cabinet_router)
    dp.include_router(agenda_flow_router)
    dp.include_router(document_router)

    try:
        bot_user = await bot.get_me()
        print(f"[BOT] @{bot_user.username} muvaffaqiyatli ishga tushdi!")
        await dp.start_polling(bot, drop_pending_updates=True)
    finally:
        cleanup_task.cancel()
        await scheduler.stop()
        await server_runner.cleanup()
        await bot.session.close()
        print("[SHUTDOWN] Bot va server to'xtatildi.")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass
