# -*- coding: utf-8 -*-
"""
SlaydHub Ultra (v4.1) — All-In-One Academic & Presentation Ecosystem
Audited, Patched & Approved by Gemini 4 Argon (Lead Architect & QA Auditor)
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

# 1. Avval .env va core/config ni yuklash (import-time env muammolarini oldini oladi)
from core.config import config  # noqa: E402
from core.database import init_db  # noqa: E402
from core.memory_guard import (  # noqa: E402
    GLOBAL_SCHEDULER,
    enforce_tmp_quota,
    PENDING_UPLOADS,
)
from billing.webhooks import (  # noqa: E402
    handle_click_prepare,
    handle_click_complete,
    handle_payme_rpc,
)
from handlers.start_cabinet import router as start_cabinet_router  # noqa: E402
from handlers.agenda_flow import router as agenda_flow_router  # noqa: E402
from handlers.document_router import router as document_router  # noqa: E402

START_TIME = time.time()


# ── YAGONA VEB-SERVER: KEEPALIVE + BILLING (PORT 10000) ───────────────────────
async def handle_ping(request: web.Request) -> web.Response:
    """Render.com 24/7 Keepalive sog'liq tekshiruvi (< 2ms, zero-lock)."""
    uptime_sec = int(time.time() - START_TIME)
    return web.json_response(
        {
            "status": "ok",
            "service": "SlaydHub Ultra v4.1 (Argon QA Verified)",
            "uptime": f"{uptime_sec // 3600}h {(uptime_sec % 3600) // 60}m {uptime_sec % 60}s",
            "pending_uploads": len(PENDING_UPLOADS),
        }
    )


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
    print("  🚀 SLAYDHUB ULTRA (v4.1) — PRODUCTION ENGINE")
    print("  Arxitektor: Gemini 4 Argon | Infratuzilma: Render 512MB RAM")
    print("=" * 65)

    bot_token = config.bot_token
    if not bot_token:
        print("[CRITICAL] BOT_TOKEN topilmadi! Render Environment Variables ga kiriting.")
        sys.exit(1)

    # 1. SQLite WAL ma'lumotlar bazasini ishga tushirish
    db_path = config.db_path
    await init_db(db_path)
    print(f"[DB] SQLite WAL bazasi tayyor: {db_path}")

    # 2. Global PriorityQueue vazifalar rejalashtiruvchisini boshlash
    await GLOBAL_SCHEDULER.start()
    print("[SCHEDULER] SlaydHub GLOBAL_SCHEDULER (2 Workers, Priority 1..3) faollashtirildi")

    # 3. Yagona aiohttp veb-serverni ishga tushirish (PORT 10000)
    port = config.port
    server_runner = await start_unified_web_server(port=port)

    # 4. /tmp/soff_ready tozalash foni
    perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
    os.makedirs(perm_dir, exist_ok=True)
    cleanup_task = asyncio.create_task(background_cleanup_loop(perm_dir))

    # 5. Aiogram 3.x bot va routerlarni ulash
    bot = Bot(
        token=bot_token,
        default=DefaultBotProperties(parse_mode=ParseMode.HTML),
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
        await GLOBAL_SCHEDULER.stop()
        await server_runner.cleanup()
        await bot.session.close()
        print("[SHUTDOWN] Bot va server to'xtatildi.")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        pass
