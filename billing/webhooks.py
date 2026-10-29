# -*- coding: utf-8 -*-
import os
import time
import hmac
import base64
import hashlib
import aiosqlite
from aiohttp import web

CLICK_SERVICE_ID = os.getenv("CLICK_SERVICE_ID", "")
CLICK_SECRET_KEY = os.getenv("CLICK_SECRET_KEY", "")
PAYME_MERCHANT_KEY = os.getenv("PAYME_MERCHANT_KEY", "")
DB_PATH = os.getenv("SLAYDHUB_DB_PATH", "slaydhub_ultra.db")


async def _credit_user_with_cashback(db: aiosqlite.Connection, user_id: int, amount_uzs: int) -> None:
    """Asosiy balansni va 10% keshbekni atomik ravishda qo'shadi."""
    cashback = int(amount_uzs * 0.10)
    await db.execute(
        """
        UPDATE users
        SET balance_uzs = balance_uzs + ? + ?,
            cashback_uzs = cashback_uzs + ?
        WHERE user_id = ?;
        """,
        (amount_uzs, cashback, cashback, user_id),
    )


# ── 1. CLICK PREPARE & COMPLETE WEBHOOKS ──────────────────────────────────────
async def handle_click_prepare(request: web.Request) -> web.Response:
    data = await request.post()
    click_trans_id = str(data.get("click_trans_id", ""))
    service_id = str(data.get("service_id", ""))
    merchant_trans_id = str(data.get("merchant_trans_id", ""))  # Telegram user_id
    amount = str(data.get("amount", "0"))
    action = str(data.get("action", "0"))
    sign_time = str(data.get("sign_time", ""))
    sign_string = str(data.get("sign_string", ""))

    raw = f"{click_trans_id}{service_id}{CLICK_SECRET_KEY}{merchant_trans_id}{amount}{action}{sign_time}"
    expected_sign = hashlib.md5(raw.encode("utf-8")).hexdigest()

    if CLICK_SECRET_KEY and not hmac.compare_digest(expected_sign, sign_string):
        return web.json_response({"error": -1, "error_note": "SIGN CHECK FAILED!"})

    try:
        user_id = int(merchant_trans_id)
        amount_uzs = int(float(amount))
    except ValueError:
        return web.json_response({"error": -5, "error_note": "User does not exist"})

    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("BEGIN IMMEDIATE;")
        cur = await db.execute("SELECT user_id FROM users WHERE user_id = ?;", (user_id,))
        if not await cur.fetchone():
            await db.rollback()
            return web.json_response({"error": -5, "error_note": "User does not exist"})

        now_ms = int(time.time() * 1000)
        await db.execute(
            """
            INSERT OR IGNORE INTO payment_transactions
            (provider, provider_tx_id, user_id, amount_uzs, state, create_time_ms)
            VALUES ('click', ?, ?, ?, 1, ?);
            """,
            (click_trans_id, user_id, amount_uzs, now_ms),
        )
        cur_tx = await db.execute(
            "SELECT id, state FROM payment_transactions WHERE provider='click' AND provider_tx_id=?;",
            (click_trans_id,),
        )
        tx_row = await cur_tx.fetchone()
        await db.commit()

    if tx_row and tx_row[1] == 2:
        return web.json_response({"error": -4, "error_note": "Already paid"})

    return web.json_response(
        {
            "click_trans_id": click_trans_id,
            "merchant_trans_id": merchant_trans_id,
            "merchant_prepare_id": tx_row[0] if tx_row else 1,
            "error": 0,
            "error_note": "Success",
        }
    )


async def handle_click_complete(request: web.Request) -> web.Response:
    data = await request.post()
    click_trans_id = str(data.get("click_trans_id", ""))
    service_id = str(data.get("service_id", ""))
    merchant_trans_id = str(data.get("merchant_trans_id", ""))
    merchant_prepare_id = str(data.get("merchant_prepare_id", ""))
    amount = str(data.get("amount", "0"))
    action = str(data.get("action", "1"))
    sign_time = str(data.get("sign_time", ""))
    sign_string = str(data.get("sign_string", ""))
    click_error = int(data.get("error", 0))

    raw = (
        f"{click_trans_id}{service_id}{CLICK_SECRET_KEY}"
        f"{merchant_trans_id}{merchant_prepare_id}{amount}{action}{sign_time}"
    )
    expected_sign = hashlib.md5(raw.encode("utf-8")).hexdigest()

    if CLICK_SECRET_KEY and not hmac.compare_digest(expected_sign, sign_string):
        return web.json_response({"error": -1, "error_note": "SIGN CHECK FAILED!"})

    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("BEGIN IMMEDIATE;")
        cur = await db.execute(
            "SELECT id, user_id, amount_uzs, state FROM payment_transactions WHERE provider='click' AND provider_tx_id=?;",
            (click_trans_id,),
        )
        tx = await cur.fetchone()
        if not tx:
            await db.rollback()
            return web.json_response({"error": -6, "error_note": "Transaction does not exist"})

        tx_id, user_id, amount_uzs, state = tx
        if state == 2:
            await db.rollback()
            return web.json_response({"error": -4, "error_note": "Already paid"})
        if state < 0:
            await db.rollback()
            return web.json_response({"error": -9, "error_note": "Transaction cancelled"})

        now_ms = int(time.time() * 1000)
        if click_error < 0:
            await db.execute(
                "UPDATE payment_transactions SET state=-1, cancel_time_ms=? WHERE id=?;",
                (now_ms, tx_id),
            )
            await db.commit()
            return web.json_response({"error": -9, "error_note": "Transaction cancelled"})

        # Muvaffaqiyatli to'lov: state=2 va 10% keshbek bilan balansni oshirish
        await db.execute(
            "UPDATE payment_transactions SET state=2, perform_time_ms=? WHERE id=?;",
            (now_ms, tx_id),
        )
        await _credit_user_with_cashback(db, user_id, amount_uzs)
        await db.commit()

    return web.json_response(
        {
            "click_trans_id": click_trans_id,
            "merchant_trans_id": merchant_trans_id,
            "merchant_confirm_id": tx_id,
            "error": 0,
            "error_note": "Success",
        }
    )


# ── 2. PAYME JSON-RPC 2.0 WEBHOOK ────────────────────────────────────────────
def _verify_payme_auth(request: web.Request) -> bool:
    if not PAYME_MERCHANT_KEY:
        return True
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Basic "):
        return False
    try:
        decoded = base64.b64decode(auth_header[6:]).decode("utf-8")
        _, key = decoded.split(":", 1)
        return hmac.compare_digest(key, PAYME_MERCHANT_KEY)
    except Exception:
        return False


async def handle_payme_rpc(request: web.Request) -> web.Response:
    try:
        body = await request.json()
    except Exception:
        return web.json_response({"jsonrpc": "2.0", "error": {"code": -32700, "message": "Parse error"}})

    rpc_id = body.get("id")

    if not _verify_payme_auth(request):
        return web.json_response(
            {"jsonrpc": "2.0", "id": rpc_id, "error": {"code": -32504, "message": "Insufficient privilege"}}
        )

    method = body.get("method")
    params = body.get("params", {})

    if method == "CheckPerformTransaction":
        account = params.get("account", {})
        user_id = int(account.get("user_id", -1))
        amount_tiyin = int(params.get("amount", 0))
        if amount_tiyin < 1000 * 100:
            return web.json_response(
                {"jsonrpc": "2.0", "id": rpc_id, "error": {"code": -31001, "message": "Incorrect amount"}}
            )
        async with aiosqlite.connect(DB_PATH) as db:
            cur = await db.execute("SELECT user_id FROM users WHERE user_id=?;", (user_id,))
            if not await cur.fetchone():
                return web.json_response(
                    {"jsonrpc": "2.0", "id": rpc_id, "error": {"code": -31050, "message": "User not found"}}
                )
        return web.json_response({"jsonrpc": "2.0", "id": rpc_id, "result": {"allow": True}})

    elif method == "CreateTransaction":
        payme_id = str(params.get("id"))
        user_id = int(params.get("account", {}).get("user_id", -1))
        amount_uzs = int(params.get("amount", 0)) // 100
        create_time = int(params.get("time", int(time.time() * 1000)))

        async with aiosqlite.connect(DB_PATH) as db:
            await db.execute("BEGIN IMMEDIATE;")
            await db.execute(
                """
                INSERT OR IGNORE INTO payment_transactions
                (provider, provider_tx_id, user_id, amount_uzs, state, create_time_ms)
                VALUES ('payme', ?, ?, ?, 1, ?);
                """,
                (payme_id, user_id, amount_uzs, create_time),
            )
            cur = await db.execute(
                "SELECT id, state, create_time_ms FROM payment_transactions WHERE provider='payme' AND provider_tx_id=?;",
                (payme_id,),
            )
            row = await cur.fetchone()
            await db.commit()

        return web.json_response(
            {
                "jsonrpc": "2.0",
                "id": rpc_id,
                "result": {
                    "create_time": row[2] if row else create_time,
                    "transaction": str(row[0] if row else 1),
                    "state": row[1] if row else 1,
                },
            }
        )

    elif method == "PerformTransaction":
        payme_id = str(params.get("id"))
        now_ms = int(time.time() * 1000)

        async with aiosqlite.connect(DB_PATH) as db:
            await db.execute("BEGIN IMMEDIATE;")
            cur = await db.execute(
                "SELECT id, user_id, amount_uzs, state, perform_time_ms FROM payment_transactions WHERE provider='payme' AND provider_tx_id=?;",
                (payme_id,),
            )
            row = await cur.fetchone()
            if not row:
                await db.rollback()
                return web.json_response(
                    {"jsonrpc": "2.0", "id": rpc_id, "error": {"code": -31003, "message": "Transaction not found"}}
                )
            tx_id, user_id, amount_uzs, state, perform_time_ms = row
            if state == 1:
                perform_time_ms = now_ms
                await db.execute(
                    "UPDATE payment_transactions SET state=2, perform_time_ms=? WHERE id=?;",
                    (perform_time_ms, tx_id),
                )
                await _credit_user_with_cashback(db, user_id, amount_uzs)
                await db.commit()
                state = 2
            else:
                await db.rollback()

        return web.json_response(
            {
                "jsonrpc": "2.0",
                "id": rpc_id,
                "result": {
                    "transaction": str(tx_id),
                    "perform_time": perform_time_ms,
                    "state": state,
                },
            }
        )

    return web.json_response(
        {"jsonrpc": "2.0", "id": rpc_id, "error": {"code": -32601, "message": "Method not found"}}
    )
