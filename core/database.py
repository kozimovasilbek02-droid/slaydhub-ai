# -*- coding: utf-8 -*-
import os
import time
import json
from typing import Optional, Dict, Any
import aiosqlite


def get_db_path() -> str:
    return os.getenv("SLAYDHUB_DB_PATH", "slaydhub_ultra.db")


async def init_db(db_path: Optional[str] = None) -> None:
    """SQLite WAL rejimida foydalanuvchilar, tranzaksiyalar va qoralama jadvallarini ishga tushiradi."""
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        await db.execute("PRAGMA journal_mode=WAL;")
        await db.execute("PRAGMA synchronous=NORMAL;")
        await db.execute("PRAGMA busy_timeout=5000;")

        # 1. Foydalanuvchilar jadvali (balans, keshbek, referal)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                balance_uzs INTEGER NOT NULL DEFAULT 0,
                cashback_uzs INTEGER NOT NULL DEFAULT 0,
                referred_by INTEGER,
                referral_count INTEGER NOT NULL DEFAULT 0,
                free_slide_credits INTEGER NOT NULL DEFAULT 1,
                created_at INTEGER NOT NULL
            );
        """)

        # 2. To'lov tranzaksiyalari (Click va Payme)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS payment_transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                provider TEXT NOT NULL,
                provider_tx_id TEXT NOT NULL,
                user_id INTEGER NOT NULL,
                amount_uzs INTEGER NOT NULL,
                state INTEGER NOT NULL, -- 1: created, 2: completed, -1/-2: cancelled
                create_time_ms INTEGER NOT NULL,
                perform_time_ms INTEGER NOT NULL DEFAULT 0,
                cancel_time_ms INTEGER NOT NULL DEFAULT 0,
                reason INTEGER,
                UNIQUE(provider, provider_tx_id)
            );
        """)

        # 3. Foydalanuvchi qoralamalari (Agenda, DOCX bo'limlari)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS user_drafts (
                user_id INTEGER NOT NULL,
                doc_type TEXT NOT NULL,
                title TEXT NOT NULL,
                stage TEXT NOT NULL,
                data_json TEXT NOT NULL,
                updated_at INTEGER NOT NULL,
                PRIMARY KEY(user_id, doc_type)
            );
        """)
        await db.commit()


async def get_or_create_user(
    user_id: int, referred_by: Optional[int] = None, db_path: Optional[str] = None
) -> Dict[str, Any]:
    """Foydalanuvchini oladi yoki yangisini yaratib beradi (referal mukofotlari bilan)."""
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        db.row_factory = aiosqlite.Row
        await db.execute("BEGIN IMMEDIATE;")
        cur = await db.execute("SELECT * FROM users WHERE user_id = ?;", (user_id,))
        row = await cur.fetchone()
        if row:
            await db.commit()
            return dict(row)

        now = int(time.time())
        valid_ref = None
        if referred_by and referred_by != user_id:
            cur_ref = await db.execute("SELECT user_id FROM users WHERE user_id = ?;", (referred_by,))
            if await cur_ref.fetchone():
                valid_ref = referred_by
                await db.execute("""
                    UPDATE users
                    SET referral_count = referral_count + 1,
                        free_slide_credits = free_slide_credits + (CASE WHEN (referral_count + 1) % 3 = 0 THEN 1 ELSE 0 END)
                    WHERE user_id = ?;
                """, (referred_by,))

        await db.execute("""
            INSERT INTO users (user_id, balance_uzs, cashback_uzs, referred_by, referral_count, free_slide_credits, created_at)
            VALUES (?, 0, 0, ?, 0, 1, ?);
        """, (user_id, valid_ref, now))
        await db.commit()

        cur2 = await db.execute("SELECT * FROM users WHERE user_id = ?;", (user_id,))
        new_row = await cur2.fetchone()
        return dict(new_row) if new_row else {}


async def credit_user_with_cashback(
    user_id: int, amount_uzs: int, cashback_percent: float = 0.10, db_path: Optional[str] = None
) -> None:
    """Foydalanuvchiga to'lov summasi va 10% keshbekni atomik qo'shish."""
    target_db = db_path or get_db_path()
    cashback = int(amount_uzs * cashback_percent)
    async with aiosqlite.connect(target_db) as db:
        await db.execute("BEGIN IMMEDIATE;")
        await db.execute("""
            UPDATE users
            SET balance_uzs = balance_uzs + ? + ?,
                cashback_uzs = cashback_uzs + ?
            WHERE user_id = ?;
        """, (amount_uzs, cashback, cashback, user_id))
        await db.commit()


async def debit_user_balance(user_id: int, amount_uzs: int, db_path: Optional[str] = None) -> bool:
    """Balansdan xizmat narxini yechib olish. Yetarli bo'lsa True, aks holda False."""
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        await db.execute("BEGIN IMMEDIATE;")
        cur = await db.execute("SELECT balance_uzs FROM users WHERE user_id = ?;", (user_id,))
        row = await cur.fetchone()
        if not row:
            await db.rollback()
            return False
        balance = row[0]
        if balance >= amount_uzs:
            await db.execute(
                "UPDATE users SET balance_uzs = balance_uzs - ? WHERE user_id = ?;",
                (amount_uzs, user_id),
            )
            await db.commit()
            return True
        await db.rollback()
        return False


async def refund_user_balance(user_id: int, amount_uzs: int, db_path: Optional[str] = None) -> None:
    """Generatsiya xatolik bilan tugaganda foydalanuvchi balansini qaytarish."""
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        await db.execute("BEGIN IMMEDIATE;")
        await db.execute(
            "UPDATE users SET balance_uzs = balance_uzs + ? WHERE user_id = ?;",
            (amount_uzs, user_id),
        )
        await db.commit()


async def use_free_credit(user_id: int, db_path: Optional[str] = None) -> bool:
    """Bepul slayd kreditidan foydalanish."""
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        await db.execute("BEGIN IMMEDIATE;")
        cur = await db.execute("SELECT free_slide_credits FROM users WHERE user_id = ?;", (user_id,))
        row = await cur.fetchone()
        if row and row[0] > 0:
            await db.execute(
                "UPDATE users SET free_slide_credits = free_slide_credits - 1 WHERE user_id = ?;",
                (user_id,),
            )
            await db.commit()
            return True
        await db.rollback()
        return False


async def refund_free_credit(user_id: int, db_path: Optional[str] = None) -> None:
    """Generatsiya xatolik bilan tugaganda bepul kreditni joyiga qaytarish."""
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        await db.execute("BEGIN IMMEDIATE;")
        await db.execute(
            "UPDATE users SET free_slide_credits = free_slide_credits + 1 WHERE user_id = ?;",
            (user_id,),
        )
        await db.commit()


async def save_draft(
    user_id: int, doc_type: str, title: str, stage: str, data: Dict[str, Any], db_path: Optional[str] = None
) -> None:
    target_db = db_path or get_db_path()
    now = int(time.time())
    data_str = json.dumps(data, ensure_ascii=False)
    async with aiosqlite.connect(target_db) as db:
        await db.execute("""
            INSERT OR REPLACE INTO user_drafts (user_id, doc_type, title, stage, data_json, updated_at)
            VALUES (?, ?, ?, ?, ?, ?);
        """, (user_id, doc_type, title, stage, data_str, now))
        await db.commit()


async def get_draft(user_id: int, doc_type: str, db_path: Optional[str] = None) -> Optional[Dict[str, Any]]:
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        db.row_factory = aiosqlite.Row
        cur = await db.execute(
            "SELECT * FROM user_drafts WHERE user_id = ? AND doc_type = ?;", (user_id, doc_type)
        )
        row = await cur.fetchone()
        if not row:
            return None
        res = dict(row)
        try:
            res["data"] = json.loads(res.get("data_json", "{}"))
        except Exception:
            res["data"] = {}
        return res


async def delete_draft(user_id: int, doc_type: str, db_path: Optional[str] = None) -> None:
    target_db = db_path or get_db_path()
    async with aiosqlite.connect(target_db) as db:
        await db.execute(
            "DELETE FROM user_drafts WHERE user_id = ? AND doc_type = ?;", (user_id, doc_type)
        )
        await db.commit()
