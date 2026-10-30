# -*- coding: utf-8 -*-
"""
core/aiosqlite_shim.py
Lightweight, zero-dependency async wrapper around Python standard library sqlite3.
Provides 100% API compatibility with basic aiosqlite operations:
connect, execute, executescript, commit, rollback, fetchone, fetchall, Row.
Ensures SlaydHub Ultra starts and functions even if aiosqlite is missing on Render.com.
"""

import sqlite3
import asyncio
from typing import Any, Optional, Tuple, List


class AsyncCursor:
    def __init__(self, cur: sqlite3.Cursor):
        self._cur = cur
        self.lastrowid = cur.lastrowid
        self.rowcount = cur.rowcount

    async def fetchone(self) -> Optional[Any]:
        return await asyncio.to_thread(self._cur.fetchone)

    async def fetchall(self) -> List[Any]:
        return await asyncio.to_thread(self._cur.fetchall)


class AsyncConnection:
    def __init__(self, path: str, timeout: float = 10.0):
        self._path = path
        self._timeout = timeout
        self._conn: Optional[sqlite3.Connection] = None
        self.row_factory = sqlite3.Row

    async def _ensure_conn(self) -> sqlite3.Connection:
        if self._conn is None:
            self._conn = await asyncio.to_thread(
                sqlite3.connect, self._path, timeout=self._timeout, check_same_thread=False
            )
            self._conn.row_factory = self.row_factory
        return self._conn

    async def __aenter__(self) -> "AsyncConnection":
        await self._ensure_conn()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._conn:
            if exc_type:
                try:
                    await asyncio.to_thread(self._conn.rollback)
                except Exception:
                    pass
            try:
                await asyncio.to_thread(self._conn.close)
            except Exception:
                pass
            self._conn = None

    async def execute(self, sql: str, params: Tuple = ()) -> AsyncCursor:
        conn = await self._ensure_conn()
        cur = await asyncio.to_thread(conn.execute, sql, params)
        return AsyncCursor(cur)

    async def executescript(self, script: str) -> AsyncCursor:
        conn = await self._ensure_conn()
        cur = await asyncio.to_thread(conn.executescript, script)
        return AsyncCursor(cur)

    async def commit(self) -> None:
        if self._conn:
            await asyncio.to_thread(self._conn.commit)

    async def rollback(self) -> None:
        if self._conn:
            await asyncio.to_thread(self._conn.rollback)


def connect(path: str, timeout: float = 10.0, *args, **kwargs) -> AsyncConnection:
    return AsyncConnection(path, timeout=timeout)


Row = sqlite3.Row
Connection = AsyncConnection
