# -*- coding: utf-8 -*-
import os
import gc
import time
import ctypes
import asyncio
import tempfile
from dataclasses import dataclass, field
from typing import Callable, Any, Awaitable, Dict

# Linux glibc arenalarini cheklash (Render.com konteyneri uchun)
os.environ.setdefault("MALLOC_ARENA_MAX", "2")
os.environ.setdefault("MALLOC_TRIM_THRESHOLD_", "65536")

try:
    _libc = ctypes.CDLL("libc.so.6")
except Exception:
    _libc = None

# 1. Ikki bosqichli semaforlar
RAM_COMPUTE_LOCK = asyncio.Semaphore(1)      # Faqat CPU/RAM og'ir C/XML amallar uchun (1-4s)
NETWORK_IO_SEMAPHORE = asyncio.Semaphore(5)  # Gemini API va HTTP so'rovlar uchun

# 2. Vaqtinchalik fayllar registri va 60 MB qat'iy kvota
MAX_TMP_STORAGE_BYTES = 60 * 1024 * 1024  # 60 MB hard limit on Render tmpfs
UPLOAD_CACHE_TTL_SECONDS = 1800           # 30 daqiqa
PENDING_UPLOADS: Dict[str, Dict[str, Any]] = {}


def release_c_and_python_memory() -> None:
    """PyMuPDF C-keshini, Python GC avlodlarini va glibc C-heap sahifalarini OS ga qaytaradi."""
    try:
        import fitz
        fitz.TOOLS.store_shrink(100)
    except Exception:
        pass
    gc.collect()
    if _libc is not None:
        try:
            _libc.malloc_trim(0)
        except Exception:
            pass


async def run_ram_heavy_sync(func: Callable, *args, **kwargs) -> Any:
    """
    Sinxron C/XML funksiyalarini (fitz, python-docx, python-pptx)
    faqat RAM_COMPUTE_LOCK ichida, alohida oqimda bajaradi va darhol RAMni tozalaydi.
    """
    async with RAM_COMPUTE_LOCK:
        try:
            return await asyncio.to_thread(func, *args, **kwargs)
        finally:
            release_c_and_python_memory()


def enforce_tmp_quota(target_dir: str) -> None:
    """Agar /tmp/soff_ready 60 MB dan oshsa, eng eski fayllarni LRU tartibida o'chiradi."""
    if not os.path.exists(target_dir):
        return
    now = time.time()
    # 1. Muddati o'tganlarni o'chirish (TTL: 30 min)
    for uid, meta in list(PENDING_UPLOADS.items()):
        if now - meta.get("created_at", now) > UPLOAD_CACHE_TTL_SECONDS:
            fp = meta.get("file_path")
            if fp and os.path.exists(fp):
                try:
                    os.remove(fp)
                except OSError:
                    pass
            PENDING_UPLOADS.pop(uid, None)

    # 2. Hajm kvotasini tekshirish (LRU Eviction)
    entries = []
    total_bytes = 0
    for uid, meta in list(PENDING_UPLOADS.items()):
        fp = meta.get("file_path")
        if fp and os.path.exists(fp):
            sz = os.path.getsize(fp)
            total_bytes += sz
            entries.append((meta.get("created_at", 0.0), uid, fp, sz))

    entries.sort(key=lambda x: x[0])  # Eng eskisi birinchi
    while total_bytes > MAX_TMP_STORAGE_BYTES and entries:
        _, uid, fp, sz = entries.pop(0)
        try:
            os.remove(fp)
        except OSError:
            pass
        PENDING_UPLOADS.pop(uid, None)
        total_bytes -= sz


@dataclass(order=True)
class PrioritizedJob:
    priority: int  # 1 = VIP/Paid B2C, 2 = Standard Interactive, 3 = Heavy B2B Book Batch
    created_at: float = field(compare=True)
    job_id: str = field(compare=False)
    coro_factory: Callable[[], Awaitable[Any]] = field(compare=False)
    future: asyncio.Future = field(compare=False)


class SlaydHubTaskScheduler:
    """Render 512MB RAM uchun In-Process Priority Queue va Worker tizimi."""

    def __init__(self, max_queue_size: int = 50, concurrency: int = 2):
        self.queue: asyncio.PriorityQueue[PrioritizedJob] = asyncio.PriorityQueue(maxsize=max_queue_size)
        self.concurrency = concurrency
        self._workers: list[asyncio.Task] = []
        self._running = False

    async def start(self) -> None:
        if self._running:
            return
        self._running = True
        for i in range(self.concurrency):
            self._workers.append(asyncio.create_task(self._worker_loop(i)))

    async def submit(self, priority: int, job_id: str, coro_factory: Callable[[], Awaitable[Any]]) -> Any:
        loop = asyncio.get_running_loop()
        fut = loop.create_future()
        job = PrioritizedJob(
            priority=priority,
            created_at=time.monotonic(),
            job_id=job_id,
            coro_factory=coro_factory,
            future=fut,
        )
        await self.queue.put(job)
        return await fut

    async def _worker_loop(self, worker_id: int) -> None:
        while self._running:
            job = await self.queue.get()
            try:
                result = await job.coro_factory()
                if not job.future.done():
                    job.future.set_result(result)
            except Exception as exc:
                if not job.future.done():
                    job.future.set_exception(exc)
            finally:
                release_c_and_python_memory()
                self.queue.task_done()

    async def stop(self) -> None:
        self._running = False
        for w in self._workers:
            w.cancel()
