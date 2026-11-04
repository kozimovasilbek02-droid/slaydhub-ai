# -*- coding: utf-8 -*-
import os
import gc
import time
import ctypes
import asyncio
from dataclasses import dataclass, field
from typing import Callable, Any, Awaitable, Dict, Optional

# Linux glibc arenalarini cheklash (Render.com 512MB RAM konteyneri uchun)
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
    Sinxron C/XML funksiyalarini (fitz, python-docx, python-pptx, reportlab)
    faqat RAM_COMPUTE_LOCK ichida, alohida oqimda bajaradi va darhol RAMni tozalaydi.
    DIQQAT: Bu funksiyaga faqat sinxron (def) funksiya uzatilishi shart!
    """
    if asyncio.iscoroutinefunction(func):
        raise TypeError(
            f"run_ram_heavy_sync faqat sinxron funksiyalarni qabul qiladi, '{func.__name__}' esa async korutina!"
        )
    async with RAM_COMPUTE_LOCK:
        try:
            return await asyncio.to_thread(func, *args, **kwargs)
        finally:
            release_c_and_python_memory()


def enforce_tmp_quota(target_dir: str) -> None:
    """/tmp/soff_ready (shu jumladan cloud_templates) hajmini 60 MB kvota ichida ushlab turadi."""
    if not os.path.exists(target_dir):
        return
    now = time.time()

    # 1. PENDING_UPLOADS dagi muddati o'tgan fayllarni tozalash
    for uid, meta in list(PENDING_UPLOADS.items()):
        if now - meta.get("created_at", now) > UPLOAD_CACHE_TTL_SECONDS:
            fp = meta.get("file_path")
            if fp and os.path.exists(fp):
                try:
                    os.remove(fp)
                except OSError:
                    pass
            PENDING_UPLOADS.pop(uid, None)

    now = time.time()
    # 2. target_dir ichidagi fayllarni LRU bo'yicha tekshirish (in-flight .part va yangi uploadlarni himoyalash)
    active_upload_paths = {
        meta.get("file_path") for meta in PENDING_UPLOADS.values()
        if now - meta.get("created_at", 0) < 300  # Oxirgi 5 daqiqada yaratilgan fayllar
    }
    all_files = []
    total_bytes = 0
    for root, _, files in os.walk(target_dir):
        for fn in files:
            fp = os.path.join(root, fn)
            try:
                st = os.stat(fp)
                if fn.endswith(".part"):
                    # 5 daqiqadan oshgan osilib qolgan (zombie) .part fayllarni darhol o'chiramiz
                    if now - st.st_mtime > 300:
                        os.remove(fp)
                    continue
                total_bytes += st.st_size
                # Aktiv upload fayllariga yuqori mtime berib, eng oxirida o'chiriladigan qilamiz
                sort_key = st.st_mtime + (3600 if fp in active_upload_paths else 0)
                all_files.append((sort_key, fp, st.st_size))
            except OSError:
                pass

    all_files.sort(key=lambda x: x[0])  # Eng eski ishlatilgan fayl birinchi o'chadi
    while total_bytes > MAX_TMP_STORAGE_BYTES and all_files:
        _, fp, sz = all_files.pop(0)
        try:
            os.remove(fp)
            total_bytes -= sz
        except OSError:
            pass


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
        self.queue: Optional[asyncio.PriorityQueue[PrioritizedJob]] = None
        self.max_queue_size = max_queue_size
        self.concurrency = concurrency
        self._workers: list[asyncio.Task] = []
        self._running = False

    async def start(self) -> None:
        if self._running:
            return
        self.queue = asyncio.PriorityQueue(maxsize=self.max_queue_size)
        self._running = True
        for i in range(self.concurrency):
            self._workers.append(asyncio.create_task(self._worker_loop(i)))

    async def submit(self, priority: int, job_id: str, coro_factory: Callable[[], Awaitable[Any]]) -> Any:
        if not self._running or self.queue is None:
            await self.start()
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
        assert self.queue is not None
        while self._running:
            job = await self.queue.get()
            try:
                if job.future.cancelled():
                    continue
                result = await job.coro_factory()
                if not job.future.done():
                    job.future.set_result(result)
            except asyncio.CancelledError as cancel_exc:
                if not self._running:
                    self.queue.task_done()
                    break
                if not job.future.done():
                    job.future.cancel()
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
        self._workers.clear()


# Global Singleton Scheduler (barcha handlerlar foydalanishi uchun)
GLOBAL_SCHEDULER = SlaydHubTaskScheduler(max_queue_size=50, concurrency=2)
