# -*- coding: utf-8 -*-
import os
import time
import base64
import json
import asyncio
from typing import Dict, Any, Tuple, Optional
import aiohttp


class SoffAPIClient:
    """Soff.uz sotuvchi platformasiga haqiqiy 2-bosqichli (FormData) avtomatik yuklash mijozi."""

    BASE_URL = "https://api.soff.uz/api/v1"

    def __init__(self):
        self.token_refresh_url = os.getenv(
            "SOFF_REFRESH_URL", f"{self.BASE_URL}/auth/token/refresh/"
        )
        self.login_url = os.getenv("SOFF_LOGIN_URL", f"{self.BASE_URL}/auth/login/")
        self._auth_lock = asyncio.Lock()
        self._active_uploads: set[str] = set()
        self._cached_doc_ids: Dict[str, Any] = {}

    def is_jwt_expired(self, token: str, buffer_seconds: int = 180) -> bool:
        if not token:
            return True
        parts = token.split(".")
        if len(parts) != 3:
            return False  # JWT bo'lmagan statik token bo'lsa
        try:
            padded = parts[1] + "=" * ((4 - len(parts[1]) % 4) % 4)
            payload = json.loads(base64.urlsafe_b64decode(padded))
            exp = payload.get("exp", 0)
            return bool(exp and time.time() > (exp - buffer_seconds))
        except Exception:
            return False

    async def get_valid_token(
        self, session: aiohttp.ClientSession, force_refresh: bool = False
    ) -> Optional[str]:
        current_token = os.getenv("SELLER_TOKEN", "")
        if not force_refresh and current_token and not self.is_jwt_expired(current_token):
            return current_token

        async with self._auth_lock:
            current_token = os.getenv("SELLER_TOKEN", "")
            if not force_refresh and current_token and not self.is_jwt_expired(current_token):
                return current_token

            refresh_token = os.getenv("SOFF_REFRESH_TOKEN", "")
            if refresh_token:
                try:
                    async with session.post(
                        self.token_refresh_url,
                        json={"refresh": refresh_token},
                        timeout=aiohttp.ClientTimeout(total=15),
                    ) as resp:
                        if resp.status in (200, 201):
                            data = await resp.json()
                            new_access = data.get("access") or data.get("token")
                            if new_access:
                                os.environ["SELLER_TOKEN"] = new_access
                                if data.get("refresh"):
                                    os.environ["SOFF_REFRESH_TOKEN"] = data["refresh"]
                                return new_access
                except Exception:
                    pass

            phone = os.getenv("SOFF_PHONE", "")
            password = os.getenv("SOFF_PASSWORD", "")
            if phone and password:
                try:
                    async with session.post(
                        self.login_url,
                        json={"phone": phone, "password": password},
                        timeout=aiohttp.ClientTimeout(total=15),
                    ) as resp:
                        if resp.status in (200, 201):
                            data = await resp.json()
                            new_access = data.get("access") or data.get("token")
                            if new_access:
                                os.environ["SELLER_TOKEN"] = new_access
                                if data.get("refresh"):
                                    os.environ["SOFF_REFRESH_TOKEN"] = data["refresh"]
                                return new_access
                except Exception:
                    pass

        return os.getenv("SELLER_TOKEN")

    async def upload_product(
        self,
        pptx_path: str,
        title: str,
        description: str,
        price_uzs: int,
        category_id: int,
        tags: str = "",
    ) -> Tuple[bool, str, str]:
        upload_key = f"{pptx_path}_{title}"
        if upload_key in self._active_uploads:
            return False, "Ushbu fayl hozirda yuklanmoqda, iltimos kuting...", ""

        if not pptx_path or not os.path.exists(pptx_path):
            return False, "Fayl server keshidan topilmadi yoki muddati o'tgan.", ""

        self._active_uploads.add(upload_key)
        try:
            with open(pptx_path, "rb") as fh:
                file_bytes = fh.read()
            fn = os.path.basename(pptx_path)

            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=75)) as session:
                token = await self.get_valid_token(session)
                if not token:
                    return False, "Soff.uz SELLER_TOKEN o'rnatilmagan", ""

                # 1-Bosqich: Faylni yuklash (/seller/product-create-first/)
                doc_id = self._cached_doc_ids.get(upload_key)
                if not doc_id:
                    for attempt in range(1, 4):
                        headers = {
                            "Authorization": f"Bearer {token}",
                            "User-Agent": "SlaydHubUltra/4.1",
                        }
                        form1 = aiohttp.FormData()
                        form1.add_field("content_type", "file")
                        form1.add_field(
                            "file",
                            file_bytes,
                            filename=fn,
                            content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                        async with session.post(
                            f"{self.BASE_URL}/seller/product-create-first/",
                            headers=headers,
                            data=form1,
                        ) as resp1:
                            if resp1.status in (200, 201):
                                rj1 = await resp1.json()
                                doc_id = rj1.get("id")
                                if doc_id:
                                    self._cached_doc_ids[upload_key] = doc_id
                                    break
                            elif resp1.status == 401 and attempt == 1:
                                token = await self.get_valid_token(session, force_refresh=True) or token
                                continue
                            elif resp1.status in (429, 500, 502, 503, 504):
                                await asyncio.sleep(2 ** attempt)
                                continue
                            else:
                                err_text = await resp1.text()
                                return (
                                    False,
                                    f"1-bosqich (fayl yuklash) xatosi ({resp1.status}): {err_text[:120]}",
                                    "",
                                )

                if not doc_id:
                    return False, "Soff.uz 1-bosqichdan document ID olinmadi.", ""

                # 2-Bosqich: Mahsulot metama'lumotlarini tasdiqlash (/seller/product-create-second/)
                full_desc = description
                if tags and tags not in full_desc:
                    full_desc = f"{description}\n<p>{tags}</p>"

                for attempt in range(1, 4):
                    headers = {
                        "Authorization": f"Bearer {token}",
                        "User-Agent": "SlaydHubUltra/4.1",
                    }
                    form2 = aiohttp.FormData()
                    form2.add_field("title", str(title))
                    form2.add_field("category", str(category_id if category_id >= 100 else 1463))
                    form2.add_field("price", str(price_uzs))
                    form2.add_field("document", str(doc_id))
                    form2.add_field("description", str(full_desc))

                    async with session.post(
                        f"{self.BASE_URL}/seller/product-create-second/",
                        headers=headers,
                        data=form2,
                    ) as resp2:
                        if resp2.status in (200, 201):
                            rj2 = await resp2.json()
                            prod_id = rj2.get("id", "")
                            link = (
                                f"https://soff.uz/seller/product-update/{prod_id}/"
                                if prod_id
                                else "https://soff.uz"
                            )
                            self._cached_doc_ids.pop(upload_key, None)
                            return True, "Muvaffaqiyatli yuklandi!", link
                        elif resp2.status == 401 and attempt == 1:
                            token = await self.get_valid_token(session, force_refresh=True) or token
                            continue
                        elif resp2.status in (429, 500, 502, 503, 504):
                            await asyncio.sleep(2 ** attempt)
                            continue
                        else:
                            err_text = await resp2.text()
                            return (
                                False,
                                f"2-bosqich (karta yaratish) xatosi ({resp2.status}): {err_text[:120]}",
                                "",
                            )

                return False, "Soff.uz 2-bosqich maksimal urinishlardan so'ng bajarilmadi.", ""

        except Exception as e:
            return False, f"Yuklash jarayonida xatolik: {str(e)}", ""
        finally:
            self._active_uploads.discard(upload_key)
