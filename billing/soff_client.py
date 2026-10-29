# -*- coding: utf-8 -*-
import os
import time
import base64
import json
import asyncio
from typing import Dict, Any, Tuple, Optional
import aiohttp


class SoffAPIClient:
    """Soff.uz sotuvchi platformasiga 2-bosqichli avtomatik yuklash mijozi."""

    def __init__(self):
        self.api_url = os.getenv("SOFF_API_URL", "https://api.soff.uz/api/v1/seller/products/")
        self.token_refresh_url = os.getenv("SOFF_REFRESH_URL", "https://api.soff.uz/api/v1/auth/token/refresh/")
        self.login_url = os.getenv("SOFF_LOGIN_URL", "https://api.soff.uz/api/v1/auth/login/")
        self._auth_lock = asyncio.Lock()
        self._active_uploads: set[str] = set()

    def is_jwt_expired(self, token: str, buffer_seconds: int = 120) -> bool:
        if not token:
            return True
        parts = token.split(".")
        if len(parts) != 3:
            return True
        try:
            padded = parts[1] + "=" * ((4 - len(parts[1]) % 4) % 4)
            payload = json.loads(base64.urlsafe_b64decode(padded))
            exp = payload.get("exp", 0)
            return time.time() > (exp - buffer_seconds)
        except Exception:
            return True

    async def get_valid_token(self, session: aiohttp.ClientSession) -> Optional[str]:
        current_token = os.getenv("SELLER_TOKEN", "")
        if not self.is_jwt_expired(current_token):
            return current_token

        async with self._auth_lock:
            current_token = os.getenv("SELLER_TOKEN", "")
            if not self.is_jwt_expired(current_token):
                return current_token

            refresh_token = os.getenv("SOFF_REFRESH_TOKEN", "")
            if refresh_token:
                try:
                    async with session.post(self.token_refresh_url, json={"refresh": refresh_token}) as resp:
                        if resp.status == 200:
                            data = await resp.json()
                            new_access = data.get("access")
                            if new_access:
                                os.environ["SELLER_TOKEN"] = new_access
                                return new_access
                except Exception:
                    pass

            phone = os.getenv("SOFF_PHONE", "")
            password = os.getenv("SOFF_PASSWORD", "")
            if phone and password:
                try:
                    async with session.post(self.login_url, json={"phone": phone, "password": password}) as resp:
                        if resp.status == 200:
                            data = await resp.json()
                            new_access = data.get("access") or data.get("token")
                            if new_access:
                                os.environ["SELLER_TOKEN"] = new_access
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
        tags: str,
    ) -> Tuple[bool, str, str]:
        upload_key = f"{pptx_path}_{title}"
        if upload_key in self._active_uploads:
            return False, "Ushbu fayl allaqachon yuklanmoqda", ""

        self._active_uploads.add(upload_key)
        try:
            async with aiohttp.ClientSession(timeout=aiohttp.ClientTimeout(total=90)) as session:
                token = await self.get_valid_token(session)
                if not token:
                    return False, "Soff.uz SELLER_TOKEN olinmadi (avtorizatsiya xatosi)", ""

                headers = {"Authorization": f"Bearer {token}"}

                # 1-Bosqich: Faylni yuklash
                data = aiohttp.FormData()
                data.add_field(
                    "file",
                    open(pptx_path, "rb"),
                    filename=os.path.basename(pptx_path),
                    content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

                async with session.post("https://api.soff.uz/api/v1/files/upload/", headers=headers, data=data) as resp1:
                    if resp1.status not in (200, 201):
                        err_text = await resp1.text()
                        return False, f"1-bosqich (fayl yuklash) xatosi: HTTP {resp1.status} - {err_text[:100]}", ""
                    res1_data = await resp1.json()
                    doc_id = res1_data.get("id") or res1_data.get("file_id")

                # 2-Bosqich: Mahsulot kartasini yaratish
                product_payload = {
                    "title": title,
                    "description": description,
                    "price": price_uzs,
                    "category": category_id,
                    "tags": [t.strip() for t in tags.split(",") if t.strip()],
                    "file_id": doc_id,
                }

                async with session.post(self.api_url, headers=headers, json=product_payload) as resp2:
                    if resp2.status in (200, 201):
                        res2_data = await resp2.json()
                        prod_id = res2_data.get("id", "")
                        link = f"https://soff.uz/product/{prod_id}" if prod_id else "https://soff.uz"
                        return True, "Muvaffaqiyatli yuklandi!", link
                    else:
                        err_text = await resp2.text()
                        return False, f"2-bosqich (karta yaratish) xatosi: HTTP {resp2.status} - {err_text[:100]}", ""

        except Exception as e:
            return False, f"Yuklash jarayonida istisno: {str(e)}", ""
        finally:
            self._active_uploads.discard(upload_key)
