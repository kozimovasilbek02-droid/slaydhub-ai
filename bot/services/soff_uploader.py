# -*- coding: utf-8 -*-
import os
import sys
import json
import asyncio
from typing import Dict, Any, Optional
from playwright.async_api import async_playwright

from bot.config import SOFF_SESSION_FILE


class SoffUploaderService:
    """
    Taqdimotlarni Playwright va oldindan saqlangan sessiya orqali
    seller.soff.uz ga avtomatik yuklash xizmati.
    """

    @staticmethod
    def is_session_available() -> bool:
        return SOFF_SESSION_FILE.exists()

    @staticmethod
    async def upload_presentation(
        pptx_path: str,
        title: str,
        price: int = 25000,
        category: str = "Taqdimot",
        tags: Optional[list] = None
    ) -> Dict[str, Any]:
        """
        Faylni seller.soff.uz ga avtomatik yuklaydi.
        """
        if not os.path.exists(pptx_path):
            return {"success": False, "error": f"Fayl topilmadi: {pptx_path}"}

        if not SOFF_SESSION_FILE.exists():
            return {"success": False, "error": "Soff.uz sessiya fayli topilmadi."}

        if not tags:
            tags = ["Taqdimot", "Prezentatsiya", "Slayd", "Ilmiy Ish"]

        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True)
                # Saqlangan sessiya bilan context ochish
                context = await browser.new_context(storage_state=str(SOFF_SESSION_FILE))
                page = await context.new_page()

                # Yuklash sahifasiga o'tish
                await page.goto("https://seller.soff.uz/seller/upload/scientific-works/create", timeout=45000)
                await page.wait_for_load_state("networkidle")
                await asyncio.sleep(2)

                # 1. Faylni yuklash
                file_input = page.locator('input[type="file"]')
                if await file_input.count() == 0:
                    await browser.close()
                    return {"success": False, "error": "Soff.uz ga kirib bo'lmadi yoki sessiya eskirgan."}

                await page.set_input_files('input[type="file"]', pptx_path)
                await asyncio.sleep(5)

                # 2. Kategoriya tanlash
                await page.locator(".ant-select-selector").first.click()
                await asyncio.sleep(0.5)
                await page.keyboard.type(category, delay=50)
                await asyncio.sleep(1.5)

                first_opt = page.locator(".ant-select-item-option").first
                if await first_opt.count() > 0:
                    await first_opt.click()
                else:
                    await page.keyboard.type("Pedagogika", delay=50)
                    await asyncio.sleep(1)
                    opt = page.locator(".ant-select-item-option").first
                    if await opt.count() > 0:
                        await opt.click()

                await asyncio.sleep(1)

                # 3. Sarlavha
                title_input = page.locator('input[placeholder*="mahsulot nomi"]')
                if await title_input.count() > 0:
                    await title_input.fill(title)

                # 4. Bepul emas (Pullik) switch
                bepul_switch = page.locator(".ant-switch")
                if await bepul_switch.count() > 0:
                    is_checked = "ant-switch-checked" in (await bepul_switch.get_attribute("class") or "")
                    if is_checked:
                        await bepul_switch.click()
                        await asyncio.sleep(0.5)

                # 5. Narx
                price_input = page.locator('input[type="number"], input.ant-input-number-input').first
                if await price_input.count() > 0:
                    await price_input.fill(str(price))

                # 6. CKEditor tavsif
                desc_html = f"<h2>{title}</h2><p>Mazkur taqdimot chuqur ilmiy va akademik asosda, 2026-yilgi standartlarga mos ravishda tayyorlangan. Mavzuga oid barcha tushunchalar, zamonaviy tahlillar va amaliy misollar jamlangan.</p>"
                await page.evaluate("""(htmlContent) => {
                    const el = document.querySelector('.ck-editor__editable');
                    if (el && el.ckeditorInstance) {
                        el.ckeditorInstance.setData(htmlContent);
                    }
                }""", desc_html)

                # 7. Teglar
                tag_search = page.locator(".ant-select-selection-search-input").nth(1)
                if await tag_search.count() > 0:
                    for t in tags[:5]:
                        await tag_search.fill(t)
                        await tag_search.press("Enter")
                        await asyncio.sleep(0.15)

                await asyncio.sleep(1)

                # 8. Yuborish (Submit)
                submit_btn = page.locator('button[type="submit"]')
                if await submit_btn.count() > 0:
                    await submit_btn.click()
                    await asyncio.sleep(6)

                await browser.close()
                return {"success": True, "message": "Taqdimot Soff.uz ga muvaffaqiyatli yuklandi va moderatsiyaga topshirildi!"}

        except Exception as e:
            return {"success": False, "error": f"Soff.uz ga yuklashda xatolik yuz berdi: {str(e)}"}
