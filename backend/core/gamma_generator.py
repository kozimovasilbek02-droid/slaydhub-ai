# -*- coding: utf-8 -*-
"""
Gamma-Style Ultra-HD Presentation Generator
-------------------------------------------
HTML + CSS (Glassmorphism, Neon Gradients, Modern Cards) shablonlari asosida
Playwright yordamida 1920x1080 pikselli Ultra-HD Gamma-uslubidagi
taqdimotlarni to'liq generatsiya qiluvchi dvigatel.
"""

import os
import re
import asyncio
import tempfile
from typing import Dict, List, Any, Optional
from pptx import Presentation
from pptx.util import Inches

from backend.core.deck_builder import DeckBuilder


THEME_BACKGROUNDS = {
    "uzbek_blue": "linear-gradient(135deg, #0f172a 0%, #1e1b4b 50%, #090d16 100%)",
    "emerald_teal": "linear-gradient(135deg, #0f172a 0%, #064e3b 50%, #022c22 100%)",
    "dark_slate": "linear-gradient(135deg, #090d16 0%, #1e293b 50%, #0f172a 100%)",
    "modern_purple": "linear-gradient(135deg, #0f172a 0%, #3b0764 50%, #18022e 100%)",
    "crimson_ruby": "linear-gradient(135deg, #0f172a 0%, #450a0a 50%, #1c0404 100%)"
}


class GammaGenerator:
    """Gamma.app uslubidagi Ultra-HD taqdimot generatori."""

    @classmethod
    async def build(cls, spec: Dict[str, Any], output_path: str) -> str:
        """
        Playwright orqali HTML slaydlarni render qilib, 16:9 PPTX faylga yig'adi.
        Playwright mavjud bo'lmasa, avtomatik ravishda DeckBuilder ga o'tadi.
        """
        try:
            from playwright.async_api import async_playwright
        except ImportError:
            # Fallback to vector deck builder
            return DeckBuilder.build(spec, output_path)

        slides_data = spec.get("slides", [])
        theme_name = spec.get("theme", "uzbek_blue")
        default_bg = THEME_BACKGROUNDS.get(theme_name, THEME_BACKGROUNDS["uzbek_blue"])

        temp_dir = tempfile.mkdtemp(prefix="gamma_slides_")
        img_paths = []

        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True)
                page = await browser.new_page(viewport={"width": 1920, "height": 1080})

                for idx, s in enumerate(slides_data, 1):
                    html_content = cls._generate_slide_html(s, idx, len(slides_data), default_bg)
                    html_file = os.path.join(temp_dir, f"slide_{idx}.html")
                    with open(html_file, "w", encoding="utf-8") as f:
                        f.write(html_content)

                    await page.goto(f"file:///{html_file.replace(os.sep, '/')}")
                    await page.wait_for_timeout(100)
                    
                    img_file = os.path.join(temp_dir, f"slide_{idx}.png")
                    await page.screenshot(path=img_file, full_page=False)
                    img_paths.append(img_file)

                await browser.close()

            # PPTX ga joylash
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
            blank_layout = prs.slide_layouts[6]

            for img in img_paths:
                slide = prs.slides.add_slide(blank_layout)
                slide.shapes.add_picture(img, Inches(0), Inches(0), Inches(13.333), Inches(7.5))

            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            prs.save(output_path)
            return output_path

        except Exception as e:
            # Xatolik yuz bersa (masalan Playwright brauzeri topilmasa), xavfsiz Vektorli DeckBuilder ga o'tamiz
            return DeckBuilder.build(spec, output_path)
        finally:
            # Tozalash
            try:
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass

    @classmethod
    def _generate_slide_html(cls, data: Dict[str, Any], slide_num: int, total_slides: int, default_bg: str) -> str:
        badge = data.get("badge", "O'ZBEKISTONNING ENG YANGI TARIXI").upper()
        title = data.get("title", "")
        subtitle = data.get("subtitle", "")
        stype = data.get("type", "cards")
        bg = data.get("bg_gradient") or default_bg

        content_html = ""
        if stype == "hero":
            content_html = f"""
            <div style="margin-top: 80px;">
                <div style="font-size: 56px; font-weight: 800; line-height: 1.25; background: linear-gradient(90deg, #ffffff, #93c5fd); -webkit-background-clip: text; -webkit-text-fill-color: transparent; max-width: 1550px; margin-bottom: 32px;">{title}</div>
                <div style="font-size: 24px; color: #94a3b8; max-width: 1350px; line-height: 1.6;">{subtitle}</div>
            </div>
            """
        elif stype == "agenda":
            items_html = ""
            for item in data.get("items", []):
                num = item.get("num", "01")
                t = item.get("title", "")
                d = item.get("desc", "")
                items_html += f"""
                <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 20px; padding: 20px 28px; display: flex; align-items: center; gap: 24px;">
                    <div style="background: linear-gradient(135deg, #38bdf8, #818cf8); color: #0f172a; font-size: 22px; font-weight: 800; width: 52px; height: 52px; border-radius: 14px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">{num}</div>
                    <div>
                        <div style="font-size: 22px; font-weight: 700; color: #f8fafc; margin-bottom: 4px;">{t}</div>
                        <div style="font-size: 15px; color: #94a3b8;">{d}</div>
                    </div>
                </div>
                """
            content_html = f"""
            <div>
                <div style="font-size: 42px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 25px;">{title}</div>
                <div style="display: flex; flex-direction: column; gap: 14px;">
                    {items_html}
                </div>
            </div>
            """
        elif stype == "stats":
            stats_html = ""
            for m in data.get("metrics", []):
                v = m.get("value", "100%")
                lbl = m.get("label", "")
                dsc = m.get("desc", "")
                stats_html += f"""
                <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); backdrop-filter: blur(16px); border-radius: 24px; padding: 40px; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; box-shadow: 0 20px 40px rgba(0,0,0,0.3);">
                    <div style="font-size: 58px; font-weight: 800; background: linear-gradient(135deg, #38bdf8, #818cf8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 16px;">{v}</div>
                    <div style="font-size: 22px; font-weight: 700; color: #f8fafc; margin-bottom: 12px;">{lbl}</div>
                    <div style="font-size: 16px; color: #94a3b8; line-height: 1.5;">{dsc}</div>
                </div>
                """
            content_html = f"""
            <div>
                <div style="font-size: 40px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 35px;">{title}</div>
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 28px; height: 520px;">
                    {stats_html}
                </div>
            </div>
            """
        elif stype == "comparative":
            col1_pts = "".join([f"<div style='margin-bottom: 14px;'>• {p}</div>" for p in data.get("col1_points", [])])
            col2_pts = "".join([f"<div style='margin-bottom: 14px;'>• {p}</div>" for p in data.get("col2_points", [])])
            content_html = f"""
            <div>
                <div style="font-size: 40px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 35px;">{title}</div>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 32px; height: 520px;">
                    <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 24px; padding: 36px;">
                        <div style="font-size: 26px; font-weight: 700; color: #38bdf8; margin-bottom: 24px;">{data.get("col1_title", "1-Yo'nalish")}</div>
                        <div style="font-size: 18px; color: #cbd5e1; line-height: 1.6;">{col1_pts}</div>
                    </div>
                    <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 24px; padding: 36px;">
                        <div style="font-size: 26px; font-weight: 700; color: #818cf8; margin-bottom: 24px;">{data.get("col2_title", "2-Yo'nalish")}</div>
                        <div style="font-size: 18px; color: #cbd5e1; line-height: 1.6;">{col2_pts}</div>
                    </div>
                </div>
            </div>
            """
        elif stype == "conclusion":
            pts = "".join([f"<div style='margin-bottom: 18px; display: flex; align-items: center; gap: 12px;'><span style='color: #38bdf8;'>✔</span> {p}</div>" for p in data.get("points", [])])
            content_html = f"""
            <div>
                <div style="font-size: 42px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 35px;">{title}</div>
                <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); border-radius: 24px; padding: 48px;">
                    <div style="font-size: 22px; color: #e2e8f0; line-height: 1.7; margin-bottom: 30px;">
                        {pts}
                    </div>
                    <div style="font-size: 24px; font-weight: 700; color: #38bdf8; margin-top: 20px;">
                        {data.get("thank_you", "E'tiboringiz uchun rahmat!")}
                    </div>
                </div>
            </div>
            """
        else:  # cards
            cards_html = ""
            for c in data.get("cards", []):
                ctitle = c.get("title", "")
                p1 = c.get("p1", "")
                p2 = c.get("p2", "")
                cards_html += f"""
                <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.08); backdrop-filter: blur(16px); border-radius: 24px; padding: 36px; display: flex; flex-direction: column; justify-content: space-between; box-shadow: 0 20px 40px rgba(0,0,0,0.3);">
                    <div>
                        <div style="display: flex; align-items: center; gap: 10px; color: #38bdf8; font-size: 13px; font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; margin-bottom: 16px;">
                            <span style="width: 8px; height: 8px; background: #38bdf8; border-radius: 50%;"></span> BO'LIM KONSEPSIYASI
                        </div>
                        <div style="font-size: 24px; font-weight: 700; color: #f8fafc; margin-bottom: 20px; line-height: 1.3;">{ctitle}</div>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 14px;">
                        <div style="background: rgba(255,255,255,0.04); border-left: 4px solid #38bdf8; padding: 14px 18px; border-radius: 8px; font-size: 15px; color: #cbd5e1; line-height: 1.5;">{p1}</div>
                        <div style="background: rgba(255,255,255,0.04); border-left: 4px solid #818cf8; padding: 14px 18px; border-radius: 8px; font-size: 15px; color: #cbd5e1; line-height: 1.5;">{p2}</div>
                    </div>
                </div>
                """
            content_html = f"""
            <div>
                <div style="font-size: 40px; font-weight: 800; color: #ffffff; margin-top: 15px; margin-bottom: 35px;">{title}</div>
                <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 28px; height: 520px;">
                    {cards_html}
                </div>
            </div>
            """

        return f"""<!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', system-ui, -apple-system, sans-serif; }}
                body {{
                    width: 1920px;
                    height: 1080px;
                    background: {bg};
                    color: #ffffff;
                    padding: 80px 100px;
                    display: flex;
                    flex-direction: column;
                    justify-content: space-between;
                    overflow: hidden;
                    position: relative;
                }}
                .badge {{
                    display: inline-flex;
                    align-items: center;
                    gap: 8px;
                    background: rgba(56, 189, 248, 0.1);
                    border: 1px solid rgba(56, 189, 248, 0.3);
                    color: #38bdf8;
                    padding: 8px 18px;
                    border-radius: 30px;
                    font-size: 14px;
                    font-weight: 700;
                    letter-spacing: 1.5px;
                    text-transform: uppercase;
                    width: fit-content;
                }}
                .footer {{
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    border-top: 1px solid rgba(255,255,255,0.1);
                    padding-top: 24px;
                    color: #64748b;
                    font-size: 16px;
                }}
            </style>
        </head>
        <body>
            <div>
                <div class="badge">{badge}</div>
                {content_html}
            </div>
            <div class="footer">
                <div>AI Precision Studio & Gamma Engine</div>
                <div>{slide_num} / {total_slides}</div>
            </div>
        </body>
        </html>
        """
