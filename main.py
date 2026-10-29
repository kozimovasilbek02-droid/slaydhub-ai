# -*- coding: utf-8 -*-
"""
SlaydHub & Soff.uz 24/7 Telegram Bot
Production-ready for Render.com & 24/7 Hosting
"""

import os
import re
import sys
import json
import asyncio
import tempfile
import requests
import fitz  # PyMuPDF
from aiohttp import web
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import CommandStart, Command
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ── CONFIGURATION & SECRETS (Render Environment Variables) ───────────────────
BOT_TOKEN = os.getenv("BOT_TOKEN")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SELLER_TOKEN = os.getenv("SELLER_TOKEN")
PORT = int(os.getenv("PORT", 10000))

# Fallback for local testing from local files
if not SELLER_TOKEN:
    local_token_file = r"C:\Users\user\Desktop\Ish\Soff.uz\scripts\token.txt"
    if os.path.exists(local_token_file):
        try:
            with open(local_token_file, "r", encoding="utf-8") as f:
                SELLER_TOKEN = f.read().strip()
        except Exception:
            pass

if not BOT_TOKEN:
    print("[OGOHLANTIRISH] BOT_TOKEN muhit o'zgaruvchisi topilmadi!")

bot = Bot(token=BOT_TOKEN or "DUMMY_TOKEN")
dp = Dispatcher()

# Cache for generated presentations waiting for upload
PENDING_UPLOADS = {}

# ── HELPER: GEMINI ANALYSIS ──────────────────────────────────────────────────
def analyze_book_with_gemini(sample_text: str, filename: str) -> list:
    """Gemini API orqali kitobni semantik taqdimotlarga ajratish."""
    if not GEMINI_API_KEY:
        print("[XATO] GEMINI_API_KEY mavjud emas!")
        return []

    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
    
    prompt = f"""
Sen professional taqdimotlar arxitektori va ilmiy muharrirsan.
Quyidagi kitob/hujjatdan olingan matnni chuqur tahlil qilib, uni Soff.uz ta'limiy platformasi uchun bir nechta (hajmiga qarab 2 tadan 6 tagacha) mustaqil, to'liq ochib berilgan PowerPoint taqdimotlariga ajratib ber.

Fayl nomi: {filename}
Kitobdan olingan matn parchalari:
\"\"\"
{sample_text[:18000]}
\"\"\"

TALABLAR:
1. Har bir taqdimot toza o'zbek adabiy tilida (lotin yozuvida) bo'lsin.
2. Har bir taqdimotda 6 tadan 10 tagacha mustaqil slayd bo'lsin.
3. Har bir slaydda 3-4 tadan tushunarli, mazmunli asosiy punktlar bo'lsin.
4. Javobingni FAQAT va FAQAT to'g'ri JSON formatida qaytar (hech qanday qo'shimcha so'z yozma):

{{
  "presentations": [
    {{
      "title": "Taqdimot nomi (Lotin yozuvida)",
      "category_id": 1463,
      "subtitle": "Qisqa izoh",
      "slides": [
        {{"heading": "Slayd 1 sarlavhasi", "points": ["punkt 1", "punkt 2", "punkt 3"]}},
        {{"heading": "Slayd 2 sarlavhasi", "points": ["punkt 1", "punkt 2", "punkt 3"]}}
      ],
      "tags": "#teg1 #teg2 #teg3 #soff #taqdimot",
      "description": "<h2>Mavzu</h2><p>Batafsil tavsif...</p>"
    }}
  ]
}}
"""
    try:
        resp = requests.post(
            url,
            json={"contents": [{"parts": [{"text": prompt}]}]},
            timeout=60
        )
        if resp.status_code == 200:
            raw_text = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
            # Clean JSON markdown fences
            clean_json = re.sub(r"```json\s*", "", raw_text)
            clean_json = re.sub(r"```\s*$", "", clean_json).strip()
            data = json.loads(clean_json)
            return data.get("presentations", [])
        else:
            print(f"Gemini API javobi: {resp.status_code} -> {resp.text[:120]}")
    except Exception as e:
        print(f"Gemini API xatosi: {e}")
    return []

# ── HELPER: PPTX GENERATION ──────────────────────────────────────────────────
def generate_pptx_file(spec: dict, out_path: str) -> int:
    """Slayd ma'lumotlari asosida toza 16:9 formatdagi PPTX yaratadi."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    
    # 1. Titul slaydi
    s1 = prs.slides.add_slide(blank)
    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(3.0))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = spec.get("title", "Taqdimot")
    p1.font.size = Pt(38)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(30, 58, 138)
    
    sub = tf1.add_paragraph()
    sub.text = spec.get("subtitle", "Ilmiy-amaliy taqdimot | O'zbek tilida")
    sub.font.size = Pt(20)
    sub.font.color.rgb = RGBColor(100, 116, 139)
    sub.space_before = Pt(14)
    
    # 2. Mazmun slaydlari
    for s_info in spec.get("slides", []):
        sl = prs.slides.add_slide(blank)
        
        # Sarlavha
        tb = sl.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.3), Inches(1.0))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = s_info.get("heading", "Mavzu")
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)
        
        # Punktlar
        tb_c = sl.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(4.8))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True
        for pt in s_info.get("points", []):
            pp = tf_c.add_paragraph()
            pp.text = f"•  {pt}"
            pp.font.size = Pt(18)
            pp.font.color.rgb = RGBColor(51, 65, 85)
            pp.space_after = Pt(14)
            
    prs.save(out_path)
    return len(prs.slides)

# ── BOT HANDLERS ─────────────────────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: types.Message):
    await message.answer(
        "👋 **Assalomu alaykum! SlaydHub & Soff.uz avtonom botiga xush kelibsiz!**\n\n"
        "📚 Menga istalgan **PDF kitob yoki o'quv qo'llanmani** yuboring.\n\n"
        "⚡ **Bot nima qiladi:**\n"
        "1. Kitobni tahlil qilib, optimal mustaqil taqdimot mavzulariga ajratadi;\n"
        "2. Toza o'zbek tilida (lotin alifbosida) **PowerPoint (.pptx)** slaydlarini yaratadi;\n"
        "3. Sizga tayyor fayllarni tashlab, bitta tugma bilan **Soff.uz** ga joylash imkoniyatini beradi!",
        parse_mode="Markdown"
    )

@dp.message(F.document)
async def handle_document(message: types.Message):
    doc = message.document
    if not doc.file_name.lower().endswith(".pdf"):
        await message.answer("⚠️ Iltimos, faqat **PDF** formatidagi kitob yoki hujjat yuboring.")
        return

    status_msg = await message.answer(f"⏳ **«{doc.file_name}»** qabul qilindi. Matn ajratilmoqda...")
    
    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_path = os.path.join(tmp_dir, doc.file_name)
        
        # 1. Download
        file_info = await bot.get_file(doc.file_id)
        await bot.download_file(file_info.file_path, pdf_path)
        
        # 2. Extract sample text
        try:
            pdf_doc = fitz.open(pdf_path)
            total_pages = len(pdf_doc)
            samples = []
            for p_no in range(min(40, total_pages)):
                txt = pdf_doc[p_no].get_text()
                if txt.strip():
                    samples.append(txt)
            pdf_doc.close()
            sample_text = "\n".join(samples)
        except Exception as e:
            await status_msg.edit_text(f"❌ PDF faylni o'qishda xatolik: {e}")
            return
            
        await status_msg.edit_text(f"🧠 **Gemini 4 Argon** kitobni ({total_pages} sahifa) tahlil qilib, taqdimotlarga ajratmoqda...")
        
        # 3. Analyze with Gemini
        specs = analyze_book_with_gemini(sample_text, doc.file_name)
        if not specs:
            await status_msg.edit_text("❌ Kitobni mavzularga ajratishda xatolik yuz berdi. Iltimos, qayta urinib ko'ring.")
            return
            
        await status_msg.edit_text(f"✨ Kitobdan **{len(specs)} ta mustaqil taqdimot** aniqlandi! Slaydlar (.pptx) shakllantirilmoqda...")
        
        # 4. Generate PPTX and send
        for idx, spec in enumerate(specs, 1):
            title = spec.get("title", f"Taqdimot {idx}")
            safe_title = "".join(c for c in title if c.isalnum() or c in (" ", "_", "-")).strip()
            pptx_filename = f"{safe_title}.pptx"
            pptx_path = os.path.join(tmp_dir, pptx_filename)
            
            slide_count = generate_pptx_file(spec, pptx_path)
            price = 3000 if slide_count <= 6 else 7000 if slide_count <= 12 else 10000
            
            # Save into pending uploads
            upload_id = f"up_{idx}_{int(asyncio.get_event_loop().time())}"
            full_desc = f"{spec.get('description', '')}\n<p><strong>Jami slaydlar soni:</strong> {slide_count} ta slayd</p>\n<p>{spec.get('tags', '')}</p>"
            
            # Permanent copy for upload
            perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
            os.makedirs(perm_dir, exist_ok=True)
            perm_pptx = os.path.join(perm_dir, pptx_filename)
            import shutil
            shutil.copy2(pptx_path, perm_pptx)
            
            PENDING_UPLOADS[upload_id] = {
                "file_path": perm_pptx,
                "title": title,
                "category_id": spec.get("category_id", 1463),
                "price": price,
                "description": full_desc
            }
            
            kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🚀 Soff.uz ga yuklash", callback_data=f"soff_{upload_id}")]
            ])
            
            caption = (
                f"📊 **{idx}-Taqdimot: {title}**\n\n"
                f"📄 Slaydlar: **{slide_count} ta**\n"
                f"💰 Narxi: **{price:,} so'm**\n"
                f"🏷️ Teglar: {spec.get('tags', '')}\n\n"
                f"👇 Quyidagi tugma orqali Soff.uz ga bir klikda joylashingiz mumkin:"
            )
            
            await message.answer_document(
                FSInputFile(perm_pptx, filename=pptx_filename),
                caption=caption,
                reply_markup=kb,
                parse_mode="Markdown"
            )
            
        await status_msg.delete()

@dp.callback_query(F.data.startswith("soff_"))
async def handle_soff_upload(call: types.CallbackQuery):
    upload_id = call.data.replace("soff_", "")
    item = PENDING_UPLOADS.get(upload_id)
    if not item:
        await call.answer("⚠️ Ushbu fayl ma'lumoti eskirgan yoki topilmadi.", show_alert=True)
        return

    await call.answer("🚀 Soff.uz ga yuklanmoqda...", show_alert=False)
    
    # Upload to Soff.uz API
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {SELLER_TOKEN}"})
    
    fp = item["file_path"]
    fn = os.path.basename(fp)
    
    # Step 1
    doc_id = None
    try:
        with open(fp, "rb") as fh:
            r1 = session.post(
                "https://api.soff.uz/api/v1/seller/product-create-first/",
                files={"file": (fn, fh, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
                data={"content_type": "file"},
                timeout=45
            )
        if r1.status_code in (200, 201):
            doc_id = r1.json().get("id")
    except Exception as e:
        await call.message.reply(f"❌ Yuklashda ulanish xatosi: {e}")
        return

    if not doc_id:
        await call.message.reply(f"❌ Soff.uz Step 1 xatosi: {r1.status_code if 'r1' in locals() else 'error'}")
        return

    # Step 2
    try:
        mp = [
            ("title", (None, item["title"])),
            ("category", (None, str(item["category_id"]))),
            ("price", (None, str(item["price"]))),
            ("document", (None, str(doc_id))),
            ("description", (None, item["description"]))
        ]
        r2 = session.post("https://api.soff.uz/api/v1/seller/product-create-second/", files=mp, timeout=30)
        if r2.status_code in (200, 201):
            new_id = r2.json().get("id")
            kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Soff.uz da ko'rish", url=f"https://soff.uz/seller/product-update/{new_id}/")]
            ])
            await call.message.edit_reply_markup(reply_markup=kb)
            await call.message.reply(f"🎉 **Muvaffaqiyatli yuklandi!**\nSoff.uz Mahsulot ID: `{new_id}`", parse_mode="Markdown")
        else:
            await call.message.reply(f"❌ Soff.uz Step 2 xatosi: {r2.text[:120]}")
    except Exception as e:
        await call.message.reply(f"❌ Xatolik yuz berdi: {e}")

# ── WEB HEALTHCHECK SERVER (RENDER KEEPALIVE) ──────────────────────────────
async def handle_ping(request):
    return web.json_response({"status": "ok", "service": "SlaydHub & Soff.uz Bot", "uptime": "24/7"})

async def start_web_server():
    app = web.Application()
    app.router.add_get("/", handle_ping)
    app.router.add_get("/health", handle_ping)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    print(f"🚀 Web Health-Check server port {PORT} da ishga tushdi.")

# ── MAIN RUNNER ─────────────────────────────────────────────────────────────
async def main():
    print("🤖 SlaydHub Telegram bot ishga tushmoqda...")
    # 1. Start web server for Render keepalive
    await start_web_server()
    # 2. Start bot polling
    print("✅ Bot polling boshlandi.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
