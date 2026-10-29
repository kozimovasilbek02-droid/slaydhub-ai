# -*- coding: utf-8 -*-
"""
SlaydHub & Soff.uz 24/7 Telegram Bot
Production-ready for Render.com & 24/7 Autonomous Hosting
Optimized with Gemini Multi-Model Fallback & Smart PDF Sampling
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

if not GEMINI_API_KEY:
    local_env_file = r"C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot\.env"
    if os.path.exists(local_env_file):
        try:
            with open(local_env_file, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("GEMINI_API_KEY="):
                        GEMINI_API_KEY = line.split("=", 1)[1].strip()
        except Exception:
            pass

if not BOT_TOKEN:
    print("[OGOHLANTIRISH] BOT_TOKEN muhit o'zgaruvchisi topilmadi!")

bot = Bot(token=BOT_TOKEN or "DUMMY_TOKEN")
dp = Dispatcher()

# Cache for generated presentations waiting for upload
PENDING_UPLOADS = {}

# Active Gemini models in priority order (Fast, Quota-Safe, Fallback)
GEMINI_MODELS = [
    "gemini-3.1-flash-lite",
    "gemini-flash-lite-latest",
    "gemini-3.8-flash",
    "gemini-flash-latest",
    "gemini-4-argon"
]

# ── HELPER: SMART PDF EXTRACTION ─────────────────────────────────────────────
def extract_smart_pdf_text(pdf_path: str) -> tuple[str, int]:
    """
    8.9 MB va undan katta PDF kitoblardan mundarija, kirish va
    asosiy boblardan namunalar ajratib oluvchi aqlli funksiya.
    """
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    
    sections = []
    
    # 1. Mundarija (Table of Contents) tekshirish
    toc = doc.get_toc()
    if toc:
        toc_lines = [f"- {item[1]} (bet {item[2]})" for item in toc[:40]]
        sections.append("MUNDARIJA VA ASOSIY BOBLAR:\n" + "\n".join(toc_lines))
    
    # 2. Dastlabki sahifalar (Mundarija va Kirish) - birinchi 15 sahifa
    start_samples = []
    for p_no in range(min(15, total_pages)):
        text = doc[p_no].get_text().strip()
        if text:
            start_samples.append(text)
    if start_samples:
        sections.append("[KITOB BOSHI VA KIRISH]:\n" + "\n".join(start_samples)[:12000])
        
    # 3. O'rta sahifalar (Asosiy mavzular)
    if total_pages > 20:
        mid_samples = []
        step = max(1, (total_pages - 20) // 10)
        for p_no in range(16, total_pages - 10, step):
            text = doc[p_no].get_text().strip()
            if text:
                mid_samples.append(f"--- Sahifa {p_no + 1} ---\n" + text[:1500])
            if len(mid_samples) >= 8:
                break
        if mid_samples:
            sections.append("[ASOSIY BOBLAR NAMUNALARI]:\n" + "\n".join(mid_samples)[:15000])
            
    # 4. Oxirgi sahifalar (Xulosalar)
    if total_pages > 30:
        end_samples = []
        for p_no in range(max(0, total_pages - 6), total_pages):
            text = doc[p_no].get_text().strip()
            if text:
                end_samples.append(text)
        if end_samples:
            sections.append("[XULOSA VA YAKUNIY QISMLAR]:\n" + "\n".join(end_samples)[:6000])
            
    doc.close()
    
    combined_text = "\n\n".join(sections)
    return combined_text, total_pages

# ── HELPER: GEMINI ANALYSIS ──────────────────────────────────────────────────
def analyze_book_with_gemini(sample_text: str, filename: str) -> tuple[list, str]:
    """Gemini API orqali kitobni semantik taqdimotlarga ajratish (Multi-Model Fallback)."""
    if not GEMINI_API_KEY:
        return [], "GEMINI_API_KEY o'rnatilmagan"

    prompt = f"""
Sen professional taqdimotlar arxitektori va ilmiy muharrirsan.
Quyidagi kitob/hujjatdan olingan matnni chuqur tahlil qilib, uni Soff.uz ta'limiy platformasi uchun 2 tadan 5 tagacha mustaqil, to'liq ochib berilgan PowerPoint taqdimotlariga ajratib ber.

Fayl nomi: {filename}
Kitobdan olingan mundarija va matn parchalari:
\"\"\"
{sample_text[:28000]}
\"\"\"

TALABLAR:
1. Har bir taqdimot toza o'zbek adabiy tilida (lotin yozuvida) bo'lsin.
2. Har bir taqdimotda 6 tadan 10 tagacha mustaqil, chuqur tahliliy slayd bo'lsin.
3. Har bir slaydda 3-4 tadan aniq, mazmunli asosiy punktlar (tushunchalar, formulalar, faktlar) bo'lsin.
4. Javobingni FAQAT quyidagi JSON formatida qaytar:

{{
  "presentations": [
    {{
      "title": "Taqdimot nomi (Lotin yozuvida)",
      "category_id": 1463,
      "subtitle": "Qisqa ilmiy-amaliy izoh",
      "slides": [
        {{"heading": "Slayd 1 sarlavhasi", "points": ["punkt 1", "punkt 2", "punkt 3"]}},
        {{"heading": "Slayd 2 sarlavhasi", "points": ["punkt 1", "punkt 2", "punkt 3"]}}
      ],
      "tags": "#tibbiyot #fan #soff #taqdimot #talaba",
      "description": "<h2>Mavzu</h2><p>Batafsil tushuntirish va o'quv dasturi...</p>"
    }}
  ]
}}
"""
    last_error = ""
    
    # Multi-model fallback loop
    for model in GEMINI_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GEMINI_API_KEY}"
        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "responseMimeType": "application/json",
                "temperature": 0.2
            }
        }
        try:
            print(f"🤖 Gemini so'rovi yuborilmoqda (model: {model})...")
            resp = requests.post(url, json=payload, timeout=50)
            
            if resp.status_code == 200:
                raw_text = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
                
                # Robust JSON extraction
                clean_text = raw_text.strip()
                match = re.search(r"(\{[\s\S]*\})", clean_text)
                if match:
                    clean_text = match.group(1)
                
                data = json.loads(clean_text)
                presentations = data.get("presentations", [])
                if presentations:
                    print(f"✅ Muvaffaqiyatli: {model} orqali {len(presentations)} ta taqdimot ajratildi!")
                    return presentations, "OK"
            else:
                last_error = f"{model} status {resp.status_code}: {resp.text[:120]}"
                print(f"⚠️ {last_error}")
        except Exception as e:
            last_error = f"{model} xatosi: {e}"
            print(f"⚠️ {last_error}")

    return [], last_error

# ── HELPER: PPTX GENERATION ──────────────────────────────────────────────────
def generate_pptx_file(spec: dict, out_path: str) -> int:
    """Slayd ma'lumotlari asosida toza 16:9 zamonaviy PPTX yaratadi."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    
    # 1. Titul slaydi (Cover Slide)
    s1 = prs.slides.add_slide(blank)
    
    # Yuqori ko'k chiziq / aksent
    top_bar = s1.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(0.1))
    
    tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = spec.get("title", "Ilmiy Taqdimot")
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(30, 58, 138)  # Deep Navy
    
    sub = tf1.add_paragraph()
    sub.text = spec.get("subtitle", "Ilmiy-amaliy ta'limiy taqdimot | O'zbek tilida")
    sub.font.size = Pt(20)
    sub.font.color.rgb = RGBColor(100, 116, 139)  # Slate Gray
    sub.space_before = Pt(16)
    
    info_p = tf1.add_paragraph()
    info_p.text = "SlaydHub AI & Soff.uz | Avtonom ta'limiy platformasi"
    info_p.font.size = Pt(14)
    info_p.font.color.rgb = RGBColor(148, 163, 184)
    info_p.space_before = Pt(24)
    
    # 2. Mazmun slaydlari
    slides_info = spec.get("slides", [])
    for idx, s_info in enumerate(slides_info, 1):
        sl = prs.slides.add_slide(blank)
        
        # Kichik indeks nishoni
        num_tb = sl.shapes.add_textbox(Inches(1.0), Inches(0.4), Inches(2.0), Inches(0.4))
        np = num_tb.text_frame.paragraphs[0]
        np.text = f"BOB {idx:02d} / {len(slides_info):02d}"
        np.font.size = Pt(12)
        np.font.bold = True
        np.font.color.rgb = RGBColor(37, 99, 235)  # Royal Blue
        
        # Sarlavha
        tb = sl.shapes.add_textbox(Inches(1.0), Inches(0.8), Inches(11.333), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s_info.get("heading", f"Mavzu {idx}")
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)  # Dark Charcoal
        
        # Punktlar
        tb_c = sl.shapes.add_textbox(Inches(1.0), Inches(2.1), Inches(11.333), Inches(4.8))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True
        points = s_info.get("points", [])
        for pt_idx, pt in enumerate(points):
            pp = tf_c.paragraphs[0] if pt_idx == 0 else tf_c.add_paragraph()
            pp.text = f"•   {pt}"
            pp.font.size = Pt(18)
            pp.font.color.rgb = RGBColor(51, 65, 85)
            pp.space_after = Pt(14)
            
    # 3. Yakuniy Xulosa slaydi
    s_end = prs.slides.add_slide(blank)
    tb_end = s_end.shapes.add_textbox(Inches(1.0), Inches(2.5), Inches(11.333), Inches(2.5))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True
    pe1 = tf_end.paragraphs[0]
    pe1.text = "E'tiboringiz uchun rahmat!"
    pe1.font.size = Pt(36)
    pe1.font.bold = True
    pe1.font.color.rgb = RGBColor(30, 58, 138)
    
    pe2 = tf_end.add_paragraph()
    pe2.text = f"{spec.get('title', '')} bo'yicha savollar va munozara uchun"
    pe2.font.size = Pt(18)
    pe2.font.color.rgb = RGBColor(100, 116, 139)
    pe2.space_before = Pt(14)
            
    prs.save(out_path)
    return len(prs.slides)

# ── BOT HANDLERS ─────────────────────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: types.Message):
    await message.answer(
        "👋 **Assalomu alaykum! SlaydHub & Soff.uz avtonom botiga xush kelibsiz!**\n\n"
        "📚 Menga istalgan **PDF kitob yoki o'quv qo'llanmani** yuboring.\n\n"
        "⚡ **Bot imkoniyatlari:**\n"
        "1. Katta kitoblarni aqlli tahlil qilib, mustaqil taqdimotlarga ajratadi;\n"
        "2. Toza o'zbek tilida (lotin yozuvida) professional **PowerPoint (.pptx)** yaratadi;\n"
        "3. Tayyor fayllarni taqdim etib, bitta tugma bilan **Soff.uz** ga to'g'ridan-to'g'ri joylaydi!\n\n"
        "📥 *Boshlash uchun PDF fayl yuboring!*",
        parse_mode="Markdown"
    )

@dp.message(F.document)
async def handle_document(message: types.Message):
    doc = message.document
    if not (doc.file_name and doc.file_name.lower().endswith(".pdf")):
        await message.answer("⚠️ Iltimos, faqat **PDF** formatidagi kitob yoki hujjat yuboring.")
        return

    file_size_mb = round((doc.file_size or 0) / (1024 * 1024), 1)
    status_msg = await message.answer(f"📥 **«{doc.file_name}»** ({file_size_mb} MB) qabul qilindi.\n⏳ 1/4: Fayl yuklab olinmoqda...")
    
    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_path = os.path.join(tmp_dir, doc.file_name)
        
        # 1. Download
        try:
            file_info = await bot.get_file(doc.file_id)
            await bot.download_file(file_info.file_path, pdf_path)
        except Exception as e:
            await status_msg.edit_text(f"❌ Faylni yuklab olishda xatolik: {e}")
            return
            
        # 2. Extract smart text
        await status_msg.edit_text(f"📖 2/4: Kitob sahifalari va mundarijasi o'qilmoqda...")
        try:
            sample_text, total_pages = extract_smart_pdf_text(pdf_path)
            if len(sample_text.strip()) < 80:
                await status_msg.edit_text("⚠️ Ushbu PDF faylda matn qatlami topilmadi (ehtimol skanerlangan rasm). Iltimos, raqamli PDF yuboring.")
                return
        except Exception as e:
            await status_msg.edit_text(f"❌ PDF faylni o'qishda xatolik: {e}")
            return
            
        await status_msg.edit_text(f"🧠 3/4: **Gemini AI** kitobni ({total_pages} sahifa) tahlil qilib, taqdimotlarga ajratmoqda...")
        
        # 3. Analyze with Gemini Multi-Model Fallback
        specs, err = analyze_book_with_gemini(sample_text, doc.file_name)
        if not specs:
            await status_msg.edit_text(f"❌ Kitobni mavzularga ajratishda xatolik: {err}\nIltimos, qayta urinib ko'ring.")
            return
            
        await status_msg.edit_text(f"🎨 4/4: Kitobdan **{len(specs)} ta mustaqil taqdimot** aniqlandi! Slaydlar (.pptx) shakllantirilmoqda...")
        
        # 4. Generate PPTX and send
        for idx, spec in enumerate(specs, 1):
            title = spec.get("title", f"Taqdimot {idx}")
            safe_title = "".join(c for c in title if c.isalnum() or c in (" ", "_", "-")).strip()
            pptx_filename = f"{safe_title}.pptx"
            pptx_path = os.path.join(tmp_dir, pptx_filename)
            
            slide_count = generate_pptx_file(spec, pptx_path)
            # Strict Soff.uz pricing rules: <=6 -> 3000, 7..12 -> 7000, >=13 -> 10000
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
                f"📄 Slaydlar soni: **{slide_count} ta**\n"
                f"💰 Tavsiya etilgan narx: **{price:,} so'm**\n"
                f"🏷️ Teglar: {spec.get('tags', '')}\n\n"
                f"👇 Quyidagi tugma orqali Soff.uz kabinetingizga 1-klikda joylang:"
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
    
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {SELLER_TOKEN}"})
    
    fp = item["file_path"]
    fn = os.path.basename(fp)
    
    # Step 1: Upload File
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

    # Step 2: Create Product Metadata
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
    await start_web_server()
    print("✅ Bot polling boshlandi.")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
