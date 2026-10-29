# -*- coding: utf-8 -*-
import os
import tempfile
import time
from typing import Dict, Any
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from core.database import get_or_create_user, debit_user_balance, use_free_credit
from core.gemini_client import (
    call_gemini_cascade, DOCX_SCHEMA, DTM_QUIZ_SCHEMA,
    build_academic_docx_prompt, build_dtm_quiz_prompt
)
from core.memory_guard import run_ram_heavy_sync, PENDING_UPLOADS, enforce_tmp_quota
from engines.docx_academic import build_otm_academic_docx
from engines.dtm_quiz_engine import build_dtm_quiz_pdf
from engines.pdf_decomposer import extract_smart_pdf_text, decompose_book_into_presentations
from engines.pptx_generator import generate_pptx_file
from engines.pptx_translator import translate_pptx_deck
from billing.soff_client import SoffAPIClient

router = Router(name="document_router")
soff_client = SoffAPIClient()


class DocxFSM(StatesGroup):
    waiting_for_topic = State()
    waiting_for_type = State()


class QuizFSM(StatesGroup):
    waiting_for_topic = State()


# ── 1. AKADEMIK WORD GENERATORI (.DOCX) ───────────────────────────────────────
@router.callback_query(F.data == "menu_academic_word")
async def cb_start_docx(callback: CallbackQuery, state: FSMContext):
    await state.set_state(DocxFSM.waiting_for_topic)
    txt = (
        "📄 <b>OTM STANDARTIDAGI AKADEMIK WORD (.DOCX)</b>\n\n"
        "Iltimos, yaratmoqchi bo'lgan ilmiy ishingiz <b>mavzusi nomini</b> yuboring:\n"
        "<i>Masalan: «Raqamli iqtisodiyotda blokcheyn texnologiyalarining rivojlanishi»</i>"
    )
    await callback.message.edit_text(txt, parse_mode="HTML")


@router.message(DocxFSM.waiting_for_topic)
async def process_docx_topic(message: Message, state: FSMContext):
    topic = message.text.strip()
    await state.update_data(topic=topic)
    await state.set_state(DocxFSM.waiting_for_type)

    kb = [
        [
            InlineKeyboardButton(text="📘 Referat (10-15 bet)", callback_data="docx_type_referat"),
            InlineKeyboardButton(text="📗 Mustaqil ish (8-12 bet)", callback_data="docx_type_mustaqil"),
        ],
        [
            InlineKeyboardButton(text="📙 Kurs ishi (25-35 bet)", callback_data="docx_type_kurs_ishi"),
            InlineKeyboardButton(text="📕 Ilmiy Maqola (IMRAD)", callback_data="docx_type_maqola"),
        ],
        [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="menu_home")],
    ]
    await message.answer("Qaysi turdagi akademik hujjat kerak?", reply_markup=InlineKeyboardMarkup(inline_keyboard=kb))


@router.callback_query(F.data.startswith("docx_type_"))
async def process_docx_generation(callback: CallbackQuery, state: FSMContext):
    raw_type = callback.data.replace("docx_type_", "")
    doc_type_map = {
        "referat": ("REFERAT", 12, 5000),
        "mustaqil": ("MUSTAQIL ISH", 10, 5000),
        "kurs_ishi": ("KURS ISHI", 28, 15000),
        "maqola": ("ILMIY MAQOLA", 8, 7000),
    }
    doc_type_name, pages, price = doc_type_map.get(raw_type, ("REFERAT", 12, 5000))
    user_data = await state.get_data()
    topic = user_data.get("topic", "Ilmiy mavzu")

    user_id = callback.from_user.id
    user = await get_or_create_user(user_id)

    # Balans tekshirish
    has_free = await use_free_credit(user_id)
    if not has_free:
        paid = await debit_user_balance(user_id, price)
        if not paid:
            txt = (
                f"⚠️ <b>Balansingiz yetarli emas!</b>\n\n"
                f"Xizmat: {doc_type_name} ({pages} bet)\n"
                f"Narxi: {price:,} so'm\n"
                f"Sizning balansingiz: {user.get('balance_uzs', 0):,} so'm\n\n"
                f"<i>Iltimos, Click yoki Payme orqali hisobingizni to'ldiring:</i>"
            )
            kb = [[InlineKeyboardButton(text="💳 Hamyonni to'ldirish", callback_data="menu_cabinet")]]
            await callback.message.answer(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
            await callback.answer()
            return

    status_msg = await callback.message.edit_text(
        f"⏳ <b>{doc_type_name} tayyorlanmoqda...</b>\n"
        f"<i>(OTM Titul varag'i, Times New Roman 14pt, 1.5 interval, avtomatik Mundarija va GOST adabiyotlar tuzilmoqda)</i>",
        parse_mode="HTML"
    )

    prompt = build_academic_docx_prompt(topic, doc_type=doc_type_name, pages=pages)
    docx_spec, err = await call_gemini_cascade(prompt, schema=DOCX_SCHEMA, max_output_tokens=8192)

    if not docx_spec or "chapters" not in docx_spec:
        await status_msg.edit_text(f"❌ Word matnini generatsiya qilishda xatolik: {err}")
        await state.clear()
        return

    perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
    os.makedirs(perm_dir, exist_ok=True)
    out_docx_path = os.path.join(perm_dir, f"{user_id}_{doc_type_name.replace(' ', '_')}.docx")

    await run_ram_heavy_sync(build_otm_academic_docx, docx_spec, out_docx_path)
    enforce_tmp_quota(perm_dir)

    await state.clear()
    doc_file = FSInputFile(out_docx_path, filename=f"{topic[:35]}_{doc_type_name}.docx")
    await callback.message.answer_document(
        doc_file,
        caption=f"✅ <b>{doc_type_name} muvaffaqiyatli tayyorlandi!</b>\n\n"
                f"📄 <b>Mavzu:</b> «{topic}»\n"
                f"🎓 <b>Standart:</b> O'zbekiston OTM talablari (Times New Roman 14pt, 1.5 interval, chap 3cm, avto-mundarija)\n"
                f"✨ <i>10% keshbek hisobingizga qo'shildi!</i>",
        parse_mode="HTML"
    )
    try:
        await status_msg.delete()
    except Exception:
        pass


# ── 2. DTM TEST & VIKTORINA GENERATORI ────────────────────────────────────────
@router.callback_query(F.data == "menu_dtm_quiz")
async def cb_start_quiz(callback: CallbackQuery, state: FSMContext):
    await state.set_state(QuizFSM.waiting_for_topic)
    txt = (
        "📋 <b>DTM TEST VA VIKTORINA TUZISH</b>\n\n"
        "Qaysi mavzuda yoki fan bo'yicha DTM test savollari tuzish kerak?\n"
        "Mavzuni yozing yoki qisqa matn yuboring:"
    )
    await callback.message.edit_text(txt, parse_mode="HTML")


@router.message(QuizFSM.waiting_for_topic)
async def process_quiz_generation(message: Message, state: FSMContext):
    topic = message.text.strip()
    status_msg = await message.answer("⏳ <i>DTM standarti bo'yicha 10 ta test savollari va doirachali javoblar varaqasi tayyorlanmoqda...</i>", parse_mode="HTML")

    prompt = build_dtm_quiz_prompt(topic, question_count=10)
    quiz_spec, err = await call_gemini_cascade(prompt, schema=DTM_QUIZ_SCHEMA, max_output_tokens=4096)

    if not quiz_spec or "questions" not in quiz_spec:
        await status_msg.edit_text(f"❌ Test tuzishda xatolik yuz berdi: {err}")
        await state.clear()
        return

    perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
    os.makedirs(perm_dir, exist_ok=True)
    out_pdf_path = os.path.join(perm_dir, f"{message.from_user.id}_DTM_Test.pdf")

    await run_ram_heavy_sync(build_dtm_quiz_pdf, quiz_spec, out_pdf_path)
    enforce_tmp_quota(perm_dir)

    await state.clear()
    doc_file = FSInputFile(out_pdf_path, filename=f"DTM_Test_{topic[:30]}.pdf")
    await message.answer_document(
        doc_file,
        caption=f"✅ <b>«{topic}»</b> bo'yicha DTM nazorat testlari va javoblar varaqasi (A4 Blanka) tayyor!",
        parse_mode="HTML"
    )

    # 1-savolni Telegram viktorinasi sifatida chiqarish (Demo)
    first_q = quiz_spec["questions"][0]
    await message.answer_poll(
        question=f"1-savol: {first_q['question'][:250]}",
        options=first_q["options"][:4],
        type="quiz",
        correct_option_id=first_q.get("correct_index", 0),
        explanation=first_q.get("explanation", "")[:200],
        is_anonymous=False,
    )
    try:
        await status_msg.delete()
    except Exception:
        pass


# ── 3. HUJJAT YUKLASH: KATTA PDF KITOB VA PPTX TARJIMA ────────────────────────
@router.message(F.document)
async def handle_document_upload(message: Message):
    doc = message.document
    fn = doc.file_name.lower() if doc.file_name else ""
    user_id = message.from_user.id

    if fn.endswith(".pptx"):
        # Slayd Tarjimoni
        status_msg = await message.answer("⏳ <i>Taqdimot qabul qilindi. 100% shakl va dizaynni saqlagan holda o'zbek tiliga o'girilmoqda...</i>", parse_mode="HTML")
        file_io = await message.bot.download(doc)
        
        perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
        os.makedirs(perm_dir, exist_ok=True)
        in_pptx = os.path.join(perm_dir, f"in_{doc.file_name}")
        out_pptx = os.path.join(perm_dir, f"uz_{doc.file_name}")

        with open(in_pptx, "wb") as f:
            f.write(file_io.read())

        await run_ram_heavy_sync(translate_pptx_deck, in_pptx, out_pptx)
        enforce_tmp_quota(perm_dir)

        doc_file = FSInputFile(out_pptx, filename=f"Uzbek_{doc.file_name}")
        await message.answer_document(
            doc_file,
            caption="✅ <b>Taqdimot to'liq o'zbek tiliga o'girildi!</b>\nBarcha shakllar, shriftlar, jadvallar va koordinatalar 100% saqlandi.",
            parse_mode="HTML"
        )
        try:
            await status_msg.delete()
        except Exception:
            pass

    elif fn.endswith(".pdf"):
        # Katta Kitob Dekompozitori
        status_msg = await message.answer("📚 <i>Darslik qabul qilindi. Mundarija (TOC) tahlil qilinib, mustaqil taqdimotlarga ajratilmoqda...</i>", parse_mode="HTML")
        file_io = await message.bot.download(doc)

        perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
        os.makedirs(perm_dir, exist_ok=True)
        pdf_path = os.path.join(perm_dir, doc.file_name)

        with open(pdf_path, "wb") as f:
            f.write(file_io.read())

        # TOC extraction
        sample_text = await run_ram_heavy_sync(extract_smart_pdf_text, pdf_path, 27500)
        try:
            os.remove(pdf_path)
        except OSError:
            pass

        specs, err = await decompose_book_into_presentations(sample_text, doc.file_name)
        if not specs:
            await status_msg.edit_text(f"❌ Kitobni mavzularga ajratishda xatolik: {err}")
            return

        await status_msg.edit_text(f"🎨 Kitobdan <b>{len(specs)} ta mustaqil taqdimot</b> aniqlandi! Slaydlar yig'ilmoqda...", parse_mode="HTML")

        for idx, spec in enumerate(specs, 1):
            out_deck_path = os.path.join(perm_dir, f"Deck_{idx}_{user_id}.pptx")
            await run_ram_heavy_sync(generate_pptx_file, spec, out_deck_path)
            
            p_id = f"{user_id}_{idx}_{int(time.time())}"
            PENDING_UPLOADS[p_id] = {
                "file_path": out_deck_path,
                "title": spec["title"],
                "description": spec.get("description", ""),
                "price": spec.get("price_uzs", 7000),
                "category_id": spec.get("category_id", 1),
                "tags": spec.get("tags", ""),
                "created_at": time.time(),
            }

            kb = [
                [InlineKeyboardButton(text="🛒 Soff.uz do'koniga yuklash", callback_data=f"soff_upload_{p_id}")],
            ]
            doc_file = FSInputFile(out_deck_path, filename=f"{spec['title'][:40]}.pptx")
            await message.answer_document(
                doc_file,
                caption=f"📖 <b>{idx}-Qism: «{spec['title']}»</b>\n"
                        f"📊 Slaydlar: {len(spec.get('slides', [])) + 3} ta | Narxi: {spec.get('price_uzs', 7000):,} so'm",
                reply_markup=InlineKeyboardMarkup(inline_keyboard=kb),
                parse_mode="HTML"
            )

        enforce_tmp_quota(perm_dir)
        try:
            await status_msg.delete()
        except Exception:
            pass


@router.callback_query(F.data.startswith("soff_upload_"))
async def cb_soff_upload(callback: CallbackQuery):
    p_id = callback.data.replace("soff_upload_", "")
    meta = PENDING_UPLOADS.get(p_id)
    if not meta:
        await callback.answer("Fayl muddati o'tgan yoki topilmadi", show_alert=True)
        return

    await callback.message.edit_reply_markup(reply_markup=None)
    status_msg = await callback.message.answer(f"⏳ <b>«{meta['title']}»</b> Soff.uz platformasiga yuklanmoqda...", parse_mode="HTML")

    success, msg, link = await soff_client.upload_product(
        pptx_path=meta["file_path"],
        title=meta["title"],
        description=meta["description"],
        price_uzs=meta["price"],
        category_id=meta["category_id"],
        tags=meta["tags"],
    )

    if success:
        txt = (
            f"🎉 <b>Muvaffaqiyatli yuklandi!</b>\n\n"
            f"📦 <b>Mahsulot:</b> «{meta['title']}»\n"
            f"💰 <b>Narxi:</b> {meta['price']:,} so'm\n"
            f"🔗 <a href='{link}'>Soff.uz da ko'rish</a>"
        )
        await status_msg.edit_text(txt, parse_mode="HTML", disable_web_page_preview=False)
    else:
        await status_msg.edit_text(f"❌ Yuklashda xatolik: {msg}")
