# -*- coding: utf-8 -*-
import os
import time
import html
import tempfile
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from core.database import (
    get_or_create_user,
    debit_user_balance,
    refund_user_balance,
    use_free_credit,
    refund_free_credit,
)
from core.gemini_client import (
    call_gemini_cascade,
    DOCX_SCHEMA,
    DTM_QUIZ_SCHEMA,
    build_academic_docx_prompt,
    build_dtm_quiz_prompt,
)
from core.memory_guard import (
    run_ram_heavy_sync,
    PENDING_UPLOADS,
    enforce_tmp_quota,
    GLOBAL_SCHEDULER,
)
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
    kb = [[InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")]]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
    await callback.answer()


@router.message(DocxFSM.waiting_for_topic)
async def process_docx_topic(message: Message, state: FSMContext):
    topic = (message.text or "").strip()
    if len(topic) < 3:
        await message.answer("⚠️ Iltimos, mavzu nomini to'liqroq yozing:")
        return
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
    await message.answer(
        "Qaysi turdagi akademik hujjat kerak?",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb),
    )


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

    used_free_credit = await use_free_credit(user_id)
    used_paid_balance = False
    if not used_free_credit:
        used_paid_balance = await debit_user_balance(user_id, price)
        if not used_paid_balance:
            txt = (
                f"⚠️ <b>Balansingiz yetarli emas!</b>\n\n"
                f"Xizmat: {doc_type_name} ({pages} bet)\n"
                f"Narxi: {price:,} so'm\n"
                f"Sizning balansingiz: {user.get('balance_uzs', 0):,} so'm\n\n"
                f"<i>Iltimos, Click yoki Payme orqali hisobingizni to'ldiring:</i>"
            )
            kb = [[InlineKeyboardButton(text="💳 Hamyonni to'ldirish", callback_data="menu_cabinet")]]
            await callback.message.answer(
                txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML"
            )
            await callback.answer()
            return

    await callback.answer()
    status_msg = await callback.message.edit_text(
        f"⏳ <b>{doc_type_name} tayyorlanmoqda...</b>\n"
        f"<i>(OTM Titul varag'i, Times New Roman 14pt, 1.5 interval, avtomatik Mundarija va GOST adabiyotlar tuzilmoqda)</i>",
        parse_mode="HTML",
    )

    async def _build_docx_job():
        prompt = build_academic_docx_prompt(topic, doc_type=doc_type_name, pages=pages)
        docx_spec, err = await call_gemini_cascade(prompt, schema=DOCX_SCHEMA, max_output_tokens=8192)
        if not docx_spec or "chapters" not in docx_spec:
            raise RuntimeError(err or "Word matnini generatsiya qilishda xatolik")

        perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
        os.makedirs(perm_dir, exist_ok=True)
        out_docx_path = os.path.join(
            perm_dir, f"docx_{user_id}_{int(time.time() * 1000)}.docx"
        )
        await run_ram_heavy_sync(build_otm_academic_docx, docx_spec, out_docx_path)
        enforce_tmp_quota(perm_dir)
        return out_docx_path

    job_priority = 1 if used_paid_balance else 2
    try:
        out_docx_path = await GLOBAL_SCHEDULER.submit(
            priority=job_priority,
            job_id=f"docx_{user_id}_{int(time.time())}",
            coro_factory=_build_docx_job,
        )
    except Exception as exc:
        if used_free_credit:
            await refund_free_credit(user_id)
        elif used_paid_balance:
            await refund_user_balance(user_id, price)
        await status_msg.edit_text(
            f"❌ Word hujjatini yaratishda xatolik: {html.escape(str(exc))}\n"
            f"<i>Hisobingizdan mablag' yechilmadi (qaytarildi).</i>",
            parse_mode="HTML",
        )
        await state.clear()
        return

    await state.clear()
    safe_fn = "".join(c for c in topic if c.isalnum() or c in (" ", "_", "-")).strip()[:35] or "Ilmiy_ish"
    doc_file = FSInputFile(out_docx_path, filename=f"{safe_fn}_{doc_type_name}.docx")
    await callback.message.answer_document(
        doc_file,
        caption=(
            f"✅ <b>{doc_type_name} muvaffaqiyatli tayyorlandi!</b>\n\n"
            f"📄 <b>Mavzu:</b> «{html.escape(str(topic))}»\n"
            f"🎓 <b>Standart:</b> O'zbekiston OTM talablari (Times New Roman 14pt, 1.5 interval, chap 3cm, avto-mundarija)"
        ),
        parse_mode="HTML",
    )
    try:
        await status_msg.delete()
    except Exception:
        pass


# ── 2. DTM TEST & VIKTORINA GENERATORI ────────────────────────────────────────
async def _generate_and_send_dtm_quiz(target_message: Message, user_id: int, topic: str):
    status_msg = await target_message.answer(
        "⏳ <i>DTM standarti bo'yicha 10 ta test savollari va doirachali javoblar varaqasi tayyorlanmoqda...</i>",
        parse_mode="HTML",
    )

    prompt = build_dtm_quiz_prompt(topic, question_count=10)
    quiz_spec, err = await call_gemini_cascade(
        prompt, schema=DTM_QUIZ_SCHEMA, max_output_tokens=4096
    )

    if not quiz_spec or not quiz_spec.get("questions"):
        await status_msg.edit_text(f"❌ Test tuzishda xatolik yuz berdi: {html.escape(str(err))}")
        return

    perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
    os.makedirs(perm_dir, exist_ok=True)
    out_pdf_path = os.path.join(perm_dir, f"dtm_{user_id}_{int(time.time() * 1000)}.pdf")

    await run_ram_heavy_sync(build_dtm_quiz_pdf, quiz_spec, out_pdf_path)
    enforce_tmp_quota(perm_dir)

    safe_fn = "".join(c for c in topic if c.isalnum() or c in (" ", "_", "-")).strip()[:30] or "Test"
    doc_file = FSInputFile(out_pdf_path, filename=f"DTM_Test_{safe_fn}.pdf")
    await target_message.answer_document(
        doc_file,
        caption=f"✅ <b>«{html.escape(str(topic))}»</b> bo'yicha DTM nazorat testlari va javoblar varaqasi (A4 Blanka) tayyor!",
        parse_mode="HTML",
    )

    first_q = quiz_spec["questions"][0]
    options = [str(o)[:95] for o in first_q.get("options", [])[:4]]
    while len(options) < 4:
        options.append(f"Variant {len(options) + 1}")
    c_idx = first_q.get("correct_index", 0)
    if not isinstance(c_idx, int) or not (0 <= c_idx < 4):
        c_idx = 0

    await target_message.answer_poll(
        question=f"1-savol: {str(first_q.get('question', ''))[:250]}",
        options=options,
        type="quiz",
        correct_option_id=c_idx,
        explanation=str(first_q.get("explanation", ""))[:195],
        is_anonymous=False,
    )
    try:
        await status_msg.delete()
    except Exception:
        pass


@router.callback_query(F.data == "menu_dtm_quiz")
async def cb_start_quiz(callback: CallbackQuery, state: FSMContext):
    await state.set_state(QuizFSM.waiting_for_topic)
    txt = (
        "📋 <b>DTM TEST VA VIKTORINA TUZISH</b>\n\n"
        "Qaysi mavzuda yoki fan bo'yicha DTM test savollari tuzish kerak?\n"
        "Mavzuni yozing yoki qisqa matn yuboring:"
    )
    kb = [[InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")]]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
    await callback.answer()


@router.message(QuizFSM.waiting_for_topic)
async def process_quiz_generation(message: Message, state: FSMContext):
    topic = (message.text or "").strip()
    if len(topic) < 3:
        await message.answer("⚠️ Iltimos, mavzu nomini aniqroq yozing:")
        return
    await state.clear()
    await _generate_and_send_dtm_quiz(message, message.from_user.id, topic)


@router.callback_query(F.data.startswith("make_quiz_"))
async def cb_make_quiz_from_slide(callback: CallbackQuery):
    upload_id = callback.data.replace("make_quiz_", "")
    meta = PENDING_UPLOADS.get(upload_id)
    topic = meta["title"] if meta else "Ilmiy taqdimot mavzusi"
    await callback.answer("📋 DTM test tayyorlanmoqda...")
    await _generate_and_send_dtm_quiz(callback.message, callback.from_user.id, topic)


# ── 3. HUJJAT YUKLASH: KATTA PDF KITOB VA PPTX TARJIMA ────────────────────────
@router.message(F.document)
async def handle_document_upload(message: Message):
    doc = message.document
    fn = doc.file_name.lower() if doc.file_name else ""
    safe_orig_name = os.path.basename(doc.file_name or "document")
    user_id = message.from_user.id

    file_size_bytes = doc.file_size or 0
    file_size_mb = round(file_size_bytes / (1024 * 1024), 1)
    if file_size_bytes > 20 * 1024 * 1024 and not os.getenv("LOCAL_BOT_API_URL"):
        await message.answer(
            f"⚠️ <b>«{html.escape(safe_orig_name)}»</b> hajmi <b>{file_size_mb} MB</b>.\n"
            f"Standart Telegram Bot API 20 MB gacha bo'lgan fayllarni yuklab oladi. "
            f"Iltimos, faylni siqilgan holda yuboring yoki Local Bot API Server ulang.",
            parse_mode="HTML",
        )
        return

    if fn.endswith(".pptx"):
        status_msg = await message.answer(
            "⏳ <i>Taqdimot qabul qilindi. 100% shakl va dizaynni saqlagan holda o'zbek tiliga o'girilmoqda...</i>",
            parse_mode="HTML",
        )
        perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
        os.makedirs(perm_dir, exist_ok=True)
        ts = int(time.time() * 1000)
        in_pptx = os.path.join(perm_dir, f"in_{user_id}_{ts}.pptx")
        out_pptx = os.path.join(perm_dir, f"uz_{user_id}_{ts}.pptx")

        try:
            await message.bot.download(doc, destination=in_pptx)
            await GLOBAL_SCHEDULER.submit(
                priority=2,
                job_id=f"trans_{user_id}_{ts}",
                coro_factory=lambda: translate_pptx_deck(in_pptx, out_pptx),
            )
        except Exception as exc:
            await status_msg.edit_text(
                f"❌ Tarjima qilishda xatolik: {html.escape(str(exc))}", parse_mode="HTML"
            )
            return
        finally:
            if os.path.exists(in_pptx):
                try:
                    os.remove(in_pptx)
                except OSError:
                    pass

        enforce_tmp_quota(perm_dir)

        doc_file = FSInputFile(out_pptx, filename=f"Uzbek_{safe_orig_name}")
        await message.answer_document(
            doc_file,
            caption="✅ <b>Taqdimot to'liq o'zbek tiliga o'girildi!</b>\nBarcha shakllar, shriftlar, jadvallar va koordinatalar 100% saqlandi.",
            parse_mode="HTML",
        )
        try:
            await status_msg.delete()
        except Exception:
            pass

    elif fn.endswith(".pdf"):
        status_msg = await message.answer(
            f"📚 <i>«{html.escape(safe_orig_name)}» ({file_size_mb} MB) qabul qilindi. "
            f"Mundarija (TOC) tahlil qilinib, mustaqil taqdimotlarga ajratilmoqda...</i>",
            parse_mode="HTML",
        )
        perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
        os.makedirs(perm_dir, exist_ok=True)
        ts = int(time.time() * 1000)
        pdf_path = os.path.join(perm_dir, f"book_{user_id}_{ts}.pdf")

        try:
            await message.bot.download(doc, destination=pdf_path)
            sample_text = await run_ram_heavy_sync(extract_smart_pdf_text, pdf_path, 27500)
        except Exception as exc:
            await status_msg.edit_text(f"❌ PDF faylni o'qishda xatolik: {html.escape(str(exc))}")
            return
        finally:
            if os.path.exists(pdf_path):
                try:
                    os.remove(pdf_path)
                except OSError:
                    pass

        if len(sample_text.strip()) < 80:
            await status_msg.edit_text(
                "⚠️ Ushbu PDF faylda matn qatlami topilmadi (skanerlangan rasm yoki himoyalangan PDF). Iltimos, raqamli PDF yuboring."
            )
            return

        # Priority 3 = Heavy B2B Book Decomposition
        specs, err = await GLOBAL_SCHEDULER.submit(
            priority=3,
            job_id=f"book_{user_id}_{ts}",
            coro_factory=lambda: decompose_book_into_presentations(sample_text, safe_orig_name),
        )
        if not specs:
            await status_msg.edit_text(f"❌ Kitobni mavzularga ajratishda xatolik: {html.escape(str(err))}")
            return

        await status_msg.edit_text(
            f"🎨 Kitobdan <b>{len(specs)} ta mustaqil taqdimot</b> aniqlandi! Slaydlar (.pptx) shakllantirilmoqda...",
            parse_mode="HTML",
        )

        for idx, spec in enumerate(specs, 1):
            out_deck_path = os.path.join(perm_dir, f"Deck_{idx}_{user_id}_{int(time.time()*1000)}.pptx")
            await run_ram_heavy_sync(generate_pptx_file, spec, out_deck_path)

            slide_cnt = len(spec.get("slides", [])) + 3
            price = 3000 if slide_cnt <= 6 else 7000 if slide_cnt <= 12 else 10000
            p_id = f"{user_id}_{idx}_{int(time.time() * 1000)}"
            PENDING_UPLOADS[p_id] = {
                "file_path": out_deck_path,
                "title": spec.get("title", f"Taqdimot {idx}"),
                "description": spec.get("description", ""),
                "price": price,
                "category_id": spec.get("category_id", 1463),
                "tags": spec.get("tags", ""),
                "created_at": time.time(),
            }

            kb = [
                [
                    InlineKeyboardButton(
                        text="🛒 Soff.uz do'koniga yuklash", callback_data=f"soff_upload_{p_id}"
                    )
                ],
            ]
            safe_title = (
                "".join(c for c in spec.get("title", "") if c.isalnum() or c in (" ", "_", "-")).strip()[:40]
                or f"Taqdimot_{idx}"
            )
            doc_file = FSInputFile(out_deck_path, filename=f"{safe_title}.pptx")
            await message.answer_document(
                doc_file,
                caption=(
                    f"📖 <b>{idx}-Qism: «{html.escape(str(spec.get('title', '')))}»</b>\n"
                    f"📊 Slaydlar: {slide_cnt} ta | Narxi: {price:,} so'm"
                ),
                reply_markup=InlineKeyboardMarkup(inline_keyboard=kb),
                parse_mode="HTML",
            )

        enforce_tmp_quota(perm_dir)
        try:
            await status_msg.delete()
        except Exception:
            pass
    else:
        await message.answer(
            "⚠️ Iltimos, faqat <b>.pdf</b> (kitob tahlili uchun) yoki <b>.pptx</b> (tarjima uchun) fayl yuboring.",
            parse_mode="HTML",
        )


@router.callback_query(F.data.startswith(("soff_upload_", "publish_soff_")))
async def cb_soff_upload(callback: CallbackQuery):
    p_id = callback.data.replace("soff_upload_", "").replace("publish_soff_", "")
    meta = PENDING_UPLOADS.get(p_id)
    if not meta:
        await callback.answer("Fayl muddati o'tgan yoki topilmadi", show_alert=True)
        return

    await callback.answer("🚀 Soff.uz ga yuklanmoqda...")
    status_msg = await callback.message.answer(
        f"⏳ <b>«{html.escape(str(meta['title']))}»</b> Soff.uz platformasiga yuklanmoqda...",
        parse_mode="HTML",
    )

    success, msg, link = await soff_client.upload_product(
        pptx_path=meta["file_path"],
        title=meta["title"],
        description=meta["description"],
        price_uzs=meta["price"],
        category_id=meta["category_id"],
        tags=meta["tags"],
    )

    if success:
        kb_done = [[InlineKeyboardButton(text="✅ Soff.uz da ko'rish", url=link)]]
        try:
            await callback.message.edit_reply_markup(
                reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_done)
            )
        except Exception:
            pass
        txt = (
            f"🎉 <b>Muvaffaqiyatli yuklandi!</b>\n\n"
            f"📦 <b>Mahsulot:</b> «{html.escape(str(meta['title']))}»\n"
            f"💰 <b>Narxi:</b> {meta['price']:,} so'm\n"
            f"🔗 <a href='{link}'>Soff.uz da ko'rish</a>"
        )
        await status_msg.edit_text(txt, parse_mode="HTML", disable_web_page_preview=False)
    else:
        await status_msg.edit_text(
            f"❌ Yuklashda xatolik: {html.escape(str(msg))}", parse_mode="HTML"
        )
