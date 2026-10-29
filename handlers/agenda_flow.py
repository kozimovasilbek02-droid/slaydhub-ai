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
    save_draft,
    get_draft,
    delete_draft,
    use_free_credit,
    refund_free_credit,
    debit_user_balance,
    refund_user_balance,
)
from core.gemini_client import (
    call_gemini_cascade,
    STAGE1_OUTLINE_SCHEMA,
    STAGE2_SLIDES_SCHEMA,
    build_stage1_outline_prompt,
    build_stage2_locked_expansion_prompt,
)
from core.memory_guard import (
    run_ram_heavy_sync,
    PENDING_UPLOADS,
    enforce_tmp_quota,
    GLOBAL_SCHEDULER,
)
from engines.pptx_generator import generate_pptx_file

router = Router(name="agenda_flow")


class SlideCreationFSM(StatesGroup):
    waiting_for_topic = State()
    viewing_agenda = State()
    adding_agenda_item = State()


def get_agenda_editor_keyboard(theme: str) -> InlineKeyboardMarkup:
    """Rejani 0-token tahrirlash klaviaturasi."""
    kb = [
        [
            InlineKeyboardButton(text="➕ Band qo'shish", callback_data="agenda_add"),
            InlineKeyboardButton(text="🗑️ Oxirgi bandni o'chirish", callback_data="agenda_delete_last"),
        ],
        [
            InlineKeyboardButton(
                text=f"🎨 Mavzu: {theme.replace('_', ' ').title()}", callback_data="agenda_toggle_theme"
            ),
        ],
        [
            InlineKeyboardButton(
                text="🚀 Tasdiqlash va Slaydni Yaratish", callback_data="agenda_confirm_build"
            ),
        ],
        [
            InlineKeyboardButton(text="❌ Bekor qilish", callback_data="menu_home"),
        ],
    ]
    return InlineKeyboardMarkup(inline_keyboard=kb)


@router.callback_query(F.data == "menu_create_slide")
async def cb_start_slide_creation(callback: CallbackQuery, state: FSMContext):
    await state.set_state(SlideCreationFSM.waiting_for_topic)
    txt = (
        "🎨 <b>SLAYD YARATISH (SLIDEPRO ENGINE)</b>\n\n"
        "Iltimos, taqdimot qilmoqchi bo'lgan <b>mavzuingiz nomini</b> yoki qisqa matn yozib yuboring.\n"
        "<i>Masalan: «Sun'iy intellektning tibbiyotdagi o'rni va istiqbollari»</i>"
    )
    kb = [[InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")]]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
    await callback.answer()


@router.message(SlideCreationFSM.waiting_for_topic)
async def process_topic_input(message: Message, state: FSMContext):
    topic = (message.text or "").strip()
    if len(topic) < 3:
        await message.answer("⚠️ Iltimos, batafsilroq mavzu nomini kiriting:")
        return

    status_msg = await message.answer(
        "⏳ <i>1/2: Mavzu tahlil qilinmoqda va mantiqiy reja tuzilmoqda...</i>", parse_mode="HTML"
    )

    # 1-Bosqich: Outline Skeleton (~200 token)
    prompt = build_stage1_outline_prompt(topic, slide_count=10, doc_mode="pptx")
    outline_data, err = await call_gemini_cascade(
        prompt, schema=STAGE1_OUTLINE_SCHEMA, max_output_tokens=512
    )

    if not outline_data or "agenda" not in outline_data:
        await status_msg.edit_text(
            f"❌ Reja tuzishda xatolik yuz berdi: {html.escape(str(err))}\nIltimos, qayta urinib ko'ring."
        )
        return

    title = outline_data.get("title", topic)
    agenda = outline_data.get("agenda", [])
    theme = outline_data.get("recommended_theme", "academic_navy")
    cat_id = outline_data.get("category_id", 1463)
    if not isinstance(cat_id, int) or cat_id < 100:
        cat_id = 1463

    draft_data = {
        "title": title,
        "subtitle": outline_data.get("subtitle", ""),
        "category_id": cat_id,
        "theme": theme,
        "agenda": agenda,
        "raw_topic": topic,
    }
    await save_draft(message.from_user.id, "pptx", title, "outline", draft_data)
    await state.set_state(SlideCreationFSM.viewing_agenda)

    agenda_str = "\n".join(f"<b>{i}.</b> {html.escape(str(item))}" for i, item in enumerate(agenda, 1))
    msg_text = (
        f"📝 <b>Taqdimot:</b> «{html.escape(str(title))}»\n\n"
        f"<b>Mantiqiy Reja ({len(agenda)} ta slayd):</b>\n"
        f"{agenda_str}\n\n"
        f"<i>Siz ushbu rejaga yangi band qo'shishingiz, ortiqchasini o'chirishingiz yoki dizayn mavzusini o'zgartirishingiz mumkin (0 token).</i>"
    )
    await status_msg.edit_text(
        msg_text, reply_markup=get_agenda_editor_keyboard(theme), parse_mode="HTML"
    )


@router.callback_query(F.data == "agenda_toggle_theme")
async def cb_toggle_theme(callback: CallbackQuery):
    draft = await get_draft(callback.from_user.id, "pptx")
    if not draft:
        await callback.answer("Qoralama topilmadi", show_alert=True)
        return

    data = draft["data"]
    themes = ["academic_navy", "emerald_science", "tech_minimal", "warm_philosophy", "cyber_dark"]
    curr = data.get("theme", "academic_navy")
    next_idx = (themes.index(curr) + 1) % len(themes) if curr in themes else 0
    data["theme"] = themes[next_idx]

    await save_draft(callback.from_user.id, "pptx", draft["title"], draft["stage"], data)

    agenda = data.get("agenda", [])
    agenda_str = "\n".join(f"<b>{i}.</b> {html.escape(str(item))}" for i, item in enumerate(agenda, 1))
    msg_text = (
        f"📝 <b>Taqdimot:</b> «{html.escape(str(draft['title']))}»\n\n"
        f"<b>Mantiqiy Reja ({len(agenda)} ta slayd):</b>\n"
        f"{agenda_str}\n\n"
        f"<i>Dizayn mavzusi yangilandi: <b>{data['theme'].replace('_', ' ').title()}</b></i>"
    )
    await callback.message.edit_text(
        msg_text, reply_markup=get_agenda_editor_keyboard(data["theme"]), parse_mode="HTML"
    )
    await callback.answer(f"Mavzu: {data['theme']}")


@router.callback_query(F.data == "agenda_delete_last")
async def cb_delete_last_item(callback: CallbackQuery):
    draft = await get_draft(callback.from_user.id, "pptx")
    if not draft:
        await callback.answer("Qoralama topilmadi", show_alert=True)
        return

    data = draft["data"]
    agenda = data.get("agenda", [])
    if len(agenda) <= 3:
        await callback.answer("Eng kamida 3 ta reja bandi qolishi shart!", show_alert=True)
        return

    removed = agenda.pop()
    data["agenda"] = agenda
    await save_draft(callback.from_user.id, "pptx", draft["title"], draft["stage"], data)

    agenda_str = "\n".join(f"<b>{i}.</b> {html.escape(str(item))}" for i, item in enumerate(agenda, 1))
    msg_text = (
        f"📝 <b>Taqdimot:</b> «{html.escape(str(draft['title']))}»\n\n"
        f"<b>Mantiqiy Reja ({len(agenda)} ta slayd):</b>\n"
        f"{agenda_str}\n\n"
        f"<i>«{html.escape(str(removed))}» bandi o'chirildi.</i>"
    )
    await callback.message.edit_text(
        msg_text,
        reply_markup=get_agenda_editor_keyboard(data.get("theme", "academic_navy")),
        parse_mode="HTML",
    )
    await callback.answer("O'chirildi")


@router.callback_query(F.data == "agenda_add")
async def cb_add_item_prompt(callback: CallbackQuery, state: FSMContext):
    await state.set_state(SlideCreationFSM.adding_agenda_item)
    await callback.message.answer("➕ Qo'shmoqchi bo'lgan yangi reja bandi nomini yozib yuboring:")
    await callback.answer()


@router.message(SlideCreationFSM.adding_agenda_item)
async def process_add_item(message: Message, state: FSMContext):
    new_item = (message.text or "").strip()
    if not new_item:
        await message.answer("⚠️ Iltimos, band nomini matn ko'rinishida yuboring:")
        return

    draft = await get_draft(message.from_user.id, "pptx")
    if not draft:
        await message.answer("Qoralama topilmadi, boshqatdan boshlang.")
        await state.clear()
        return

    data = draft["data"]
    agenda = data.get("agenda", [])
    agenda.append(new_item)
    data["agenda"] = agenda

    await save_draft(message.from_user.id, "pptx", draft["title"], draft["stage"], data)
    await state.set_state(SlideCreationFSM.viewing_agenda)

    agenda_str = "\n".join(f"<b>{i}.</b> {html.escape(str(item))}" for i, item in enumerate(agenda, 1))
    msg_text = (
        f"📝 <b>Taqdimot:</b> «{html.escape(str(draft['title']))}»\n\n"
        f"<b>Mantiqiy Reja ({len(agenda)} ta slayd):</b>\n"
        f"{agenda_str}\n\n"
        f"<i>Yangi band muvaffaqiyatli qo'shildi!</i>"
    )
    await message.answer(
        msg_text,
        reply_markup=get_agenda_editor_keyboard(data.get("theme", "academic_navy")),
        parse_mode="HTML",
    )


@router.callback_query(F.data == "agenda_confirm_build")
async def cb_confirm_build(callback: CallbackQuery, state: FSMContext):
    user_id = callback.from_user.id
    draft = await get_draft(user_id, "pptx")
    if not draft:
        await callback.answer("Qoralama topilmadi", show_alert=True)
        return

    user = await get_or_create_user(user_id)
    used_free_credit = await use_free_credit(user_id)
    used_paid_balance = False
    slide_price = 3000

    if not used_free_credit:
        used_paid_balance = await debit_user_balance(user_id, slide_price)
        if not used_paid_balance:
            txt = (
                f"⚠️ <b>Balansingiz yetarli emas!</b>\n\n"
                f"Slayd narxi: {slide_price:,} so'm\n"
                f"Sizning balansingiz: {user.get('balance_uzs', 0):,} so'm\n\n"
                f"<i>Iltimos, Click yoki Payme orqali hisobingizni to'ldiring yoki do'stlaringizni taklif qilib bepul kreditga ega bo'ling!</i>"
            )
            kb = [
                [InlineKeyboardButton(text="💳 Hamyonni to'ldirish", callback_data="menu_cabinet")],
                [InlineKeyboardButton(text="👥 Referal orqali bepul olish", callback_data="menu_referral")],
            ]
            await callback.message.answer(
                txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML"
            )
            await callback.answer()
            return

    await callback.answer()
    status_msg = await callback.message.edit_text(
        "🎨 <i>2/2: Slaydlar to'liq dizayn qilinmoqda (Penpot palitralari, p16:morph, spiker nutqi)...</i>",
        parse_mode="HTML",
    )

    data = draft["data"]
    title = data.get("title", "Taqdimot")
    agenda = data.get("agenda", [])
    theme = data.get("theme", "academic_navy")
    category_id = data.get("category_id", 1463)

    async def _build_slide_job():
        stage2_prompt = build_stage2_locked_expansion_prompt(
            title, locked_agenda=agenda, context_excerpt=data.get("raw_topic", "")
        )
        slides_res, err = await call_gemini_cascade(
            stage2_prompt, schema=STAGE2_SLIDES_SCHEMA, max_output_tokens=4096
        )
        if not slides_res or "slides" not in slides_res:
            raise RuntimeError(err or "Slayd matnlarini generatsiya qilishda xatolik")

        full_spec = {
            "title": title,
            "subtitle": data.get("subtitle", ""),
            "theme": theme,
            "slides": slides_res.get("slides", []),
            "description": slides_res.get("description", ""),
            "tags": slides_res.get("tags", ""),
        }

        perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
        os.makedirs(perm_dir, exist_ok=True)
        out_path = os.path.join(perm_dir, f"slide_{user_id}_{int(time.time() * 1000)}.pptx")

        await run_ram_heavy_sync(generate_pptx_file, full_spec, out_path)
        enforce_tmp_quota(perm_dir)
        return full_spec, out_path

    # Priority 1 = Paid B2C, Priority 2 = Free Credit
    job_priority = 1 if used_paid_balance else 2
    try:
        full_spec, out_path = await GLOBAL_SCHEDULER.submit(
            priority=job_priority,
            job_id=f"pptx_{user_id}_{int(time.time())}",
            coro_factory=_build_slide_job,
        )
    except Exception as exc:
        # Xatolik yuz berganda foydalanuvchi pulini yoki bepul kreditini 100% qaytarish!
        if used_free_credit:
            await refund_free_credit(user_id)
        elif used_paid_balance:
            await refund_user_balance(user_id, slide_price)
        await status_msg.edit_text(
            f"❌ Slayd yaratishda xatolik: {html.escape(str(exc))}\n"
            f"<i>Hisobingizdan mablag' yechilmadi (qaytarildi). Qayta urinib ko'ring.</i>",
            parse_mode="HTML",
        )
        return

    # Soff.uz va DTM Test tugmalari ishlashi uchun PENDING_UPLOADS ga ro'yxatdan o'tkazish
    total_slides = len(full_spec["slides"]) + 3
    soff_price = 3000 if total_slides <= 6 else 7000 if total_slides <= 12 else 10000
    upload_id = f"{user_id}_{int(time.time())}"
    PENDING_UPLOADS[upload_id] = {
        "file_path": out_path,
        "title": title,
        "description": full_spec.get("description", f"{title} bo'yicha taqdimot"),
        "price": soff_price,
        "category_id": category_id,
        "tags": full_spec.get("tags", ""),
        "created_at": time.time(),
    }

    await delete_draft(user_id, "pptx")
    await state.clear()

    safe_fn = "".join(c for c in title if c.isalnum() or c in (" ", "_", "-")).strip()[:45] or "Taqdimot"
    doc_file = FSInputFile(out_path, filename=f"{safe_fn}.pptx")
    kb_after = [
        [
            InlineKeyboardButton(
                text="📋 Shu mavzuda DTM Test yaratish", callback_data=f"make_quiz_{upload_id}"
            )
        ],
        [
            InlineKeyboardButton(
                text="🛒 Soff.uz do'koniga sotuvga qo'yish", callback_data=f"soff_upload_{upload_id}"
            )
        ],
        [InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")],
    ]

    await callback.message.answer_document(
        doc_file,
        caption=(
            f"✅ <b>«{html.escape(str(title))}»</b> taqdimoti muvaffaqiyatli tayyorlandi!\n\n"
            f"🎨 <b>Mavzu:</b> {theme.replace('_', ' ').title()}\n"
            f"📊 <b>Slaydlar soni:</b> {total_slides} ta (Muqova, Reja va Xulosa bilan)\n"
            f"💰 <b>Soff.uz tavsiya narxi:</b> {soff_price:,} so'm"
        ),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_after),
        parse_mode="HTML",
    )
    try:
        await status_msg.delete()
    except Exception:
        pass
