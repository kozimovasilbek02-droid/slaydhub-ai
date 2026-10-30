# -*- coding: utf-8 -*-
"""
handlers/agenda_flow.py
SlaydHub Ultra v4.3 - Professional Template-First Presentation Flow
1. Topic Input -> 2. Smart Category Suggestion -> 3. Cover Photo Carousel (1,000+ Templates)
4. Exact Slide Count Selection (5, 8, 10, 15, 20) -> 5. Tailored Agenda Editing -> 6. Prune & Build
"""

import os
import time
import html
import tempfile
from typing import Dict, Any, List

from aiogram import Router, F
from aiogram.types import (
    Message,
    CallbackQuery,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    FSInputFile,
    InputMediaPhoto,
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from core.database import (
    get_or_create_user,
    debit_user_balance,
    refund_user_balance,
    use_free_credit,
    refund_free_credit,
    save_draft,
    get_draft,
    delete_draft,
)
from core.gemini_client import (
    call_gemini_cascade,
    build_stage1_outline_prompt,
    build_stage2_locked_expansion_prompt,
    STAGE1_OUTLINE_SCHEMA,
    STAGE2_SLIDES_SCHEMA,
)
from core.memory_guard import (
    run_ram_heavy_sync,
    PENDING_UPLOADS,
    enforce_tmp_quota,
    GLOBAL_SCHEDULER,
)
from core.template_catalog import GLOBAL_TEMPLATE_CATALOG, CATEGORIES
from engines.pptx_generator import generate_pptx_file

router = Router(name="agenda_flow")


class SlideCreationFSM(StatesGroup):
    waiting_for_topic = State()
    selecting_category = State()
    selecting_template = State()
    selecting_slide_count = State()
    viewing_agenda = State()
    adding_agenda_item = State()


# ── 1. KLAVIATURALAR ──────────────────────────────────────────────────────────

def get_category_selection_keyboard(rec_cat: str) -> InlineKeyboardMarkup:
    """Toifalar menyusi (tavsiya etilgan toifa ⭐ belgisi bilan)."""
    kb = []
    cat_items = list(CATEGORIES.items())
    for i in range(0, len(cat_items), 2):
        row = []
        for key, info in cat_items[i : i + 2]:
            star = " ⭐ (Tavsiya)" if key == rec_cat else ""
            row.append(
                InlineKeyboardButton(
                    text=f"{info['emoji']} {info['badge']}{star}",
                    callback_data=f"sel_cat_{key}",
                )
            )
        kb.append(row)
    kb.append([InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")])
    return InlineKeyboardMarkup(inline_keyboard=kb)


def get_template_carousel_keyboard(
    cat_key: str, curr_idx: int, total_tpls: int, tpl_id: str
) -> InlineKeyboardMarkup:
    """Karusel boshqaruvi: Oldingi / Keyingi va Tanlash."""
    nav_row = [
        InlineKeyboardButton(
            text="◀️ Oldingi", callback_data=f"tpl_nav_{cat_key}_{(curr_idx - 1) % total_tpls}"
        ),
        InlineKeyboardButton(text=f"{curr_idx + 1} / {total_tpls}", callback_data="noop"),
        InlineKeyboardButton(
            text="Keyingi ▶️", callback_data=f"tpl_nav_{cat_key}_{(curr_idx + 1) % total_tpls}"
        ),
    ]
    choose_row = [
        InlineKeyboardButton(
            text="✅ Ushbu shablonni tanlash", callback_data=f"tpl_choose_{tpl_id}"
        )
    ]
    back_row = [
        InlineKeyboardButton(text="🗂️ Boshqa toifani tanlash", callback_data="agenda_change_cat")
    ]
    return InlineKeyboardMarkup(inline_keyboard=[nav_row, choose_row, back_row])


def get_slide_count_keyboard(tpl_id: str) -> InlineKeyboardMarkup:
    """Slaydlar sonini tanlash klaviaturasi."""
    kb = [
        [
            InlineKeyboardButton(text="⚡ 5 ta slayd", callback_data="count_5"),
            InlineKeyboardButton(text="📄 8 ta slayd", callback_data="count_8"),
        ],
        [
            InlineKeyboardButton(text="📊 10 ta slayd", callback_data="count_10"),
            InlineKeyboardButton(text="📚 15 ta slayd", callback_data="count_15"),
        ],
        [
            InlineKeyboardButton(text="🎓 20 ta slayd", callback_data="count_20"),
        ],
        [
            InlineKeyboardButton(text="◀️ Shablonni almashtirish", callback_data="agenda_change_cat"),
        ],
    ]
    return InlineKeyboardMarkup(inline_keyboard=kb)


def get_agenda_editor_keyboard(template_title: str) -> InlineKeyboardMarkup:
    """Rejani tahrirlash va tasdiqlash klaviaturasi."""
    short_title = template_title[:24] + "..." if len(template_title) > 24 else template_title
    kb = [
        [
            InlineKeyboardButton(text="➕ Band qo'shish", callback_data="agenda_add"),
            InlineKeyboardButton(text="🗑️ Oxirgi bandni o'chirish", callback_data="agenda_delete_last"),
        ],
        [
            InlineKeyboardButton(
                text=f"🎨 Shablon: {short_title}", callback_data="agenda_change_cat"
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


# ── 2. OQIM BOSHQARUVI ────────────────────────────────────────────────────────

@router.callback_query(F.data == "menu_create_slide")
async def cb_start_slide_creation(callback: CallbackQuery, state: FSMContext):
    await state.set_state(SlideCreationFSM.waiting_for_topic)
    txt = (
        "🎨 <b>PROFESSIONAL SLAYD YARATISH</b>\n\n"
        "Taqdimot qilmoqchi bo'lgan <b>mavzuingiz nomini</b> yoki qisqa annotatsiyani yuboring.\n\n"
        "<i>Misol: «Kardio-reabilitatsiyaning zamonaviy bosqichlari» yoki «Startap biznes rejasi»</i>"
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

    # Tavsiya etilgan toifani aniqlash
    rec_cat = GLOBAL_TEMPLATE_CATALOG.detect_category(topic)
    await state.update_data(raw_topic=topic, rec_cat=rec_cat)
    await state.set_state(SlideCreationFSM.selecting_category)

    rec_info = CATEGORIES.get(rec_cat, CATEGORIES["minimal"])
    txt = (
        f"🎯 <b>«{html.escape(topic)}»</b> mavzusi qabul qilindi!\n\n"
        f"Tizim mavzuingizga mos toifani aniqladi: <b>{rec_info['emoji']} {rec_info['name']}</b> ⭐\n\n"
        f"Quyidagi toifalardan birini tanlang va tayyor 1000+ shablonlar galereyasidan o'zingizga ma'qul dizaynni tanlang:"
    )
    await message.answer(
        txt, reply_markup=get_category_selection_keyboard(rec_cat), parse_mode="HTML"
    )


@router.callback_query(F.data.startswith("sel_cat_"))
async def cb_select_category(callback: CallbackQuery, state: FSMContext):
    cat_key = callback.data.split("_")[-1]
    await state.update_data(selected_cat=cat_key)
    await show_template_card(callback, state, cat_key, tpl_idx=0)
    await callback.answer()


@router.callback_query(F.data == "agenda_change_cat")
async def cb_reselect_category(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    rec_cat = data.get("rec_cat", "minimal")
    txt = "🗂️ Taqdimotingiz uchun dizayn toifasini tanlang:"
    try:
        await callback.message.delete()
    except Exception:
        pass
    await callback.message.answer(
        txt, reply_markup=get_category_selection_keyboard(rec_cat), parse_mode="HTML"
    )
    await state.set_state(SlideCreationFSM.selecting_category)
    await callback.answer()


async def show_template_card(
    callback: CallbackQuery, state: FSMContext, cat_key: str, tpl_idx: int
) -> None:
    """Muqova rasmi va ma'lumotlari bilan shablon kartochkasini ko'rsatadi."""
    templates = GLOBAL_TEMPLATE_CATALOG.get_templates_by_category(cat_key)
    if not templates:
        await callback.message.answer("Ushbu toifada hozircha shablonlar mavjud emas.")
        return

    curr_idx = tpl_idx % len(templates)
    tpl = templates[curr_idx]
    cat_info = CATEGORIES.get(cat_key, CATEGORIES["minimal"])

    caption = (
        f"🎨 <b>Shablon:</b> {html.escape(tpl['title'])} ({curr_idx + 1}/{len(templates)})\n"
        f"📁 <b>Toifasi:</b> {cat_info['emoji']} {cat_info['name']}\n"
        f"📊 <b>Mavjud slaydlar:</b> {tpl.get('total_slides', 30)} ta\n"
        f"✨ <i>Haqiqiy 3D grafikalar, toza vektor ikonlar va fotosuratlar</i>\n\n"
        f"<i>Oldingi/keyingi tugmalari orqali shablonlarni ko'ring va «✅ Tanlash» tugmasini bosing:</i>"
    )
    kb = get_template_carousel_keyboard(cat_key, curr_idx, len(templates), tpl["id"])

    # Muqova rasmini olish (Telegram Cloud file_id yoki lokal diskdan)
    thumb_media = None
    if tpl.get("thumb_file_id"):
        thumb_media = tpl["thumb_file_id"]
    else:
        pptx_path = GLOBAL_TEMPLATE_CATALOG.resolve_template_pptx_path(tpl["id"])
        if pptx_path:
            perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
            os.makedirs(perm_dir, exist_ok=True)
            cand_thumb = os.path.join(perm_dir, f"thumb_{tpl['id']}.jpeg")
            if os.path.exists(cand_thumb) or GLOBAL_TEMPLATE_CATALOG.extract_cover_thumbnail(pptx_path, cand_thumb):
                thumb_media = FSInputFile(cand_thumb)

    await state.set_state(SlideCreationFSM.selecting_template)

    # Agar xabar rasm bo'lsa va yana rasm bo'lsa -> edit_media
    if callback.message.photo and thumb_media:
        try:
            media = InputMediaPhoto(
                media=thumb_media, caption=caption, parse_mode="HTML"
            )
            await callback.message.edit_media(media=media, reply_markup=kb)
            return
        except Exception:
            pass

    # Yangi rasm xabari yuborish
    try:
        await callback.message.delete()
    except Exception:
        pass

    if thumb_media:
        await callback.message.answer_photo(
            photo=thumb_media,
            caption=caption,
            reply_markup=kb,
            parse_mode="HTML",
        )
    else:
        await callback.message.answer(caption, reply_markup=kb, parse_mode="HTML")


@router.callback_query(F.data.startswith("tpl_nav_"))
async def cb_template_nav(callback: CallbackQuery, state: FSMContext):
    parts = callback.data.split("_")
    cat_key = parts[2]
    tpl_idx = int(parts[3])
    await show_template_card(callback, state, cat_key, tpl_idx)
    await callback.answer()


@router.callback_query(F.data.startswith("tpl_choose_"))
async def cb_template_choose(callback: CallbackQuery, state: FSMContext):
    tpl_id = callback.data.replace("tpl_choose_", "")
    tpl = GLOBAL_TEMPLATE_CATALOG.get_template(tpl_id)
    title = tpl["title"] if tpl else "Tanlangan Shablon"

    await state.update_data(selected_tpl_id=tpl_id, selected_tpl_title=title)
    await state.set_state(SlideCreationFSM.selecting_slide_count)

    txt = (
        f"✅ <b>«{html.escape(title)}»</b> shabloni tanlandi!\n\n"
        f"📊 Taqdimotingiz <b>necha slayddan</b> iborat bo'lishini xohlaysiz?\n"
        f"<i>Shablon avtomatik ravishda tanlangan slaydlar soniga moslab qisqartiriladi va eng kerakli maketlar saralanadi:</i>"
    )
    kb = get_slide_count_keyboard(tpl_id)

    try:
        await callback.message.delete()
    except Exception:
        pass

    await callback.message.answer(txt, reply_markup=kb, parse_mode="HTML")
    await callback.answer()


@router.callback_query(F.data.startswith("count_"))
async def cb_slide_count_choose(callback: CallbackQuery, state: FSMContext):
    target_count = int(callback.data.replace("count_", ""))
    await state.update_data(target_slide_count=target_count)

    data = await state.get_data()
    topic = data.get("raw_topic", "Taqdimot")
    tpl_title = data.get("selected_tpl_title", "Shablon")
    tpl_id = data.get("selected_tpl_id", "tpl_med_heartbeat")

    status_msg = await callback.message.edit_text(
        f"⏳ <i>«{html.escape(topic)}» mavzusi va {target_count} ta slayd uchun mantiqiy reja tuzilmoqda...</i>",
        parse_mode="HTML",
    )

    # 1-Bosqich: Aynan target_count ta slayd uchun mantiqiy reja (Agenda)
    prompt = build_stage1_outline_prompt(topic, slide_count=target_count, doc_mode="pptx")
    outline_data, err = await call_gemini_cascade(
        prompt, schema=STAGE1_OUTLINE_SCHEMA, max_output_tokens=512
    )

    if not outline_data or "agenda" not in outline_data:
        await status_msg.edit_text(
            f"❌ Reja tuzishda xatolik: {html.escape(str(err))}\nIltimos, qayta urinib ko'ring."
        )
        return

    agenda = outline_data.get("agenda", [])
    title = outline_data.get("title", topic)
    category_id = outline_data.get("category_id", 1463)
    if not isinstance(category_id, int) or category_id < 100:
        category_id = 1463

    draft_data = {
        "title": title,
        "subtitle": outline_data.get("subtitle", ""),
        "category_id": category_id,
        "theme": tpl_id,
        "template_id": tpl_id,
        "template_title": tpl_title,
        "target_slide_count": target_count,
        "agenda": agenda,
        "raw_topic": topic,
    }

    await save_draft(callback.from_user.id, "pptx", title, "outline", draft_data)
    await state.set_state(SlideCreationFSM.viewing_agenda)

    agenda_str = "\n".join(f"<b>{i}.</b> {html.escape(str(item))}" for i, item in enumerate(agenda, 1))
    msg_text = (
        f"📝 <b>Taqdimot:</b> «{html.escape(title)}»\n"
        f"🎨 <b>Shablon:</b> {html.escape(tpl_title)}\n"
        f"📊 <b>Rejalashtirilgan slaydlar:</b> {len(agenda)} ta\n\n"
        f"<b>Mantiqiy Reja:</b>\n"
        f"{agenda_str}\n\n"
        f"<i>Siz rejaga band qo'shishingiz, ortiqchasini o'chirishingiz yoki to'g'ridan-to'g'ri yaratishni boshlashingiz mumkin:</i>"
    )
    await status_msg.edit_text(
        msg_text, reply_markup=get_agenda_editor_keyboard(tpl_title), parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(F.data == "agenda_delete_last")
async def cb_delete_last_item(callback: CallbackQuery):
    draft = await get_draft(callback.from_user.id, "pptx")
    if not draft:
        await callback.answer("Qoralama topilmadi", show_alert=True)
        return

    data = draft["data"]
    agenda = data.get("agenda", [])
    if len(agenda) <= 3:
        await callback.answer("Kamida 3 ta slayd bo'lishi shart!", show_alert=True)
        return

    removed = agenda.pop()
    data["agenda"] = agenda
    await save_draft(callback.from_user.id, "pptx", draft["title"], draft["stage"], data)

    tpl_title = data.get("template_title", "Shablon")
    agenda_str = "\n".join(f"<b>{i}.</b> {html.escape(str(item))}" for i, item in enumerate(agenda, 1))
    msg_text = (
        f"📝 <b>Taqdimot:</b> «{html.escape(str(draft['title']))}»\n"
        f"🎨 <b>Shablon:</b> {html.escape(tpl_title)}\n\n"
        f"<b>Mantiqiy Reja ({len(agenda)} ta slayd):</b>\n"
        f"{agenda_str}\n\n"
        f"<i>«{html.escape(str(removed))}» bandi o'chirildi.</i>"
    )
    await callback.message.edit_text(
        msg_text,
        reply_markup=get_agenda_editor_keyboard(tpl_title),
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

    tpl_title = data.get("template_title", "Shablon")
    agenda_str = "\n".join(f"<b>{i}.</b> {html.escape(str(item))}" for i, item in enumerate(agenda, 1))
    msg_text = (
        f"📝 <b>Taqdimot:</b> «{html.escape(str(draft['title']))}»\n"
        f"🎨 <b>Shablon:</b> {html.escape(tpl_title)}\n\n"
        f"<b>Mantiqiy Reja ({len(agenda)} ta slayd):</b>\n"
        f"{agenda_str}\n\n"
        f"<i>Yangi band muvaffaqiyatli qo'shildi!</i>"
    )
    await message.answer(
        msg_text,
        reply_markup=get_agenda_editor_keyboard(tpl_title),
        parse_mode="HTML",
    )


# ── 3. YAKUNIY SLAYD GENERATSIYASI ───────────────────────────────────────────

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
        "🎨 <i>Slaydlar tayyorlanmoqda (Haqiqiy shablon, aqlli qisqartirish va AI matn joylashtirish)...</i>",
        parse_mode="HTML",
    )

    data = draft["data"]
    title = data.get("title", "Taqdimot")
    agenda = data.get("agenda", [])
    tpl_id = data.get("template_id", "tpl_med_heartbeat")
    tpl_title = data.get("template_title", "Shablon")
    target_count = data.get("target_slide_count", len(agenda))
    category_id = data.get("category_id", 1463)

    async def _build_slide_job():
        # 1. Matnlarni generatsiya qilish (Stage 2)
        stage2_prompt = build_stage2_locked_expansion_prompt(
            title, locked_agenda=agenda, context_excerpt=data.get("raw_topic", "")
        )
        slides_res, err = await call_gemini_cascade(
            stage2_prompt, schema=STAGE2_SLIDES_SCHEMA, max_output_tokens=4096
        )
        if not slides_res or "slides" not in slides_res:
            raise RuntimeError(err or "Slayd matnlarini generatsiya qilishda xatolik")

        # Cloud shablon mavjudligini ta'minlash (agar lokal bo'lmasa, Telegram Cloud'dan yuklaydi)
        resolved_pptx = await GLOBAL_TEMPLATE_CATALOG.ensure_template_pptx_available(tpl_id, bot=callback.bot)

        full_spec = {
            "title": title,
            "topic": title,
            "template_id": tpl_id,
            "resolved_pptx_path": resolved_pptx,
            "target_slide_count": target_count,
            "slides": slides_res.get("slides", []),
            "description": slides_res.get("description", ""),
            "tags": slides_res.get("tags", ""),
        }

        perm_dir = os.path.join(tempfile.gettempdir(), "soff_ready")
        os.makedirs(perm_dir, exist_ok=True)
        out_path = os.path.join(perm_dir, f"slide_{user_id}_{int(time.time() * 1000)}.pptx")

        # 2. Sinxron C/XML dvigatel orqali shablonni kesish va matnni joylashtirish
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

    # Soff.uz va DTM Test registratsiyasi
    actual_slides_count = len(full_spec.get("slides", [])) + 2
    soff_price = 3000 if actual_slides_count <= 6 else 7000 if actual_slides_count <= 12 else 10000
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
            f"🎨 <b>Shablon:</b> {html.escape(tpl_title)}\n"
            f"📊 <b>Slaydlar soni:</b> {actual_slides_count} ta (Muqova, Reja va Xulosa bilan)\n"
            f"💰 <b>Soff.uz tavsiya narxi:</b> {soff_price:,} so'm"
        ),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb_after),
        parse_mode="HTML",
    )
