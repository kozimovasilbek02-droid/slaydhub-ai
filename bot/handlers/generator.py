# -*- coding: utf-8 -*-
import os
import uuid
from pathlib import Path
from aiogram import Router, types, F
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder

from bot.config import OUTPUT_DIR
from bot.services.notebooklm_service import NotebookLMService
from bot.services.slide_builder import SlideBuilderService

router = Router()
notebooklm_service = NotebookLMService()


class GeneratorStates(StatesGroup):
    waiting_for_topic = State()
    waiting_for_slide_count = State()


@router.callback_query(F.data == "menu_generate")
async def cb_start_generate(callback: types.CallbackQuery, state: FSMContext):
    await state.set_state(GeneratorStates.waiting_for_topic)
    builder = InlineKeyboardBuilder()
    builder.button(text="⬅️ Bekor qilish", callback_data="menu_main")

    await callback.message.edit_text(
        "🧠 <b>Yangi Taqdimot Yaratish (NotebookLM AI)</b>\n\n"
        "Iltimos, taqdimot yaratmoqchi bo'lgan <b>mavzuingizni</b> yozib yuboring:\n\n"
        "<i>Masalan: Sun'iy Intellekt va Neyrotarmoqlar, O'zbekistonning investitsion jozibadorligi, Bioetika muammolari...</i>",
        parse_mode="HTML",
        reply_markup=builder.as_markup()
    )
    await callback.answer()


@router.message(GeneratorStates.waiting_for_topic, F.text)
async def handle_topic_input(message: types.Message, state: FSMContext):
    topic = message.text.strip()
    if len(topic) < 3:
        await message.answer("⚠️ Iltimos, mavzuni to'liqroq yozing.")
        return

    await state.update_data(topic=topic)
    await state.set_state(GeneratorStates.waiting_for_slide_count)

    builder = InlineKeyboardBuilder()
    builder.button(text="📄 5 ta slayd", callback_data="count_5")
    builder.button(text="📊 10 ta slayd (Tavsiya)", callback_data="count_10")
    builder.button(text="📑 15 ta slayd", callback_data="count_15")
    builder.button(text="📚 20 ta slayd", callback_data="count_20")
    builder.button(text="⬅️ Bekor qilish", callback_data="menu_main")
    builder.adjust(2, 2, 1)

    await message.answer(
        f"🎯 <b>Mavzu:</b> <i>{topic}</i>\n\n"
        "Taqdimot nechta slayddan iborat bo'lsin? Quyidagi variantlardan birini tanlang:",
        parse_mode="HTML",
        reply_markup=builder.as_markup()
    )


@router.callback_query(GeneratorStates.waiting_for_slide_count, F.data.startswith("count_"))
async def handle_slide_count(callback: types.CallbackQuery, state: FSMContext):
    slide_count = int(callback.data.split("count_")[1])
    data = await state.get_data()
    topic = data.get("topic", "Taqdimot")
    await state.clear()

    status_msg = await callback.message.edit_text(
        f"🧠 <b>NotebookLM AI '{topic}' bo'yicha {slide_count} ta slayd uchun ilmiy reja va faktlarni tahlil qilmoqda...</b>\n\n"
        "<i>Bu jarayon 10-25 soniya vaqt olishi mumkin...</i>",
        parse_mode="HTML"
    )

    try:
        # 1. NotebookLM orqali reja va ilmiy kontent olish
        spec = await notebooklm_service.generate_slide_content(topic, slide_count=slide_count)

        await status_msg.edit_text(
            "🎨 <b>Ilmiy kontent tayyorlandi!</b>\n\n"
            "Endi mavzuga mos premium slayd shabloni tanlanmoqda va ortiqcha slaydlar tozalanmoqda...",
            parse_mode="HTML"
        )

        # 2. Slayd builder orqali PPTX ga yozish va ortiqcha slaydlarni o'chirish
        file_id = str(uuid.uuid4())[:6]
        clean_topic = "".join(c for c in topic if c.isalnum() or c in (" ", "_", "-")).strip()[:40]
        out_filename = f"{clean_topic}_{slide_count}slayd_{file_id}.pptx"
        final_pptx_path = OUTPUT_DIR / out_filename

        final_path = SlideBuilderService.build_presentation(
            spec=spec,
            output_path=str(final_pptx_path)
        )

        # 3. Yuborish
        await status_msg.edit_text("✅ <b>Taqdimot tayyor! Fayl yuklanmoqda...</b>", parse_mode="HTML")

        builder = InlineKeyboardBuilder()
        builder.button(text="📤 Soff.uz ga yuklash (Avtomatik)", callback_data=f"soff_up:{final_pptx_path.name}")
        builder.button(text="🧠 Yangi mavzuda yaratish", callback_data="menu_generate")
        builder.button(text="🏠 Asosiy menyu", callback_data="menu_main")
        builder.adjust(1)

        agenda_items = "\n".join([f"• {p}" for p in spec.get("slides", [{}])[1].get("points", [])[:4]]) if len(spec.get("slides", [])) > 1 else ""

        caption = (
            f"🎉 <b>Yangi Taqdimot Muvaffaqiyatli Yaratildi!</b>\n\n"
            f"📌 <b>Mavzu:</b> {topic}\n"
            f"📊 <b>Slaydlar soni:</b> {slide_count} ta slayd\n"
            f"🎓 <b>Tuzilishi (NotebookLM):</b>\n{agenda_items}\n\n"
            f"<i>Barcha keraksiz shablon qoldiqlari tozalangan va tartiblangan.</i>"
        )

        await callback.message.answer_document(
            document=types.FSInputFile(str(final_pptx_path), filename=out_filename),
            caption=caption,
            parse_mode="HTML",
            reply_markup=builder.as_markup()
        )
        await status_msg.delete()

    except Exception as e:
        await status_msg.edit_text(f"❌ Taqdimot yaratishda xatolik yuz berdi: {str(e)}")

    await callback.answer()
