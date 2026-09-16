# -*- coding: utf-8 -*-
import os
import uuid
import asyncio
from pathlib import Path
from aiogram import Router, types, F
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder

from bot.config import OUTPUT_DIR
from bot.services.notebooklm_service import NotebookLMService
from bot.services.slide_builder import SlideBuilderService
from backend.core.ai_presentation_generator import AIPresentationGenerator
from backend.core.gamma_generator import GammaGenerator
from backend.core.deck_builder import DeckBuilder

router = Router()
notebooklm_service = NotebookLMService()
ai_generator = AIPresentationGenerator()


class GeneratorStates(StatesGroup):
    waiting_for_topic = State()
    waiting_for_slide_count = State()
    waiting_for_style = State()
    waiting_for_theme = State()


@router.callback_query(F.data == "menu_generate")
async def cb_start_generate(callback: types.CallbackQuery, state: FSMContext):
    await state.set_state(GeneratorStates.waiting_for_topic)
    builder = InlineKeyboardBuilder()
    builder.button(text="⬅️ Bekor qilish", callback_data="menu_main")

    await callback.message.edit_text(
        "🧠 <b>Yangi Taqdimot Yaratish (AI Presentation Studio)</b>\n\n"
        "Iltimos, taqdimot yaratmoqchi bo'lgan <b>mavzuingizni</b> yozib yuboring:\n\n"
        "<i>Masalan: Sun'iy Intellekt va Raqamli Iqtisodiyot, O'zbekistonning investitsion jozibadorligi, Bioetika muammolari...</i>",
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
    await state.update_data(slide_count=slide_count)
    await state.set_state(GeneratorStates.waiting_for_style)

    builder = InlineKeyboardBuilder()
    builder.button(text="🎨 Gamma-Style (Ultra-HD Zamonaviy)", callback_data="style_gamma")
    builder.button(text="📊 Native Vector (100% Tahrirlanuvchi)", callback_data="style_vector")
    builder.button(text="📚 Standart Ilmiy (NotebookLM)", callback_data="style_standard")
    builder.button(text="⬅️ Bekor qilish", callback_data="menu_main")
    builder.adjust(1)

    await callback.message.edit_text(
        f"🎨 <b>Dizayn Uslubini Tanlang:</b>\n\n"
        f"1. <b>Gamma-Style</b> — Gamma.app uslubidagi qora-gradient neon kartochkalar va Ultra-HD visual effektlar.\n"
        f"2. <b>Native Vector</b> — Microsoft PowerPoint'da 100% tahrirlanuvchi, rangli shakllar va professional infografika.\n"
        f"3. <b>Standart Ilmiy</b> — NotebookLM tahlili asosidagi klassik o'quv shabloni.",
        parse_mode="HTML",
        reply_markup=builder.as_markup()
    )
    await callback.answer()


@router.callback_query(GeneratorStates.waiting_for_style, F.data.startswith("style_"))
async def handle_style_choice(callback: types.CallbackQuery, state: FSMContext):
    style = callback.data.split("style_")[1]
    await state.update_data(style=style)

    if style in ["gamma", "vector"]:
        await state.set_state(GeneratorStates.waiting_for_theme)
        builder = InlineKeyboardBuilder()
        builder.button(text="🔵 O'zbek Ko'k (Royal Blue)", callback_data="theme_uzbek_blue")
        builder.button(text="🟢 Zumrad Yashil (Emerald)", callback_data="theme_emerald_teal")
        builder.button(text="⚫ Qora Slate (Dark Tech)", callback_data="theme_dark_slate")
        builder.button(text="🟣 Zamonaviy Binafsha", callback_data="theme_modern_purple")
        builder.button(text="🔴 Qizil Yoqut (Crimson)", callback_data="theme_crimson_ruby")
        builder.button(text="⬅️ Bekor qilish", callback_data="menu_main")
        builder.adjust(2, 2, 1, 1)

        await callback.message.edit_text(
            "🎨 <b>Ranglar Palitrasini Tanlang:</b>\n\n"
            "Taqdimotingiz qaysi asosiy ranglarda yaratilishini xohlaysiz?",
            parse_mode="HTML",
            reply_markup=builder.as_markup()
        )
        await callback.answer()
        return

    # If standard
    await execute_generation(callback, state, theme="uzbek_blue")


@router.callback_query(GeneratorStates.waiting_for_theme, F.data.startswith("theme_"))
async def handle_theme_choice(callback: types.CallbackQuery, state: FSMContext):
    theme = callback.data.split("theme_")[1]
    await execute_generation(callback, state, theme=theme)


async def execute_generation(callback: types.CallbackQuery, state: FSMContext, theme: str = "uzbek_blue"):
    data = await state.get_data()
    topic = data.get("topic", "Taqdimot")
    slide_count = data.get("slide_count", 10)
    style = data.get("style", "vector")
    await state.clear()

    status_msg = await callback.message.edit_text(
        f"🧠 <b>AI '{topic}' bo'yicha {slide_count} ta slayd uchun ilmiy reja va faktlarni tahlil qilmoqda...</b>\n\n"
        "<i>Bu jarayon 10-25 soniya vaqt olishi mumkin...</i>",
        parse_mode="HTML"
    )

    try:
        file_id = str(uuid.uuid4())[:6]
        clean_topic = "".join(c for c in topic if c.isalnum() or c in (" ", "_", "-")).strip()[:40]

        if style == "standard":
            # 1. NotebookLM
            spec = await notebooklm_service.generate_slide_content(topic, slide_count=slide_count)
            await status_msg.edit_text(
                "🎨 <b>Ilmiy kontent tayyorlandi!</b>\n\n"
                "Endi mavzuga mos premium slayd shabloni tanlanmoqda va shakllantirilmoqda...",
                parse_mode="HTML"
            )
            out_filename = f"{clean_topic}_{slide_count}slayd_Standard_{file_id}.pptx"
            final_pptx_path = OUTPUT_DIR / out_filename
            final_path = SlideBuilderService.build_presentation(spec=spec, output_path=str(final_pptx_path))
        
        elif style == "gamma":
            # 2. Gamma Style (Ultra-HD)
            await status_msg.edit_text(
                f"🎨 <b>Gamma AI '{topic}' bo'yicha Ultra-HD slaydlarni loyihalashtirmoqda...</b>\n\n"
                "<i>Glassmorphism va neon vizual effektlar chizilmoqda...</i>",
                parse_mode="HTML"
            )
            spec = await ai_generator.generate_deck_spec(topic=topic, slide_count=slide_count, theme=theme)
            out_filename = f"{clean_topic}_{slide_count}slayd_Gamma_{file_id}.pptx"
            final_pptx_path = OUTPUT_DIR / out_filename
            final_path = await GammaGenerator.build(spec=spec, output_path=str(final_pptx_path))
        
        else:  # vector
            # 3. Native Vector Deck Builder (100% Editable)
            await status_msg.edit_text(
                f"📊 <b>Vektorli AI '{topic}' bo'yicha 100% tahrirlanuvchi slaydlarni chizmoqda...</b>\n\n"
                "<i>Ranglar palitrasi, kartochkalar va sarlavhalar joylashtirilmoqda...</i>",
                parse_mode="HTML"
            )
            spec = await ai_generator.generate_deck_spec(topic=topic, slide_count=slide_count, theme=theme)
            out_filename = f"{clean_topic}_{slide_count}slayd_Vector_{file_id}.pptx"
            final_pptx_path = OUTPUT_DIR / out_filename
            final_path = DeckBuilder.build(spec=spec, output_path=str(final_pptx_path))

        await status_msg.edit_text("✅ <b>Taqdimot tayyor! Fayl yuklanmoqda...</b>", parse_mode="HTML")

        builder = InlineKeyboardBuilder()
        builder.button(text="📤 Soff.uz ga yuklash (Avtomatik)", callback_data=f"soff_up:{final_pptx_path.name}")
        builder.button(text="🧠 Yangi mavzuda yaratish", callback_data="menu_generate")
        builder.button(text="🏠 Asosiy menyu", callback_data="menu_main")
        builder.adjust(1)

        style_title = "Gamma-Style (Ultra-HD)" if style == "gamma" else ("Native Vector (100% Tahrirlanuvchi)" if style == "vector" else "Standart Ilmiy (NotebookLM)")

        caption = (
            f"🎉 <b>Yangi Taqdimot Muvaffaqiyatli Yaratildi!</b>\n\n"
            f"📌 <b>Mavzu:</b> {topic}\n"
            f"📊 <b>Slaydlar soni:</b> {slide_count} ta slayd\n"
            f"🎨 <b>Uslub:</b> {style_title}\n"
            f"🌈 <b>Mavzu palitrasi:</b> {theme}\n\n"
            f"<i>Barcha slaydlar professional tarzda tartiblangan va tayyor.</i>"
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
