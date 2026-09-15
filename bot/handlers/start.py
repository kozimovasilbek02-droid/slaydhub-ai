# -*- coding: utf-8 -*-
from aiogram import Router, types, F
from aiogram.filters import CommandStart, Command
from aiogram.utils.keyboard import InlineKeyboardBuilder

router = Router()


def get_main_keyboard():
    builder = InlineKeyboardBuilder()
    builder.button(text="🌐 Slayd Tarjima Qilish (PPTX)", callback_data="menu_translate")
    builder.button(text="🧠 Yangi Slayd Yaratish (NotebookLM)", callback_data="menu_generate")
    builder.button(text="ℹ️ Bot Haqida & Yordam", callback_data="menu_help")
    builder.adjust(1)
    return builder.as_markup()


@router.message(CommandStart())
async def cmd_start(message: types.Message):
    welcome_text = (
        f"Assalomu alaykum, <b>{message.from_user.full_name}</b>! 👋\n\n"
        "🚀 <b>TaqdimotUz AI</b> (@SlaydHubUz_bot) ga xush kelibsiz!\n\n"
        "Men taqdimotlar bilan ishlash bo'yicha sizning yagona intellektual yordamchingizman.\n\n"
        "📌 <b>Asosiy imkoniyatlar:</b>\n"
        "1️⃣ <b>Slayd Tarjima qilish:</b> Tayyor inglizcha PowerPoint (.pptx) faylni yuborsangiz, format va shriftlarini buzmasdan o'zbekchaga tarjima qilib beraman.\n"
        "2️⃣ <b>Yangi Slayd Yaratish:</b> Istalgan mavzuni yozasiz, NotebookLM tizimi orqali chuqur ilmiy reja tuzib, professional shablonda tayyor taqdimot yaratib beraman.\n"
        "3️⃣ <b>Soff.uz ga Yuklash:</b> Tayyor bo'lgan taqdimotni bir bosishda Soff.uz platformasiga avtomatik joylashtirishingiz mumkin.\n\n"
        "Quyidagi tugmalardan birini tanlang:"
    )
    await message.answer(welcome_text, parse_mode="HTML", reply_markup=get_main_keyboard())


@router.callback_query(F.data == "menu_main")
async def cb_main(callback: types.CallbackQuery):
    try:
        await callback.message.edit_text(
            "Asosiy menyudasiz. Kerakli xizmatni tanlang:",
            reply_markup=get_main_keyboard()
        )
    except Exception:
        await callback.message.answer(
            "Asosiy menyudasiz. Kerakli xizmatni tanlang:",
            reply_markup=get_main_keyboard()
        )
    await callback.answer()


@router.callback_query(F.data == "menu_help")
async def cb_help(callback: types.CallbackQuery):
    help_text = (
        "📖 <b>Qo'llanma va Ko'rsatmalar:</b>\n\n"
        "• <b>Tarjima uchun:</b> 'Slayd Tarjima Qilish' tugmasini bosing va .pptx faylingizni botga yuboring.\n"
        "• <b>Yangi taqdimot uchun:</b> 'Yangi Slayd Yaratish' tugmasini bosing, mavzuni kiriting va slaydlar sonini belgilang.\n"
        "• Barcha yaratilgan va tarjima qilingan slaydlar professional PowerPoint (.pptx) formatida taqdim etiladi."
    )
    builder = InlineKeyboardBuilder()
    builder.button(text="⬅️ Orqaga", callback_data="menu_main")
    await callback.message.edit_text(help_text, parse_mode="HTML", reply_markup=builder.as_markup())
    await callback.answer()
