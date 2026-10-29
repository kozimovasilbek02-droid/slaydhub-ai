# -*- coding: utf-8 -*-
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.filters import CommandStart, Command
from core.database import get_or_create_user

router = Router(name="start_cabinet")


def get_main_menu_keyboard() -> InlineKeyboardMarkup:
    """SlaydHub Ultra universal bosh menyu klaviaturasi."""
    kb = [
        [
            InlineKeyboardButton(text="🎨 Slayd Yaratish (SlidePro)", callback_data="menu_create_slide"),
            InlineKeyboardButton(text="📚 Kitob Tahlili (PDF)", callback_data="menu_book_split"),
        ],
        [
            InlineKeyboardButton(text="🌐 Slayd Tarjimon AI", callback_data="menu_translate_slide"),
            InlineKeyboardButton(text="📄 Akademik Word (OTM)", callback_data="menu_academic_word"),
        ],
        [
            InlineKeyboardButton(text="📋 DTM Test & Viktorina", callback_data="menu_dtm_quiz"),
            InlineKeyboardButton(text="💳 Hamyon & Keshbek", callback_data="menu_cabinet"),
        ],
        [
            InlineKeyboardButton(text="👥 Do'stlarni taklif qilish (Referal)", callback_data="menu_referral"),
            InlineKeyboardButton(text="ℹ️ Yordam & Qo'llanma", callback_data="menu_help"),
        ],
    ]
    return InlineKeyboardMarkup(inline_keyboard=kb)


@router.message(CommandStart())
async def cmd_start(message: Message):
    user_id = message.from_user.id
    ref_id = None
    
    # Referal parametrini ajratish (/start 123456)
    args = message.text.split()
    if len(args) > 1 and args[1].isdigit():
        ref_id = int(args[1])

    user = await get_or_create_user(user_id, referred_by=ref_id)
    first_name = message.from_user.first_name or "Foydalanuvchi"

    greeting = (
        f"Assalomu alaykum, <b>{first_name}</b>!\n\n"
        f"🚀 <b>SlaydHub Ultra (All-In-One Academic Ecosystem)</b> ga xush kelibsiz!\n\n"
        f"Bu yerda siz:\n"
        f"• <b>Professional Slaydlar:</b> Rejani oldindan ko'rib va tahrirlab yaratish;\n"
        f"• <b>Katta Kitoblar Tahlili:</b> 8-50 MB darsliklarni 4-10 ta slaydga ajratish;\n"
        f"• <b>Taqdimot Tarjimoni:</b> Slayd shakli va shriftini 100% buzmasdan o'girish;\n"
        f"• <b>Akademik Word:</b> OTM standartida Referat, Kurs ishi, Diplom va Maqolalar;\n"
        f"• <b>DTM Testlar:</b> A4 doirachali javoblar varaqasi (PDF) va Telegram viktorina;\n"
        f"• <b>Soff.uz Do'koni:</b> Slaydlaringizni sotib doimiy daromad olishingiz mumkin.\n\n"
        f"🎁 <i>Sizga 1 ta bepul boshlang'ich generatsiya krediti berildi!</i>\n"
        f"Quyidagi bo'limlardan birini tanlang:"
    )
    await message.answer(greeting, reply_markup=get_main_menu_keyboard(), parse_mode="HTML")


@router.callback_query(F.data == "menu_cabinet")
async def cb_cabinet(callback: CallbackQuery):
    user_id = callback.from_user.id
    user = await get_or_create_user(user_id)

    balans = user.get("balance_uzs", 0)
    keshbek = user.get("cashback_uzs", 0)
    referallar = user.get("referral_count", 0)
    bepul_kredit = user.get("free_slide_credits", 0)

    txt = (
        f"💼 <b>SIZNING SHAXSIY KABINETINGIZ</b>\n\n"
        f"💰 <b>Asosiy balans:</b> {balans:,} so'm\n"
        f"💎 <b>Yig'ilgan keshbek (10%):</b> {keshbek:,} so'm\n"
        f"🎁 <b>Bepul generatsiya imkoniyati:</b> {bepul_kredit} ta\n"
        f"👥 <b>Taklif qilingan do'stlar:</b> {referallar} ta\n\n"
        f"<i>Narxlar:</i>\n"
        f"• 1 ta slayd taqdimot = 3,000 so'm\n"
        f"• 1 ta Referat / Mustaqil ish = 5,000 so'm\n"
        f"• 1 ta Kurs ishi (25-35 bet) = 15,000 so'm\n"
        f"• Har bir to'lovdan 10% keshbek qaytadi!"
    )

    kb = [
        [
            InlineKeyboardButton(text="💳 Click orqali to'lash", callback_data="pay_click"),
            InlineKeyboardButton(text="💳 Payme orqali to'lash", callback_data="pay_payme"),
        ],
        [InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")],
    ]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")


@router.callback_query(F.data == "menu_referral")
async def cb_referral(callback: CallbackQuery):
    user_id = callback.from_user.id
    bot_info = await callback.bot.get_me()
    ref_link = f"https://t.me/{bot_info.username}?start={user_id}"

    txt = (
        f"👥 <b>REFERAL DASTURI (DO'STLARNI TAKLIF QILISH)</b>\n\n"
        f"Do'stlaringizga o'z referal havolangizni yuboring:\n"
        f"👉 <code>{ref_link}</code>\n\n"
        f"🎁 <b>Sovg'a qoidasi:</b>\n"
        f"Har 3 ta do'stingiz botga kirganda, hisobingizga <b>1 TA BEPUL PROFESSIONAL SLAYD GENERATSIYASI</b> beriladi!\n"
        f"Ular amalga oshirgan to'lovlardan sizga ham keshbek tushadi."
    )
    kb = [[InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")]]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")


@router.callback_query(F.data == "menu_home")
async def cb_home(callback: CallbackQuery):
    await callback.message.edit_text(
        "Asosiy menyu bo'limini tanlang:",
        reply_markup=get_main_menu_keyboard(),
    )


@router.callback_query(F.data == "menu_help")
async def cb_help(callback: CallbackQuery):
    txt = (
        f"ℹ️ <b>SLAYDHUB ULTRA QO'LLANMASI:</b>\n\n"
        f"1. <b>Slayd yaratish:</b> Mavzu nomini yuboring. Bot reja tuzib beradi, uni tahrirlab tasdiqlasangiz, 15 soniyada tayyor bo'ladi.\n"
        f"2. <b>Kitob bo'lish:</b> 8-50 MB darslik PDF faylini yuboring. Bot kitobni mavzularga bo'lib, tayyor .pptx qiladi.\n"
        f"3. <b>Slayd tarjima:</b> Ruscha yoki inglizcha .pptx fayl yuborsangiz, shakl va dizaynni buzmasdan o'zbekchaga o'giradi.\n"
        f"4. <b>Akademik Word:</b> OTM talablariga mos Titul, avtomatik Mundarija va GOST adabiyotlar ro'yxati bilan .docx tayyorlab beradi."
    )
    kb = [[InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")]]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
