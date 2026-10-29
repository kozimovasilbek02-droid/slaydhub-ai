# -*- coding: utf-8 -*-
import os
import html
import base64
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
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
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    ref_id = None

    args = (message.text or "").split()
    if len(args) > 1 and args[1].isdigit():
        ref_id = int(args[1])

    await get_or_create_user(user_id, referred_by=ref_id)
    first_name = html.escape(message.from_user.first_name or "Foydalanuvchi")

    greeting = (
        f"Assalomu alaykum, <b>{first_name}</b>!\n\n"
        f"🚀 <b>SlaydHub Ultra (All-In-One Academic Ecosystem)</b> ga xush kelibsiz!\n\n"
        f"Bu yerda siz:\n"
        f"• <b>Professional Slaydlar:</b> Rejani oldindan ko'rib va tahrirlab yaratish;\n"
        f"• <b>Katta Kitoblar Tahlili:</b> 8-50 MB darsliklarni mustaqil slaydga ajratish;\n"
        f"• <b>Taqdimot Tarjimoni:</b> Slayd shakli va shriftini 100% buzmasdan o'girish;\n"
        f"• <b>Akademik Word:</b> OTM standartida Referat, Kurs ishi, Diplom va Maqolalar;\n"
        f"• <b>DTM Testlar:</b> A4 doirachali javoblar varaqasi (PDF) va Telegram viktorina;\n"
        f"• <b>Soff.uz Do'koni:</b> Slaydlaringizni sotib doimiy daromad olishingiz mumkin.\n\n"
        f"🎁 <i>Sizga 1 ta bepul boshlang'ich generatsiya krediti berildi!</i>\n"
        f"Quyidagi bo'limlardan birini tanlang:"
    )
    await message.answer(greeting, reply_markup=get_main_menu_keyboard(), parse_mode="HTML")


@router.callback_query(F.data == "menu_book_split")
async def cb_book_split_info(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    txt = (
        "📚 <b>KATTA KITOBLAR DEKOMPOZITORI (PDF -> MUSTAQIL SLAYDLAR)</b>\n\n"
        "Menga <b>.pdf</b> formatidagi darslik, o'quv qo'llanma yoki monografiyani yuboring!\n\n"
        "⚡ <b>Qanday ishlaydi:</b>\n"
        "1. Bot kitobning Mundarijasi (TOC), Kirish, asosiy Boblari va Xulosasini aqlli ajratib oladi;\n"
        "2. Kitobni mustaqil, tayyor <b>.pptx</b> taqdimotlarga bo'lib beradi;\n"
        "3. Har bir tayyor slaydni bitta tugma bilan <b>Soff.uz</b> do'koniga sotuvga chiqarishingiz mumkin!\n\n"
        "📥 <i>Hozir shu chatga PDF faylingizni yuboring:</i>"
    )
    kb = [[InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")]]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
    await callback.answer()


@router.callback_query(F.data == "menu_translate_slide")
async def cb_translate_slide_info(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    txt = (
        "🌐 <b>UNIVERSAL TAQDIMOT TARJIMONI (.PPTX)</b>\n\n"
        "Menga ingliz, rus yoki boshqa tildagi tayyor <b>.pptx</b> taqdimot faylini yuboring!\n\n"
        "⚡ <b>Afzalliklari:</b>\n"
        "• OpenXML 10-qatlamli rekursiv tahlil (guruhlangan shakllar va jadvallar);\n"
        "• Slayd dizayni, ranglari, shrift o'lchamlari va koordinatalar 100% saqlanadi;\n"
        "• <code>latinLnBrk='0'</code> himoyasi orqali o'zbekcha so'zlar o'rtasidan bo'linib ketmaydi.\n\n"
        "📥 <i>Hozir shu chatga .pptx faylingizni yuboring:</i>"
    )
    kb = [[InlineKeyboardButton(text="◀️ Asosiy menyu", callback_data="menu_home")]]
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
    await callback.answer()


@router.callback_query(F.data == "menu_cabinet")
async def cb_cabinet(callback: CallbackQuery):
    user_id = callback.from_user.id
    user = await get_or_create_user(user_id)

    balans = user.get("balance_uzs", 0)
    keshbek = user.get("cashback_uzs", 0)
    referallar = user.get("referral_count", 0)
    bepul_kredit = user.get("free_slide_credits", 0)

    txt = (
        f"💼 <b>SIZNING SHAXSIY KABINETINGIZ</b>\n"
        f"🆔 <b>ID raqamingiz:</b> <code>{user_id}</code>\n\n"
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
    await callback.answer()


@router.callback_query(F.data.in_({"pay_click", "pay_payme"}))
async def cb_payment_provider(callback: CallbackQuery):
    user_id = callback.from_user.id
    is_click = callback.data == "pay_click"
    provider_name = "Click" if is_click else "Payme"

    amounts = [3000, 5000, 10000, 15000]
    kb = []

    if is_click:
        service_id = os.getenv("CLICK_SERVICE_ID", "00000")
        merchant_id = os.getenv("CLICK_MERCHANT_ID", "00000")
        for amt in amounts:
            url = (
                f"https://my.click.uz/services/pay?service_id={service_id}"
                f"&merchant_id={merchant_id}&amount={amt}&transaction_param={user_id}"
            )
            kb.append([InlineKeyboardButton(text=f"💳 Click: {amt:,} so'm (+10% keshbek)", url=url)])
    else:
        merchant_id = os.getenv("PAYME_MERCHANT_ID", "000000000000000000000000")
        for amt in amounts:
            params_str = f"m={merchant_id};ac.user_id={user_id};a={amt * 100}"
            b64_params = base64.b64encode(params_str.encode("utf-8")).decode("utf-8")
            url = f"https://checkout.paycom.uz/{b64_params}"
            kb.append([InlineKeyboardButton(text=f"💳 Payme: {amt:,} so'm (+10% keshbek)", url=url)])

    kb.append([InlineKeyboardButton(text="◀️ Kabinetga qaytish", callback_data="menu_cabinet")])

    txt = (
        f"💳 <b>{provider_name.upper()} ORQALI HISOBNI TO'LDIRISH</b>\n\n"
        f"🆔 Sizning hisob ID raqamingiz: <code>{user_id}</code>\n"
        f"💎 <b>Bonus:</b> To'langan har bir summaga <b>+10% keshbek</b> avtomatik qo'shiladi!\n\n"
        f"Quyidagi to'lov miqdorlaridan birini tanlang:"
    )
    await callback.message.edit_text(txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=kb), parse_mode="HTML")
    await callback.answer()


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
    await callback.answer()


@router.callback_query(F.data == "menu_home")
async def cb_home(callback: CallbackQuery, state: FSMContext):
    await state.clear()
    await callback.message.edit_text(
        "🏠 <b>Asosiy menyu bo'limini tanlang:</b>",
        reply_markup=get_main_menu_keyboard(),
        parse_mode="HTML",
    )
    await callback.answer()


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
    await callback.answer()
