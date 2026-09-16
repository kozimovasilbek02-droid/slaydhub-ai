# -*- coding: utf-8 -*-
import os
import uuid
from pathlib import Path
from aiogram import Router, types, F, Bot
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder

from bot.config import TEMP_DIR, OUTPUT_DIR
from bot.services.translator_service import TranslatorService
from bot.services.soff_uploader import SoffUploaderService
from bot.services.file_cache import register_file, get_file

router = Router()
translator_service = TranslatorService()


class TranslationStates(StatesGroup):
    waiting_for_file = State()


@router.callback_query(F.data == "menu_translate")
async def cb_start_translate(callback: types.CallbackQuery, state: FSMContext):
    await state.set_state(TranslationStates.waiting_for_file)
    builder = InlineKeyboardBuilder()
    builder.button(text="⬅️ Bekor qilish", callback_data="menu_main")
    
    await callback.message.edit_text(
        "📄 <b>PowerPoint (.pptx) taqdimot faylingizni yuboring:</b>\n\n"
        "1️⃣ <b>Telegram orqali:</b> Faylni chatga yuboring (20 MB gacha bo'lsa);\n"
        "2️⃣ <b>Yoki og'ir fayllar uchun:</b> Faylning kompyuteringizdagi manzilini matn qilib yuboring:\n"
        "<i>Masalan: C:\\Users\\user\\Desktop\\Leonardo_DaVinci.pptx</i>\n\n"
        "Gemini AI formatni buzmasdan o'zbek tiliga tarjima qilib beradi.",
        parse_mode="HTML",
        reply_markup=builder.as_markup()
    )
    await callback.answer()


async def process_translation(message: types.Message, input_path: Path, original_name: str):
    import asyncio
    status_msg = await message.answer("📥 Slayd tahlil qilinmoqda...", parse_mode="HTML")
    
    # 1. Taqdimot va fayl nomini o'zbek tiliga chiroyli tarjima qilish
    clean_title = await asyncio.to_thread(
        translator_service.get_clean_presentation_title,
        original_name
    )
    output_filename = f"{clean_title}.pptx"
    output_path = OUTPUT_DIR / output_filename

    try:
        async def update_progress(text: str):
            try:
                await status_msg.edit_text(f"⏳ <b>{text}</b>", parse_mode="HTML")
            except Exception:
                pass

        # Tarjima qilish
        result = await translator_service.translate_presentation(
            input_pptx=str(input_path),
            output_pptx=str(output_path),
            target_script="latin",
            progress_callback=update_progress,
            presentation_title=clean_title
        )

        if not result.get("success"):
            err_msg = result.get("error", "Noma'lum xato")
            await status_msg.edit_text(f"❌ Xatolik yuz berdi: {err_msg}")
            return

        # Tayyor faylni yuborish
        size_mb = round(os.path.getsize(output_path) / (1024 * 1024), 1)
        wait_note = f"\n⏳ <i>Fayl hajmi {size_mb} MB bo'lgani uchun Telegramga yuklash bir oz vaqt olishi mumkin...</i>\n" if size_mb > 15 else ""
        await status_msg.edit_text(
            f"✅ <b>Tarjima muvaffaqiyatli yakunlandi!</b>\n{wait_note}\n"
            f"💾 <b>Fayl kompyuteringizda saqlandi:</b>\n<code>{output_path.resolve()}</code>\n\n"
            f"📤 <i>Hozir fayl chatga ham yuborilmoqda...</i>",
            parse_mode="HTML"
        )

        # Soff.uz yuklash tugmasi
        short_id = register_file(output_path)
        builder = InlineKeyboardBuilder()
        builder.button(text="📤 Soff.uz ga yuklash (Avtomatik)", callback_data=f"soff_up:{short_id}")
        builder.button(text="🏠 Asosiy menyu", callback_data="menu_main")
        builder.adjust(1)

        final_slides = result.get("final_slides", 0)
        cleaned_ad = result.get("cleaned_ad_slides", 0)
        ad_text = f"\n🧹 Tozalangan reklama slaydlari: {cleaned_ad} ta" if cleaned_ad > 0 else ""

        caption = (
            f"🎉 <b>Taqdimotingiz tayyor!</b>\n\n"
            f"📊 Slaydlar soni: {final_slides} ta{ad_text}\n"
            f"📝 Tarjima qilingan matnlar: {result.get('total_items', 0)} ta\n"
            f"✨ Til: O'zbek tili (Lotin alifbosi)\n\n"
            f"<i>Faylni bir bosishda Soff.uz platformasiga ham yuklashingiz mumkin:</i>"
        )

        await message.answer_document(
            document=types.FSInputFile(str(output_path), filename=output_filename),
            caption=caption,
            parse_mode="HTML",
            reply_markup=builder.as_markup()
        )
        await status_msg.delete()

    except Exception as e:
        if output_path.exists():
            await status_msg.edit_text(
                f"⚠️ <b>Tarjima yakunlandi va fayl saqlandi</b>, lekin Telegram orqali yuborishda xatolik bo'ldi ({str(e)}).\n\n"
                f"💾 <b>Kompyuteringizdagi fayl manzili:</b>\n<code>{output_path.resolve()}</code>\n\n"
                f"<i>Faylni to'g'ridan-to'g'ri ko'rsatilgan manzildan ochishingiz mumkin.</i>",
                parse_mode="HTML"
            )
        else:
            await status_msg.edit_text(f"❌ Xatolik yuz berdi: {str(e)}")



@router.message(TranslationStates.waiting_for_file, F.text)
async def handle_pptx_path_text(message: types.Message, state: FSMContext):
    path_str = message.text.strip().strip('"').strip("'")
    if not (path_str.lower().endswith(".pptx") and os.path.exists(path_str)):
        await message.answer(
            "⚠️ Kiritilgan manzil bo'yicha fayl topilmadi yoki bu .pptx fayl emas!\n\n"
            "Iltimos, faylni to'g'ridan-to'g'ri tashlang yoki to'g'ri manzil kiriting:\n"
            "<i>Masalan: C:\\Users\\user\\Desktop\\taqdimot.pptx</i>",
            parse_mode="HTML"
        )
        return

    await state.clear()
    input_path = Path(path_str)
    await process_translation(message, input_path=input_path, original_name=input_path.name)


@router.message(TranslationStates.waiting_for_file, F.document)
async def handle_pptx_file(message: types.Message, state: FSMContext, bot: Bot):
    doc = message.document
    if not doc.file_name.lower().endswith(".pptx"):
        await message.answer("⚠️ Iltimos, faqat <b>.pptx</b> formatidagi PowerPoint fayl yuboring!", parse_mode="HTML")
        return

    # Telegram Bot API 20MB limiti tekshiruvi
    file_size_mb = round((doc.file_size or 0) / (1024 * 1024), 1)
    if doc.file_size and doc.file_size > 20 * 1024 * 1024:
        warning_text = (
            f"⚠️ <b>Fayl hajmi juda katta ({file_size_mb} MB)!</b>\n\n"
            "Telegram Bot API rasmiy qoidalariga ko'ra, botlar chat orqali faqat <b>20 MB gacha</b> bo'lgan fayllarni yuklab olishi mumkin.\n\n"
            "💡 <b>Katta faylni tarjima qilishning 2 ta oson yo'li:</b>\n"
            "1️⃣ <b>Fayl manzilini yuboring (Eng osoni):</b> Fayl shu kompyuteringizda bo'lsa, uning manzilini botga matn qilib yozib yuboring (masalan: <code>C:\\Users\\user\\Desktop\\{doc.file_name}</code>). Bot uni diskdan to'g'ridan-to'g'ri ochib beradi!\n"
            "2️⃣ <b>Yoki rasmlarni siqish:</b> PowerPoint'da ochib, rasmni bosib <i>'Сжать рисунки' -> 'Для Интернета (150 ppi)'</i> qilib saqlasangiz, hajmi 5-10 MB ga tushadi."
        )
        await message.answer(warning_text, parse_mode="HTML")
        return

    await state.clear()
    status_msg = await message.answer("📥 Fayl Telegram serveridan yuklab olinmoqda...", parse_mode="HTML")

    file_id = str(uuid.uuid4())[:8]
    input_filename = f"orig_{file_id}_{doc.file_name}"
    input_path = TEMP_DIR / input_filename

    try:
        file = await bot.get_file(doc.file_id)
        await bot.download_file(file.file_path, destination=str(input_path))
        await status_msg.delete()
        await process_translation(message, input_path=input_path, original_name=doc.file_name)
    except Exception as e:
        await status_msg.edit_text(f"❌ Xatolik yuz berdi: {str(e)}")


@router.callback_query(F.data.startswith("soff_up:"))
async def cb_soff_upload(callback: types.CallbackQuery):
    try:
        await callback.answer()
    except Exception:
        pass

    short_id = callback.data.split("soff_up:", 1)[1]
    file_path = get_file(short_id, OUTPUT_DIR)

    if not file_path or not file_path.exists():
        await callback.message.reply("⚠️ Fayl topilmadi yoki kesh muddati o'tgan. Iltimos, qayta urinib ko'ring.")
        return

    try:
        await callback.message.edit_reply_markup(reply_markup=None)
    except Exception:
        pass
    status_msg = await callback.message.reply("🚀 <b>Soff.uz ga yuklash boshlandi... Kutilmoqda...</b>", parse_mode="HTML")

    title = file_path.stem.replace("Uzbek_", "").replace("_Taqdimot", "").replace("_", " ")
    res = await SoffUploaderService.upload_presentation(
        pptx_path=str(file_path),
        title=title,
        price=25000,
        category="Taqdimot"
    )

    if res.get("success"):
        builder = InlineKeyboardBuilder()
        builder.button(text="🏠 Asosiy menyu", callback_data="menu_main")
        await status_msg.edit_text(
            f"✅ <b>Tabriklaymiz!</b>\n\n{res.get('message')}\n\n"
            f"📄 Mahsulot nomi: <i>{title}</i>\n"
            f"💰 Narxi: 25,000 so'm",
            parse_mode="HTML",
            reply_markup=builder.as_markup()
        )
    else:
        await status_msg.edit_text(f"❌ Soff.uz ga yuklashda xatolik: {res.get('error')}")
