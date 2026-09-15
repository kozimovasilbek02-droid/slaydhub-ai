# 🌟 SlaydHub AI — Professional AI-Powered PowerPoint Presentation Platform & Telegram Bot

[![Telegram Bot](https://img.shields.io/badge/Telegram-@SlaydHubUz__bot-blue?logo=telegram)](https://t.me/SlaydHubUz_bot)
[![Python](https://img.shields.io/badge/Python-3.10+-yellow?logo=python)](https://python.org)
[![PowerPoint](https://img.shields.io/badge/PowerPoint-python--pptx-orange?logo=microsoftpowerpoint)](https://python-pptx.readthedocs.io/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

**SlaydHub AI** — sun'iy intellekt (NotebookLM va Google Gemini) yordamida professional, zamonaviy va chiroyli PowerPoint taqdimotlarini avtomatik yaratuvchi to'liq ekotizim. Telegram bot (`@SlaydHubUz_bot`), Web Studio (Streamlit) va Desktop ilovani o'z ichiga oladi.

---

## ✨ Asosiy Imkoniyatlar

- 🤖 **Telegram Bot (`@SlaydHubUz_bot`)**: Har qanday mavzu bo'yicha ilmiy va amaliy chuqur kontentga ega taqdimotlarni bir zumda yaratadi.
- 🧠 **Gibrid AI Dvigateli**:
  - **NotebookLM Avtomatizatsiyasi**: Google NotebookLM orqali mavzuni chuqur tahlil qilish va manbalarga tayangan holda professional tuzilma olish.
  - **Gemini 3.6 Flash Fallback**: Tezkor va ishonchli generatsiya zaxirasi.
- 🎨 **Zamonaviy Tipografiya va Premium Dizayn**:
  - **Zero Watermark**: Shablonlardagi barcha begona reklama va logotiplar (PresentationGO va boshqalar) tubdan tozalanadi.
  - **Unboxed Clean Cards**: Matnlarni bo'g'ib turuvchi og'ir qutilar o'rniga zamonaviy ochiq tipografik kartalar, aksentli chegara chiziqlari va yuqori kontrast.
  - **Smart Agenda & Grid**: Taqdimot rejalari uchun ketma-ketlik raqamlangan chiroyli qatorlar va proporsional ko'p ustunli bloklar.
  - **Unified Q&A**: Savol-javob sahifasi uchun yaxlit va diqqatni jamlovchi premium kartalar.
- 🖼️ **Rasmlar va Vektor Shakllar**: Avtomatik mavzuga mos vizual elementlar va vektor shakllarni joylashtirish.
- 💻 **Ko'p Platformali Interfeys**: Telegram bot, Desktop App va Streamlit Web Studio.

---

## 🚀 O'rnatish va Ishga Tushirish

### 1. Loyihani yuklab olish
```bash
git clone https://github.com/kozimovasilbek02-droid/slaydhub-ai.git
cd slaydhub-ai
```

### 2. Virtual muhit va bog'liqliklarni o'rnatish
```bash
python -m venv venv
venv\Scripts\activate  # Windows
pip install -r requirements.txt
```

### 3. Muhit o'zgaruvchilarini sozlash (`.env`)
`.env.example` faylidan nusxa olib `.env` yarating:
```env
TELEGRAM_BOT_TOKEN=your_telegram_bot_token_here
GEMINI_API_KEY=your_gemini_api_key_here
```

### 4. Telegram Botni ishga tushirish
```bash
python bot/bot_main.py
```
yoki Windows uchun qulay bat fayl orqali:
```bash
START_BOT.bat
```

### 5. Web Studio (Streamlit)
```bash
streamlit run app.py
# yoki
START_WEB_APP.bat
```

---

## 📁 Loyiha Strukturasi

```text
├── bot/
│   ├── bot_main.py                 # Telegram bot asosiy kirish nuqtasi
│   ├── config.py                   # Bot sozlamalari va .env yuklash
│   ├── handlers/                   # Telegram xabarlar va buyruqlar ishlovchilari
│   │   ├── common.py
│   │   └── presentation.py
│   └── services/                   # Asosiy xizmatlar
│       ├── notebooklm_service.py   # NotebookLM + Gemini gibrid AI xizmati
│       ├── slide_builder.py        # Professional PPTX generatsiyasi va dizayn
│       ├── template_selector.py    # Mavzuga mos shablon tanlash
│       └── soff_client.py          # Savdo va to'lov integratsiyasi
├── core/                           # Asosiy konvertatsiya va kommutatsiya yadrosi
├── ppt_engine/                     # Slayd yaratish algoritmlari
├── app.py                          # Streamlit Web Studio
├── desktop_app.py                  # PyQt/Desktop interfeysi
├── requirements.txt                # Python bog'liqliklari
├── .gitignore                      # Git istisnolar ro'yxati
└── README.md                       # Loyiha hujjatlari
```

---

## 🛡️ Xavfsizlik va Maxfiylik

- Barcha maxfiy kalitlar va tokenlar `.env` faylida xavfsiz saqlanadi va Git repozitoriyasiga kiritilmaydi.
- Foydalanuvchi taqdimotlari va vaqtinchalik fayllari vaqti-vaqti bilan tozalanadi.

---

## 👨‍💻 Muallif

- **Muallif**: [Asilbek Kozimov](https://github.com/kozimovasilbek02-droid)
- **Telegram Bot**: [@SlaydHubUz_bot](https://t.me/SlaydHubUz_bot)
