#!/usr/bin/env bash
# exit on error
set -o errexit

echo "==> [BUILD] Pip yangilanmoqda..."
pip install --upgrade pip

echo "==> [BUILD] Barcha bog'liqliklar requirements.txt dan o'rnatilmoqda..."
pip install -r requirements.txt

echo "==> [BUILD] O'rnatilgan paketlar tekshirilmoqda..."
python -c "import aiogram, aiohttp, pptx, docx, reportlab, fitz, google.genai; print('✅ Barcha modellar va kutubxonalar muvaffaqiyatli yuklandi!')"
