@echo off
chcp 65001 > nul
echo ========================================================
echo   SlaydHub Ultra v4.1 -> GitHub Push (Render Avto-Deploy)
echo ========================================================
echo.
cd /d "C:\Users\user\Desktop\Ish\Soff.uz\telegram_bot"
git status
echo.
echo [1/3] O'zgarishlar Git indeksiga qo'shilmoqda...
git add main.py requirements.txt .env.example README.md PUSH_TO_GITHUB.bat START_BOT_LOCAL.bat core/ engines/ billing/ handlers/
echo.
echo [2/3] Commit yaratilmoqda...
git commit -m "feat: SlaydHub Ultra v4.1 Argon Production Release"
echo.
echo [3/3] GitHub ga push qilinmoqda (origin main)...
git push origin main
echo.
echo ========================================================
echo   Muvaffaqiyatli yakunlandi! Render avtomatik qayta ishga tushadi.
echo ========================================================
pause
