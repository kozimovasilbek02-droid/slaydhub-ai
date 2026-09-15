@echo off
chcp 65001 > nul
title SlaydHubUz Bot - Taqdimot & Tarjima AI
echo ===================================================
echo   TaqdimotUz AI Bot (@SlaydHubUz_bot)
echo   Ishga tushirilmoqda...
echo ===================================================
cd /d "%~dp0"
python bot/bot_main.py
pause
