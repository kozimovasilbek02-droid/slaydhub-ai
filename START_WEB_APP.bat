@echo off
title SlideAI Studio - Precision Vector Slide Recreator
echo ========================================================
echo   SlideAI Studio (Web App - React + FastAPI)
echo ========================================================
echo.
echo Veb-ilova ishga tushirilmoqda...
cd /d "c:\Users\user\Desktop\Antigravity\Power Point"
start "" "http://localhost:8000"
python -m uvicorn backend.main:app --host 0.0.0.0 --port 8000 --reload
pause
