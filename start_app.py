# -*- coding: utf-8 -*-
"""
SlideTranslate AI — One-click Complete Launcher
Starts FastAPI Backend (Port 8000) and Vite React Frontend (Port 5173).
"""

import os
import sys
import time
import subprocess
import webbrowser

# UTF-8 encoding support on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def print_banner():
    print("=" * 65)
    print("      🚀 SlideTranslate AI — Professional Slayd Tarjimoni")
    print("      FastAPI (Backend) + React / Vite (Frontend) + Gemini AI")
    print("=" * 65)

def main():
    print_banner()
    root_dir = os.path.abspath(os.path.dirname(__file__))
    frontend_dir = os.path.join(root_dir, "frontend")

    print("[1/3] Backend tekshirilmoqda...")
    backend_cmd = [sys.executable, "-m", "uvicorn", "backend.main:app", "--host", "0.0.0.0", "--port", "8000", "--reload"]
    print("      Backend buyrug'i: " + " ".join(backend_cmd))
    backend_proc = subprocess.Popen(backend_cmd, cwd=root_dir)

    print("[2/3] Frontend (Vite) ishga tushirilmoqda...")
    # On Windows npm is a cmd/bat
    npm_cmd = "npm.cmd" if os.name == "nt" else "npm"
    frontend_proc = subprocess.Popen([npm_cmd, "run", "dev", "--", "--host"], cwd=frontend_dir)

    time.sleep(2)
    url = "http://localhost:5173"
    print(f"\n[3/3] Barcha xizmatlar muvaffaqiyatli ishga tushdi!")
    print(f"      🌐 Veb-ilova manzili: {url}")
    print(f"      📡 Backend API: http://localhost:8000/docs")
    print(f"\nBrauzer avtomatik ochilmoqda...")
    
    try:
        webbrowser.open(url)
    except Exception:
        pass

    print("\n[Dasturni to'xtatish uchun Ctrl+C tugmalarini bosing]\n")
    try:
        backend_proc.wait()
        frontend_proc.wait()
    except KeyboardInterrupt:
        print("\nTo'xtatilmoqda...")
        backend_proc.terminate()
        frontend_proc.terminate()
        print("Xizmatlar to'xtatildi.")

if __name__ == "__main__":
    main()
