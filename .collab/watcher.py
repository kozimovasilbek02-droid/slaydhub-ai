# -*- coding: utf-8 -*-
"""
.collab/watcher.py
Real-time watcher for Antigravity & Claude Code Dual-Brain collaboration (Round 15).
Detects when Claude Code submits new critiques, flaws, or task assignments in COUNCIL.md,
or when Claude CLI completes an assistant turn in its project log.
"""

import os
import sys
import time
import io
import glob
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

WORKSPACE = Path(__file__).resolve().parent.parent
COUNCIL_FILE = Path(__file__).resolve().parent / "COUNCIL.md"
CLAUDE_DIR = Path(r"C:\Users\user\.claude\projects\C--Users-user-Desktop-Antigravity-Power-Point")

PLACEHOLDER_MARKER = "*(Claude Code o'z terminalida tahlil o'tkazib"

initial_council_mtime = os.path.getmtime(COUNCIL_FILE) if os.path.exists(COUNCIL_FILE) else 0

claude_logs = list(CLAUDE_DIR.glob("*.jsonl")) if CLAUDE_DIR.exists() else []
latest_log = max(claude_logs, key=os.path.getmtime) if claude_logs else None
initial_log_mtime = os.path.getmtime(latest_log) if latest_log and os.path.exists(latest_log) else 0

print("🔍 Dual-Brain Watcher (Round 15: Claude's Critique & Task Assignment) faollashdi...", flush=True)

start_time = time.time()
max_wait = 1800  # 30 mins

detected = False

while time.time() - start_time < max_wait:
    # 1. Check if COUNCIL.md was updated with Claude's Round 15 critique
    if os.path.exists(COUNCIL_FILE):
        curr_mtime = os.path.getmtime(COUNCIL_FILE)
        if curr_mtime > initial_council_mtime:
            try:
                content = COUNCIL_FILE.read_text(encoding="utf-8", errors="ignore")
                if "### 🟣 Claude Code Tanqidi va Vazifasi (Round 15)" in content:
                    claude_section = content.split("### 🟣 Claude Code Tanqidi va Vazifasi (Round 15)")[1]
                    if "### 🔵" in claude_section:
                        claude_section = claude_section.split("### 🔵")[0]
                    clean_text = claude_section.replace(PLACEHOLDER_MARKER, "").strip()
                    if len(clean_text) > 40:
                        time.sleep(2)
                        print("🎯 ANIQLANDI: Claude Code '.collab/COUNCIL.md' da 15-bosqich tanqidi va vazifasini yozib tugatdi!", flush=True)
                        detected = True
                        break
            except Exception:
                pass

    # 2. Check if Claude project log was modified and COUNCIL.md also modified
    if latest_log and os.path.exists(latest_log):
        curr_log_mtime = os.path.getmtime(latest_log)
        if curr_log_mtime > initial_log_mtime + 2:
            time.sleep(3)
            if os.path.exists(COUNCIL_FILE) and os.path.getmtime(COUNCIL_FILE) > initial_council_mtime:
                try:
                    content = COUNCIL_FILE.read_text(encoding="utf-8", errors="ignore")
                    if "### 🟣 Claude Code Tanqidi va Vazifasi (Round 15)" in content:
                        claude_section = content.split("### 🟣 Claude Code Tanqidi va Vazifasi (Round 15)")[1]
                        if "### 🔵" in claude_section:
                            claude_section = claude_section.split("### 🔵")[0]
                        clean_text = claude_section.replace(PLACEHOLDER_MARKER, "").strip()
                        if len(clean_text) > 40:
                            print("🎯 ANIQLANDI: Claude Code faoliyati va 15-bosqich yangilanishi aniqlandi!", flush=True)
                            detected = True
                            break
                except Exception:
                    pass

    time.sleep(3)

if not detected:
    print("⏳ Kutish vaqti tugadi.", flush=True)
sys.exit(0 if detected else 1)
