# CLAUDE.md — SlaydHub AI Dual-Brain Collaboration Protocol

> ## 🚨 ACTIVE TASK DIRECTIVE FROM DUAL-BRAIN COORDINATOR (ROUND 15)
> **Salom, Claude!** 14-bosqich muvaffaqiyatli ratifikatsiya qilindi! Endi navbat yana to'liq senga keldi!
> 
> **Sening vazifang (Round 15):**
> 1. `.collab/COUNCIL.md` dagi **"15-Bosqich (Round 15)"** bo'limini o'qib chiq.
> 2. Kod bazamizni sinchiklab tahlil qil va yangi bir arxitekturaviy xatolik, chekka holat (edge-case) yoki renderlashdagi jiddiy kamchilikni top.
> 3. Topgan kamchiliging bo'yicha Antigravity bajarishi kerak bo'lgan aniq texnik vazifani `.collab/COUNCIL.md` ning `### 🟣 Claude Code Tanqidi va Vazifasi (Round 15)` bo'limiga yozib saqla!
> 4. Sen topshiriqni yozib bo'lishing bilan, Antigravity kodini o'zi ishlab chiqadi, sinovdan o'tkazadi va senga tekshiruv uchun topshiradi!




## 🛡️ Core Invariants (Zero-Defect Rules)
Whenever you inspect, refactor, or propose code for this project, you MUST strictly adhere to these 4 architectural rules:
1. **Zero Shape Stickers**: NEVER add `MSO_SHAPE.ROUNDED_RECTANGLE` or opaque cards over authentic templates. Text must be placed in native placeholders or transparent text boxes in computed safe zones.
2. **Safe-Margin Boundary Detection**: Templates frequently feature dark sidebars, diagonal ribbons, or header bands. Never place text overlapping these boundaries.
3. **Luminance & Contrast Awareness**: Background luminance ($Y = 0.299R + 0.587G + 0.114B$) determines text color ($Y < 128 \rightarrow$ white text, $Y \ge 128 \rightarrow$ dark text).
4. **Knuth-Plass Micro-Typography & Dual-Track**: Headlines must be balanced without orphan words (`va`, `bilan`, `uchun`, `and`, `to`, `in`). Slides must feature rich speaker notes (`slide.notes_slide`) and OpenXML native transitions (`p16:morph`).

## 🤝 How You and Antigravity Deliberate & Consult
We communicate and cross-review through the **`.collab/`** directory:
- **`.collab/COUNCIL.md`**: The live discussion board.
  - When the user asks you: *"Loyihani tekshir va muammolarni top"* or *"Antigravity bilan maslahatlash"*:
    1. Read `.collab/COUNCIL.md` to see recent architectural discussions and Antigravity's latest changes.
    2. Conduct your deep technical audit (search for edge cases, memory leaks, concurrency, or algorithmic flaws).
    3. Write your critique, suggestions, and proposals under `### 🟣 Claude Code Tahlili` in `.collab/COUNCIL.md`.
    4. Inform the user: *"Tahlilni `.collab/COUNCIL.md` fayliga yozdim, endi Antigravity bilan birgalikda yechimni ko'rib chiqishingiz mumkin."*

## 📁 Key Architecture Files
- `core/pptx_text_replacer.py`: Primary template text replacer, layout matcher, and OpenXML deck assembler.
- `core/micro_typography.py`: Knuth-Plass headline balancer and orphan word binder.
- `core/visual_rhetoric.py`: Visual rhetoric detector and Dual-Track speaker notes synthesizer.
- `core/slide_morph.py`: PowerPoint OpenXML native Morph and slide transitions engine.
- `core/content_planner.py`: AppFlowy Kanban content planner.
- `core/penpot_themes.py`: Penpot design tokens and palettes.
- `app.py`: Streamlit 5-tab web dashboard.
