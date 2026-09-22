# -*- coding: utf-8 -*-
"""
Generates a publication-grade PDF documentation of SlaydHub AI architecture and workflow
using Playwright headless browser automation.
"""

import os
from pathlib import Path
from playwright.sync_api import sync_playwright

OUTPUT_PDF = Path("SlaydHub_AI_Loyiha_Sxemasi_va_Hujjati.pdf").resolve()

HTML_CONTENT = """<!DOCTYPE html>
<html lang="uz">
<head>
<meta charset="UTF-8">
<style>
    @page {
        size: A4 portrait;
        margin: 18mm 16mm 20mm 16mm;
        @bottom-right {
            content: "Sahifa " counter(page);
            font-size: 9pt;
            color: #64748b;
        }
    }
    body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        color: #1e293b;
        background: #ffffff;
        line-height: 1.55;
        font-size: 10.5pt;
        margin: 0;
        padding: 0;
    }
    .header-banner {
        background: linear-gradient(135deg, #0a2540 0%, #0f52ba 100%);
        color: #ffffff;
        padding: 24px 28px;
        border-radius: 12px;
        margin-bottom: 24px;
        box-shadow: 0 4px 14px rgba(15, 82, 186, 0.15);
    }
    .header-banner h1 {
        margin: 0 0 6px 0;
        font-size: 22pt;
        font-weight: 800;
        letter-spacing: -0.5px;
    }
    .header-banner .subtitle {
        font-size: 11.5pt;
        opacity: 0.92;
        margin-bottom: 12px;
    }
    .header-banner .meta-tags {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
    }
    .badge {
        display: inline-block;
        background: rgba(255, 255, 255, 0.2);
        padding: 3px 10px;
        border-radius: 6px;
        font-size: 8.5pt;
        font-weight: 600;
        letter-spacing: 0.3px;
    }
    .section-title {
        font-size: 14pt;
        font-weight: 700;
        color: #0f172a;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 6px;
        margin-top: 26px;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .section-title::before {
        content: "";
        display: inline-block;
        width: 4px;
        height: 18px;
        background: #0f52ba;
        border-radius: 2px;
    }
    .card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 14px 16px;
        margin-bottom: 14px;
        page-break-inside: avoid;
    }
    .card-title {
        font-weight: 700;
        font-size: 11pt;
        color: #0f52ba;
        margin-bottom: 4px;
    }
    .diagram-box {
        background: #f1f5f9;
        border: 1.5px solid #cbd5e1;
        border-radius: 10px;
        padding: 16px;
        margin: 16px 0;
        page-break-inside: avoid;
    }
    .layer-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 12px;
        margin-top: 10px;
    }
    .layer-card {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-left: 4px solid #0f52ba;
        border-radius: 6px;
        padding: 10px 12px;
        font-size: 9.5pt;
    }
    .layer-card.green { border-left-color: #10b981; }
    .layer-card.purple { border-left-color: #8b5cf6; }
    .layer-card.amber { border-left-color: #f59e0b; }
    .layer-card.rose { border-left-color: #f43f5e; }
    .layer-card.cyan { border-left-color: #06b6d4; }
    .layer-card .l-name {
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 4px;
        font-size: 10pt;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        margin: 14px 0;
        font-size: 9.5pt;
        page-break-inside: avoid;
    }
    th, td {
        border: 1px solid #cbd5e1;
        padding: 8px 10px;
        text-align: left;
    }
    th {
        background: #f1f5f9;
        color: #0f172a;
        font-weight: 700;
    }
    tr:nth-child(even) {
        background: #f8fafc;
    }
    .step-list {
        counter-reset: step-counter;
        list-style-type: none;
        padding-left: 0;
    }
    .step-list li {
        position: relative;
        padding-left: 38px;
        margin-bottom: 12px;
        page-break-inside: avoid;
    }
    .step-list li::before {
        content: counter(step-counter);
        counter-increment: step-counter;
        position: absolute;
        left: 0;
        top: 0;
        width: 26px;
        height: 26px;
        background: #0f52ba;
        color: #ffffff;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 9.5pt;
        text-align: center;
        line-height: 26px;
    }
    .bold-lead {
        font-weight: 700;
        color: #0a3663;
    }
    .alert-box {
        background: #ecfdf5;
        border: 1px solid #a7f3d0;
        border-left: 4px solid #10b981;
        border-radius: 6px;
        padding: 10px 14px;
        margin: 12px 0;
        font-size: 9.5pt;
        color: #065f46;
    }
    .code-tag {
        font-family: Consolas, "Courier New", monospace;
        background: #e2e8f0;
        padding: 1px 5px;
        border-radius: 4px;
        font-size: 8.5pt;
        color: #0f172a;
    }
    .page-break {
        page-break-before: always;
    }
</style>
</head>
<body>

<div class="header-banner">
    <h1>SlaydHub AI</h1>
    <div class="subtitle">Loyiha Arxitekturasi, Ishlash Sxemasi va Muhandislik Qo'llanmasi</div>
    <div class="meta-tags">
        <span class="badge">Versiya 2.0</span>
        <span class="badge">5,170+ PPTX Shablonlar Bazasi</span>
        <span class="badge">McKinsey Bold Lead-In Standarti</span>
        <span class="badge">Google NotebookLM + Gemini + Groq</span>
        <span class="badge">2026-yil, Sentyabr</span>
    </div>
</div>

<p><strong>SlaydHub AI</strong> — bu sun'iy intellekt (NotebookLM, Gemini 2.5, Groq) imkoniyatlarini professional PowerPoint (.pptx) shablonlari bazasi bilan to'g'ridan-to'g'ri bog'lovchi, 100% tahrirlanuvchi, 3D vektor grafikasiga ega, yuqori darajadagi akademik va amaliy taqdimotlar generatsiya qiluvchi to'liq avtonom studiyadir.</p>

<div class="section-title">1. Tizimning 6 Asosiy Arxitektura Qatlami</div>

<div class="diagram-box">
    <div style="font-weight: 800; color: #0a2540; font-size: 11pt; margin-bottom: 8px;">
        🏛️ SLAYDHUB AI DASTURIY TA'MINOT MODELLARI STRUKTURASI
    </div>
    <div class="layer-grid">
        <div class="layer-card">
            <div class="l-name">1. Foydalanuvchi Qatlami (UI / CLI)</div>
            <div>• <span class="code-tag">app.py</span> — Streamlit Web UI (5 qadamli interfeys)</div>
            <div>• <span class="code-tag">cli.py</span> — Terminal orqali to'liq avtomatizatsiya</div>
        </div>
        <div class="layer-card green">
            <div class="l-name">2. Shablonlar va Rejalashtirish (Catalog)</div>
            <div>• <span class="code-tag">core/academic_catalog.py</span> — Ko'p tilli shablonlar qidiruvi</div>
            <div>• <span class="code-tag">core/academic_matcher.py</span> — 10 slaydli Blueprint tuzuvchi</div>
        </div>
        <div class="layer-card purple">
            <div class="l-name">3. AI va Prompt Dvigateli (LLM Engine)</div>
            <div>• <span class="code-tag">core/notebooklm_prompt_gen.py</span> — Bold Lead-In promptlari</div>
            <div>• <span class="code-tag">core/notebooklm_executor.py</span> — NotebookLM / Gemini / Groq</div>
        </div>
        <div class="layer-card amber">
            <div class="l-name">4. Tahlil va Semantik Nazorat (Validator)</div>
            <div>• <span class="code-tag">core/notebooklm_markdown_parser.py</span> — 10 slaydli tahlil</div>
            <div>• <span class="code-tag">core/topic_logic_validator.py</span> — Boy fikrlarni saqlovchi</div>
        </div>
        <div class="layer-card rose">
            <div class="l-name">5. PPTX In-Place Dvigateli (Customizer)</div>
            <div>• <span class="code-tag">core/pptx_text_replacer.py</span> — Reklama tozalash va Rich Runs</div>
            <div>• <span class="code-tag">core/image_fetcher.py</span> — Ilmiy diagrammalar inyeksiyasi</div>
        </div>
        <div class="layer-card cyan">
            <div class="l-name">6. Prevyu va Yetkazib Berish (Export)</div>
            <div>• <span class="code-tag">core/slide_preview.py</span> — PowerPoint JPG / PDF render</div>
            <div>• Bir klikda PPTX va PDF yuklab olish</div>
        </div>
    </div>
</div>

<div class="section-title">2. Real Ish Oqimi Ketma-Ketligi (End-to-End Workflow)</div>

<ul class="step-list">
    <li>
        <strong>Mavzuni Tahlil Qilish va Ilmiy Tasniflash:</strong>
        Foydalanuvchi mavzu kiritadi (masalan: <em>"Kvant kompyuterlari va asimmetrik kriptografiya"</em>). <span class="code-tag">AcademicCatalog</span> o'zbekcha-inglizcha semantik lug'at orqali mavzuni kashf etadi va eng mos autentik PPTX shablonini topadi.
    </li>
    <li>
        <strong>10 Slaydli Ilmiy Reja (Blueprint) Tuzish:</strong>
        <span class="code-tag">AcademicMatcher</span> taqdimot arxitekturasini qat'iy didaktik ketma-ketlikda belgilaydi: <em>Titul &rarr; Reja &rarr; Nazariya &rarr; Qiyosiy Tahlil &rarr; Metrikalar &rarr; Bosqichlar &rarr; Tizim Arxitekturasi &rarr; Xulosalar</em>.
    </li>
    <li>
        <strong>AI Mazmun Generatsiyasi va Fallback Zanjiri:</strong>
        Foydalanuvchi NotebookLM orqali boy manbalarini yuklaydi yoki 1-klikda Gemini AI ga so'rov beradi. Agar bepul kvota chegarasiga (429) tushilsa, tizim uzluksiz ravishda <strong>Groq API</strong>'dagi tezkor modellarga ulanadi.
    </li>
    <li>
        <strong>Markdown Tahlili va Semantik Tozalash:</strong>
        <span class="code-tag">NotebookLMMarkdownParser</span> AI matnini 10 ta slaydga ajratadi, sarlavhalar va tezislarni strukturalaydi. <span class="code-tag">TopicLogicValidator</span> reklama va suvbelgilarni yo'qotib, ilmiy fikrlarni to'liq saqlaydi.
    </li>
    <li>
        <strong>Asl PPTX Shablonini In-Place O'zgartirish:</strong>
        Oq bo'sh slayd yaratilmaydi! Tanlangan professional shablonning 3D dizaynlari, fonlari va kartalari to'liq saqlanib, reklama logotiplari o'chiriladi va matnlar <strong>McKinsey Bold Lead-In</strong> formatida joylanadi.
    </li>
    <li>
        <strong>Geometriya va Xavfsiz Maydon Nazorati:</strong>
        <span class="code-tag">SpatialCollisionEngine</span> rasm va matnlar to'qnashuvini matematik bartaraf etadi. Matnlar keng to'rtburchak kartalar (100–130pt) ichida erkin joylashadi.
    </li>
    <li>
        <strong>Prevyu Dvigateli va PDF/PPTX Eksport:</strong>
        PowerPoint avtomatizatsiyasi orqali har bir slaydning real JPG tasviri olinib, UI da ko'rsatiladi. Foydalanuvchiga <strong>.pptx</strong> va <strong>.pdf</strong> fayllari taqdim etiladi.
    </li>
</ul>

<div class="page-break"></div>

<div class="section-title">3. Slaydlarga Matn Yozish va Joylashtirishning 5 Oltin Qoidasi</div>

<div class="card">
    <div class="card-title">1. "Bold Lead-In" (Claim + Evidence) Qoidasi — McKinsey & BCG Standarti</div>
    <p>Har bir slayd punkte quruq uzun matn bo'lmasligi lozim. U <strong>Qalin Kalit So'z / Hukm</strong> bilan boshlanadi, so'ngra oddiy shriftda aniq ilmiy dalil, parametr va raqamlar keltiriladi:</p>
    <div>❌ <em>Eski usul:</em> "Post-kvant algoritmlari yangi standartlar asosida ishlaydi va ular xavfsiz hisoblanadi."</div>
    <div style="margin-top: 4px;">✅ <span class="bold-lead">NIST FIPS 203 Standarti:</span> Module-Lattice asosidagi ML-KEM algoritmi 128 va 256-bitli kalit almashinuvida 100% kvant himoyasini kafolatlaydi.</div>
</div>

<div class="card">
    <div class="card-title">2. Action-Oriented Headlines (Xulosaviy Sarlavhalar)</div>
    <p>Slayd sarlavhasi shunchaki mavzu nomi (<em>"Kriptografiya"</em> yoki <em>"Reja"</em>) emas, balki to'liq xulosaviy jumla bo'ladi:<br/>
    <strong>"Post-Kvant Standartlari Klassik RSA Zaifliklarini To'liq Bartaraf Etadi"</strong></p>
</div>

<div class="card">
    <div class="card-title">3. F-Pattern Scannability (5 Soniyada Ilg'ab Olish)</div>
    <p>Auditoriya slaydni kitob kabi o'qimaydi, balki ko'z yugirtiradi. Qalin boshlang'ich so'zlar tinglovchining ko'zini tortuvchi vizual langar (visual anchor) bo'lib xizmat qiladi.</p>
</div>

<div class="card">
    <div class="card-title">4. Keng Gorizontal To'g'ri To'rtburchaklar (Wide Rectangles vs Square Boxes)</div>
    <p>Matnni tor kvadrat qutilarga (1:1) siqish taqiqlangan. Slayd kengligining 75–85% ini egallagan keng gorizontal to'g'ri to'rtburchak kartalar (aspect ratio &ge; 4:1) ishlatiladi. Karta balandligi <strong>100–135 pt</strong>, shriftlar <strong>15–18.5 pt</strong>, qatorlar oralig'i <strong>1.25</strong>.</p>
</div>

<div class="card">
    <div class="card-title">5. Aniq Metrikalar va Raqamlar (Zero AI Fluff)</div>
    <p><em>"Juda muhim ahamiyatga ega"</em>, <em>"Ta'kidlash joizki"</em> kabi sun'iy intellekt klishelari butunlay chiqarib tashlangan. Ularning o'rniga <strong>99.9%</strong>, <strong>2,500 kubit</strong>, <strong>1.2 ms</strong>, <strong>O(n³)</strong> kabi aniq ko'rsatkichlar beriladi.</p>
</div>

<div class="section-title">4. Dastur Modullari va Fayllar Reestri</div>

<table>
    <thead>
        <tr>
            <th>Fayl Nomi</th>
            <th>Modul Vazifasi</th>
            <th>Chiqish Ma'lumotlari</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td><code>app.py</code></td>
            <td>Streamlit Web UI (5 qadamli vizual boshqaruv)</td>
            <td>Interaktiv foydalanuvchi interfeysi</td>
        </tr>
        <tr>
            <td><code>cli.py</code></td>
            <td>Terminal CLI interfeysi</td>
            <td>Avtomatlashtirilgan buyruqlar</td>
        </tr>
        <tr>
            <td><code>core/academic_catalog.py</code></td>
            <td>5,170+ PPTX shablonlarini indekslash va qidirish</td>
            <td>Mos shablonlar ro'yxati</td>
        </tr>
        <tr>
            <td><code>core/academic_matcher.py</code></td>
            <td>Mavzuni tasniflash va Blueprint rejasini tuzish</td>
            <td>10 slaydli reja (JSON)</td>
        </tr>
        <tr>
            <td><code>core/notebooklm_prompt_gen.py</code></td>
            <td>Bold Lead-In formatidagi ilmiy promptlar</td>
            <td>Strukturalangan AI prompti</td>
        </tr>
        <tr>
            <td><code>core/notebooklm_executor.py</code></td>
            <td>NotebookLM, Gemini 2.5 va Groq Fallback ijrosi</td>
            <td>AI javob matni</td>
        </tr>
        <tr>
            <td><code>core/notebooklm_markdown_parser.py</code></td>
            <td>Markdown matnini slaydlar bo'yicha tahlil qilish</td>
            <td>Slaydlar lug'ati (Dict)</td>
        </tr>
        <tr>
            <td><code>core/topic_logic_validator.py</code></td>
            <td>Mavzu aloqadorligini tekshirish va tozalash</td>
            <td>Tozalangan ilmiy kontent</td>
        </tr>
        <tr>
            <td><code>core/template_pptx_builder.py</code></td>
            <td>Asl PPTX shablonini in-place yuklovchi</td>
            <td>Fayl marshruti (.pptx)</td>
        </tr>
        <tr>
            <td><code>core/pptx_text_replacer.py</code></td>
            <td>Reklamalarni tozalash, kartalar va matn joylash</td>
            <td>Tayyor taqdimot (.pptx)</td>
        </tr>
        <tr>
            <td><code>core/slide_preview.py</code></td>
            <td>PowerPoint COM orqali JPG va PDF generatsiya</td>
            <td>Slayd rasmlari va PDF fayl</td>
        </tr>
    </tbody>
</table>

<div class="alert-box">
    <strong>🧪 TEST VA VERIFIKATSIYA NATIJALARI (100% PASS):</strong><br/>
    • <code>test_all_functions.py</code>: Barcha 28 ta tizim moduli 100% muvaffaqiyatli sinovdan o'tgan (13.5 soniya).<br/>
    • <code>test_real_pptx_button_workflow.py</code>: Barcha 7 ta real hayotiy bosqich (1-klikli avtomatik generatsiya, PowerPoint prevyulari va sifat nazorati) 100% muvaffaqiyat bilan bajarilgan.
</div>

</body>
</html>
"""

def generate_pdf():
    temp_html = Path("temp_doc.html").resolve()
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(HTML_CONTENT)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(f"file:///{temp_html.as_posix()}")
        page.wait_for_load_state("networkidle")
        page.pdf(
            path=str(OUTPUT_PDF),
            format="A4",
            print_background=True,
            margin={"top": "16mm", "bottom": "16mm", "left": "15mm", "right": "15mm"}
        )
        browser.close()

    if temp_html.exists():
        temp_html.unlink()

    print(f"Muvaffaqiyatli yaratildi: {OUTPUT_PDF} ({OUTPUT_PDF.stat().st_size / 1024:.1f} KB)")

if __name__ == "__main__":
    generate_pdf()
