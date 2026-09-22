# -*- coding: utf-8 -*-
"""
SlaydHub AI — Professional Academic Presentation Studio (Native PPTX Template Edition)
Customizes 5,170+ authentic, pre-designed PPTX templates with NotebookLM and Gemini AI content.
"""

import os
import sys
import glob
import json
import asyncio
from pathlib import Path
from PIL import Image
import streamlit as st

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from core.config import config
from core.academic_catalog import (
    get_academic_catalog,
    ACADEMIC_CATEGORIES,
    SLIDE_LAYOUT_TYPES
)
from core.slide_archetypes import SLIDE_ARCHETYPES
from core.academic_matcher import get_academic_matcher
from core.notebooklm_prompt_gen import NotebookLMPromptGenerator, generate_notebooklm_prompt
from core.notebooklm_markdown_parser import NotebookLMMarkdownParser
from core.notebooklm_executor import get_notebooklm_executor
from core.template_pptx_builder import get_academic_pptx_builder
from core.pptx_text_replacer import replace_presentation_content
from core.slide_preview import get_template_thumbnail, get_presentation_previews
from core.content_planner import AppFlowyContentPlanner
from core.fooocus_engine import FooocusImageGenerator
from core.penpot_themes import PenpotThemeManager
import slide_manager

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="SlaydHub AI — Academic Studio Pro",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== PREMIUM MODERN STYLING ====================
st.markdown("""
<style>
    /* Global Layout & Padding */
    .block-container {
        padding-top: 2.8rem !important;
        padding-bottom: 3.5rem !important;
        max-width: 1400px;
    }
    
    /* Modern Glassmorphic Hero Banner */
    .hero-banner {
        background: linear-gradient(135deg, #0b1329 0%, #1e293b 50%, #172554 100%);
        border: 1px solid rgba(59, 130, 246, 0.25);
        color: #f8fafc;
        padding: 26px 32px;
        border-radius: 18px;
        margin-bottom: 24px;
        box-shadow: 0 12px 30px -8px rgba(0, 0, 0, 0.45);
        position: relative;
        overflow: hidden;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(37, 99, 235, 0.25);
        border: 1px solid rgba(96, 165, 250, 0.4);
        color: #93c5fd;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        margin-bottom: 12px;
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        margin-bottom: 8px;
        letter-spacing: -0.02em;
        background: linear-gradient(135deg, #ffffff 0%, #cbd5e1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .hero-subtitle {
        font-size: 1.02rem;
        color: #94a3b8;
        max-width: 860px;
        line-height: 1.55;
    }

    /* Stepper Navigation Indicator */
    .stepper-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: rgba(15, 23, 42, 0.65);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 14px;
        padding: 12px 20px;
        margin-bottom: 26px;
        gap: 12px;
        overflow-x: auto;
    }
    .step-item {
        display: flex;
        align-items: center;
        gap: 10px;
        font-size: 0.88rem;
        font-weight: 600;
        color: #64748b;
        white-space: nowrap;
    }
    .step-item.active {
        color: #38bdf8;
    }
    .step-item.done {
        color: #10b981;
    }
    .step-num {
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.8rem;
        font-weight: 700;
        background: #1e293b;
        color: #94a3b8;
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .step-item.active .step-num {
        background: #0284c7;
        color: #ffffff;
        border-color: #38bdf8;
        box-shadow: 0 0 12px rgba(56, 189, 248, 0.4);
    }
    .step-item.done .step-num {
        background: #059669;
        color: #ffffff;
        border-color: #10b981;
    }
    .step-arrow {
        color: #334155;
        font-size: 0.9rem;
    }

    /* Section Cards */
    .section-card {
        background: rgba(15, 23, 42, 0.45);
        border: 1px solid rgba(255, 255, 255, 0.07);
        border-radius: 16px;
        padding: 22px;
        margin-bottom: 22px;
    }
    .section-header {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 16px;
    }
    .section-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #f1f5f9;
        margin: 0;
    }
    .section-subtitle {
        font-size: 0.86rem;
        color: #94a3b8;
        margin-top: 2px;
    }

    /* Slide Card Grid */
    .slide-preview-box {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        padding: 10px;
        margin-bottom: 12px;
        transition: transform 0.2s ease, border-color 0.2s ease, box-shadow 0.2s ease;
    }
    .slide-preview-box:hover {
        transform: translateY(-2px);
        border-color: rgba(56, 189, 248, 0.4);
        box-shadow: 0 8px 20px -6px rgba(0, 0, 0, 0.5);
    }
    .slide-chip {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 0.72rem;
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    .chip-cover { background: #3b82f6; color: #ffffff; }
    .chip-theory { background: #8b5cf6; color: #ffffff; }
    .chip-vs { background: #ec4899; color: #ffffff; }
    .chip-metrics { background: #10b981; color: #ffffff; }
    .chip-grid { background: #06b6d4; color: #ffffff; }
    .chip-process { background: #f59e0b; color: #ffffff; }
    .chip-timeline { background: #6366f1; color: #ffffff; }
    .chip-conclusion { background: #14b8a6; color: #ffffff; }

    /* Topic Pill Buttons */
    .chip-btn {
        background: rgba(30, 41, 59, 0.7);
        border: 1px solid rgba(255, 255, 255, 0.1);
        color: #cbd5e1;
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 0.82rem;
        cursor: pointer;
        display: inline-block;
        margin: 4px;
        transition: all 0.2s ease;
    }

    /* Modern Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        margin-top: 0.5rem;
        margin-bottom: 1.5rem;
        padding: 6px;
        background: rgba(15, 23, 42, 0.65);
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.08);
        overflow-x: auto;
    }
    .stTabs [data-baseweb="tab"] {
        height: auto;
        padding: 10px 18px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.95rem;
        color: #94a3b8;
        white-space: normal;
        line-height: 1.4;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
        color: #ffffff !important;
        font-weight: 700;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.35);
    }

    /* Buttons */
    .stButton>button {
        border-radius: 10px;
        font-weight: 700;
        padding: 10px 20px;
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.25);
    }

    /* Custom Result Card */
    .success-result-box {
        background: linear-gradient(135deg, rgba(16, 185, 129, 0.12) 0%, rgba(5, 150, 105, 0.06) 100%);
        border: 1px solid rgba(16, 185, 129, 0.35);
        border-radius: 16px;
        padding: 24px;
        margin: 20px 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize singletons
catalog = get_academic_catalog()
matcher = get_academic_matcher()
executor = get_notebooklm_executor()
builder = get_academic_pptx_builder()

# Initialize session state
if "blueprint" not in st.session_state:
    st.session_state.blueprint = None
if "generated_prompt" not in st.session_state:
    st.session_state.generated_prompt = ""
if "compact_notebooklm_prompt" not in st.session_state:
    st.session_state.compact_notebooklm_prompt = ""
if "notebooklm_response" not in st.session_state:
    st.session_state.notebooklm_response = ""
if "parsed_content" not in st.session_state:
    st.session_state.parsed_content = None
if "created_pptx_path" not in st.session_state:
    st.session_state.created_pptx_path = None
if "md_input_version" not in st.session_state:
    st.session_state.md_input_version = 0

# ==================== SIDEBAR ====================
with st.sidebar:
    st.markdown("""
    <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 14px;">
        <div style="background: linear-gradient(135deg, #2563eb, #1d4ed8); width: 48px; height: 48px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 26px; box-shadow: 0 4px 14px rgba(37, 99, 235, 0.45);">
            🎓
        </div>
        <div>
            <div style="font-size: 1.25rem; font-weight: 800; line-height: 1.15; color: #f8fafc;">SlaydHub AI</div>
            <div style="font-size: 0.78rem; color: #38bdf8; font-weight: 600;">Academic Studio Pro v3.5</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Quick Stats Metric Box
    st.markdown(f"""
    <div style="background: rgba(15, 23, 42, 0.6); border: 1px solid rgba(255, 255, 255, 0.08); border-radius: 12px; padding: 12px; margin-bottom: 16px;">
        <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
            <span style="font-size: 0.8rem; color: #94a3b8;">📚 Haqiqiy Shablonlar:</span>
            <span style="font-size: 0.85rem; font-weight: 700; color: #38bdf8;">{len(catalog.templates):,} ta</span>
        </div>
        <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
            <span style="font-size: 0.8rem; color: #94a3b8;">🧩 Vizual Arxetiplar:</span>
            <span style="font-size: 0.85rem; font-weight: 700; color: #10b981;">12 ta</span>
        </div>
        <div style="display: flex; justify-content: space-between;">
            <span style="font-size: 0.8rem; color: #94a3b8;">⚡ Yig'ish Dvigatellari:</span>
            <span style="font-size: 0.85rem; font-weight: 700; color: #f59e0b;">3 ta</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # API Status
    api_key_ok = bool(config.gemini_api_key)
    api_status_html = """
    <div style="display: flex; align-items: center; gap: 8px; font-size: 0.8rem; margin-bottom: 16px;">
        <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #10b981; box-shadow: 0 0 8px #10b981;"></span>
        <span style="color: #cbd5e1;">Gemini AI: <b>Ulangan va Faol</b></span>
    </div>
    """ if api_key_ok else """
    <div style="display: flex; align-items: center; gap: 8px; font-size: 0.8rem; margin-bottom: 16px;">
        <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #ef4444;"></span>
        <span style="color: #cbd5e1;">Gemini API: <i>Kalit kiritilmagan</i></span>
    </div>
    """
    st.markdown(api_status_html, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### ⚙️ **Konfiguratsiya**")
    selected_lang = st.selectbox(
        "🌐 Taqdimot Tili:",
        ["uz", "ru", "en"],
        format_func=lambda x: {"uz": "🇺🇿 O'zbekcha", "ru": "🇷🇺 Русский", "en": "🇬🇧 English"}[x],
        index=0
    )

    theme_category = st.selectbox(
        "🏛️ Asosiy Soha (Kategoriya):",
        ["auto"] + list(ACADEMIC_CATEGORIES.keys()),
        format_func=lambda x: "✨ Avtomatik Aniqlash" if x == "auto" else ACADEMIC_CATEGORIES[x][f"name_{selected_lang}"]
    )

    st.markdown("---")
    if st.button("🔄 Yangi Taqdimot Boshlash", use_container_width=True):
        st.session_state.blueprint = None
        st.session_state.generated_prompt = ""
        st.session_state.compact_notebooklm_prompt = ""
        st.session_state.notebooklm_response = ""
        st.session_state.parsed_content = None
        st.session_state.created_pptx_path = None
        st.session_state.md_input_version = st.session_state.get("md_input_version", 0) + 1
        st.rerun()


def render_presentation_preview_ui(pptx_file: str, key_suffix: str = ""):
    """
    Renders an interactive high-resolution slide deck preview gallery in Streamlit.
    """
    if not pptx_file or not os.path.exists(pptx_file):
        return

    st.markdown("---")
    st.markdown("""
    <div class="section-header">
        <span style="font-size: 1.3rem;">👁️</span>
        <div>
            <h4 class="section-title">Tayyor Bo'lgan Taqdimot Slaydlarini Ko'rish (Prevyu)</h4>
            <div class="section-subtitle">Slaydlarni brauzerda to'g'ridan-to'g'ri ko'rib chiqishingiz mumkin</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("⏳ Slaydlar prevyusi tayyorlanmoqda..."):
        slide_images = get_presentation_previews(pptx_file)

    if slide_images:
        st.markdown(f"**Jami tayyorlangan slaydlar:** `{len(slide_images)} ta`")

        skey = f"preview_slide_idx_{key_suffix}"
        if skey not in st.session_state:
            st.session_state[skey] = 0

        cur_idx = min(st.session_state[skey], len(slide_images) - 1)

        c_prev_nav1, c_prev_nav2, c_prev_nav3 = st.columns([1, 2, 1])
        with c_prev_nav1:
            if st.button("⬅️ Oldingi Slayd", key=f"prev_btn_{key_suffix}", use_container_width=True, disabled=(cur_idx == 0)):
                st.session_state[skey] = max(0, cur_idx - 1)
                st.rerun()
        with c_prev_nav2:
            selected_slide_num = st.selectbox(
                "Slaydni tanlang:",
                options=list(range(1, len(slide_images) + 1)),
                index=cur_idx,
                format_func=lambda n: f"📄 #{n:02d}-Slayd",
                key=f"slide_sel_{key_suffix}",
                label_visibility="collapsed"
            )
            if selected_slide_num - 1 != cur_idx:
                st.session_state[skey] = selected_slide_num - 1
                st.rerun()
        with c_prev_nav3:
            if st.button("Keyingi Slayd ➡️", key=f"next_btn_{key_suffix}", use_container_width=True, disabled=(cur_idx == len(slide_images) - 1)):
                st.session_state[skey] = min(len(slide_images) - 1, cur_idx + 1)
                st.rerun()

            st.image(active_img, caption=f"📄 Slayd #{cur_idx + 1} / {len(slide_images)} — {os.path.basename(pptx_file)}", use_container_width=True)

            # Dual-Track Speaker Notes Preview
            try:
                from pptx import Presentation as PPTXPrs
                prs_p = PPTXPrs(pptx_file)
                if cur_idx < len(prs_p.slides):
                    sl = prs_p.slides[cur_idx]
                    if sl.has_notes_slide:
                        notes_txt = sl.notes_slide.notes_text_frame.text.strip()
                        if notes_txt:
                            with st.expander(f"🎙️ #{cur_idx + 1}-Slayd Uchun Spiker Nutq Ssenariysi va Dalillar (Dual-Track)", expanded=False):
                                st.text(notes_txt)
            except Exception:
                pass

        with st.expander("🖼️ Barcha Slaydlarning Mini-Eskizlarini Ko'rish (Galereya)", expanded=False):
            g_cols = 4
            g_chunks = [slide_images[i:i + g_cols] for i in range(0, len(slide_images), g_cols)]
            for g_chunk in g_chunks:
                cols = st.columns(len(g_chunk))
                for idx_c, img_path in enumerate(g_chunk):
                    real_s_num = slide_images.index(img_path) + 1
                    with cols[idx_c]:
                        st.markdown(f"**Slayd #{real_s_num:02d}**")
                        if os.path.exists(img_path):
                            st.image(img_path, use_container_width=True)
    else:
        st.info("ℹ️ Slayd prevyulari PowerPoint orqali avtomatik keshlanadi.")


# ==================== MAIN TABS ====================
tab_generator, tab_appflowy, tab_penpot, tab_catalog, tab_converter = st.tabs([
    "🚀 Slayd Studiyasi (NotebookLM & Native PPTX)",
    "📋 AppFlowy Rejalar & Doska (Kanban)",
    "🎨 Penpot & Fooocus AI Studiyasi",
    "📚 5,170+ PPTX Shablonlar Bazasi",
    "🖼️ OCR & Maxsus Shablon Konvertori"
])


# =========================================================================
# TAB 1: NATIVE PPTX SLIDE GENERATOR & NOTEBOOKLM STUDIO
# =========================================================================
with tab_generator:
    # Hero Banner
    st.markdown("""
    <div class="hero-banner">
        <div class="hero-badge">⚡ Native PPTX In-Place Customizer</div>
        <div class="hero-title">Professional Ilmiy va Akademik Slayd Studiyasi</div>
        <div class="hero-subtitle">
            Haqiqiy PowerPoint shablonlarining 3D grafikasi, kartalari, fonlari va vektor elementlarini 100% buzilmagan holda saqlab, 
            NotebookLM va Gemini AI orqali "Bir slayd — bir fikr" qoidasiga asoslangan mukammal taqdimotlar yarating.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Dynamic Visual Stepper Progress Bar
    has_bp = bool(st.session_state.blueprint)
    has_content = bool(st.session_state.notebooklm_response.strip() or st.session_state.parsed_content)
    has_pptx = bool(st.session_state.created_pptx_path and os.path.exists(st.session_state.created_pptx_path))

    s1_cls = "done" if has_bp else "active"
    s2_cls = "done" if (has_bp and has_content) else ("active" if has_bp else "")
    s3_cls = "done" if (has_content and has_pptx) else ("active" if (has_bp and not has_pptx) else "")
    s4_cls = "done" if has_pptx else ("active" if has_content else "")

    st.markdown(f"""
    <div class="stepper-container">
        <div class="step-item {s1_cls}">
            <div class="step-num">1</div>
            <span>Mavzu & Format</span>
        </div>
        <div class="step-arrow">➔</div>
        <div class="step-item {s2_cls}">
            <div class="step-num">2</div>
            <span>Shablon & Arxetiplar</span>
        </div>
        <div class="step-arrow">➔</div>
        <div class="step-item {s3_cls}">
            <div class="step-num">3</div>
            <span>NotebookLM & AI Kontent</span>
        </div>
        <div class="step-arrow">➔</div>
        <div class="step-item {s4_cls}">
            <div class="step-num">4</div>
            <span>Tayyor PPTX Eksport</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ---------------------------------------------------------------------
    # STEP 1: TOPIC, OBJECTIVE & PRESENTER METADATA
    # ---------------------------------------------------------------------
    st.markdown("""
    <div class="section-header">
        <span style="font-size: 1.4rem;">🎯</span>
        <div>
            <h3 class="section-title">1-Qadam: Mavzu va Akademik Formatni Belgilang</h3>
            <div class="section-subtitle">Tadqiqot mavzusi, formati, slaydlar soni va titul slayd rekvizitlarini kiriting</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_input1, col_input2, col_input3 = st.columns([2.5, 1.5, 1])
    with col_input1:
        default_topic = st.session_state.get("last_topic", "Kvant kompyuterlari va asimmetrik kriptografiya")
        topic_input = st.text_input(
            "📝 Taqdimot Mavzusi / Tadqiqot Nomi:",
            placeholder="Masalan: Kvant kompyuterlari, Biotibbiyot, Neyron tarmoqlar...",
            value=default_topic,
            help="Taqdimotning to'liq akademik yoki biznes mavzusi"
        )
    with col_input2:
        presentation_mode_input = st.selectbox(
            "🎯 Taqdimot Formati / Maqsadi:",
            options=["scientific_defense_bmi", "general_academic", "business_project"],
            format_func=lambda x: {
                "scientific_defense_bmi": "🎓 Ilmiy Himoya / BMI / Dissertatsiya",
                "general_academic": "📚 Akademik Ma'ruza / Darslik",
                "business_project": "💼 Biznes / Startap Taqdimoti"
            }[x],
            index=0
        )
    with col_input3:
        slide_count_input = st.selectbox("📊 Slaydlar Soni:", [5, 8, 10, 12, 15], index=2)

    # Titul Slaydi Rekvizitlari (Presenter Metadata)
    with st.expander("📋 Titul Slaydi Ma'lumotlari (Muallif, Tashkilot, Ilmiy Rahbar)", expanded=False):
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            p_author = st.text_input("👤 Muallif F.I.Sh.:", value="Alisher Vohidov", placeholder="Masalan: Alisher Vohidov")
            p_org = st.text_input("🏢 Tashkilot / Universitet:", value="O'zbekiston Milliy Universiteti / Axborot xavfsizligi kafedrasi")
        with c_m2:
            p_supervisor = st.text_input("👨‍🏫 Ilmiy Rahbar (unvoni bilan):", value="Ilmiy rahbar: DSc., prof. A. Karimov")
            p_email = st.text_input("📧 Tadbir / Pochta / Sana:", value="alisher@univ.edu.uz | Himoya 2026")

    sources_notes_input = st.text_area(
        "📚 Mavzuga oid manbalar, ilmiy maqolalar yoki asosiy tezislar (ixtiyoriy):",
        placeholder="NotebookLM ga yuklangan manbalar bo'yicha qisqacha ma'lumotlar, tadqiqot gipotezasi yoki asosiy xulosalar...",
        height=70
    )

    with st.expander("ℹ️ Akademik Taqdimot Metodologiyasi & Sifat Standartlari", expanded=False):
        st.markdown("""
        - **✨ Boy Mazmun & Ko'p Qirrali Vizual Tahlil**: Slaydlar qisqa yoki yuzaki bo'lmasdan, chuqur ilmiy/amaliy tahlil, aniq ko'rsatkichlar, formulalar va faktlarga boy tuziladi.
        - **🎨 Rang-barang Vizual Arxetiplar**: Bir xil monoton qoliplar o'rniga har bir slayd o'z turiga mos ravishda KPI metrikalar, xronologik jarayonlar, ko'p ustunli tahlil yoki qiyosiy bloklarda namoyon bo'ladi.
        - **🎯 Action-Oriented (Xulosaviy) Sarlavhalar**: Slayd sarlavhasi shunchaki mavzuni emas, balki asosiy analitik xulosani aks ettiradi (Masalan: *"ML-KEM Kvant Xavfsizligini 100% Kafolatlaydi"*).
        - **🛡️ Format va Sifat Kafolati**: Slayd maydonidan chiqib ketish, matnlar ustma-ust tushishi va reklama elementlari to'liq bartaraf etilgan.
        """)

    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
    btn_match = st.button("🔍 Shablonni Tanlash va Ishchi Loyihani Boshlash", use_container_width=True, type="primary")

    if btn_match and topic_input.strip():
        st.session_state.last_topic = topic_input.strip()
        cat_param = None if theme_category == "auto" else theme_category
        presenter_dict = {
            "author_name": p_author.strip(),
            "organization": p_org.strip(),
            "supervisor": p_supervisor.strip(),
            "email": p_email.strip(),
            "event_name": p_email.strip()
        }
        with st.spinner("⏳ Mavzuga mos eng yuqori sifatli haqiqiy PPTX shablon tanlanmoqda..."):
            bp = matcher.build_blueprint(
                topic=topic_input.strip(),
                slide_count=slide_count_input,
                language=selected_lang,
                selected_category=cat_param,
                presentation_mode=presentation_mode_input,
                presenter_info=presenter_dict
            )
            st.session_state.blueprint = bp
            st.session_state.generated_prompt = generate_notebooklm_prompt(bp)
            
            category_title = ACADEMIC_CATEGORIES.get(cat_param, {}).get("name_uz", "Akademik / Ilmiy") if cat_param else "Akademik / Ilmiy"
            st.session_state.compact_notebooklm_prompt = NotebookLMPromptGenerator.generate_compact_markdown_prompt(
                topic=topic_input.strip(),
                domain=category_title,
                author_info=f"{p_author.strip()} ({p_org.strip()})",
                sources_summary=sources_notes_input.strip(),
                lang=selected_lang
            )
            st.session_state.created_pptx_path = None
            st.rerun()

    # ---------------------------------------------------------------------
    # STEP 2: BLUEPRINT, TEMPLATE SELECTION & ARCHETYPES
    # ---------------------------------------------------------------------
    if st.session_state.blueprint:
        bp = st.session_state.blueprint
        st.markdown("---")
        st.markdown("""
        <div class="section-header">
            <span style="font-size: 1.4rem;">🎨</span>
            <div>
                <h3 class="section-title">2-Qadam: Shablon va Slayd Shakllari (Arxetiplar)</h3>
                <div class="section-subtitle">Har bir slaydning vizual shaklini (qiyoslash, metrika, vaqt xaritasi) 1-klik orqali erkin o'zgartirishingiz mumkin</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Template Details
        c_bp_info1, c_bp_info2 = st.columns([3, 1])
        with c_bp_info1:
            st.markdown(f"**Asosiy Tanlangan PPTX Shablon:** `{bp['primary_deck_title']}`")
            st.caption(f"📁 Asl fayl: `{os.path.basename(bp['primary_pptx_path'])}`")
        with c_bp_info2:
            st.markdown(f"🏷️ Soha: **{bp['category_name']}**")

        # Alternative Templates Selector & Live Preview
        matching_decks = bp.get("matching_decks", [])
        if len(matching_decks) > 1:
            deck_options = {d["pptx_path"]: f"{d['title']} ({os.path.basename(d['pptx_path'])})" for d in matching_decks}
            cur_sel_idx = list(deck_options.keys()).index(bp["primary_pptx_path"]) if bp["primary_pptx_path"] in deck_options else 0
            
            c_sel_dd, c_sel_thumb = st.columns([3, 1])
            with c_sel_dd:
                selected_pptx = st.selectbox(
                    "🔄 Boshqa mos keluvchi shablonga almashtirish:",
                    options=list(deck_options.keys()),
                    format_func=lambda x: deck_options[x],
                    index=cur_sel_idx
                )
                if selected_pptx != bp["primary_pptx_path"]:
                    bp["primary_pptx_path"] = selected_pptx
                    bp["primary_deck_title"] = deck_options[selected_pptx]
                    st.session_state.blueprint = bp
                    st.rerun()
            with c_sel_thumb:
                tmpl_thumb = get_template_thumbnail(bp["primary_pptx_path"])
                if tmpl_thumb and os.path.exists(tmpl_thumb):
                    st.image(tmpl_thumb, caption="👁️ Tanlangan Shablon", use_container_width=True)
                else:
                    st.markdown("<div class='slide-preview-box' style='height:75px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-size:11px;'>📄 Native PPTX</div>", unsafe_allow_html=True)

        # Splicing Architecture Selector
        cur_engine = bp.get("engine_mode", st.session_state.get("engine_mode", "harmonized"))
        engine_idx = 0 if cur_engine == "harmonized" else (1 if cur_engine == "multi_master" else 2)
        
        c_eng, c_eng_desc = st.columns([2, 2])
        with c_eng:
            selected_engine = st.radio(
                "⚙️ Slayd Yig'ish Arxitekturasi (Splicing Engine):",
                options=["harmonized", "multi_master", "modular_stacker"],
                format_func=lambda x: {
                    "harmonized": "🎨 Master Theme Harmonizer (Yagona rang va shrift)",
                    "multi_master": "🎭 Multi-Master Deck (Asl dizayn va fonlarni saqlash)",
                    "modular_stacker": "🧩 Modular Component Stacker (Infografika inyeksiyasi)"
                }[x],
                index=engine_idx,
                key="splicing_engine_radio"
            )
            if selected_engine != bp.get("engine_mode"):
                bp["engine_mode"] = selected_engine
                st.session_state.engine_mode = selected_engine
        with c_eng_desc:
            engine_explanations = {
                "harmonized": "ℹ️ **Master Theme Harmonizer**: Barcha slaydlarning rang palitrasi va shriftlarini avtomatik ravishda titul slaydiga moslashtiradi (Yagona korporativ/akademik brend uslubi).",
                "multi_master": "ℹ️ **Multi-Master Deck**: Har bir slayd o'zining original foni, teksturalari, 3D grafikasi va master dizaynini to'liq saqlab qoladi (Maksimal vizual boylik).",
                "modular_stacker": "ℹ️ **Modular Component Stacker**: Asosiy slaydga boy diagramma va kartalarni to'g'ridan-to'g'ri inyeksiya qiladi."
            }
            st.info(engine_explanations[selected_engine])

        st.markdown("##### 🖼️ Slaydlar Arxitekturasi va Shakllari:")
        st.caption("Pastdagi har bir slayd shaklini o'zingiz xohlagan ilmiy shaklga (jadval, qiyoslash, vaqt xaritasi va h.k.) almashtirishingiz mumkin:")

        # Display Slide Thumbnails / Structure with 1-Click Archetype Customizer
        cols_per_row = 5
        slide_rows = [bp['slides'][i:i + cols_per_row] for i in range(0, len(bp['slides']), cols_per_row)]
        archetype_keys = list(SLIDE_ARCHETYPES.keys())

        for r_idx, row in enumerate(slide_rows):
            cols = st.columns(len(row))
            for c_idx, slide_item in enumerate(row):
                with cols[c_idx]:
                    s_num = slide_item["slide_number"]
                    l_name = slide_item["layout_name"]
                    t_hint = slide_item["title_hint"]
                    img_p = slide_item.get("image_path", "")
                    current_arch = slide_item.get("layout_type", "cards_grid")

                    st.markdown(f"**Slayd #{s_num:02d}**")
                    
                    if img_p and os.path.exists(img_p) and "slide_" in os.path.basename(img_p).lower():
                        st.image(img_p, use_container_width=True)
                    else:
                        t_pptx = slide_item.get("pptx_path") or bp.get("primary_pptx_path")
                        t_thumb = get_template_thumbnail(t_pptx) if t_pptx else None
                        if t_thumb and os.path.exists(t_thumb):
                            st.image(t_thumb, use_container_width=True)
                        else:
                            arch_icon = SLIDE_ARCHETYPES.get(current_arch, {}).get("icon", "📊")
                            st.markdown(f"""
                            <div class="slide-preview-box" style="height: 96px; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center;">
                                <div style="font-size: 24px; margin-bottom: 2px;">{arch_icon}</div>
                                <div style="font-size: 11px; font-weight: 700; color: #38bdf8;">{l_name}</div>
                                <div style="font-size: 9px; color: #64748b;">Haqiqiy 3D Qolip</div>
                            </div>
                            """, unsafe_allow_html=True)

                    # 1-Click Archetype Customizer
                    cur_arch_idx = archetype_keys.index(current_arch) if current_arch in archetype_keys else 0
                    bp_arch_ver = bp.get("archetype_version", 0)
                    new_arch = st.selectbox(
                        f"Slayd {s_num} shakli",
                        options=archetype_keys,
                        index=cur_arch_idx,
                        format_func=lambda k: f"{SLIDE_ARCHETYPES[k]['icon']} {SLIDE_ARCHETYPES[k]['name_uz']}",
                        key=f"slide_arch_pick_{s_num}_{bp_arch_ver}",
                        label_visibility="collapsed"
                    )

                    if new_arch != current_arch:
                        slide_item["layout_type"] = new_arch
                        slide_item["layout_name"] = SLIDE_ARCHETYPES[new_arch]["name_uz"]
                        slide_item["title_hint"] = matcher.generate_slide_title_hint(
                            s_num, new_arch, bp["topic"], bp.get("language", "uz"), bp.get("presentation_mode", "general_academic")
                        )
                        # Re-query slide matching new archetype
                        matches = catalog.search_slides_by_archetype(new_arch, category=bp.get("category"), topic=bp.get("topic"), limit=1)
                        if matches:
                            slide_item["matched_slide_ref"] = matches[0]
                            slide_item["pptx_path"] = matches[0]["pptx_path"]
                            slide_item["template_title"] = matches[0]["template_title"]
                            slide_item["template_slide_number"] = matches[0]["slide_index"] + 1
                            slide_item["image_path"] = matches[0].get("image_path", "")
                        # Refresh prompt
                        st.session_state.generated_prompt = generate_notebooklm_prompt(bp)
                        st.rerun()

                    st.caption(f"📌 {t_hint[:38]}..." if len(t_hint) > 38 else f"📌 {t_hint}")

        # -----------------------------------------------------------------
        # STEP 3: NOTEBOOKLM PROMPT & AI CONTENT STUDIO
        # -----------------------------------------------------------------
        st.markdown("---")
        st.markdown("""
        <div class="section-header">
            <span style="font-size: 1.4rem;">🧠</span>
            <div>
                <h3 class="section-title">3-Qadam: Taqdimot Matnini Olish (NotebookLM yoki Gemini AI)</h3>
                <div class="section-subtitle">NotebookLM manbalaringizni chuqur tahlil qilib, 'Bir slayd — bir fikr' va Action-Oriented ilmiy xulosalarni beradi</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        tab_c1, tab_c2, tab_c3 = st.tabs([
            "📋 1. NotebookLM Rejimi (Ixcham Prompt & Markdown)",
            "⚡ 2. Gemini AI orqali 1-Klikda Yaratish",
            "📝 3. Tayyor JSON / Matn Kiritish"
        ])

        with tab_c1:
            st.markdown("##### 📌 NotebookLM uchun Tayyorlangan Ixcham Prompt:")
            st.caption("Ushbu prompt NotebookLM cheklovlaridan oshmaydi (< 1,800 belgi). Blokning yuqori o'ng burchagidagi nusxalash tugmasini bosing:")
            
            prompt_to_show = st.session_state.get("compact_notebooklm_prompt") or st.session_state.generated_prompt
            st.code(prompt_to_show, language="markdown")

            st.markdown("##### 📥 NotebookLM qaytargan Markdown matnini shu yerga joylashtiring (Ctrl+V):")
            md_ver = st.session_state.get("md_input_version", 0)
            md_input = st.text_area(
                "NotebookLM Markdown javobi:",
                value=st.session_state.notebooklm_response,
                height=220,
                placeholder="### 1-slayd: Titul va Kirish\n* **Action-oriented sarlavha:** ...\n* **Mantiqiy maqsadi:** ...\n* **Tarkibiy tezislar:**\n  * ...",
                key=f"notebooklm_md_input_{md_ver}"
            )
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                btn_parse_md = st.button("🔍 1. Matnni Tahlil Qilish & Slaydlarga Bog'lash", type="secondary", use_container_width=True)
            with col_b2:
                btn_parse_and_build = st.button("🚀 2. Tahlil Qilish va Darhol PPTX Yaratish", type="primary", use_container_width=True)

            if (btn_parse_md or btn_parse_and_build) and md_input.strip():
                st.session_state.notebooklm_response = md_input.strip()
                with st.spinner("⏳ Markdown matn tahlil qilinmoqda va vizual arxetiplar aniqlanmoqda..."):
                    try:
                        parsed = NotebookLMMarkdownParser.parse(
                            md_input.strip(),
                            topic=bp.get("topic", ""),
                            default_metadata=bp.get("presenter_info", {})
                        )
                        st.session_state.parsed_content = parsed
                        
                        # Dynamically update blueprint slides with detected archetypes
                        slides_p = parsed.get("slides", [])
                        for idx, s_p in enumerate(slides_p):
                            arch = s_p.get("layout_type", "cards_grid")
                            if idx < len(bp["slides"]):
                                s_num = bp["slides"][idx]["slide_number"]
                                bp["slides"][idx]["layout_type"] = arch
                                bp["slides"][idx]["layout_name"] = SLIDE_ARCHETYPES.get(arch, {}).get("name_uz", arch)
                                bp["slides"][idx]["title_hint"] = s_p.get("title", bp["slides"][idx].get("title_hint", ""))

                                matches = catalog.search_slides_by_archetype(arch, category=bp.get("category"), topic=bp.get("topic"), limit=1)
                                if matches:
                                    bp["slides"][idx]["matched_slide_ref"] = matches[0]
                                    bp["slides"][idx]["pptx_path"] = matches[0]["pptx_path"]
                                    bp["slides"][idx]["template_title"] = matches[0]["template_title"]
                                    bp["slides"][idx]["template_slide_number"] = matches[0]["slide_index"] + 1
                                    bp["slides"][idx]["image_path"] = matches[0].get("image_path", "")
                            else:
                                s_num = idx + 1
                                new_slide_item = {
                                    "slide_number": s_num,
                                    "layout_type": arch,
                                    "layout_name": SLIDE_ARCHETYPES.get(arch, {}).get("name_uz", arch),
                                    "title_hint": s_p.get("title", f"{s_num}-Slayd")
                                }
                                matches = catalog.search_slides_by_archetype(arch, category=bp.get("category"), topic=bp.get("topic"), limit=1)
                                if matches:
                                    new_slide_item["matched_slide_ref"] = matches[0]
                                    new_slide_item["pptx_path"] = matches[0]["pptx_path"]
                                    new_slide_item["template_title"] = matches[0]["template_title"]
                                    new_slide_item["template_slide_number"] = matches[0]["slide_index"] + 1
                                    new_slide_item["image_path"] = matches[0].get("image_path", "")
                                bp["slides"].append(new_slide_item)
                        
                        bp["archetype_version"] = bp.get("archetype_version", 0) + 1
                        st.session_state.blueprint = bp

                        if btn_parse_and_build:
                            chosen_engine = st.session_state.get("engine_mode", bp.get("engine_mode", "harmonized"))
                            out_pptx = builder.create_presentation(parsed, st.session_state.blueprint, engine_mode=chosen_engine)
                            st.session_state.created_pptx_path = out_pptx
                            st.toast("🎉 PPTX Taqdimot muvaffaqiyatli yaratildi!", icon="✅")
                        else:
                            st.toast(f"✅ {len(slides_p)} ta slayd tahlil qilindi!", icon="🎯")

                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Xatolik: {e}")

        with tab_c2:
            st.markdown("##### ⚡ Gemini AI orqali bir zumda taqdimot kontentini yaratish:")
            st.caption("NotebookLM ga o'tmasdan, o'rnatilgan Gemini AI orqali to'liq 10 betlik Action-Oriented akademik taqdimot matnini generatsiya qilish.")
            
            btn_gemini_direct = st.button("🤖 1-Klikda To'liq Taqdimot Matnini Generatsiya Qilish", type="primary", use_container_width=True)
            if btn_gemini_direct:
                with st.spinner("⏳ Gemini AI mavzuni va ilmiy jihatlarni chuqur tahlil qilmoqda..."):
                    try:
                        p_info = bp.get("presenter_info", {})
                        author_str = f"{p_info.get('author_name', '')} ({p_info.get('organization', '')})"
                        sources_note = sources_notes_input.strip() if 'sources_notes_input' in locals() and sources_notes_input else ""
                        gemini_text = NotebookLMPromptGenerator.generate_with_gemini_direct(
                            topic=bp.get("topic", ""),
                            domain=bp.get("category_name", "Akademik / Ilmiy"),
                            author_info=author_str,
                            sources_summary=sources_note,
                            lang=bp.get("language", "uz")
                        )
                        st.session_state.notebooklm_response = gemini_text
                        st.session_state.md_input_version = st.session_state.get("md_input_version", 0) + 1
                        
                        parsed = NotebookLMMarkdownParser.parse(
                            gemini_text,
                            topic=bp.get("topic", ""),
                            default_metadata=p_info
                        )
                        st.session_state.parsed_content = parsed
                        
                        # Dynamically update blueprint slides
                        slides_p = parsed.get("slides", [])
                        for idx, s_p in enumerate(slides_p):
                            arch = s_p.get("layout_type", "cards_grid")
                            if idx < len(bp["slides"]):
                                s_num = bp["slides"][idx]["slide_number"]
                                bp["slides"][idx]["layout_type"] = arch
                                bp["slides"][idx]["layout_name"] = SLIDE_ARCHETYPES.get(arch, {}).get("name_uz", arch)
                                bp["slides"][idx]["title_hint"] = s_p.get("title", bp["slides"][idx].get("title_hint", ""))

                                matches = catalog.search_slides_by_archetype(arch, category=bp.get("category"), topic=bp.get("topic"), limit=1)
                                if matches:
                                    bp["slides"][idx]["matched_slide_ref"] = matches[0]
                                    bp["slides"][idx]["pptx_path"] = matches[0]["pptx_path"]
                                    bp["slides"][idx]["template_title"] = matches[0]["template_title"]
                                    bp["slides"][idx]["template_slide_number"] = matches[0]["slide_index"] + 1
                                    bp["slides"][idx]["image_path"] = matches[0].get("image_path", "")
                            else:
                                s_num = idx + 1
                                new_slide_item = {
                                    "slide_number": s_num,
                                    "layout_type": arch,
                                    "layout_name": SLIDE_ARCHETYPES.get(arch, {}).get("name_uz", arch),
                                    "title_hint": s_p.get("title", f"{s_num}-Slayd")
                                }
                                matches = catalog.search_slides_by_archetype(arch, category=bp.get("category"), topic=bp.get("topic"), limit=1)
                                if matches:
                                    new_slide_item["matched_slide_ref"] = matches[0]
                                    new_slide_item["pptx_path"] = matches[0]["pptx_path"]
                                    new_slide_item["template_title"] = matches[0]["template_title"]
                                    new_slide_item["template_slide_number"] = matches[0]["slide_index"] + 1
                                    new_slide_item["image_path"] = matches[0].get("image_path", "")
                                bp["slides"].append(new_slide_item)
                        
                        bp["archetype_version"] = bp.get("archetype_version", 0) + 1
                        st.session_state.blueprint = bp
                        st.success("✅ Ilmiy taqdimot matni Gemini orqali yaratildi va shablonlarga bog'landi!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Gemini generatsiyasida xatolik: {e}")

        with tab_c3:
            st.markdown("##### 📝 Tayyor JSON yoki strukturalangan matn kiriting:")
            json_input = st.text_area(
                "JSON formatidagi matn:",
                value=st.session_state.notebooklm_response if st.session_state.notebooklm_response.startswith("{") else "",
                height=180,
                placeholder='{"topic": "...", "slides": [...]}'
            )
            if st.button("Formatni Yangilash", key="btn_update_raw_json"):
                if json_input.strip():
                    st.session_state.notebooklm_response = json_input.strip()
                    try:
                        st.session_state.parsed_content = executor.parse_notebooklm_response(json_input.strip())
                        st.success("✅ JSON muvaffaqiyatli qabul qilindi!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Xato: {e}")

        # -----------------------------------------------------------------
        # STEP 4: FINAL GENERATION & DOWNLOAD
        # -----------------------------------------------------------------
        st.markdown("---")
        st.markdown("""
        <div class="section-header">
            <span style="font-size: 1.4rem;">🚀</span>
            <div>
                <h3 class="section-title">4-Qadam: Native PPTX Taqdimotni Generatsiya Qilish</h3>
                <div class="section-subtitle">Tanlangan shablon slaydlariga matnlarni joylashtirish va yakuniy taqdimot faylini yaratish</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        c_r1, c_r2 = st.columns([3, 1])
        with c_r1:
            chosen_engine = st.session_state.get("engine_mode", bp.get("engine_mode", "harmonized"))
            engine_title = {
                "harmonized": "🎨 Master Theme Harmonizer (Titul ranglari va shriftlari bilan uyg'unlashtirilgan)",
                "multi_master": "🎭 Multi-Master Deck (Asl shablon fonlari va dizaynlari bilan)",
                "modular_stacker": "🧩 Modular Component Stacker (Infografika inyeksiyasi bilan)"
            }.get(chosen_engine, chosen_engine)
            st.info(f"Faol yig'ish arxitekturasi: **{engine_title}**")
        with c_r2:
            st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
            btn_render = st.button("🚀 PPTX Taqdimotni Yaratish", use_container_width=True, type="primary")

        if btn_render:
            raw_data_to_use = (
                (md_input.strip() if 'md_input' in locals() and md_input else "") or
                st.session_state.get("notebooklm_response", "").strip()
            )

            # Auto-generate content on the fly if user hasn't provided it yet
            if not raw_data_to_use and not st.session_state.get("parsed_content"):
                with st.spinner("🤖 Gemini AI mavzu bo'yicha ilmiy taqdimot kontentini tayyorlamoqda..."):
                    try:
                        p_info = bp.get("presenter_info", {})
                        author_str = f"{p_info.get('author_name', '')} ({p_info.get('organization', '')})"
                        sources_note = sources_notes_input.strip() if 'sources_notes_input' in locals() and sources_notes_input else ""
                        gemini_text = NotebookLMPromptGenerator.generate_with_gemini_direct(
                            topic=bp.get("topic", ""),
                            domain=bp.get("category_name", "Akademik / Ilmiy"),
                            author_info=author_str,
                            sources_summary=sources_note,
                            lang=bp.get("language", "uz")
                        )
                        st.session_state.notebooklm_response = gemini_text
                        st.session_state.parsed_content = NotebookLMMarkdownParser.parse(
                            gemini_text,
                            topic=bp.get("topic", ""),
                            default_metadata=p_info
                        )
                        raw_data_to_use = gemini_text
                    except Exception as e:
                        st.error(f"❌ Kontentni avtomatik generatsiya qilishda xatolik: {e}")

            if raw_data_to_use or st.session_state.get("parsed_content"):
                with st.spinner("⏳ 5,170+ haqiqiy PPTX shablonlaridan tanlangan slaydlar yig'ilmoqda va matnlar to'ldirilmoqda..."):
                    try:
                        if st.session_state.get("parsed_content"):
                            parsed_content = st.session_state.parsed_content
                        else:
                            parsed_content = executor.parse_notebooklm_response(raw_data_to_use)
                            st.session_state.parsed_content = parsed_content
                            st.session_state.notebooklm_response = raw_data_to_use

                        # Update blueprint slide archetypes to match parsed content
                        slides_p = parsed_content.get("slides", [])
                        for idx, s_p in enumerate(slides_p):
                            arch = s_p.get("layout_type", "cards_grid")
                            if idx < len(bp["slides"]):
                                s_num = bp["slides"][idx]["slide_number"]
                                bp["slides"][idx]["layout_type"] = arch
                                bp["slides"][idx]["layout_name"] = SLIDE_ARCHETYPES.get(arch, {}).get("name_uz", arch)
                                bp["slides"][idx]["title_hint"] = s_p.get("title", bp["slides"][idx].get("title_hint", ""))
                                matches = catalog.search_slides_by_archetype(arch, category=bp.get("category"), topic=bp.get("topic"), limit=1)
                                if matches:
                                    bp["slides"][idx]["matched_slide_ref"] = matches[0]
                                    bp["slides"][idx]["pptx_path"] = matches[0]["pptx_path"]
                                    bp["slides"][idx]["template_title"] = matches[0]["template_title"]
                                    bp["slides"][idx]["template_slide_number"] = matches[0]["slide_index"] + 1
                                    bp["slides"][idx]["image_path"] = matches[0].get("image_path", "")
                            else:
                                s_num = idx + 1
                                new_slide_item = {
                                    "slide_number": s_num,
                                    "layout_type": arch,
                                    "layout_name": SLIDE_ARCHETYPES.get(arch, {}).get("name_uz", arch),
                                    "title_hint": s_p.get("title", f"{s_num}-Slayd")
                                }
                                matches = catalog.search_slides_by_archetype(arch, category=bp.get("category"), topic=bp.get("topic"), limit=1)
                                if matches:
                                    new_slide_item["matched_slide_ref"] = matches[0]
                                    new_slide_item["pptx_path"] = matches[0]["pptx_path"]
                                    new_slide_item["template_title"] = matches[0]["template_title"]
                                    new_slide_item["template_slide_number"] = matches[0]["slide_index"] + 1
                                    new_slide_item["image_path"] = matches[0].get("image_path", "")
                                bp["slides"].append(new_slide_item)

                        bp["archetype_version"] = bp.get("archetype_version", 0) + 1
                        st.session_state.blueprint = bp

                        chosen_engine = st.session_state.get("engine_mode", bp.get("engine_mode", "harmonized"))
                        out_pptx = builder.create_presentation(parsed_content, st.session_state.blueprint, engine_mode=chosen_engine)
                        st.session_state.created_pptx_path = out_pptx
                        st.balloons()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Generatsiyada xatolik: {e}")

        # -----------------------------------------------------------------
        # STEP 5: DOWNLOAD & OPEN IN POWERPOINT
        # -----------------------------------------------------------------
        if st.session_state.get("created_pptx_path") and os.path.exists(st.session_state.created_pptx_path):
            pptx_file = st.session_state.created_pptx_path
            file_name = os.path.basename(pptx_file)
            file_size_kb = os.path.getsize(pptx_file) / 1024
            
            st.markdown(f"""
            <div class="success-result-box" style="margin-top: 24px;">
                <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 8px;">
                    <span style="font-size: 28px;">🎉</span>
                    <div>
                        <div style="font-size: 1.35rem; font-weight: 800; color: #10b981;">Taqdimot Muvaffaqiyatli Tayyorlandi!</div>
                        <div style="font-size: 0.88rem; color: #cbd5e1;">Fayl: <code>{file_name}</code> ({file_size_kb:.1f} KB)</div>
                    </div>
                </div>
                <div style="font-size: 0.85rem; color: #94a3b8;">
                    ✅ Taqdimotda asl shablonning barcha 3D elementlari, vektor kartalari va fonlari to'liq saqlanib, yangi ilmiy kontent bilan to'ldirildi.
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            col_d1, col_d2, col_d3 = st.columns([1.2, 1.2, 1.0])
            with col_d1:
                with open(pptx_file, "rb") as f:
                    pptx_data = f.read()
                st.download_button(
                    label="📥 PowerPoint (.pptx) Yuklab Olish",
                    data=pptx_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    type="primary"
                )
            with col_d2:
                pdf_file = os.path.splitext(pptx_file)[0] + ".pdf"
                if not os.path.exists(pdf_file):
                    from core.slide_preview import export_pptx_to_pdf
                    export_pptx_to_pdf(pptx_file, pdf_file)

                if os.path.exists(pdf_file):
                    with open(pdf_file, "rb") as f_pdf:
                        pdf_data = f_pdf.read()
                    st.download_button(
                        label="📄 PDF Hujjatini (.pdf) Yuklab Olish",
                        data=pdf_data,
                        file_name=os.path.basename(pdf_file),
                        mime="application/pdf",
                        use_container_width=True
                    )
                else:
                    st.button("📄 PDF Tayyorlanmoqda...", disabled=True, use_container_width=True)

            with col_d3:
                if st.button("📂 PowerPointda Ochish", use_container_width=True):
                    try:
                        os.startfile(pptx_file)
                        st.toast("PowerPoint dasturida ochildi!", icon="✅")
                    except Exception as e:
                        st.error(f"Ochishda xatolik: {e}")

            # Render full visual slide deck preview gallery
            render_presentation_preview_ui(pptx_file, key_suffix="step5")



# =========================================================================
# TAB 2: TEMPLATE DATABASE EXPLORER & 1-CLICK GENERATOR
# =========================================================================
with tab_catalog:
    st.markdown("""
    <div class="section-header">
        <span style="font-size: 1.4rem;">📚</span>
        <div>
            <h3 class="section-title">5,170+ Haqiqiy PPTX Shablonlar Katalogi</h3>
            <div class="section-subtitle">Professional shablonlarni ko'rish, tanlash va to'g'ridan-to'g'ri yangi mavzu bo'yicha taqdimot generatsiya qilish</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.session_state.get("created_pptx_path") and os.path.exists(st.session_state.created_pptx_path):
        cat_pptx = st.session_state.created_pptx_path
        st.markdown(f"""
        <div class="success-result-box" style="margin-bottom: 24px;">
            <div style="font-size: 1.25rem; font-weight: 700; color: #10b981; margin-bottom: 4px;">🎉 Taqdimot Tayyor!</div>
            <div style="font-size: 0.9rem; color: #cbd5e1;">Fayl: <code>{os.path.basename(cat_pptx)}</code> ({os.path.getsize(cat_pptx)/1024:.1f} KB)</div>
        </div>
        """, unsafe_allow_html=True)
        c_cat_d1, c_cat_d2 = st.columns(2)
        with c_cat_d1:
            with open(cat_pptx, "rb") as f_cat_down:
                cat_data = f_cat_down.read()
            st.download_button(
                label="📥 Tayyor PPTX Faylini Yuklab Olish",
                data=cat_data,
                file_name=os.path.basename(cat_pptx),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary",
                key="down_cat_pptx"
            )
        with c_cat_d2:
            if st.button("📂 PowerPoint Dasturida Ochish", key="open_cat_pptx", use_container_width=True):
                try:
                    os.startfile(cat_pptx)
                    st.toast("PowerPoint dasturida ochildi!", icon="✅")
                except Exception as e:
                    st.error(f"Ochishda xatolik: {e}")

        render_presentation_preview_ui(cat_pptx, key_suffix="tab2")
        st.markdown("---")

    col_t2_1, col_t2_2, col_t2_3 = st.columns([2.5, 1.5, 1])
    with col_t2_1:
        cat_topic_input = st.text_input(
            "📝 Taqdimot Mavzusi (Shablonni to'ldirish uchun):",
            value=st.session_state.get("last_topic", "Kvant kompyuterlari va asimmetrik kriptografiya"),
            key="cat_topic_input"
        )
    with col_t2_2:
        cat_filter = st.selectbox(
            "🏛️ Soha (Kategoriya):",
            ["all"] + list(ACADEMIC_CATEGORIES.keys()),
            format_func=lambda x: "Barcha Kategoriyalar" if x == "all" else ACADEMIC_CATEGORIES[x]["name_uz"],
            key="cat_filter_box"
        )
    with col_t2_3:
        cat_slide_count = st.selectbox("📊 Slaydlar:", [5, 8, 10, 12, 15], index=2, key="cat_slide_count")

    c_s1, c_s2 = st.columns([3, 1])
    with c_s1:
        search_query = st.text_input("🔍 Kalit so'z bo'yicha qidiruv (O'zbek, Rus, Ingliz):", placeholder="Masalan: Quantum, Medical, Physics, Timeline, Kiberxavfsizlik, Business...", key="cat_search_q")
    with c_s2:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        st.caption(f"📚 Jami mavjud shablonlar: **{len(catalog.templates):,} ta**")

    target_cat = None if cat_filter == "all" else cat_filter
    search_results = catalog.search_templates(query=search_query, category=target_cat, limit=28)

    st.markdown(f"**Topilgan professional PPTX shablonlar:** `{len(search_results)} ta`")

    if search_results:
        t_cols = 4
        chunks = [search_results[i:i + t_cols] for i in range(0, len(search_results), t_cols)]
        
        for chunk in chunks:
            cols = st.columns(len(chunk))
            for idx, tmpl in enumerate(chunk):
                with cols[idx]:
                    st.markdown(f"**{tmpl['title'][:28]}**")
                    st.caption(f"📁 `{os.path.basename(tmpl['pptx_path'])[:26]}`")
                    
                    cov_img = tmpl.get("cover_image", "")
                    if not cov_img or not os.path.exists(cov_img):
                        cov_img = get_template_thumbnail(tmpl.get("pptx_path", ""))

                    if cov_img and os.path.exists(cov_img):
                        st.image(cov_img, use_container_width=True)
                    else:
                        st.markdown("<div class='slide-preview-box' style='height:110px; display:flex; align-items:center; justify-content:center; color:#94a3b8; font-size:12px; text-align:center;'>📄 Native PPTX Template</div>", unsafe_allow_html=True)
                    
                    btn_gen_this = st.button("🚀 Slayd Yaratish", key=f"btn_gen_{tmpl['id']}", use_container_width=True, type="primary")
                    if btn_gen_this:
                        st.session_state.last_topic = cat_topic_input.strip()
                        with st.spinner(f"⏳ '{tmpl['title']}' shablonida taqdimot generatsiya qilinmoqda..."):
                            try:
                                bp = matcher.build_blueprint(
                                    topic=cat_topic_input.strip(),
                                    slide_count=cat_slide_count,
                                    language=selected_lang,
                                    selected_category=target_cat
                                )
                                bp["primary_pptx_path"] = tmpl["pptx_path"]
                                bp["primary_deck_title"] = tmpl["title"]
                                prompt = generate_notebooklm_prompt(bp)
                                
                                content_json = asyncio.run(executor.generate_content_auto(prompt))
                                st.session_state.notebooklm_response = json.dumps(content_json, ensure_ascii=False, indent=2)
                                
                                clean_t = "".join(c for c in cat_topic_input if c.isalnum() or c in (" ", "_", "-")).strip()[:40]
                                out_p = os.path.join(str(config.OUTPUT_STUDIO_DIR), f"{clean_t}_Customized.pptx")
                                replace_presentation_content(tmpl["pptx_path"], content_json, out_p, blueprint=bp)
                                
                                st.session_state.created_pptx_path = out_p
                                st.session_state.blueprint = bp
                                st.session_state.generated_prompt = prompt
                                st.balloons()
                                st.success(f"🎉 '{tmpl['title']}' shablonida taqdimot muvaffaqiyatli yaratildi!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Xatolik: {e}")

                    if os.path.exists(tmpl["pptx_path"]):
                        with open(tmpl["pptx_path"], "rb") as f_tmpl:
                            tmpl_data = f_tmpl.read()
                        st.download_button(
                            label="📥 Shablon (.pptx)",
                            data=tmpl_data,
                            file_name=os.path.basename(tmpl["pptx_path"]),
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            key=f"dl_{tmpl['id']}",
                            use_container_width=True
                        )
    else:
        st.info("Kiritilgan mezon bo'yicha hech qanday shablon topilmadi.")


# =========================================================================
# TAB 3: CUSTOM PPTX UPLOADER & OCR / IMAGE RECREATOR
# =========================================================================
with tab_converter:
    st.markdown("""
    <div class="section-header">
        <span style="font-size: 1.4rem;">🖼️</span>
        <div>
            <h3 class="section-title">Maxsus PPTX Shablon Yuklash va Rasm/OCR Konvertori</h3>
            <div class="section-subtitle">O'zingizdagi PowerPoint faylini yuklab unga yangi matnlarni joylashtiring yoki slayd rasmlarini vektor PPTX ga aylantiring</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    conv_mode = st.radio(
        "🛠️ Ish rejimi:",
        [
            "📁 O'z PPTX Shablonimni Yuklab Slayd Yaratish (Native PPTX Replacer)",
            "🖼️ Slayd Rasmlari / Papkadan Vektor Taqdimot Yaratish (OCR & Recreator)"
        ],
        index=0
    )

    if "O'z PPTX Shablonimni Yuklab" in conv_mode:
        st.markdown("#### 📁 O'z PPTX Faylingizni Yuklang va Slayd Generatsiya Qiling:")
        c_up1, c_up2 = st.columns([2, 1])
        with c_up1:
            uploaded_pptx = st.file_uploader(
                "Kompyuteringizdan PPTX faylni tanlang:",
                type=["pptx"],
                key="custom_pptx_uploader"
            )
            custom_topic = st.text_input(
                "📝 Yangi Taqdimot Mavzusi:",
                value=st.session_state.get("last_topic", "Kvant kompyuterlari va asimmetrik kriptografiya"),
                key="custom_topic_input"
            )
        with c_up2:
            custom_slide_cnt = st.selectbox("📊 Slaydlar soni:", [5, 8, 10, 12, 15], index=2, key="custom_slide_cnt")
            custom_lang = st.selectbox("🌐 Til:", ["uz", "ru", "en"], index=0, key="custom_lang", format_func=lambda x: {"uz": "🇺🇿 O'zbekcha", "ru": "🇷🇺 Русский", "en": "🇬🇧 English"}[x])

        if uploaded_pptx is not None:
            temp_dir = config.OUTPUT_ROOT / "uploaded_templates"
            temp_dir.mkdir(parents=True, exist_ok=True)
            saved_uploaded_path = str(temp_dir / uploaded_pptx.name)
            with open(saved_uploaded_path, "wb") as f_up:
                f_up.write(uploaded_pptx.getbuffer())

            c_up_stat1, c_up_stat2 = st.columns([3, 1])
            with c_up_stat1:
                st.success(f"✅ Shablon yuklandi: `{uploaded_pptx.name}` ({os.path.getsize(saved_uploaded_path):,} bytes)")
            with c_up_stat2:
                up_thumb = get_template_thumbnail(saved_uploaded_path)
                if up_thumb and os.path.exists(up_thumb):
                    st.image(up_thumb, caption="👁️ Shablon Ko'rinishi", use_container_width=True)

            if st.button("🚀 Yuklangan Shablon Bilan Slayd Yaratish", type="primary", use_container_width=True):
                with st.spinner("⏳ Yuklangan shablon tahlil qilinmoqda va yangi kontent joylashtirilmoqda..."):
                    try:
                        bp = matcher.build_blueprint(
                            topic=custom_topic.strip(),
                            slide_count=custom_slide_cnt,
                            language=custom_lang
                        )
                        bp["primary_pptx_path"] = saved_uploaded_path
                        bp["primary_deck_title"] = os.path.splitext(uploaded_pptx.name)[0]
                        prompt = generate_notebooklm_prompt(bp)
                        
                        content_json = asyncio.run(executor.generate_content_auto(prompt))
                        clean_topic = "".join(c for c in custom_topic if c.isalnum() or c in (" ", "_", "-")).strip()[:40]
                        out_custom_p = os.path.join(str(config.OUTPUT_STUDIO_DIR), f"{clean_topic}_Customized.pptx")
                        
                        replace_presentation_content(saved_uploaded_path, content_json, out_custom_p, blueprint=bp)
                        
                        st.session_state.created_pptx_path = out_custom_p
                        st.session_state.blueprint = bp
                        st.session_state.notebooklm_response = json.dumps(content_json, ensure_ascii=False, indent=2)
                        st.session_state.generated_prompt = prompt
                        st.balloons()
                        st.success(f"🎉 Taqdimot muvaffaqiyatli yaratildi: `{os.path.basename(out_custom_p)}`!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Xatolik yuz berdi: {e}")

            if st.session_state.get("created_pptx_path") and os.path.exists(st.session_state.created_pptx_path):
                c_pptx = st.session_state.created_pptx_path
                st.markdown(f"""
                <div class="success-result-box" style="margin-top: 20px;">
                    <div style="font-size: 1.2rem; font-weight: 700; color: #10b981;">🎉 Taqdimot Tayyor!</div>
                    <div style="font-size: 0.85rem; color: #cbd5e1;">Fayl: <code>{os.path.basename(c_pptx)}</code> ({os.path.getsize(c_pptx)/1024:.1f} KB)</div>
                </div>
                """, unsafe_allow_html=True)
                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    with open(c_pptx, "rb") as f_down:
                        c_data = f_down.read()
                    st.download_button(
                        label="📥 PPTX Faylini Yuklab Olish",
                        data=c_data,
                        file_name=os.path.basename(c_pptx),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                        type="primary",
                        key="down_custom_pptx"
                    )
                with c_d2:
                    if st.button("📂 PowerPoint Dasturida Ochish", key="open_custom_pptx", use_container_width=True):
                        try:
                            os.startfile(c_pptx)
                            st.toast("PowerPoint dasturida ochildi!", icon="✅")
                        except Exception as e:
                            st.error(f"Ochishda xatolik: {e}")

                render_presentation_preview_ui(c_pptx, key_suffix="tab3")

    else:
        st.markdown("#### 🖼️ Slayd Rasmlari yoki Papkadan Vektor Taqdimot Yaratish:")
        
        c_img_up1, c_img_up2 = st.columns([2, 1])
        with c_img_up1:
            uploaded_images = st.file_uploader(
                "Slayd rasmlarini yuklang (PNG, JPG):",
                type=["png", "jpg", "jpeg"],
                accept_multiple_files=True,
                key="img_batch_uploader"
            )
        with c_img_up2:
            preset_folders = slide_manager.get_cached_folders()
            preset_keys = list(preset_folders.keys())
            selected_conv_key = st.selectbox("Yoki tayyor skanerlangan papkani tanlang:", preset_keys if preset_keys else ["Mavjud emas"], index=0)

        if uploaded_images:
            st.info(f"📸 {len(uploaded_images)} ta rasm yuklandi.")
            temp_img_dir = str(config.OUTPUT_ROOT / "uploaded_images_session")
            os.makedirs(temp_img_dir, exist_ok=True)
            for img_file in uploaded_images:
                with open(os.path.join(temp_img_dir, img_file.name), "wb") as f_img:
                    f_img.write(img_file.getbuffer())

            if st.button("⚡ Yuklangan Rasmlardan Vektor Taqdimot Yaratish", type="primary", use_container_width=True):
                with st.spinner("⏳ Rasmlar tahlil qilinib OCR va vektor shakllarga o'girilmoqda..."):
                    ok, msg = slide_manager.convert_presentation(temp_img_dir)
                    if ok:
                        out_p = slide_manager.get_folder_pptx_path(temp_img_dir)
                        if out_p and os.path.exists(out_p):
                            st.session_state.created_pptx_path = out_p
                        st.balloons()
                        st.success("✅ Vektor taqdimot muvaffaqiyatli yaratildi!")
                        st.rerun()
                    else:
                        st.error(f"Xatolik: {msg}")

        elif selected_conv_key and selected_conv_key in preset_folders:
            target_conv_folder = preset_folders[selected_conv_key]
            if os.path.exists(target_conv_folder):
                orig_slides = slide_manager.get_folder_slides(target_conv_folder)
                p_path = slide_manager.get_folder_pptx_path(target_conv_folder)
                has_p = (p_path is not None and os.path.exists(p_path))

                c1, c2 = st.columns([3, 1])
                with c1:
                    st.markdown(f"**Papka:** `{target_conv_folder}` | **Slaydlar:** `{len(orig_slides)} ta`")
                with c2:
                    if has_p:
                        st.success("✅ PPTX Mavjud")
                    else:
                        st.warning("⏳ PPTX Yaratilmagan")

                if st.button("🚀 Ushbu Papkadan Vektor Taqdimot Yaratish", type="primary", use_container_width=True):
                    with st.spinner("⏳ OCR va vektor konvertatsiya bajarilmoqda..."):
                        ok, msg = slide_manager.convert_presentation(target_conv_folder)
                        if ok:
                            out_p = slide_manager.get_folder_pptx_path(target_conv_folder)
                            if out_p and os.path.exists(out_p):
                                st.session_state.created_pptx_path = out_p
                            st.balloons()
                            st.success("✅ Muvaffaqiyatli yakunlandi!")
                            st.rerun()
                        else:
                            st.error(f"Xatolik: {msg}")


# =========================================================================
# TAB: APPFLOWY & PROJECT MANAGEMENT (KANBAN)
# =========================================================================
with tab_appflowy:
    st.markdown("""
    <div class="hero-banner">
        <div class="hero-badge">📋 AppFlowy & Project Management</div>
        <div class="hero-title">Taqdimot Rejalashtirish va Kanban Doskasi</div>
        <div class="hero-subtitle">
            Taqdimotlaringizning mantiqiy tuzilishini (Action-Oriented arxiteplar) oldindan rejalashtiring,
            AppFlowy va Markdown formatlarida eksport qiling.
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_ap1, col_ap2 = st.columns([3, 1])
    with col_ap1:
        ap_topic = st.text_input("📝 Rejalashtiriladigan Taqdimot Mavzusi:", value="Antigravity Texnologiyalari va Zamonaviy Fizika", key="ap_topic_input")
    with col_ap2:
        ap_slides = st.selectbox("📊 Slaydlar Soni:", [6, 8, 10, 12, 15], index=1, key="ap_slides_count")

    if st.button("🚀 AppFlowy Rejasini Shakllantirish", type="primary", use_container_width=True):
        plan = AppFlowyContentPlanner.generate_plan(ap_topic.strip(), slide_count=ap_slides)
        st.session_state["active_appflowy_plan"] = plan
        board_file = AppFlowyContentPlanner.export_to_appflowy_markdown(ap_topic.strip(), plan)
        st.session_state["active_appflowy_file"] = str(board_file)
        st.success(f"✅ AppFlowy Rejasi yaratildi: `{board_file.name}`")

    if st.session_state.get("active_appflowy_plan"):
        cur_plan = st.session_state["active_appflowy_plan"]
        st.markdown(f"#### 📌 Slaydlar Arxitekturasi ({len(cur_plan)} ta slayd):")

        # Display as cards in columns
        for row_start in range(0, len(cur_plan), 3):
            cols = st.columns(3)
            for i in range(3):
                idx = row_start + i
                if idx < len(cur_plan):
                    sl = cur_plan[idx]
                    with cols[i]:
                        st.markdown(f"""
                        <div style="background: rgba(30, 41, 59, 0.75); border: 1px solid rgba(59, 130, 246, 0.3); border-radius: 12px; padding: 16px; margin-bottom: 12px; min-height: 220px;">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                                <span style="font-weight: 700; color: #38bdf8;">#{sl.slide_number} Slayd</span>
                                <span style="background: #2563eb; color: #fff; font-size: 0.72rem; padding: 2px 8px; border-radius: 6px; font-weight: 600;">{sl.archetype.upper()}</span>
                            </div>
                            <div style="font-size: 1rem; font-weight: 700; color: #f8fafc; margin-bottom: 4px;">{sl.title}</div>
                            <div style="font-size: 0.8rem; color: #94a3b8; margin-bottom: 8px;">{sl.subtitle}</div>
                            <div style="font-size: 0.78rem; color: #cbd5e1; border-top: 1px solid rgba(255,255,255,0.08); padding-top: 6px;">
                                <b>🎨 Visual:</b> <i>{sl.visual_cue[:60]}...</i>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)

        if st.session_state.get("active_appflowy_file") and os.path.exists(st.session_state["active_appflowy_file"]):
            with open(st.session_state["active_appflowy_file"], "r", encoding="utf-8") as f_md:
                md_bytes = f_md.read().encode("utf-8")
            st.download_button(
                "📥 AppFlowy Markdown Doskasini Yuklab Olish (.md)",
                data=md_bytes,
                file_name=os.path.basename(st.session_state["active_appflowy_file"]),
                mime="text/markdown",
                use_container_width=True
            )


# =========================================================================
# TAB: PENPOT & FOOOCUS AI ASSET STUDIO
# =========================================================================
with tab_penpot:
    st.markdown("""
    <div class="hero-banner">
        <div class="hero-badge">🎨 Penpot Design System & Fooocus AI</div>
        <div class="hero-title">Dizayn Tizimi va Lokal AI Rasm Studiyasi</div>
        <div class="hero-subtitle">
            Penpot ochiq kodli dizayn tizimining professional ranglar palitrasi va
            Fooocus AI yordamida taqdimotlar uchun 16:9 premium illyustratsiyalar yaratish markazi.
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("#### 🎨 1. Penpot Dizayn Tizimi Palitralari")
    p_cols = st.columns(4)
    themes = PenpotThemeManager.list_themes()
    for i, t_key in enumerate(themes):
        t_data = PenpotThemeManager.get_theme(t_key)
        with p_cols[i]:
            st.markdown(f"""
            <div style="background: {t_data['card_bg']}; border: 2px solid {t_data['primary']}; border-radius: 12px; padding: 16px; color: {t_data['text_title']};">
                <div style="font-weight: 700; font-size: 1rem; margin-bottom: 6px;">{t_data['name']}</div>
                <div style="display: flex; gap: 6px; margin: 10px 0;">
                    <div style="width: 24px; height: 24px; border-radius: 6px; background: {t_data['primary']};" title="Primary"></div>
                    <div style="width: 24px; height: 24px; border-radius: 6px; background: {t_data['secondary']};" title="Secondary"></div>
                    <div style="width: 24px; height: 24px; border-radius: 6px; background: {t_data['accent']};" title="Accent"></div>
                </div>
                <div style="font-size: 0.75rem; opacity: 0.8;">Shrift: {t_data['font_title']}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 🖼️ 2. Fooocus AI Rasm Promptlari Generator")
    fc_col1, fc_col2 = st.columns([2, 1])
    with fc_col1:
        fc_topic = st.text_input("Taqdimot Mavzusi:", value="Kvant kompyuterlari va fizika", key="fc_topic_in")
        fc_title = st.text_input("Slayd Sarlavhasi:", value="Kvant Superpozitsiyasi va Algoritmlar", key="fc_title_in")
    with fc_col2:
        fc_type = st.selectbox("Asset Turi:", ["background", "illustration", "icon", "diagram"], index=1)

    generated_fc_prompt = FooocusImageGenerator.build_slide_prompt(fc_topic, fc_title, fc_type)
    st.markdown("##### 📌 Fooocus uchun Optimallashtirilgan Prompt:")
    st.code(generated_fc_prompt, language="text")

    f_client = FooocusImageGenerator()
    if f_client.is_available():
        if st.button("⚡ Fooocus orqali Rasm Yaratish (Lokal GPU)", type="primary"):
            with st.spinner("⏳ Fooocus AI rasm generatsiya qilmoqda..."):
                img_path = f_client.generate(generated_fc_prompt)
                if img_path and os.path.exists(img_path):
                    st.image(str(img_path), caption="Generatsiya qilingan AI rasm")
                    st.success("✅ Rasm tayyor!")
    else:
        st.info("💡 Fooocus dasturi lokalda (http://127.0.0.1:7865) ishga tushirilganda, ushbu tugma orqali to'g'ridan-to'g'ri rasm generatsiya qilishingiz mumkin.")
