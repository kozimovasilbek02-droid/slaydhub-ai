import streamlit as st
import os
import sys
import glob
from PIL import Image

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import slide_manager

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="AI PowerPoint Precision Vector Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Clean Styling
st.markdown("""
<style>
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2rem;
    }
    .status-badge {
        display: inline-block;
        padding: 5px 14px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 0.95rem;
    }
    .status-ready {
        background: #DCFCE7;
        color: #166534;
        border: 1px solid #BBF7D0;
    }
    .status-pending {
        background: #FEF3C7;
        color: #92400E;
        border: 1px solid #FDE68A;
    }
    .stButton>button {
        background: linear-gradient(135deg, #0A3B75 0%, #1E40AF 100%);
        color: white;
        font-weight: 700;
        font-size: 1.05rem;
        border-radius: 10px;
        padding: 12px 28px;
        border: none;
        box-shadow: 0 4px 12px rgba(10, 59, 117, 0.15);
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #F15A24 0%, #EA580C 100%);
        color: white;
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)

# ==================== SIDEBAR ====================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/microsoft-powerpoint-2019.png", width=52)
    st.header("📁 Taqdimot Tanlash")
    
    preset_folders = slide_manager.get_cached_folders()
    preset_keys = list(preset_folders.keys())
    
    selected_key = st.selectbox("Tanlangan Taqdimot:", preset_keys, index=0)
    target_folder = preset_folders[selected_key]
    
    st.markdown("---")
    st.markdown("🔍 **Taqdimotlarni qidirish (597+ taqdimot):**")
    search_q = st.text_input("Nomi bo'yicha qidiring:", placeholder="Masalan: Leadership, Marketing, 5G...")
    if search_q:
        results = slide_manager.search_drive_folders(search_q)
        if results:
            chosen_res = st.selectbox("Qidiruv natijalari:", results)
            if st.button("Ushbu papkani ochish"):
                target_folder = chosen_res
                st.session_state['custom_folder'] = target_folder
                st.rerun()
        else:
            st.caption("Hech qanday papka topilmadi.")
            
    if 'custom_folder' in st.session_state and os.path.exists(st.session_state['custom_folder']):
        target_folder = st.session_state['custom_folder']
        
    st.markdown("---")
    manual_path = st.text_input("yoki To'liq manzilni kiriting:", value=target_folder)
    if manual_path and os.path.exists(manual_path):
        target_folder = manual_path

# ==================== MAIN CONTENT ====================
st.title("🎯 AI PowerPoint Vector Precision Converter")
st.markdown("Slayd rasmlaridan **OCR matn qutilari**, **vektor kartalar** va **ranglar palitrasini** ajratib olib, 100% tahrir qilinadigan professional PPTX yaratish tizimi.")

# Validate folder
if not os.path.exists(target_folder):
    st.error(f"⚠️ Tanlangan papka topilmadi: `{target_folder}`")
    st.stop()

folder_name = os.path.basename(os.path.normpath(target_folder))
original_slides = slide_manager.get_folder_slides(target_folder)
pptx_path = slide_manager.get_folder_pptx_path(target_folder)
pptx_ready = (pptx_path is not None and os.path.exists(pptx_path))

# Info Overview Bar
c_info1, c_info2, c_info3 = st.columns([2, 1, 1])
with c_info1:
    st.markdown(f"📂 **Faol Taqdimot:** `{folder_name}`")
with c_info2:
    st.markdown(f"🖼️ **Original Slaydlar:** `{len(original_slides)} ta`")
with c_info3:
    if pptx_ready:
        st.markdown('<span class="status-badge status-ready">✅ Vector PPTX Tayyor</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="status-badge status-pending">⏳ PPTX Yaratilmagan</span>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 1-Action: Convert Button
c_act1, c_act2 = st.columns([2, 1])
with c_act1:
    btn_generate = st.button("🚀 Ushbu Papkadan Vektor Taqdimot Yaratish (AI Smart Generate)", use_container_width=True)

if btn_generate:
    with st.spinner("⏳ OCR tahlil, vektor shakllar va matn qutilari yaratilmoqda..."):
        success, msg = slide_manager.convert_presentation(target_folder)
        if success:
            st.success("✅ Taqdimot 100% muvaffaqiyatli yaratildi!")
            st.rerun()
        else:
            st.error(f"❌ Xatolik: {msg}")

st.markdown("---")

# 2-Action: Download & Open
if pptx_ready:
    st.markdown("### 📊 Tayyor Vektor Taqdimot")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        with open(pptx_path, "rb") as f:
            st.download_button(
                label="📥 PPTX Faylni Yuklab Olish",
                data=f,
                file_name=os.path.basename(pptx_path),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
    with col_d2:
        if st.button("📂 PowerPoint Dasturida Ochish", use_container_width=True):
            try:
                os.startfile(pptx_path)
                st.toast("PowerPoint dasturida ochildi!", icon="✅")
            except Exception as e:
                st.error(f"Ochishda xatolik: {e}")

    st.markdown("---")
    
    # 3-Action: Side-by-side Comparison Gallery
    st.markdown(f"### 🖼️ Slaydlarni Taqqoslash: `{folder_name}`")
    st.caption("Chap tomonda original rasm, o'ng tomonda AI tomonidan noldan chizilgan tahrir qilinadigan vektor slayd:")

    out_folder = os.path.join(current_dir, "output", "presentations", folder_name)
    comparisons = sorted(glob.glob(os.path.join(out_folder, "comparisons", "Comparison_Slide_*.png")), key=slide_manager.natural_sort_key)
    generated_previews = slide_manager.export_pptx_previews(pptx_path, folder_name)
    
    total_display = max(len(original_slides), len(generated_previews), len(comparisons))

    if total_display > 0:
        tab_titles = [f"Slayd {i+1}" for i in range(total_display)]
        tabs = st.tabs(tab_titles)
        
        for idx, tab in enumerate(tabs):
            with tab:
                if idx < len(comparisons):
                    st.image(comparisons[idx], width="stretch")
                else:
                    col_left, col_right = st.columns(2)
                    with col_left:
                        st.markdown(f"#### 📄 Original Slayd {idx+1}")
                        if idx < len(original_slides):
                            st.image(original_slides[idx], width="stretch")
                        else:
                            st.info("Original rasm mavjud emas.")
                            
                    with col_right:
                        st.markdown(f"#### ✨ Qayta Yaratilgan Vektor Natija {idx+1}")
                        if idx < len(generated_previews):
                            st.image(generated_previews[idx], width="stretch")
                        else:
                            st.warning("Ushbu slayd uchun prevyu mavjud emas.")
else:
    st.info("💡 Ushbu papka uchun hali taqdimot yaratilmagan. Yuqoridagi **'Ushbu Papkadan Vektor Taqdimot Yaratish'** tugmasini bosing.")
