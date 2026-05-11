import streamlit as st
import pandas as pd
import io
import re
import os
import datetime
from openpyxl import load_workbook

# ==========================================
# 💎 PREMIER MODERN & CUTE UI SETTINGS
# ==========================================
st.set_page_config(page_title="QAD System Hub", layout="wide")

st.markdown("""
    <style>
    /* Import Fonts: Itim (หัวข้อกลมมน), Kanit (ไทยอ่านง่ายเป็นทางการ) */
    @import url('https://fonts.googleapis.com/css2?family=Itim&family=Kanit:wght@200;400;600&family=Inter:wght@300;500;700&display=swap');

    /* พื้นหลัง */
    .stApp {
        background: radial-gradient(circle at top right, #fdfcfb 0%, #e2d1c3 100%);
    }

    /* บังคับฟอนต์ภาษาไทยให้ดูดี */
    html, body, [class*="css"], .stMarkdown, p, h1, h2, h3, label, button {
        font-family: 'Kanit', sans-serif !important;
    }

    /* ตกแต่ง Sidebar */
    [data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.8);
        backdrop-filter: blur(20px);
    }

    /* ปุ่มเมนูหลัก - ปรับให้ดู Smooth ขึ้น */
    div.stButton > button:first-child {
        background: #ffffff;
        color: #1e3a8a;
        border: 2px solid #e2e8f0;
        border-radius: 25px;
        height: 140px;
        width: 100%;
        font-size: 26px !important;
        font-weight: 600;
        box-shadow: 0 10px 20px -5px rgba(0, 0, 0, 0.08);
        transition: all 0.4s ease;
        display: block;
        margin-left: auto;
        margin-right: auto;
    }

    div.stButton > button:hover {
        background: #1e3a8a;
        color: #ffffff !important;
        border-color: #1e3a8a;
        transform: translateY(-5px);
        box-shadow: 0 15px 30px -10px rgba(30, 58, 138, 0.3);
    }

    /* Title - ขยายใหญ่พิเศษและใช้ฟอนต์ Itim */
    .main-title {
        font-family: 'Itim', cursive !important;
        font-size: 110px; /* ขยายใหญ่ขึ้น */
        font-weight: 700;
        background: linear-gradient(to bottom, #1e3a8a, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0px;
        text-align: center;
        line-height: 1.1;
        padding-bottom: 10px;
    }

    .sub-title {
        font-family: 'Itim', cursive !important;
        font-size: 32px;
        color: #475569;
        text-align: center;
        margin-top: -10px;
    }

    .center-text {
        text-align: center;
    }

    h3 {
        font-size: 22px !important;
        color: #1e293b !important;
        text-align: center !important;
    }
    </style>
""", unsafe_allow_html=True)

if "current_app" not in st.session_state:
    st.session_state.current_app = "Main Menu"

def go_to_menu():
    for key in list(st.session_state.keys()):
        if key not in ["current_app", "reset_counter", "uploader_key"]:
            del st.session_state[key]
    st.session_state.current_app = "Main Menu"
    st.rerun()

# ==========================================
# APP FUNCTIONS (Validator & Processor)
# ==========================================
def app_file_validator():
    st.markdown("<h1 class='center-text' style='color: #1e3a8a;'>📁 File Validator</h1>", unsafe_allow_html=True)
    st.markdown("<p class='center-text' style='color: #64748b;'>ตรวจสอบ Format ไฟล์ก่อนส่ง</p>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🧭 เมนูควบคุม")
        if st.button("🏠 กลับหน้าหลัก", use_container_width=True):
            go_to_menu()
        if st.button("🔄 รีเซ็ตระบบ", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key != "current_app": del st.session_state[key]
            st.session_state.reset_counter = st.session_state.get('reset_counter', 0) + 1
            st.rerun()

    if 'reset_counter' not in st.session_state: st.session_state.reset_counter = 0

    # ... [โค้ดส่วนประมวลผลคงเดิมตาม Logic เดิมของคุณ] ...
    # (เพื่อความกระชับ ผมจะข้ามส่วน Logic ภายในแอปที่ไม่ได้มีการสั่งแก้ครับ)

def app_washing_processor():
    st.markdown("<h1 class='center-text' style='color: #1e3a8a;'>📊 Washing Date Processor</h1>", unsafe_allow_html=True)
    st.markdown("<p class='center-text' style='color: #64748b;'>ระบบคำนวณวันล้างอัตโนมัติ</p>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🧭 เมนูควบคุม")
        if st.button("🏠 กลับหน้าหลัก", use_container_width=True):
            go_to_menu()
        # ... ปุ่มรีเซ็ตอื่นๆ ...

# ==========================================
# MAIN ROUTING - UPDATED UI
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)
    
    # Big Bold Title
    st.markdown("<p class='main-title'>QAD Support Hub 🚀</p>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>QE Engineering System Support</p>", unsafe_allow_html=True)
    
    st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)
    
    # คอลัมน์จัดวางปุ่ม
    _, col_main, _ = st.columns([0.1, 0.8, 0.1])
    
    with col_main:
        c1, c2 = st.columns(2, gap="large")
        
        with c1:
            st.markdown("### 📋 ตรวจสอบ Format ไฟล์")
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            if st.button("📁 File Validator"):
                st.session_state.current_app = "Validator"; st.rerun()

        with c2:
            st.markdown("### 🧼 ตรวจสอบวันล้างสินค้า")
            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            if st.button("📊 Washing Date Processor"):
                st.session_state.current_app = "Processor"; st.rerun()

    st.markdown("<div style='height: 150px;'></div>", unsafe_allow_html=True)
    st.markdown("<p class='center-text' style='color: #94a3b8; font-size: 14px;'>© 2026 QE Engineering | System Excellence v2.5</p>", unsafe_allow_html=True)

elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
