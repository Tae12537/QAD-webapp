import streamlit as st
import pandas as pd
import io
import re
import os
import datetime
from openpyxl import load_workbook

# ==========================================
# 💎 UI: LUXURY BRIGHT & BOLD TITLE EDITION
# ==========================================
st.set_page_config(page_title="QAD System Hub", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;500;800&family=Inter:wght@400;700;900&display=swap');

    /* พื้นหลังสว่าง สะอาดตา (Luxury Off-White) */
    .stApp {
        background-color: #f8fafc;
        color: #1e293b;
        font-family: 'Inter', 'Kanit', sans-serif;
    }

    /*Sidebar ขาวคลีน */
    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e2e8f0;
    }

    /* ปุ่มเมนูหลัก: เน้นขอบหนาและเงา (Industrial Luxury) */
    div.stButton > button:first-child {
        background: #ffffff;
        color: #0f172a;
        border: 3px solid #0f172a; 
        border-radius: 0px; 
        height: 160px;
        font-size: 32px !important;
        font-weight: 800;
        text-transform: uppercase;
        transition: all 0.3s ease;
        box-shadow: 10px 10px 0px rgba(15, 23, 42, 0.05);
    }

    div.stButton > button:hover {
        background: #0f172a;
        color: #ffffff !important;
        transform: translate(-3px, -3px);
        box-shadow: 15px 15px 0px #3b82f6;
    }

    /* 🔥 ชื่อแอปหลัก - ขยายใหญ่พิเศษตามคำขอ 🔥 */
    .main-title {
        font-family: 'Kanit', sans-serif;
        font-size: 120px; /* ใหญ่สะใจ อ่านง่ายแน่นอน */
        font-weight: 800;
        color: #0f172a;
        text-align: center;
        line-height: 1.0;
        margin-top: 40px;
        margin-bottom: 0px;
        letter-spacing: -4px;
    }

    .sub-title {
        text-align: center;
        color: #3b82f6; 
        font-size: 30px;
        font-weight: 500;
        margin-bottom: 4rem;
        letter-spacing: 2px;
    }

    /* ขนาดตัวหนังสือคงเดิมตามสั่ง */
    p, label, .stMarkdown {
        font-size: 19px !important;
    }
    
    h3 {
        font-size: 28px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        border-left: 6px solid #3b82f6;
        padding-left: 15px;
    }

    /* Tabs สไตล์ Modern */
    .stTabs [data-baseweb="tab"] {
        height: 60px;
        font-size: 20px;
        background-color: #ffffff;
        color: #475569;
        border: 1px solid #e2e8f0;
        padding: 0 40px;
    }

    .stTabs [aria-selected="true"] {
        background-color: #0f172a !important;
        color: #ffffff !important;
    }

    /* File Uploader */
    .stFileUploader section {
        border: 2px dashed #0f172a;
        background-color: #ffffff;
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

# --- ส่วนแอปต่างๆ (Logic และคำกำกับเดิมของคุณ) ---
def app_file_validator():
    st.markdown("<h1 style='text-align: center; color: #0f172a; font-size: 60px; font-weight: 800;'>📁 File Validator</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 22px; color: #64748b;'>ตรวจสอบ format ไฟล์ก่อนส่ง ✨</p>", unsafe_allow_html=True)
    st.divider()
    # ... (ใส่ส่วนที่เหลือของ Validator ของคุณตรงนี้) ...

def app_washing_processor():
    st.markdown("<h1 style='text-align: center; color: #0f172a; font-size: 60px; font-weight: 800;'>📊 Washing Date Processor</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 22px; color: #64748b;'>ตรวจสอบวันล้างจาก lot no. ⚡</p>", unsafe_allow_html=True)
    st.divider()
    # ... (ใส่ส่วนที่เหลือของ Processor ของคุณตรงนี้) ...

# ==========================================
# MAIN ROUTING
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<p class='main-title'>QAD SYSTEM HUB</p>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>📂 Quality Engineering Support Application 🚀</p>", unsafe_allow_html=True)
    
    _, col_main, _ = st.columns([0.1, 0.8, 0.1])
    
    with col_main:
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown("### 📋 ตรวจเช็ค Format ไฟล์")
            if st.button("📁 File Validator", use_container_width=True):
                st.session_state.current_app = "Validator"; st.rerun()
        with c2:
            st.markdown("### 🧼 ตรวจวันล้างสินค้า")
            if st.button("📊 Washing Date Processor", use_container_width=True):
                st.session_state.current_app = "Processor"; st.rerun()

    st.markdown("<div style='height: 100px;'></div>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 16px;'>© 2026 Quality Engineering | Systems v2.0 ✨</p>", unsafe_allow_html=True)

elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
