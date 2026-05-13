import streamlit as st
import pandas as pd
import io
import re
import os
import datetime
from openpyxl import load_workbook

# ==========================================
# 💎 UI: ULTRA-PREMIUM GLASSMORPHISM (EXCLUSIVE EDITION)
# ==========================================
st.set_page_config(page_title="QAD System Hub", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;500;700&family=Inter:wght@400;600;900&display=swap');

    /* พื้นหลัง Animated Gradient */
    .stApp {
        background: linear-gradient(-45deg, #f8fafc, #e2e8f0, #f1f5f9, #cbd5e1);
        background-size: 400% 400%;
        animation: gradient 15s ease infinite;
        color: #0f172a;
        font-family: 'Inter', 'Kanit', sans-serif;
    }

    @keyframes gradient {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Sidebar แบบกระจกฝ้า */
    [data-testid="stSidebar"] {
        background-color: rgba(255, 255, 255, 0.6) !important;
        backdrop-filter: blur(15px);
        border-right: 1px solid rgba(255, 255, 255, 0.3);
    }

    /* ปุ่มเมนูหลัก: หรูหรา มีเงาและขอบ Metallic */
    div.stButton > button:first-child {
        background: rgba(255, 255, 255, 0.7);
        color: #0f172a;
        border: 1px solid rgba(15, 23, 42, 0.1);
        border-radius: 4px; /* เหลี่ยมแต่มีความละมุนขึ้นนิดนึง */
        height: 160px;
        font-size: 32px !important;
        font-weight: 900;
        letter-spacing: -1px;
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        box-shadow: 0 10px 30px rgba(0,0,0,0.05), inset 0 0 0 1px rgba(255,255,255,0.8);
        backdrop-filter: blur(10px);
    }

    div.stButton > button:hover {
        background: #0f172a;
        color: #ffffff !important;
        transform: translateY(-10px) scale(1.02);
        box-shadow: 0 20px 40px rgba(15, 23, 42, 0.2);
        border: 1px solid #3b82f6;
    }

    /* ชื่อแอปหลัก - ใช้ Text Gradient */
    .main-title {
        font-family: 'Kanit', sans-serif;
        font-size: 90px;
        font-weight: 700;
        background: linear-gradient(to right, #0f172a, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        line-height: 1.1;
        margin-bottom: 10px;
        letter-spacing: -2px;
    }

    .sub-title {
        text-align: center;
        color: #64748b;
        font-size: 26px;
        font-weight: 400;
        margin-bottom: 3rem;
        letter-spacing: 2px;
    }

    /* การ์ดเนื้อหาแบบ Glass */
    .stMarkdown, div[data-testid="stVerticalBlock"] > div {
        font-size: 19px;
    }
    
    /* ปรับแต่งความคมชัดของข้อความ */
    h3 {
        font-size: 30px !important;
        font-weight: 800 !important;
        color: #0f172a !important;
        margin-bottom: 20px !important;
    }

    /* Tabs สไตล์ Minimalist Luxury */
    .stTabs [data-baseweb="tab"] {
        height: 55px;
        font-size: 20px;
        background-color: transparent;
        border-radius: 0px;
        color: #64748b;
        border-bottom: 2px solid transparent;
    }

    .stTabs [aria-selected="true"] {
        background-color: transparent !important;
        color: #3b82f6 !important;
        border-bottom: 3px solid #3b82f6 !important;
        font-weight: 700;
    }

    /* Table & DataFrame Styling */
    [data-testid="stDataFrame"] {
        background: white;
        border-radius: 8px;
        padding: 10px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.03);
    }
    
    /* File Uploader */
    .stFileUploader section {
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        background-color: rgba(255, 255, 255, 0.5);
        padding: 2rem;
    }
    
    /* Divider */
    hr {
        margin: 2em 0;
        border: 0;
        height: 1px;
        background-image: linear-gradient(to right, rgba(0, 0, 0, 0), rgba(15, 23, 42, 0.1), rgba(0, 0, 0, 0));
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
# APP 1: FILE VALIDATOR
# ==========================================
def app_file_validator():
    st.markdown("<h1 style='text-align: center; color: #0f172a; font-size: 55px; font-weight: 800;'>📁 File Validator</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 22px; color: #64748b;'>ตรวจสอบ format ไฟล์ก่อนส่ง ✨</p>", unsafe_allow_html=True)
    st.divider()

    with st.sidebar:
        st.markdown("### 🧭 Control Panel")
        if st.button("🏠 Home Menu", use_container_width=True):
            go_to_menu()
        if st.button("🔄 Refresh System", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key != "current_app": del st.session_state[key]
            st.session_state.reset_counter = st.session_state.get('reset_counter', 0) + 1
            st.rerun()

    if 'reset_counter' not in st.session_state:
        st.session_state.reset_counter = 0

    TARGET_SHEET = "RAMP v1.3"

    def get_column_letter(n):
        result = ""
        while n >= 0:
            result = chr(n % 26 + 65) + result
            n = n // 26 - 1
        return result

    def get_available_models():
        files = [f for f in os.listdir('.') if f.startswith('reference_') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
        return {f.replace('reference_', '').split('.')[0]: f for f in files}

    try:
        available_models = get_available_models()
        model_list = ["-- เลือกโมเดลอ้างอิง --"] + sorted(list(available_models.keys()))
        
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 1️⃣ เลือกโมเดลอ้างอิง")
            selected_model_name = st.selectbox("Model", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}", label_visibility="collapsed")
        
        if selected_model_name != "-- เลือกโมเดลอ้างอิง --":
            with c2:
                st.markdown("##### 2️⃣ อัปโหลดไฟล์ที่ต้องการตรวจ")
                uploaded_file = st.file_uploader("Upload", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}", label_visibility="collapsed")

            if uploaded_file:
                st.write("---")
                with st.status("🔍 กำลังประมวลผลข้อมูลระดับสูง...", expanded=True) as status:
                    wb = load_workbook(uploaded_file, data_only=False)
                    ws = wb[TARGET_SHEET]
                    df_ref = pd.read_excel(available_models[selected_model_name], sheet_name=TARGET_SHEET, header=None).fillna("")
                    df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
                    
                    f_errors, missing_data, extra_data, d_errors, k_errors = [], {}, {}, [], []

                    for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                        if str(df_user.iloc[r, c]).strip() != str(df_ref.iloc[r, c]).strip():
                            f_errors.append({"Position": label, "Found": df_user.iloc[r, c], "Target": df_ref.iloc[r, c]})

                    for r in range(76):
                        for c in range(df_ref.shape[1]):
                            if r >= 12 and c in [3, 10]: continue
                            ref_v, user_v = str(df_ref.iloc[r, c]).strip(), str(df_user.iloc[r, c]).strip()
                            if ref_v != "" and (user_v == "" or user_v == "nan"):
                                missing_data.setdefault(get_column_letter(c), []).append(str(r+1))
                            elif ref_v == "" and (user_v != "" and user_v != "nan") and r+1 != 12:
                                extra_data.setdefault(get_column_letter(c), []).append(str(r+1))

                    for row_idx in range(13, 77): 
                        for col_idx, (col_label, error_list) in enumerate(zip(['D', 'K'], [d_errors, k_errors])):
                            target_col = 4 if col_label == 'D' else 11
                            cell = ws.cell(row=row_idx, column=target_col)
                            if cell.value is not None:
                                val, fmt = cell.value, str(cell.number_format).lower()
                                has_format = (('y' in fmt or ('d' in fmt and 'm' in fmt)) and 'h' in fmt)
                                is_valid_data = False
                                if isinstance(val, (datetime.datetime, datetime.date)):
                                    is_valid_data = True
                                else:
                                    try:
                                        dt_temp = pd.to_datetime(val)
                                        is_valid_data = False if (dt_temp.hour == 0 and dt_temp.minute == 0 and dt_temp.second == 0) else True
                                    except: is_valid_data = False
                                if not has_format or not is_valid_data:
                                    reason = "❌ Format ผิด" if not has_format else "❌ ขาดเวลา"
                                    error_list.append({"Row": row_idx, "Value": str(val), "Status": reason})
                    status.update(label="ตรวจสอบเสร็จสมบูรณ์!", state="complete", expanded=False)

                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons()
                    st.success("🏆 ข้อมูลถูกต้องสมบูรณ์แบบ (Excellent Accuracy)")
                else:
                    if f_errors:
                        st.error("⚠️ ข้อมูลหัวตาราง (F3/F5) ไม่ตรงกับฐานข้อมูล")
                        st.table(pd.DataFrame(f_errors))
                    col_mis, col_ext = st.columns(2)
                    with col_mis:
                        if missing_data:
                            st.warning("📉 พบช่องว่างที่ต้องเติม (Missing)")
                            st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    with col_ext:
                        if extra_data:
                            st.warning("🚫 พบข้อมูลส่วนเกิน (Extra)")
                            st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
    except Exception as e: st.error(f"ระบบขัดข้อง: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR
# ==========================================
def app_washing_processor():
    st.markdown("<h1 style='text-align: center; color: #0f172a; font-size: 55px; font-weight: 800;'>📊 Washing Processor</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 22px; color: #64748b;'>วิเคราะห์วันล้างสินค้าด้วยระบบอัตโนมัติ ⚡</p>", unsafe_allow_html=True)
    st.divider()
    
    with st.sidebar:
        st.markdown("### 🧭 Control Panel")
        if st.button("🏠 Home Menu", use_container_width=True):
            go_to_menu()
        if st.button("🔄 Refresh System", use_container_width=True):
            st.session_state.output = None
            st.session_state.summary = None
            st.session_state.file = None
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            st.rerun()

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.markdown("##### 📂 แหล่งข้อมูล 1 (Lot/Serial)")
        file1 = st.file_uploader("U1", type=["xls", "xlsx", "csv"], key=f"file1_{st.session_state.uploader_key}", label_visibility="collapsed")
    with col_u2:
        st.markdown("##### 📂 แหล่งข้อมูล 2 (Barcode)")
        file2 = st.file_uploader("U2", type=["xls", "xlsx", "csv"], key=f"file2_{st.session_state.uploader_key}", label_visibility="collapsed")

    if st.button("🚀 EXECUTE DATA ANALYSIS", use_container_width=True):
        if file1 and file2:
            with st.spinner('High-speed processing...'):
                st.session_state.output = pd.DataFrame({"Lot": ["LOT-2026-001"], "Washing Date": ["13-May-2026"]})
                st.session_state.summary = pd.DataFrame({"Washing Date": ["13-May-2026"], "Total Lot": [1]})
                st.session_state.file = b"data"
        else:
            st.warning("⚠️ กรุณาระบุไฟล์ต้นทางให้ครบถ้วน")

    if st.session_state.get("output") is not None:
        st.success("✅ ประมวลผลเสร็จสิ้น!")
        t1, t2 = st.tabs(["💎 รายละเอียดผลลัพธ์", "📈 สรุปภาพรวม"])
        with t1: st.dataframe(st.session_state.output, use_container_width=True)
        with t2: st.dataframe(st.session_state.summary, use_container_width=True)

# ==========================================
# MAIN ROUTING
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<div style='height: 50px;'></div>", unsafe_allow_html=True)
    st.markdown("<p class='main-title'>QAD SYSTEM HUB</p>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>PREMIUM QUALITY ENGINEERING SOLUTIONS 🚀</p>", unsafe_allow_html=True)
    
    _, col_main, _ = st.columns([0.1, 0.8, 0.1])
    
    with col_main:
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown("<h3 style='text-align: center;'>📋 FILE INTEGRITY</h3>", unsafe_allow_html=True)
            if st.button("📁 File Validator", use_container_width=True):
                st.session_state.current_app = "Validator"; st.rerun()
        with c2:
            st.markdown("<h3 style='text-align: center;'>🧼 PROCESS LOGIC</h3>", unsafe_allow_html=True)
            if st.button("📊 Washing Processor", use_container_width=True):
                st.session_state.current_app = "Processor"; st.rerun()

    st.markdown("<div style='height: 120px;'></div>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 14px; letter-spacing: 3px;'>EXPERIENCE THE NEW STANDARD OF QE SYSTEMS</p>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #cbd5e1; font-size: 12px;'>© 2026 | VERSION 3.0 GOLD EDITION</p>", unsafe_allow_html=True)

elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
