import streamlit as st
import pandas as pd
import io
import re
import os
import datetime
from openpyxl import load_workbook

# ==========================================
# 💎 PREMIER LUXURY & CUTE UI SETTINGS
# ==========================================
st.set_page_config(page_title="QAD System Hub", layout="wide")

st.markdown("""
    <style>
    /* Import Fonts: Itim (กลมมนน่ารักแต่อ่านง่ายกว่า), Mali (ใช้ในส่วนรอง) */
    @import url('https://fonts.googleapis.com/css2?family=Mali:wght@300;500;700&family=Itim&family=Inter:wght@300;500;700&display=swap');

    /* พื้นหลังแบบ Soft Gradient */
    .stApp {
        background: radial-gradient(circle at top right, #fdfcfb 0%, #e2d1c3 100%);
        font-family: 'Itim', 'Mali', sans-serif;
    }

    /* ฟอนต์ภาษาไทยหลักใช้ Itim เพื่อให้อ่านง่ายขึ้นแต่ยังน่ารัก */
    html, body, [class*="css"], .stMarkdown, p, h1, h2, h3, label, button {
        font-family: 'Itim', sans-serif !important;
    }

    /* ตกแต่ง Sidebar */
    [data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.8);
        backdrop-filter: blur(20px);
        border-right: 1px solid rgba(0,0,0,0.05);
    }

    div.block-container {
        padding-top: 2rem;
    }

    /* ปุ่มเมนูหลักสไตล์ Premium */
    div.stButton > button:first-child {
        background: #ffffff;
        color: #1e3a8a;
        border: 2px solid #e2e8f0;
        border-radius: 30px;
        height: 140px;
        width: 100%;
        font-size: 26px !important;
        font-weight: 700;
        box-shadow: 0 10px 20px -5px rgba(0, 0, 0, 0.08);
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        text-transform: uppercase;
        display: block;
        margin-left: auto;
        margin-right: auto;
    }

    div.stButton > button:hover {
        background: #1e3a8a;
        color: #ffffff !important;
        border-color: #1e3a8a;
        transform: scale(1.05);
        box-shadow: 0 20px 30px -10px rgba(30, 58, 138, 0.3);
    }

    /* ชื่อแอปหลัก - ปรับขนาดใหญ่พิเศษ (110px) และเปลี่ยนฟอนต์ให้อ่านง่ายขึ้น */
    .main-title {
        font-family: 'Itim', cursive !important;
        font-size: 110px; /* ขยายขนาดใหญ่ขึ้น */
        font-weight: 700;
        background: linear-gradient(to right, #1e3a8a, #3b82f6, #1e3a8a);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 5px;
        text-align: center;
        filter: drop-shadow(3px 5px 8px rgba(0,0,0,0.12));
        line-height: 1.2;
    }

    .center-text {
        text-align: center;
    }

    h3 {
        font-size: 24px !important;
        color: #334155 !important;
        margin-bottom: 15px !important;
    }

    /* ปรับแต่ง Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }

    .stTabs [data-baseweb="tab"] {
        font-size: 18px;
        color: #64748b;
    }

    [data-testid="stDataFrame"] {
        border-radius: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
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
    st.markdown("<h1 class='center-text' style='color: #1e3a8a;'>📁 File Validator</h1>", unsafe_allow_html=True)
    st.markdown("<p class='center-text' style='color: #64748b; font-size: 20px;'>ตรวจสอบ format ไฟล์ก่อนส่ง ✨</p>", unsafe_allow_html=True)
    
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
            selected_model_name = st.selectbox("📍 Reference Model", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}")
        
        if selected_model_name != "-- เลือกโมเดลอ้างอิง --":
            with c2:
                uploaded_file = st.file_uploader("📥 อัปโหลดไฟล์ที่ต้องการตรวจ", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}")

            if uploaded_file:
                st.write("---")
                with st.status("🔍 กำลังตรวจสอบความถูกต้องของไฟล์...", expanded=True) as status:
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
                    status.update(label="ตรวจสอบเสร็จเรียบร้อย!", state="complete", expanded=False)

                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons()
                    st.success("✨ ยินดีด้วย! ข้อมูลถูกต้องสมบูรณ์ 100%")
                else:
                    if f_errors:
                        st.error("⚠️ หัวตารางไม่ตรง (F3/F5)")
                        st.table(pd.DataFrame(f_errors))

                    col_mis, col_ext = st.columns(2)
                    with col_mis:
                        if missing_data:
                            st.warning("📉 ข้อมูลที่หายไป (Missing Data)")
                            st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    with col_ext:
                        if extra_data:
                            st.warning("🚫 ข้อมูลส่วนเกิน (Extra Data)")
                            st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])

                    col_d, col_k = st.columns(2)
                    with col_d:
                        if d_errors:
                            st.error("🔍 ปัญหาที่คอลัมน์ Washing Date (D)")
                            st.dataframe(pd.DataFrame(d_errors), use_container_width=True, hide_index=True)
                    with col_k:
                        if k_errors:
                            st.error("🔍 ปัญหาที่คอลัมน์ Finish Date (K)")
                            st.dataframe(pd.DataFrame(k_errors), use_container_width=True, hide_index=True)
                            
    except Exception as e: st.error(f"เกิดข้อผิดพลาด: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR (FIXED VERSION)
# ==========================================
def app_washing_processor():
    st.markdown("<h1 class='center-text' style='color: #1e3a8a;'>📊 Washing Date Processor</h1>", unsafe_allow_html=True)
    st.markdown("<p class='center-text' style='color: #64748b; font-size: 20px;'>คำนวณวันล้างแบบรวดเร็วและแม่นยำ ⚡</p>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🧭 Control Panel")
        if st.button("🏠 Home Menu", use_container_width=True):
            go_to_menu()
        if st.button("🔄 Reset System", use_container_width=True):
            st.session_state.output = None
            st.session_state.summary = None
            st.session_state.file = None
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            st.rerun()

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    # UI ส่วนการอัปโหลด
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        file1 = st.file_uploader("📂 Upload File 1 (Lot List)", type=["xls", "xlsx", "csv"], key=f"file1_{st.session_state.uploader_key}")
    with col_f2:
        file2 = st.file_uploader("📂 Upload File 2 (Barcode Data)", type=["xls", "xlsx", "csv"], key=f"file2_{st.session_state.uploader_key}")

    def read_excel_safely(file):
        try:
            return pd.read_excel(file, engine="openpyxl", header=None)
        except:
            return pd.read_excel(file, engine="xlrd", header=None)

    # ฟังก์ชันอ่านไฟล์ 1 (Lot List)
    def process_file1(file):
        df = read_excel_safely(file)
        # อ่านจากแถว 17 (index 16) คอลัมน์ F (index 5)
        raw_data = df.iloc[16:, 5].dropna()
        lot_list = [str(val).strip() for val in raw_data if str(val).strip() != "" and str(val).lower() != "lot/serial"]
        return pd.DataFrame({"Lot": lot_list})

    # ฟังก์ชันอ่านไฟล์ 2 (Barcode/Runcard)
    def process_file2(file):
        df = read_excel_safely(file)
        header_row = None
        # วนลูปหาแถวที่มี Header จริงๆ
        for i in range(min(25, len(df))):
            row_values = df.iloc[i].astype(str).str.lower().values
            if any("runcard" in val for val in row_values) and any("barcode" in val for val in row_values):
                header_row = i
                break
        
        if header_row is None:
            return None
            
        df.columns = df.iloc[header_row]
        df = df[header_row + 1:].copy()
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        # ค้นหาชื่อคอลัมน์ที่แท้จริง (เผื่อชื่อไม่เป๊ะ)
        col_lot = next((c for c in df.columns if "runcard" in c), None)
        col_bar = next((c for c in df.columns if "barcode" in c), None)
        col_date = next((c for c in df.columns if "packed" in c and "date" in c), None)
        
        if not all([col_lot, col_bar, col_date]):
            return None
            
        df_out = df[[col_lot, col_bar, col_date]].copy()
        df_out.columns = ["Lot", "Barcode No", "Packed Date"]
        df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
        df_out["Packed Date"] = pd.to_datetime(df_out["Packed Date"], errors="coerce")
        return df_out.dropna(subset=["Lot"])

    # ฟังก์ชันดึง WW และ Day จาก Barcode
    def extract_logic(barcode):
        s = str(barcode).strip()
        match = re.search('[A-Za-z]', s)
        if not match: return pd.Series([None, None])
        start = match.start()
        code = s[start+3:start+6] # ดึงตัวเลข 3 หลักต่อจากตัวอักษร+2
        if code.isdigit() and len(code) == 3:
            return pd.Series([int(code[:2]), int(code[2])])
        return pd.Series([None, None])

    if st.button("🚀 START PROCESSING", use_container_width=True):
        if not file1 or not file2:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ช่อง")
        else:
            try:
                with st.spinner('กำลังวิเคราะห์ข้อมูล...'):
                    df1 = process_file1(file1)
                    df2 = process_file2(file2)
                    
                    if df2 is None:
                        st.error("❌ หาหัวตารางใน File 2 ไม่เจอ (ต้องมีคำว่า Runcard และ Barcode)")
                        return

                    # Merge ข้อมูล
                    merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
                    merged[['WW', 'Day']] = merged['Barcode No'].apply(extract_logic)
                    
                    # โหลด Database
                    if not os.path.exists("database.txt"):
                        st.error("❌ ไม่พบไฟล์ database.txt ในระบบ")
                        return
                        
                    db = pd.read_csv("database.txt")
                    db["Date"] = pd.to_datetime(db["Date"], format="%d-%b-%Y", errors="coerce")

                    def find_best_date(row, d_db):
                        if pd.isna(row["WW"]) or pd.isna(row["Day"]) or pd.isna(row["Packed Date"]): 
                            return None
                        cands = d_db[(d_db["WW"] == row["WW"]) & (d_db["Day"] == row["Day"])].copy()
                        if cands.empty: return None
                        cands["diff"] = (cands["Date"] - row["Packed Date"]).abs()
                        return cands.sort_values("diff").iloc[0]["Date"]

                    merged["Washing Date Raw"] = merged.apply(lambda r: find_best_date(r, db), axis=1)
                    
                    # เตรียมผลลัพธ์
                    output = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date Raw"]].copy()
                    output["Washing Date"] = pd.to_datetime(output["Washing Date Raw"]).dt.strftime("%d-%b-%Y")
                    
                    # สรุปผล
                    summary = output.dropna(subset=["Washing Date Raw"])
                    summary = summary.groupby("Washing Date")["Lot"].count().reset_index()
                    summary.columns = ["Washing Date", "Total Lot"]

                    # บันทึกลง Buffer
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        output.drop(columns=["Washing Date Raw"]).to_excel(writer, index=False, sheet_name="Result")
                        summary.to_excel(writer, index=False, sheet_name="Summary")
                    
                    st.session_state.output = output.drop(columns=["Washing Date Raw"])
                    st.session_state.summary = summary
                    st.session_state.file = buf.getvalue()
                    
            except Exception as e:
                st.error(f"💥 เกิดข้อผิดพลาดขณะประมวลผล: {str(e)}")

    # แสดงผลลัพธ์
    if st.session_state.get("output") is not None:
        st.success("✅ ประมวลผลสำเร็จ!")
        tab1, tab2 = st.tabs(["💎 Result Details", "📈 Summary View"])
        with tab1:
            st.dataframe(st.session_state.output, use_container_width=True)
        with tab2:
            st.dataframe(st.session_state.summary, use_container_width=True)
        
        st.download_button(
            label="📥 DOWNLOAD EXCEL RESULT",
            data=st.session_state.file,
            file_name=f"Washing_Date_{datetime.datetime.now().strftime('%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
# ==========================================
# MAIN ROUTING - UPDATED FOR PERFECT CENTERING
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<div style='height: 60px;'></div>", unsafe_allow_html=True)
    # ปรับหัวข้อให้ใหญ่และเปลี่ยนสไตล์ฟอนต์ตามคำขอ
    st.markdown("<p class='main-title'></p>", unsafe_allow_html=True)
    st.markdown("<p class='center-text' style='color: #475569; font-size: 28px; font-weight: 500;'>📂 QAD Support Application 🚀</p>", unsafe_allow_html=True)
    st.markdown("<div style='height: 60px;'></div>", unsafe_allow_html=True)
    
    # ดึงปุ่มมาตรงกลางหน้าจอ (Layout เดิม)
    _, col_main, _ = st.columns([0.15, 0.7, 0.15])
    
    with col_main:
        c1, c2 = st.columns(2, gap="large")
        
        with c1:
            st.markdown("### 📋 ตรวจเช็ค Format ไฟล์")
            st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
            if st.button("📁 File Validator"):
                st.session_state.current_app = "Validator"; st.rerun()

        with c2:
            st.markdown("### 🧼 ตรวจวันล้างสินค้า")
            st.markdown("<div style='height: 5px;'></div>", unsafe_allow_html=True)
            if st.button("📊 Washing Date Processor"):
                st.session_state.current_app = "Processor"; st.rerun()

    st.markdown("<div style='height: 120px;'></div>", unsafe_allow_html=True)
    st.markdown("<p class='center-text' style='color: #94a3b8; font-size: 14px;'>© 2026 QE Engineering | Excellence in Systems v2.5 ✨</p>", unsafe_allow_html=True)

elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
