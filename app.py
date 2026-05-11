import streamlit as st
import pandas as pd
import io
import re
import os
import datetime
from openpyxl import load_workbook

# ==========================================
# 💎 PREMIER LUXURY UI SETTINGS
# ==========================================
st.set_page_config(page_title="QAD System Hub", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@200;400;600&family=Inter:wght@300;500;700&display=swap');

    /* พื้นหลังแบบ Soft Gradient */
    .stApp {
        background: radial-gradient(circle at top right, #fdfcfb 0%, #e2d1c3 100%);
        font-family: 'Inter', 'Kanit', sans-serif;
    }

    /* ตกแต่ง Sidebar ให้ดูแพง */
    [data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.8);
        backdrop-filter: blur(20px);
        border-right: 1px solid rgba(0,0,0,0.05);
    }

    /* Card Container สำหรับเนื้อหา */
    div.block-container {
        padding-top: 2rem;
    }

    /* ปุ่มเมนูหลักสไตล์ Premium */
    div.stButton > button:first-child {
        background: #ffffff;
        color: #0f172a;
        border: 1px solid #e2e8f0;
        border-radius: 24px;
        height: 120px;
        font-size: 24px;
        font-weight: 700;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05);
        transition: all 0.5s cubic-bezier(0.4, 0, 0.2, 1);
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    div.stButton > button:hover {
        background: #1e3a8a;
        color: #ffffff;
        border-color: #1e3a8a;
        transform: translateY(-8px);
        box-shadow: 0 20px 25px -5px rgba(30, 58, 138, 0.25);
    }

    /* ชื่อแอปหลัก */
    .main-title {
        font-size: 64px;
        font-weight: 700;
        background: linear-gradient(to right, #1e3a8a, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0px;
        text-align: center;
    }

    /* หัวข้อภายในแอป */
    h1, h2 {
        font-weight: 600 !important;
        color: #1e293b !important;
        letter-spacing: -0.5px;
    }

    /* ปรับแต่ง Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre;
        background-color: transparent;
        border-radius: 8px;
        color: #64748b;
        font-weight: 500;
    }

    .stTabs [aria-selected="true"] {
        color: #1e3a8a !important;
        border-bottom-color: #1e3a8a !important;
    }

    /* Style สำหรับ Dataframe */
    [data-testid="stDataFrame"] {
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
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
    st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>📁 File Validator</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b;'>ตรวจสอบโครงสร้างและความถูกต้องของไฟล์</p>", unsafe_allow_html=True)
    
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
        model_list = ["-- Select Reference Model --"] + sorted(list(available_models.keys()))
        
        c1, c2 = st.columns(2)
        with c1:
            selected_model_name = st.selectbox("📍 Reference Model", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}")
        
        if selected_model_name != "-- Select Reference Model --":
            with c2:
                uploaded_file = st.file_uploader("📥 Target File", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}")

            if uploaded_file:
                st.write("---")
                with st.status("🔍 Checking File Integrity...", expanded=True) as status:
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
                    status.update(label="Checking Complete!", state="complete", expanded=False)

                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons(); st.success("✨ ข้อมูลถูกต้องสมบูรณ์ 100%")
                else:
                    if f_errors: st.warning("⚠️ หัวตารางไม่ตรง (F3/F5)"); st.table(pd.DataFrame(f_errors))
                    
                    res_c1, res_c2 = st.columns(2)
                    with res_c1:
                        if missing_data: 
                            st.markdown("📉 **Missing Data**")
                            st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                        if d_errors:
                            st.markdown("🔍 **Washing Date Issues**")
                            st.dataframe(pd.DataFrame(d_errors), use_container_width=True)
                    with res_c2:
                        if extra_data:
                            st.markdown("🚫 **Extra Data**")
                            st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                        if k_errors:
                            st.markdown("🔍 **Finish Date Issues**")
                            st.dataframe(pd.DataFrame(k_errors), use_container_width=True)
    except Exception as e: st.error(f"Error: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR
# ==========================================
def app_washing_processor():
    st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>📊 Washing Date Processor</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b;'>ตรวจสอบวันล้าง</p>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🧭 Control Panel")
        if st.button("🏠 Home Menu", use_container_width=True):
            go_to_menu()
        if st.button("🔄 Clear Buffer", use_container_width=True):
            st.session_state.output, st.session_state.summary, st.session_state.file = None, None, None
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            st.rerun()

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        file1 = st.file_uploader("📂 Lot List File (ไฟล์จาก Part Vintage)", type=["xls", "xlsx", "csv"], key=f"p1_{st.session_state.uploader_key}")
    with col_f2:
        file2 = st.file_uploader("📂 Barcode Data File (File download จากเว็บ PDD)", type=["xls", "xlsx", "csv"], key=f"p2_{st.session_state.uploader_key}")

    def read_excel(file):
        try: return pd.read_excel(file, engine="openpyxl", header=None)
        except: return pd.read_excel(file, engine="xlrd", header=None)

    if st.button("🚀 PROCESSING", use_container_width=True):
        if not file1 or not file2:
            st.warning("กรุณาอัปโหลดไฟล์ให้ครบถ้วน")
        else:
            with st.spinner('⚡ รันระบบอัจฉริยะกำลังประมวลผล...'):
                df1_raw = read_excel(file1)
                df1 = pd.DataFrame({"Lot": [str(v).strip() for v in df1_raw.iloc[16:, 5] if not pd.isna(v) and str(v).strip() != ""]})
                df2_raw = read_excel(file2)
                h_row = next((i for i in range(20) if df2_raw.iloc[i].astype(str).str.lower().str.contains("runcard").any()), None)
                if h_row is not None:
                    df2_raw.columns = df2_raw.iloc[h_row]
                    df2 = df2_raw[h_row + 1:].copy()
                    df2.columns = df2.columns.astype(str).str.strip().str.lower()
                    lot_c = [c for c in df2.columns if "runcard" in c][0]
                    bar_c = [c for c in df2.columns if "barcode" in c][0]
                    pak_c = [c for c in df2.columns if "packed" in c and "date" in c][0]
                    df2_final = df2[[lot_c, bar_c, pak_c]].dropna(subset=[lot_c])
                    df2_final.columns = ["Lot", "Barcode No", "Packed Date"]
                    df2_final["Packed Date"] = pd.to_datetime(df2_final["Packed Date"], errors="coerce")
                    merged = pd.merge(df1, df2_final, on="Lot", how="left").drop_duplicates(subset=["Lot"])
                    def ext_ww(b):
                        m = re.search('[A-Za-z]', str(b)); return (int(str(b)[m.start()+3:m.start()+5]), int(str(b)[m.start()+5])) if m and str(b)[m.start()+3:m.start()+6].isdigit() else (None, None)
                    merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(ext_ww(x)))
                    db = pd.read_csv("database.txt"); db["Date"] = pd.to_datetime(db["Date"], format="%d-%b-%Y", errors="coerce")
                    def find_d(r, d_db):
                        if pd.isna(r["WW"]) or pd.isna(r["Day"]) or pd.isna(r["Packed Date"]): return None
                        c = d_db[(d_db["WW"] == r["WW"]) & (d_db["Day"] == r["Day"])].copy()
                        if c.empty: return None
                        c["diff"] = (c["Date"] - r["Packed Date"]).abs()
                        return c.sort_values("diff").iloc[0]["Date"]
                    merged["Washing Date"] = merged.apply(lambda r: find_d(r, db), axis=1)
                    out = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date"]].copy()
                    out["Washing Date"] = pd.to_datetime(out["Washing Date"]).dt.strftime("%d-%b-%Y")
                    sum_df = out.groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        out.to_excel(writer, index=False, sheet_name="Result")
                        sum_df.to_excel(writer, index=False, sheet_name="Summary")
                    st.session_state.output, st.session_state.summary, st.session_state.file = out, sum_df, buf.getvalue()

    if st.session_state.get("output") is not None:
        st.success("✅ ระบบประมวลผลเสร็จสมบูรณ์")
        t1, t2 = st.tabs(["💎 Result Details", "📈 Summary View"])
        with t1: st.dataframe(st.session_state.output, use_container_width=True)
        with t2: st.dataframe(st.session_state.summary, use_container_width=True)
        st.download_button("📥 DOWNLOAD RESULT (.XLSX)", st.session_state.file, f"Result_{datetime.datetime.now().strftime('%H%M')}.xlsx", use_container_width=True)

# ==========================================
# MAIN ROUTING
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<div style='height: 50px;'></div>", unsafe_allow_html=True)
    st.markdown("<p class='main-title'>QAD System Hub</p>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; font-size: 20px; font-weight: 300;'>QAD Support Application</p>", unsafe_allow_html=True)
    st.markdown("<div style='height: 40px;'></div>", unsafe_allow_html=True)
    
    col_l, c1, gap, c2, col_r = st.columns([0.4, 2, 0.3, 2, 0.4])
    
    with c1:
        st.markdown("<h3 style='text-align: center; margin-bottom: 17px;'>ตรวจสอบ format ไฟล์ก่อนส่ง</h3>", unsafe_allow_html=True)
        if st.button("📁 File Validator"):
            st.session_state.current_app = "Validator"; st.rerun()
        st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 14px;'></p>", unsafe_allow_html=True)

    with c2:
        st.markdown("<h3 style='text-align: center; margin-bottom: 17px;'>ตรวจสอบวันล้าง</h3>", unsafe_allow_html=True)
        if st.button("📊 Washing Date Processor"):
            st.session_state.current_app = "Processor"; st.rerun()
        st.markdown("<p style='text-align: center; color: #94a3b8; font-size: 14px;'></p>", unsafe_allow_html=True)

    st.markdown("<div style='height: 100px;'></div>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #cbd5e1; font-size: 12px;'>© 2026 QAD Engineering | System Excellence v2.5</p>", unsafe_allow_html=True)

elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
