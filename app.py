import streamlit as st
import pandas as pd
import io
import re
import os
import datetime
from openpyxl import load_workbook

# ==========================================
# 🎨 LUXURY & CUTE UI CUSTOMIZATION
# ==========================================
st.set_page_config(page_title="QAD System Hub", layout="wide") # ปรับเป็น Wide เพื่อความโปร่ง

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;600&display=swap');

    /* พื้นหลังแบบ Gradient ไล่สีนุ่มๆ */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        font-family: 'Kanit', sans-serif;
    }

    /* ปรับแต่ง Container หลักให้ดูเหมือนกระจก */
    [data-testid="stVerticalBlock"] > div:has(div.stButton) {
        background: rgba(255, 255, 255, 0.7);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        padding: 40px;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.1);
        border: 1px solid rgba(255, 255, 255, 0.18);
    }

    /* ปุ่มเมนูหลัก - สไตล์ Luxury */
    div.stButton > button:first-child {
        width: 100%;
        height: 100px;
        border-radius: 18px;
        border: none;
        background: white;
        color: #1e3a8a;
        font-size: 22px;
        font-weight: 600;
        letter-spacing: 0.5px;
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        display: flex;
        align-items: center;
        justify-content: center;
    }

    div.stButton > button:hover {
        background: #1e3a8a;
        color: white;
        transform: translateY(-5px) scale(1.02);
        box-shadow: 0 12px 20px rgba(30, 58, 138, 0.2);
    }

    /* ตกแต่ง Sidebar */
    [data-testid="stSidebar"] {
        background-color: rgba(255, 255, 255, 0.5);
        backdrop-filter: blur(10px);
        border-right: 1px solid rgba(0,0,0,0.05);
    }

    /* หัวข้อภาษาไทยให้ดูเด่นขึ้น */
    h1 {
        font-weight: 600;
        color: #1e3a8a !important;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 30px !important;
    }
    
    h2, h3 {
        color: #334155 !important;
        font-weight: 400;
    }

    /* ปรับแต่ง File Uploader */
    .stFileUploader section {
        background-color: white;
        border-radius: 15px;
        border: 2px dashed #cbd5e1;
        padding: 20px;
    }

    /* สไตล์ตารางให้ดูสะอาด */
    .styled-table {
        border-radius: 10px;
        overflow: hidden;
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
# APP 1: FILE VALIDATOR (Smart DateTime Check)
# ==========================================
def app_file_validator():
    st.markdown("<h1 style='text-align: center;'>📁 File Validator</h1>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🛠️ Navigation")
        if st.button("⬅️ Back to Home"):
            go_to_menu()
        st.divider()
        if st.button("🔄 Reset This App"):
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
        model_list = ["-- Please Select Model --"] + sorted(list(available_models.keys()))
        
        # จัดระเบียบ Layout การเลือกไฟล์
        col1, col2 = st.columns([1, 1])
        with col1:
            selected_model_name = st.selectbox("🎯 1. เลือกโมเดลอ้างอิง", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}")
        
        if selected_model_name != "-- Please Select Model --":
            with col2:
                uploaded_file = st.file_uploader("📥 2. อัปโหลดไฟล์ที่ต้องการตรวจ", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}")

            if uploaded_file:
                st.divider()
                with st.spinner('กำลังตรวจสอบความถูกต้อง...'):
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

                    st.markdown("### 📋 ผลการตรวจสอบ")
                    if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                        st.balloons(); st.success("🎉 ยอดเยี่ยม! ข้อมูลและรูปแบบถูกต้องทั้งหมด")
                    else:
                        if f_errors: st.warning("⚠️ ข้อมูลหัวตาราง (F3/F5) ไม่ตรงกับต้นฉบับ"); st.table(pd.DataFrame(f_errors))
                        if missing_data: st.info("⚠️ ข้อมูลหายไปในบางช่อง (Missing Data)"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                        if extra_data: st.error("🚫 พบข้อมูลเกินมาในช่องที่ควรว่าง (Extra Data)"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                        
                        res_c1, res_c2 = st.columns(2)
                        with res_c1:
                            st.markdown("🔍 **Washing Date (Col D)**")
                            if d_errors: st.dataframe(pd.DataFrame(d_errors), use_container_width=True)
                            else: st.write("✅ สมบูรณ์")
                        with res_c2:
                            st.markdown("🔍 **Finish Date (Col K)**")
                            if k_errors: st.dataframe(pd.DataFrame(k_errors), use_container_width=True)
                            else: st.write("✅ สมบูรณ์")
    except Exception as e: st.error(f"ระบบขัดข้อง: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR
# ==========================================
def app_washing_processor():
    st.markdown("<h1 style='text-align: center;'>📊 Washing Date Processor</h1>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🛠️ Navigation")
        if st.button("⬅️ Back to Home"):
            go_to_menu()
        st.divider()
        if st.button("🔄 Reset Process"):
            st.session_state.output, st.session_state.summary, st.session_state.file = None, None, None
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            st.rerun()

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        file1 = st.file_uploader("📂 อัปโหลด File 1 (Lot List)", type=["xls", "xlsx", "csv"], key=f"p1_{st.session_state.uploader_key}")
    with col_f2:
        file2 = st.file_uploader("📂 อัปโหลด File 2 (Barcode Data)", type=["xls", "xlsx", "csv"], key=f"p2_{st.session_state.uploader_key}")

    def read_excel(file):
        try: return pd.read_excel(file, engine="openpyxl", header=None)
        except: return pd.read_excel(file, engine="xlrd", header=None)

    if st.button("🚀 เริ่มประมวลผลข้อมูล"):
        if not file1 or not file2:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ไฟล์ก่อนเริ่มงาน")
        else:
            with st.spinner('กำลังจับคู่ข้อมูลและคำนวณวันที่...'):
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
                        m = re.search('[A-Za-z]', str(b))
                        if not m: return None, None
                        c = str(b)[m.start()+3:m.start()+6]
                        return (int(c[:2]), int(c[2])) if (len(c)==3 and c.isdigit()) else (None, None)
                    
                    merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(ext_ww(x)))
                    db = pd.read_csv("database.txt")
                    db["Date"] = pd.to_datetime(db["Date"], format="%d-%b-%Y", errors="coerce")
                    
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
        st.success("✅ ประมวลผลเสร็จสิ้น!")
        tab1, tab2 = st.tabs(["📄 รายละเอียดข้อมูล", "📊 สรุปผล"])
        with tab1:
            st.dataframe(st.session_state.output, use_container_width=True)
        with tab2:
            st.dataframe(st.session_state.summary, use_container_width=True)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (Result)",
            data=st.session_state.file,
            file_name=f"Result_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==========================================
# MAIN ROUTING
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<h1 style='text-align: center; font-size: 50px; margin-bottom: 10px;'>🏭 QAD System Hub</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; font-size: 18px;'>Support Application for QAD</p>", unsafe_allow_html=True)
    st.write("---")
    
    # วางปุ่มใน Layout ที่สวยงาม
    empty_col_left, c1, col_gap, c2, empty_col_right = st.columns([0.5, 2, 0.2, 2, 0.5])
    
    with c1:
        st.markdown("### 📂 Data Audit")
        if st.button("📁 File Validator\n(ตรวจสอบรูปแบบไฟล์)"):
            st.session_state.current_app = "Validator"; st.rerun()
        st.caption("ใช้สำหรับตรวจเช็คความถูกต้องของ Format ก่อนส่งงาน")

    with c2:
        st.markdown("### 📈 Data Processing")
        if st.button("📊 Washing Date\nProcessor"):
            st.session_state.current_app = "Processor"; st.rerun()
        st.caption("ระบบดึงข้อมูลวันที่ล้างจาก Lot")

    st.markdown("<br><br><p style='text-align: center; color: #94a3b8;'>Production Tools v2.0 • Premium Interface</p>", unsafe_allow_html=True)

elif st.session_state.current_app == "Validator": 
    app_file_validator()
elif st.session_state.current_app == "Processor": 
    app_washing_processor()
