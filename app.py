import streamlit as st
import pandas as pd
import io
import re
import os
import datetime
from openpyxl import load_workbook

# ==========================================
# 💎 UI CUSTOMIZATION (ANGULAR LUXURY THEME)
# ==========================================
# ปรับ layout เป็น wide เพื่อความภูมิฐานและใช้พื้นที่ได้เต็มที่
st.set_page_config(page_title="QAD System Hub", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;600&family=Inter:wght@300;500;700&display=swap');

    /* พื้นหลังโทนเข้มแบบผู้บริหาร (Deep Slate) */
    .stApp {
        background-color: #111827;
        color: #f8fafc;
        font-family: 'Inter', 'Kanit', sans-serif;
    }

    /* ตกแต่ง Sidebar ให้ดูเป็นมืออาชีพและคมชัด */
    [data-testid="stSidebar"] {
        background-color: #1f2937;
        border-right: 1px solid #374151;
    }

    /* เนื้อหาหลัก */
    div.block-container {
        padding-top: 2rem;
    }

    /* =========================================
       ปุ่มเมนูหลักสไตล์ ANGULAR LUXURY
       เน้นเหลี่ยมมุม ความโปร่งแสง และเส้นขอบเงิน/ทอง
       ========================================= */
    div.stButton > button:first-child {
        background: rgba(255, 255, 255, 0.03);
        color: #f8fafc;
        border: 1px solid #374151;
        border-radius: 0px; /* บังคับเหลี่ยมเป๊ะ */
        height: 120px;
        font-size: 24px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 2px; /* เพิ่มช่องว่างระหว่างอักษรให้ดูแพง */
        transition: all 0.3s ease-in-out;
        box-shadow: none;
    }

    /* Hover Effect: เปลี่ยนเป็นสีเน้นและขยายเล็กน้อย */
    div.stButton > button:hover {
        background: rgba(59, 130, 246, 0.1);
        color: #60a5fa;
        border: 1px solid #60a5fa;
        transform: scale(1.02);
    }

    /* ปุ่มกลับหน้าหลักใน Sidebar */
    [data-testid="stSidebar"] div.stButton > button {
        height: 50px;
        font-size: 16px;
        background-color: #374151;
        border: 1px solid #4b5563;
    }

    /* ชื่อแอปหลัก - ปรับให้ดูโมเดิร์นแต่ยังเด่น */
    .main-title {
        font-size: 60px;
        font-weight: 700;
        color: #f8fafc;
        margin-bottom: 5px;
        text-align: center;
        letter-spacing: -1px;
    }

    /* คำอธิบายใต้ชื่อแอป */
    .sub-title {
        text-align: center;
        color: #94a3b8;
        font-size: 18px;
        font-weight: 300;
        margin-bottom: 2rem;
    }

    /* หัวข้อภายในแอป */
    h1, h2, h3 {
        font-weight: 600 !important;
        color: #f1f5f9 !important;
        letter-spacing: -0.5px;
        border-radius: 0px !important;
    }

    /* ปรับแต่ง Tabs ให้เป็นเหลี่ยม */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px;
        background-color: #1f2937;
        border: 1px solid #374151;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        background-color: transparent;
        border-radius: 0px;
        color: #94a3b8;
        font-weight: 500;
        border-right: 1px solid #374151;
        padding: 0 30px;
    }

    /* Tab ที่เลือก */
    .stTabs [aria-selected="true"] {
        background-color: #374151;
        color: #f8fafc !important;
        border-bottom: 2px solid #60a5fa !important;
    }

    /* ปรับแต่ง DataFrame และ Table ให้ขอบคม */
    [data-testid="stDataFrame"], .styled-table, table {
        border-radius: 0px !important;
        border: 1px solid #374151 !important;
        background-color: #1f2937;
    }

    /* ปรับแต่งสี Alert/Status */
    .stAlert {
        border-radius: 0px;
        background-color: #1f2937;
        border: 1px solid #374151;
    }
    
    /* สีทองหม่นสำหรับ Success */
    .stAlert:has(div[class*="st-b"]) {
        border-color: #eab308;
        color: #eab308;
    }
    
    /* สีฟ้าสำหรับ Info */
    .stAlert:has(div[class*="st-c"]) {
        border-color: #3b82f6;
        color: #3b82f6;
    }

    /* ปรับแต่ง File Uploader */
    .stFileUploader section {
        background-color: #1f2937;
        border: 1px solid #374151;
        border-radius: 0px;
    }
    </style>
""", unsafe_allow_html=True)

# 管理 Session State
if "current_app" not in st.session_state:
    st.session_state.current_app = "Main Menu"

# ฟังก์ชันกลับหน้าหลัก
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
    st.markdown("<h2 style='text-align: center; border-bottom: 2px solid #374151; padding-bottom: 15px;'>📁 File Validator</h2>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🛠️ Control Panel")
        if st.button("⬅️ Back to Menu"):
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
            st.markdown("#### 1. เลือกโมเดลอ้างอิง")
            selected_model_name = st.selectbox("Select Model", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}", label_visibility="collapsed")
        
        if selected_model_name != "-- Please Select Model --":
            with col2:
                st.markdown("#### 2. อัปโหลดไฟล์ที่ต้องการตรวจ")
                uploaded_file = st.file_uploader("Upload", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}", label_visibility="collapsed")

            if uploaded_file:
                st.divider()
                with st.spinner('กำลังตรวจสอบความถูกต้อง...'):
                    wb = load_workbook(uploaded_file, data_only=False)
                    ws = wb[TARGET_SHEET]
                    df_ref = pd.read_excel(available_models[selected_model_name], sheet_name=TARGET_SHEET, header=None).fillna("")
                    df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
                    
                    f_errors, missing_data, extra_data, d_errors, k_errors = [], {}, {}, [], []

                    # 1. ตรวจ Par no./ Drawing no. (ช่อง F3/F5)
                    for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                        if str(df_user.iloc[r, c]).strip() != str(df_ref.iloc[r, c]).strip():
                            f_errors.append({"Position": label, "Found": df_user.iloc[r, c], "Target": df_ref.iloc[r, c]})

                    # 2. ตรวจ Missing/Extra Data ในตารางหลัก
                    for r in range(76):
                        for c in range(df_ref.shape[1]):
                            # ยกเว้นคอลัมน์ D และ K ที่เป็นสูตร/วันที่
                            if r >= 12 and c in [3, 10]: continue
                            
                            ref_v = str(df_ref.iloc[r, c]).strip()
                            user_v = str(df_user.iloc[r, c]).strip()
                            
                            # ถ้าต้นฉบับมี แต่ไฟล์แนบไม่มี -> Missing
                            if ref_v != "" and (user_v == "" or user_v == "nan"):
                                missing_data.setdefault(get_column_letter(c), []).append(str(r+1))
                            # ถ้าต้นฉบับว่าง แต่ไฟล์แนบมี (ยกเว้นแถว 12) -> Extra
                            elif ref_v == "" and (user_v != "" and user_v != "nan") and r+1 != 12:
                                extra_data.setdefault(get_column_letter(c), []).append(str(r+1))

                    # 3. ตรวจรูปแบบวันที่และเวลาคอลัมน์ D และ K
                    for row_idx in range(13, 77): 
                        for col_idx, (col_label, error_list) in enumerate(zip(['D', 'K'], [d_errors, k_errors])):
                            target_col = 4 if col_label == 'D' else 11
                            cell = ws.cell(row=row_idx, column=target_col)
                            
                            if cell.value is not None:
                                # ตรวจสอบรูปแบบวันที่และเวลาผ่าน number_format
                                fmt = str(cell.number_format).lower()
                                # ต้องมีอย่างน้อย y (ปี) หรือ d (วัน) m (เดือน) และ h (ชั่วโมง)
                                has_date = ('y' in fmt) or ('d' in fmt and 'm' in fmt)
                                has_time = ('h' in fmt)
                                
                                if not (has_date and has_time):
                                    error_list.append({"Row": row_idx, "Format": fmt, "Status": "❌ รูปแบบผิด (ต้องมีทั้งวันที่และเวลา)"})

                st.markdown("### 📋 ผลการตรวจสอบ")
                # หากไม่พบข้อผิดพลาดเลย
                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.success("✅ ยอดเยี่ยม! ข้อมูลและรูปแบบถูกต้องทั้งหมด")
                else:
                    # แสดงข้อผิดพลาดตามหมวดหมู่
                    if f_errors: st.warning("⚠️ Par no. / Drawing no. ไม่ถูก (ช่อง F3/F5)"); st.table(pd.DataFrame(f_errors))
                    if missing_data: st.info("⚠️ พบข้อมูลที่หายไป (Missing Data)"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    if extra_data: st.error("🚫 พบข้อมูลเกินมาในช่องที่ควรว่าง (Extra Data)"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                    
                    res_c1, res_c2 = st.columns(2)
                    with res_c1:
                        st.markdown("🔍 **ตรวจสอบคอลัมน์ D (Washing Date)**")
                        if d_errors: st.dataframe(pd.DataFrame(d_errors), use_container_width=True)
                        else: st.write("✅ ปกติ")
                    with res_c2:
                        st.markdown("🔍 **ตรวจสอบคอลัมน์ K (Finish Date)**")
                        if k_errors: st.dataframe(pd.DataFrame(k_errors), use_container_width=True)
                        else: st.write("✅ ปกติ")
    except Exception as e: st.error(f"Error: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR
# ==========================================
def app_washing_processor():
    st.markdown("<h2 style='text-align: center; border-bottom: 2px solid #374151; padding-bottom: 15px;'>📊 Washing Date Processor</h2>", unsafe_allow_html=True)
    
    with st.sidebar:
        st.markdown("### 🛠️ Control Panel")
        if st.button("⬅️ Back to Menu"):
            go_to_menu()
        st.divider()
        if st.button("🔄 Reset This App"):
            # ล้างผลลัพธ์ใน session state
            st.session_state.output = None
            st.session_state.summary = None
            st.session_state.file = None
            # เปลี่ยน uploader_key เพื่อล้างไฟล์ที่อัปโหลดค้างไว้
            st.session_state.uploader_key = st.session_state.get('uploader_key', 0) + 1
            st.rerun()

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        file1 = st.file_uploader("📂 Upload File 1 (Lot/Serial)", type=["xls", "xlsx", "csv"], key=f"file1_{st.session_state.uploader_key}")
    with col_u2:
        file2 = st.file_uploader("📂 Upload File 2 (Barcode)", type=["xls", "xlsx", "csv"], key=f"file2_{st.session_state.uploader_key}")

    def read_excel(file):
        try: return pd.read_excel(file, engine="openpyxl", header=None)
        except: return pd.read_excel(file, engine="xlrd", header=None)

    # ฟังก์ชันอ่านไฟล์ 1 (Lot List)
    def read_file1(file):
        df = read_excel(file)
        # อ่านคอลัมน์ F (index 5) ตั้งแต่แถวที่ 17 (index 16)
        data = df.iloc[16:, 5]
        lot_list = []
        for val in data:
            if pd.isna(val) or str(val).strip() == "": break # หยุดอ่านเมื่อเจอช่องว่าง
            lot_list.append(str(val).strip())
        return pd.DataFrame({"Lot": lot_list})

    # ฟังก์ชันอ่านไฟล์ 2 (Barcode Data)
    def read_file2(file):
        df = read_excel(file)
        header_row = next((i for i in range(20) if df.iloc[i].astype(str).str.lower().str.contains("runcard").any()), None)
        if header_row is None:
            return None
        df.columns = df.iloc[header_row]
        df = df[header_row + 1:].copy()
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        lot_col = [c for c in df.columns if "runcard" in c][0]
        barcode_col = [c for c in df.columns if "barcode" in c][0]
        packed_col = [c for c in df.columns if "packed" in c and "date" in c][0]
        
        df_out = df[[lot_col, barcode_col, packed_col]].dropna(subset=[lot_col])
        df_out.columns = ["Lot", "Barcode No", "Packed Date"]
        df_out["Packed Date"] = pd.to_datetime(df_out["Packed Date"], errors="coerce")
        return df_out

    if st.button("🚀 ประมวลผลข้อมูล", use_container_width=True):
        if not file1 or not file2:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบทั้ง 2 ช่อง")
        else:
            with st.spinner('กำลังคำนวณ...'):
                df1 = read_file1(file1)
                df2 = read_file2(file2)
                
                if df2 is not None:
                    # จับคู่ข้อมูล
                    merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
                    
                    # ดึง WW และ Day
                    def ext_ww(b):
                        m = re.search('[A-Za-z]', str(b))
                        if not m: return None, None
                        c = str(b)[m.start()+3:m.start()+6]
                        return (int(c[:2]), int(c[2])) if (len(c)==3 and c.isdigit()) else (None, None)
                    
                    merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(ext_ww(x)))
                    
                    # โหลด Database
                    date_db = pd.read_csv("database.txt")
                    date_db["Date"] = pd.to_datetime(date_db["Date"], format="%d-%b-%Y", errors="coerce")
                    
                    # เทียบ WW/Day เพื่อหาวันที่ล้างสินค้า
                    def find_d(r, d_db):
                        if pd.isna(r["WW"]) or pd.isna(r["Day"]) or pd.isna(r["Packed Date"]): return None
                        c = d_db[(d_db["WW"] == r["WW"]) & (d_db["Day"] == r["Day"])].copy()
                        if c.empty: return None
                        c["diff"] = (c["Date"] - r["Packed Date"]).abs()
                        return c.sort_values("diff").iloc[0]["Date"]

                    merged["Washing Date"] = merged.apply(lambda r: find_d(r, date_db), axis=1)
                    
                    # เตรียมผลลัพธ์
                    out = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date"]].copy()
                    out["Washing Date"] = pd.to_datetime(out["Washing Date"]).dt.strftime("%d-%b-%Y")
                    sum_df = out.groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})
                    
                    # บันทึกลง Buffer เพื่อรอการดาวน์โหลด
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                        out.to_excel(writer, index=False, sheet_name="Result")
                        sum_df.to_excel(writer, index=False, sheet_name="Summary")
                    
                    st.session_state.output, st.session_state.summary, st.session_state.file = out, sum_df, buf.getvalue()

    if st.session_state.get("output") is not None:
        st.success("✅ ประมวลผลสำเร็จ!")
        tab1, tab2 = st.tabs(["📄 รายละเอียดข้อมูล", "📊 สรุปผล"])
        with tab1: st.dataframe(st.session_state.output, use_container_width=True)
        with tab2: st.dataframe(st.session_state.summary, use_container_width=True)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (Result)",
            data=st.session_state.file,
            file_name=f"result_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ==========================================
# MAIN ROUTING
# ==========================================
# หน้าเมนูหลัก
if st.session_state.current_app == "Main Menu":
    st.markdown("<div style='height: 50px;'></div>", unsafe_allow_html=True)
    st.markdown("<p class='main-title'>QAD System Hub</p>", unsafe_allow_html=True)
    st.markdown("<p class='sub-title'>QAD Support Application</p>", unsafe_allow_html=True)
    st.markdown("<div style='height: 40px;'></div>", unsafe_allow_html=True)
    
    # วางปุ่มใน Layout ที่สมดุล
    empty_col_left, c1, col_gap, c2, empty_col_right = st.columns([1, 4, 0.5, 4, 1])
    
    with c1:
        st.markdown("### ตรวจสอบ format ไฟล์ก่อนส่ง")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("📁 File Validator"):
            st.session_state.current_app = "Validator"
            st.rerun()

    with c2:
        st.markdown("### ตรวจสอบวันล้าง")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("📊 Washing Date Processor"):
            st.session_state.current_app = "Processor"
            st.rerun()

# หน้าแอป File Validator
elif st.session_state.current_app == "Validator":
    app_file_validator()

# หน้าแอป Washing Date Processor
elif st.session_state.current_app == "Processor":
    app_washing_processor()
