import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

st.set_page_config(page_title="File Validator Pro", layout="centered")

def reset_app():
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

def get_column_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

def get_available_models():
    files = [f for f in os.listdir('.') if f.startswith('reference_') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
    return {f.replace('reference_', '').split('.')[0]: f for f in files}

st.title("📁 File Validator")
if st.button("🔄 Reset Application"):
    reset_app()

st.divider()
TARGET_SHEET = "RAMP v1.3"

try:
    available_models = get_available_models()
    selected_model_name = st.selectbox("1️⃣ Select Model:", ["-- Please Select --"] + sorted(list(available_models.keys())))

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            # --- ส่วนการอ่านข้อมูล ---
            # ใช้ openpyxl สำหรับเช็ค Format (D, K)
            wb = load_workbook(uploaded_file, data_only=False)
            ws = wb[TARGET_SHEET]
            
            # ใช้ Pandas สำหรับเช็ค Missing/Extra (ข้อมูลทั่วไป)
            df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []

            # --- 1. Check F3 / F5 (ข้อมูลส่วนหัว) ---
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # --- 2. Check Missing/Extra Data (หาช่องที่หายหรือเกิน) ---
            # วนลูปเช็คถึงแถว 76 (ยกเว้น D และ K ในช่วงข้อมูลที่จะตรวจ Format แยก)
            for r in range(76):
                for c in range(df_ref.shape[1]):
                    if r >= 12 and c in [3, 10]: continue # ข้าม D(3), K(10) ไปตรวจในข้อ 3
                    
                    ref_v = str(df_ref.iloc[r, c]).strip() if r < df_ref.shape[0] else ""
                    user_v = str(df_user.iloc[r, c]).strip() if r < df_user.shape[0] else ""
                    
                    # ข้อมูลหาย
                    if ref_v != "" and (user_v == "" or user_v == "nan"):
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    
                    # ข้อมูลเกิน
                    elif ref_v == "" and (user_v != "" and user_v != "nan"):
                        if r+1 == 12: continue # ข้ามหัวตารางแถว 12
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # --- 3. Strict Format Check (คอลัมน์ D และ K) ---
            for row_idx in range(13, 77): 
                for col_idx, col_label in zip([4, 11], ['D', 'K']): # openpyxl เริ่มนับที่ 1
                    cell = ws.cell(row=row_idx, column=col_idx)
                    val = cell.value
                    fmt = str(cell.number_format)
                    
                    if val is not None:
                        # เช็คว่า Format มีการตั้งค่าเวลา (h) หรือไม่
                        # ถ้าเป็น Date ธรรมดาใน Excel จะไม่มีตัว h ใน number_format
                        if 'h' not in fmt.lower():
                            date_errors.append({
                                "Column": col_label,
                                "Row": row_idx,
                                "Current Format": fmt,
                                "Issue": "Format เป็น Date (ขาดข้อมูลเวลา)"
                            })

            # --- แสดงผลลัพธ์การตรวจสอบ ---
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons()
                st.success("✅ ข้อมูลและรูปแบบถูกต้องทั้งหมด!")
            else:
                if f_errors:
                    st.warning("⚠️ ข้อมูลส่วนหัวไม่ตรง (F3/F5)")
                    st.table(pd.DataFrame(f_errors))
                
                if missing_data:
                    st.warning("⚠️ พบช่องที่ข้อมูลขาดหายไป (Missing)")
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                
                if extra_data:
                    st.error("🚫 พบข้อมูลเกินจากต้นฉบับ (Extra)")
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                
                if date_errors:
                    st.error("⏰ พบข้อผิดพลาดในคอลัมน์ D หรือ K (Format ไม่ใช่ m/d/yyyy hh:mm:ss)")
                    st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")
