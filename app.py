import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบบีบตรงกลาง
st.set_page_config(page_title="File Validator", layout="centered")

# --- ฟังก์ชัน Reset ข้อมูล ---
def reset_app():
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

# ฟังก์ชันแปลง Index เป็นชื่อคอลัมน์ Excel
def get_column_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

# สแกนหาไฟล์โมเดล
def get_available_models():
    files = [f for f in os.listdir('.') if f.startswith('reference_') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
    models = {f.replace('reference_', '').split('.')[0]: f for f in files}
    return models

# --- ส่วนหัวโปรแกรม ---
st.title("📁 File Validator Before Upload")
if st.button("🔄 Reset Application"):
    reset_app()

st.divider()

TARGET_SHEET = "RAMP v1.3"

try:
    available_models = get_available_models()
    
    if not available_models:
        st.error("❌ No reference files found. (Filename must be 'reference_ModelName.xlsx')")
        st.stop()

    # 1. เลือกโมเดล
    model_names = sorted(list(available_models.keys()))
    selected_model_name = st.selectbox("1️⃣ Select Model:", ["-- Please Select --"] + model_names)

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        
        # โหลดไฟล์ Reference
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        # ดึงค่าเกณฑ์มาตรฐาน
        correct_part_no = str(df_ref.iloc[2, 5]).strip()
        correct_dwg_no = str(df_ref.iloc[4, 5]).strip()

        # 2. อัปโหลดไฟล์
        uploaded_file = st.file_uploader(f"2️⃣ Upload File to Check (Sheet: {TARGET_SHEET})", type=["xlsx", "xlsm"])

        if uploaded_file:
            try:
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            except ValueError:
                st.error(f"❌ Sheet '{TARGET_SHEET}' not found in uploaded file.")
                st.stop()

            # --- Validation Logic ---
            f_errors = [] 
            missing_data = {} 
            date_errors = [] 

            max_row = min(len(df_ref), len(df_user), 76)
            max_col = min(df_ref.shape[1], df_user.shape[1])

            # Check F3 & F5
            user_part_no = str(df_user.iloc[2, 5]).strip()
            user_dwg_no = str(df_user.iloc[4, 5]).strip()

            if user_part_no != correct_part_no:
                f_errors.append({"Position": "F3 (Part No.)", "Found": user_part_no, "Target": correct_part_no})
            if user_dwg_no != correct_dwg_no:
                f_errors.append({"Position": "F5 (Dwg No.)", "Found": user_dwg_no, "Target": correct_dwg_no})

            # Check Missing Data
            for r in range(max_row):
                for c in range(max_col):
                    ref_val = df_ref.iloc[r, c]
                    if pd.notna(ref_val) and str(ref_val).strip() != "":
                        user_val = df_user.iloc[r, c]
                        if pd.isna(user_val) or str(user_val).strip() == "":
                            col_name = get_column_letter(c)
                            if col_name not in missing_data:
                                missing_data[col_name] = []
                            missing_data[col_name].append(str(r + 1))

            # Check Date Format (D, K)
            date_format = "%m/%d/%Y %H:%M:%S"
            for row_idx in range(12, 76):
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < len(df_user):
                        val = df_user.iloc[row_idx, col_idx]
                        if pd.isna(val) or str(val).strip() == "":
                            date_errors.append({"Column": col_label, "Row": row_idx + 1, "Issue": "Empty Cell"})
                        elif not isinstance(val, datetime):
                            check_str = str(val).strip()
                            if check_str != "00:00:00":
                                try:
                                    datetime.strptime(check_str, date_format)
                                except ValueError:
                                    date_errors.append({"Column": col_label, "Row": row_idx + 1, "Issue": "Invalid Date Format"})

            # --- Display Results ---
            st.divider()
            
            if not f_errors and not missing_data and not date_errors:
                st.balloons()
                st.success(f"✅ Everything looks good for model: {selected_model_name}")
            else:
                # 1. Main Info Match
                if f_errors:
                    st.warning("⚠️ Critical Info Mismatch")
                    st.table(pd.DataFrame(f_errors))

                # 2. Missing Data Table
                if missing_data:
                    st.warning("⚠️ Missing Data Summary")
                    table_data = []
                    for col, rows in missing_data.items():
                        table_data.append({
                            "Column": col,
                            "Count": len(rows),
                            "Affected Rows": ", ".join(rows)
                        })
                    st.table(pd.DataFrame(table_data))

                # 3. Date Errors Expander
                if date_errors:
                    with st.expander("🔍 View Date Format Issues (Col D & K)"):
                        st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"Unexpected Error: {e}")
