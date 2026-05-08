import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="File Validator", layout="centered")

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
    models = {f.replace('reference_', '').split('.')[0]: f for f in files}
    return models

st.title("📁 File Validator Before Upload")
if st.button("🔄 Reset Application"):
    reset_app()

st.divider()

TARGET_SHEET = "RAMP v1.3"

try:
    available_models = get_available_models()
    if not available_models:
        st.error("❌ No reference files found.")
        st.stop()

    model_names = sorted(list(available_models.keys()))
    selected_model_name = st.selectbox("1️⃣ Select Model:", ["-- Please Select --"] + model_names)

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        correct_part_no = str(df_ref.iloc[2, 5]).strip()
        correct_dwg_no = str(df_ref.iloc[4, 5]).strip()

        uploaded_file = st.file_uploader(f"2️⃣ Upload File (Target: {selected_model_name})", type=["xlsx", "xlsm"])

        if uploaded_file:
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            
            f_errors = [] 
            missing_data = {} 
            extra_data = {}  
            date_errors = [] 

            # ปรับปรุงขอบเขตการเช็ค (เริ่มเช็คตั้งแต่แถว 1 จนถึงแถวสุดท้ายที่มีข้อมูล)
            user_rows, user_cols = df_user.shape
            ref_rows, ref_cols = df_ref.shape
            max_r = max(ref_rows, user_rows)
            max_c = max(ref_cols, user_cols)

            # 1. Check F3 & F5
            user_part_no = str(df_user.iloc[2, 5]).strip() if user_rows > 2 else ""
            user_dwg_no = str(df_user.iloc[4, 5]).strip() if user_rows > 4 else ""

            if user_part_no != correct_part_no:
                f_errors.append({"Position": "F3", "Found": user_part_no, "Target": correct_part_no})
            if user_dwg_no != correct_dwg_no:
                f_errors.append({"Position": "F5", "Found": user_dwg_no, "Target": correct_dwg_no})

            # 2. Check All Cells (Missing & Extra)
            for r in range(max_r):
                for c in range(max_c):
                    ref_val = df_ref.iloc[r, c] if (r < ref_rows and c < ref_cols) else pd.NA
                    user_val = df_user.iloc[r, c] if (r < user_rows and c < user_cols) else pd.NA
                    
                    is_ref_empty = pd.isna(ref_val) or str(ref_val).strip() == ""
                    is_user_empty = pd.isna(user_val) or str(user_val).strip() == ""
                    
                    col_name = get_column_letter(c)
                    cell_pos = f"{col_name}{r+1}"

                    # Missing: ต้นฉบับมี แต่เราไม่มี
                    if not is_ref_empty and is_user_empty:
                        if col_name not in missing_data: missing_data[col_name] = []
                        missing_data[col_name].append(str(r+1))
                    
                    # Extra: ต้นฉบับไม่มี แต่เราดันมี (เช็คเน้น F15, K15)
                    elif is_ref_empty and not is_user_empty:
                        if col_name not in extra_data: extra_data[col_name] = []
                        extra_data[col_name].append(str(r+1))

            # 3. Check Date Format (ปรับให้เช็คตั้งแต่แถว 12 เป็นต้นไป)
            # รูปแบบที่บังคับ: M/D/YYYY HH:MM:SS
            date_format = "%m/%d/%Y %H:%M:%S"
            for row_idx in range(11, 76): # เปลี่ยนเริ่มที่ 11 (แถว 12)
                for col_idx, col_label in zip([3, 10], ['D', 'K']): # D=3, K=10
                    if row_idx < user_rows:
                        val = df_user.iloc[row_idx, col_idx]
                        if not pd.isna(val) and str(val).strip() != "":
                            # ถ้าไม่ใช่ Object วันที่จาก Excel หรือไม่มีเวลาต่อท้าย ให้ Alarm
                            check_str = str(val).strip()
                            try:
                                # ลองตรวจสอบว่ามีรูปแบบ HH:MM:SS ไหม
                                if len(check_str.split()) < 2 and check_str != "00:00:00":
                                    raise ValueError
                                datetime.strptime(check_str, date_format)
                            except ValueError:
                                if not isinstance(val, datetime):
                                    date_errors.append({"Column": col_label, "Row": row_idx + 1, "Value": check_str, "Issue": "Must be M/D/YYYY HH:MM:SS"})

            # --- Display ---
            st.divider()
            if not f_errors and not missing_data and not extra_data and not date_errors:
                st.balloons()
                st.success(f"✅ All checks passed for {selected_model_name}")
            else:
                if f_errors:
                    st.warning("⚠️ Part No. / Dwg No. Mismatch")
                    st.table(pd.DataFrame(f_errors))

                if missing_data:
                    st.warning("⚠️ Missing Data")
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])

                if extra_data:
                    st.error("🚫 Unexpected Data Found (Reference is empty here)")
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])

                if date_errors:
                    with st.expander("🔍 Date Format Details (Required: M/D/YYYY HH:MM:SS)"):
                        st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"Error: {e}")
