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
    return {f.replace('reference_', '').split('.')[0]: f for f in files}

st.title("📁 File Validator Before Upload")
if st.button("🔄 Reset Application"):
    reset_app()

st.divider()
TARGET_SHEET = "RAMP v1.3"

try:
    available_models = get_available_models()
    selected_model_name = st.selectbox("1️⃣ Select Model:", ["-- Please Select --"] + sorted(list(available_models.keys())))

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File (Target: {selected_model_name})", type=["xlsx", "xlsm"])

        if uploaded_file:
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []
            user_rows, user_cols = df_user.shape
            ref_rows, ref_cols = df_ref.shape

            # 1. Check F3 & F5
            correct_f3 = str(df_ref.iloc[2, 5]).strip()
            correct_f5 = str(df_ref.iloc[4, 5]).strip()
            user_f3 = str(df_user.iloc[2, 5]).strip() if user_rows > 2 else ""
            user_f5 = str(df_user.iloc[4, 5]).strip() if user_rows > 4 else ""

            if user_f3 != correct_f3: f_errors.append({"Position": "F3", "Found": user_f3, "Target": correct_f3})
            if user_f5 != correct_f5: f_errors.append({"Position": "F5", "Found": user_f5, "Target": correct_f5})

            # 2. Loop Check All Cells (Missing & Extra)
            max_r = max(ref_rows, user_rows)
            max_c = max(ref_cols, user_cols)

            for r in range(max_r):
                for c in range(max_c):
                    ref_val = df_ref.iloc[r, c] if (r < ref_rows and c < ref_cols) else pd.NA
                    user_val = df_user.iloc[r, c] if (r < user_rows and c < user_cols) else pd.NA
                    is_ref_empty = pd.isna(ref_val) or str(ref_val).strip() == ""
                    is_user_empty = pd.isna(user_val) or str(user_val).strip() == ""
                    col_name = get_column_letter(c)

                    if not is_ref_empty and is_user_empty:
                        if col_name not in missing_data: missing_data[col_name] = []
                        missing_data[col_name].append(str(r+1))
                    elif is_ref_empty and not is_user_empty:
                        if col_name not in extra_data: extra_data[col_name] = []
                        extra_data[col_name].append(str(r+1))

            # 3. ตรวจสอบวันที่ (D, K แถว 12-76) - แก้ไขใหม่ให้ไม่ Alarm มั่ว
            date_formats = ["%m/%d/%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"]
            
            for row_idx in range(11, 76): 
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < user_rows:
                        val = df_user.iloc[row_idx, col_idx]
                        
                        if not pd.isna(val) and str(val).strip() != "":
                            # ถ้าเป็น datetime object จาก Excel อยู่แล้ว (มักจะมีเวลาพ่วงมาด้วย) ให้ข้ามการเช็คไปเลย
                            if isinstance(val, datetime):
                                continue
                            
                            raw_val = str(val).strip()
                            if raw_val == "00:00:00": continue
                            
                            # ตรวจสอบเบื้องต้น: ต้องมี "เวลา" พ่วงท้าย (เช็คว่ามี Space หรือ Colon ไหม)
                            if " " not in raw_val and ":" not in raw_val:
                                date_errors.append({"Column": col_label, "Row": row_idx + 1, "Value": raw_val, "Issue": "Missing Time (HH:MM:SS)"})
                                continue
                            
                            # ลอง Parse หลายๆ Format เผื่อกรณีสลับ วัน/เดือน
                            valid_format = False
                            for fmt in date_formats:
                                try:
                                    datetime.strptime(raw_val, fmt)
                                    valid_format = True
                                    break
                                except ValueError:
                                    continue
                            
                            if not valid_format:
                                date_errors.append({"Column": col_label, "Row": row_idx + 1, "Value": raw_val, "Issue": "Invalid Format (Use M/D/YYYY HH:MM:SS)"})

            # --- Display Results ---
            st.divider()
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ Everything is correct!")
            else:
                if f_errors:
                    st.warning("⚠️ Critical Info Mismatch"); st.table(pd.DataFrame(f_errors))
                if missing_data:
                    st.warning("⚠️ Missing Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data:
                    st.error("🚫 Extra Data Found (Reference is empty here)"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors:
                    st.error("⏰ Date/Time Format Error (D, K)"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"System Error: {e}")
