import streamlit as st
import pandas as pd
import os

# 1. ตั้งค่าหน้าเว็บแบบบีบตรงกลาง
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

# --- UI Header ---
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
            # อ่านไฟล์ User โดยบังคับให้ทุกช่องเป็น String เพื่อเช็คความยาวตัวอักษรที่แท้จริง
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl')
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []
            user_rows, user_cols = df_user.shape
            ref_rows, ref_cols = df_ref.shape

            # --- 1. Check F3 & F5 (ต้องตรงเป๊ะ) ---
            correct_f3 = str(df_ref.iloc[2, 5]).strip()
            correct_f5 = str(df_ref.iloc[4, 5]).strip()
            user_f3 = str(df_user.iloc[2, 5]).strip() if user_rows > 2 else ""
            user_f5 = str(df_user.iloc[4, 5]).strip() if user_rows > 4 else ""

            if user_f3 != correct_f3: f_errors.append({"Position": "F3", "Found": user_f3, "Target": correct_f3})
            if user_f5 != correct_f5: f_errors.append({"Position": "F5", "Found": user_f5, "Target": correct_f5})

            # --- 2. Check All Cells (Missing & Extra) ---
            max_r = max(ref_rows, user_rows)
            max_c = max(ref_cols, user_cols)

            for r in range(max_r):
                for c in range(max_c):
                    ref_val = df_ref.iloc[r, c] if (r < ref_rows and c < ref_cols) else pd.NA
                    user_val = df_user.iloc[r, c] if (r < user_rows and c < user_cols) else pd.NA
                    is_ref_empty = pd.isna(ref_val) or str(ref_val).strip() == "" or str(ref_val).strip().lower() == "nan"
                    is_user_empty = pd.isna(user_val) or str(user_val).strip() == "" or str(user_val).strip().lower() == "nan"
                    col_name = get_column_letter(c)

                    if not is_ref_empty and is_user_empty:
                        if col_name not in missing_data: missing_data[col_name] = []
                        missing_data[col_name].append(str(r+1))
                    elif is_ref_empty and not is_user_empty:
                        if r+1 == 12: continue # ข้ามหัวตาราง
                        if col_name not in extra_data: extra_data[col_name] = []
                        extra_data[col_name].append(str(r+1))

            # --- 3. Check Date Length (D13-76, K13-76) ---
            # เกณฑ์ความยาวขั้นต่ำ: 15 ตัวอักษร (เพื่อดักวันที่ที่ไม่มีเวลา)
            MIN_LENGTH = 15 

            for row_idx in range(12, 76): # แถว 13 ถึง 76
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < user_rows:
                        val = str(df_user.iloc[row_idx, col_idx]).strip()
                        
                        # เช็คเฉพาะช่องที่มีข้อมูล (ไม่ว่าง)
                        if val != "" and val.lower() != "nan":
                            # ถ้าความยาวน้อยกว่าเกณฑ์ หรือไม่มีเวลา (ไม่มีเครื่องหมาย :)
                            if len(val) < MIN_LENGTH or ":" not in val:
                                if val == "00:00:00": continue # ยกเว้นกรณีนี้
                                
                                date_errors.append({
                                    "Column": col_label,
                                    "Row": row_idx + 1,
                                    "Value": val,
                                    "Status": "ข้อมูลช่องนั้นผิดพลาด (กรุณาใส่เวลาให้ครบ)"
                                })

            # --- 4. Display Results ---
            st.divider()
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ File is Perfect!")
            else:
                if f_errors:
                    st.warning("⚠️ Critical Info Mismatch (F3/F5)"); st.table(pd.DataFrame(f_errors))
                if missing_data:
                    st.warning("⚠️ Missing Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data:
                    st.error("🚫 Unexpected Extra Data Found"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors:
                    st.error("⏰ Date/Time Data Error"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"System Error: {e}")
