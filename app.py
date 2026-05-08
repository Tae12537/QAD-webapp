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
        # อ่านไฟล์ Reference
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File (Target: {selected_model_name})", type=["xlsx", "xlsm"])

        if uploaded_file:
            # *** จุดสำคัญ: บังคับอ่านทุกอย่างเป็น String (dtype=str) เพื่อไม่ให้ Excel แอบเติมเวลาให้เอง ***
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl')
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []
            user_rows, user_cols = df_user.shape
            ref_rows, ref_cols = df_ref.shape

            # 1. Check F3 & F5 (Part No / Dwg No)
            correct_f3 = str(df_ref.iloc[2, 5]).strip()
            correct_f5 = str(df_ref.iloc[4, 5]).strip()
            user_f3 = str(df_user.iloc[2, 5]).strip() if user_rows > 2 else ""
            user_f5 = str(df_user.iloc[4, 5]).strip() if user_rows > 4 else ""

            if user_f3 != correct_f3: f_errors.append({"Position": "F3", "Found": user_f3, "Target": correct_f3})
            if user_f5 != correct_f5: f_errors.append({"Position": "F5", "Found": user_f5, "Target": correct_f5})

            # 2. Check All Cells (Missing & Extra)
            max_r = max(ref_rows, user_rows)
            max_c = max(ref_cols, user_cols)

            for r in range(max_r):
                for c in range(max_c):
                    ref_val = df_ref.iloc[r, c] if (r < ref_rows and c < ref_cols) else pd.NA
                    user_val = df_user.iloc[r, c] if (r < user_rows and c < user_cols) else pd.NA
                    
                    # ล้างช่องว่าง
                    is_ref_empty = pd.isna(ref_val) or str(ref_val).strip() == "" or str(ref_val).strip().lower() == "nan"
                    is_user_empty = pd.isna(user_val) or str(user_val).strip() == "" or str(user_val).strip().lower() == "nan"
                    
                    col_name = get_column_letter(c)

                    # Missing: ต้นฉบับมี แต่เราไม่มี
                    if not is_ref_empty and is_user_empty:
                        if col_name not in missing_data: missing_data[col_name] = []
                        missing_data[col_name].append(str(r+1))
                    
                    # Extra: ต้นฉบับไม่มี แต่เราดันมี (เน้นตรวจ F15, K15 ถ้าว่าง)
                    elif is_ref_empty and not is_user_empty:
                        if r+1 == 12: continue # ข้ามหัวตารางแถว 12
                        if col_name not in extra_data: extra_data[col_name] = []
                        extra_data[col_name].append(str(r+1))

            # 3. Check Date Format (เริ่มแถว 13 ถึง 76)
            for row_idx in range(12, 76): # Index 12 = แถว 13
                for col_idx, col_label in zip([3, 10], ['D', 'K']): # คอลัมน์ D(3) และ K(10)
                    if row_idx < user_rows:
                        val = df_user.iloc[row_idx, col_idx]
                        if not pd.isna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                            raw_val = str(val).strip()
                            
                            # ตรวจสอบ: ต้องมี "ช่องว่าง" (Space) และ "เครื่องหมายเวลา" (:)
                            # ถ้าเป็น 4/5/2026 เฉยๆ จะไม่มี Space ทำให้โดน Alarm
                            if " " not in raw_val or ":" not in raw_val:
                                # ข้อยกเว้น: ถ้ากรอกเป็น 00:00:00 เป๊ะๆ ให้ผ่าน
                                if raw_val == "00:00:00":
                                    continue
                                    
                                date_errors.append({
                                    "Column": col_label, 
                                    "Row": row_idx + 1, 
                                    "Your Value": raw_val, 
                                    "Issue": "Missing Time Component (Need HH:MM:SS)"
                                })

            # --- Display Results ---
            st.divider()
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ File is Perfect!")
            else:
                if f_errors:
                    st.warning("⚠️ Main Info Mismatch (F3/F5)"); st.table(pd.DataFrame(f_errors))
                if missing_data:
                    st.warning("⚠️ Missing Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data:
                    st.error("🚫 Unexpected Extra Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors:
                    st.error("⏰ Date Format Error (Missing Time)"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"System Error: {e}")
