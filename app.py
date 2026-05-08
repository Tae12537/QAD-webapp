import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

# 1. ปรับ Layout กลับมาเป็นแบบเล็กลงอยู่ตรงกลาง (centered)
st.set_page_config(page_title="File Validator Pro", layout="centered")

# --- Function สำหรับ Reset ระบบแบบล้างค่า Widget ---
if 'reset_counter' not in st.session_state:
    st.session_state.reset_counter = 0

def reset_app():
    # ล้างค่าทั้งหมดใน session_state
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    # เพิ่ม counter เพื่อเปลี่ยน Key ของ Selectbox บังคับ Reset UI
    st.session_state.reset_counter = st.session_state.get('reset_counter', 0) + 1
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
    model_list = ["-- Please Select --"] + sorted(list(available_models.keys()))
    
    # ใช้ key ที่เปลี่ยนไปเรื่อยๆ ตามการกด Reset เพื่อให้ Selectbox ยอมกลับไปที่ index 0
    selected_model_name = st.selectbox(
        "1️⃣ Select Model:", 
        model_list, 
        index=0, 
        key=f"model_select_{st.session_state.reset_counter}"
    )

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            wb = load_workbook(uploaded_file, data_only=False)
            ws = wb[TARGET_SHEET]
            df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
            
            f_errors, missing_data, extra_data = [], {}, {}
            d_errors, k_errors = [], []

            # --- 1. Check F3 / F5 ---
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # --- 2. Check Missing/Extra Data ---
            for r in range(76):
                for c in range(df_ref.shape[1]):
                    if r >= 12 and c in [3, 10]: continue
                    ref_v = str(df_ref.iloc[r, c]).strip() if r < df_ref.shape[0] else ""
                    user_v = str(df_user.iloc[r, c]).strip() if r < df_user.shape[0] else ""
                    
                    if ref_v != "" and (user_v == "" or user_v == "nan"):
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    elif ref_v == "" and (user_v != "" and user_v != "nan"):
                        if r+1 == 12: continue
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # --- 3. Strict Custom Format Check (D & K) ---
            for row_idx in range(13, 77): 
                # ตรวจ Column D (Washing Date)
                cell_d = ws.cell(row=row_idx, column=4)
                if cell_d.value is not None:
                    fmt_d = str(cell_d.number_format).lower()
                    if not ('h' in fmt_d and 's' in fmt_d):
                        d_errors.append({"Row": row_idx, "Format": fmt_d, "Status": "❌ ขาดเวลา"})

                # ตรวจ Column K (Finish Date)
                cell_k = ws.cell(row=row_idx, column=11)
                if cell_k.value is not None:
                    fmt_k = str(cell_k.number_format).lower()
                    if not ('h' in fmt_k and 's' in fmt_k):
                        k_errors.append({"Row": row_idx, "Format": fmt_k, "Status": "❌ ขาดเวลา"})

            # --- Display Results ---
            if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                st.balloons()
                st.success("✅ ข้อมูลและรูปแบบถูกต้องทั้งหมด!")
            else:
                if f_errors:
                    st.warning("⚠️ ข้อมูลส่วนหัวไม่ตรง (F3/F5)")
                    st.table(pd.DataFrame(f_errors))
                
                if missing_data:
                    st.warning("⚠️ ข้อมูลขาดหาย (Missing)")
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                
                if extra_data:
                    st.error("🚫 ข้อมูลเกิน (Extra)")
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])

                # ส่วนแยก D และ K แบบเล็กลง
                st.subheader("⏰ ตรวจสอบรูปแบบวันที่และเวลา")
                if d_errors:
                    st.error("❌ คอลัมน์ D (Washing Date) ผิดพลาด")
                    st.table(pd.DataFrame(d_errors))
                else:
                    st.success("✅ คอลัมน์ D ถูกต้อง")

                if k_errors:
                    st.error("❌ คอลัมน์ K (Finish Date) ผิดพลาด")
                    st.table(pd.DataFrame(k_errors))
                else:
                    st.success("✅ คอลัมน์ K ถูกต้อง")

except Exception as e:
    st.error(f"Error: {e}")
