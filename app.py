import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

st.set_page_config(page_title="File Validator Pro", layout="wide")

# --- Function สำหรับ Reset ระบบ ---
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
st.title("📁 File Validator (Strict Format Check)")
if st.button("🔄 Reset All (Clear Model & File)"):
    reset_app()

st.divider()
TARGET_SHEET = "RAMP v1.3"

try:
    available_models = get_available_models()
    
    # ใช้ Session State เพื่อให้ปุ่ม Reset ทำงานกับ Selectbox ได้
    if 'model_idx' not in st.session_state:
        st.session_state.model_idx = 0

    model_list = ["-- Please Select --"] + sorted(list(available_models.keys()))
    selected_model_name = st.selectbox("1️⃣ Select Model:", model_list, index=st.session_state.model_idx)

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            # อ่านข้อมูล
            wb = load_workbook(uploaded_file, data_only=False)
            ws = wb[TARGET_SHEET]
            df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
            
            f_errors, missing_data, extra_data = [], {}, {}
            d_errors, k_errors = [], [] # แยกกลุ่ม Error D และ K

            # --- 1. Check F3 / F5 ---
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # --- 2. Check Missing/Extra Data (Excluding D, K) ---
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

            # --- 3. Strict Custom Format Check (D & K แยกกลุ่ม) ---
            # กำหนด Format ที่ยอมรับ (m/d/yyyy hh:mm:ss ใน Excel มักจะเป็น mm/dd/yyyy... หรือคล้ายกัน)
            # เราจะเช็คว่าต้องมีทั้ง 'y', 'm', 'd', 'h', และ 's' อยู่ใน Format
            for row_idx in range(13, 77): 
                # ตรวจ Column D
                cell_d = ws.cell(row=row_idx, column=4)
                if cell_d.value is not None:
                    fmt_d = str(cell_d.number_format).lower()
                    if not ('h' in fmt_d and 's' in fmt_d):
                        d_errors.append({"Row": row_idx, "Current Format": fmt_d, "Issue": "ไม่ใช่ Custom m/d/yyyy hh:mm:ss"})

                # ตรวจ Column K
                cell_k = ws.cell(row=row_idx, column=11)
                if cell_k.value is not None:
                    fmt_k = str(cell_k.number_format).lower()
                    if not ('h' in fmt_k and 's' in fmt_k):
                        k_errors.append({"Row": row_idx, "Current Format": fmt_k, "Issue": "ไม่ใช่ Custom m/d/yyyy hh:mm:ss"})

            # --- Display Results ---
            st.divider()
            if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                st.balloons()
                st.success("✅ ข้อมูลและรูปแบบถูกต้องทั้งหมด!")
            else:
                # ส่วนหัว F3/F5
                if f_errors:
                    st.warning("⚠️ ข้อมูลส่วนหัวไม่ตรง (F3/F5)")
                    st.table(pd.DataFrame(f_errors))
                
                # ส่วนข้อมูล Missing/Extra
                col1, col2 = st.columns(2)
                with col1:
                    if missing_data:
                        st.warning("⚠️ ข้อมูลขาดหาย (Missing)")
                        st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                with col2:
                    if extra_data:
                        st.error("🚫 ข้อมูลเกิน (Extra)")
                        st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])

                # ส่วนแยก D และ K (Grouped UI)
                st.subheader("⏰ ตรวจสอบรูปแบบวันที่และเวลา (Custom Format Only)")
                d_col, k_col = st.columns(2)
                
                with d_col:
                    st.markdown("**คอลัมน์ D (Washing Date)**")
                    if d_errors:
                        st.error(f"พบ {len(d_errors)} จุดที่ผิดในคอลัมน์ D")
                        st.table(pd.DataFrame(d_errors))
                    else:
                        st.success("✅ คอลัมน์ D ถูกต้อง")

                with k_col:
                    st.markdown("**คอลัมน์ K (Finish Date)**")
                    if k_errors:
                        st.error(f"พบ {len(k_errors)} จุดที่ผิดในคอลัมน์ K")
                        st.table(pd.DataFrame(k_errors))
                    else:
                        st.success("✅ คอลัมน์ K ถูกต้อง")

except Exception as e:
    st.error(f"Error: {e}")
