import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

# Layout ตรงกลางเหมือนเดิม
st.set_page_config(page_title="File Validator Pro", layout="centered")

if 'reset_counter' not in st.session_state:
    st.session_state.reset_counter = 0

def reset_app():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
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

            # --- 3. Super Strict Format Check (D & K) ---
            for row_idx in range(13, 77): 
                for col_idx, (col_label, error_list) in enumerate(zip(['D', 'K'], [d_errors, k_errors]), start=0):
                    # openpyxl col: D=4, K=11
                    target_col = 4 if col_label == 'D' else 11
                    cell = ws.cell(row=row_idx, column=target_col)
                    
                    if cell.value is not None:
                        fmt = str(cell.number_format).lower()
                        
                        # ต้องมีองค์ประกอบครบ: y (ปี), d หรือ m (วัน/เดือน), และ h (ชั่วโมง)
                        # ถ้าเป็น Time อย่างเดียวจะไม่มี y/d
                        # ถ้าเป็น Date อย่างเดียวจะไม่มี h
                        has_date = ('y' in fmt) or ('d' in fmt and 'm' in fmt)
                        has_time = ('h' in fmt)
                        
                        if not (has_date and has_time):
                            error_list.append({
                                "Row": row_idx, 
                                "Format": fmt, 
                                "Status": "❌ รูปแบบผิด (ต้องมีทั้งวันที่และเวลา)"
                            })

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

                st.subheader("⏰ ตรวจสอบรูปแบบวันที่และเวลา")
                if d_errors:
                    st.error("❌ คอลัมน์ D (Washing Date) รูปแบบไม่ถูกต้อง")
                    st.table(pd.DataFrame(d_errors))
                else:
                    st.success("✅ คอลัมน์ D ถูกต้อง")

                if k_errors:
                    st.error("❌ คอลัมน์ K (Finish Date) รูปแบบไม่ถูกต้อง")
                    st.table(pd.DataFrame(k_errors))
                else:
                    st.success("✅ คอลัมน์ K ถูกต้อง")

except Exception as e:
    st.error(f"Error: {e}")
