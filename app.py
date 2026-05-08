import streamlit as st
import pandas as pd
import os

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
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl').fillna("")
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            # อ่านไฟล์โดยใช้ dtype=str เพื่อพยายามรักษาค่าเดิมให้มากที่สุด
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl').fillna("")
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []

            # 1. Check ข้อมูลทั่วไป
            for r in range(76):
                for c in range(df_ref.shape[1]):
                    if r >= 12 and c in [3, 10]: continue
                    ref_v = str(df_ref.iloc[r, c]).strip()
                    user_v = str(df_user.iloc[r, c]).strip()
                    if ref_v != "" and user_v == "":
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    elif ref_v == "" and user_v != "":
                        if r+1 == 12: continue
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # 2. Check คอลัมน์ D และ K (ดักจับพวกที่ไม่มีเวลา)
            for row_idx in range(12, 76): 
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < df_user.shape[0]:
                        val = str(df_user.iloc[row_idx, col_idx]).strip()
                        
                        if val != "" and val.lower() != "nan":
                            # Logic:
                            # ถ้าค่าลงท้ายด้วย 00:00:00 (ซึ่ง Excel แอบเติมให้เองเมื่อพิมพ์แต่วันที่)
                            # หรือ ถ้ามีความยาวสั้นเกินไป (ไม่มีเวลาเลย)
                            # เราจะมองว่า "ข้อมูลผิดพลาด" ทันที
                            if "00:00:00" in val or len(val) < 15:
                                date_errors.append({
                                    "Column": col_label,
                                    "Row": row_idx + 1,
                                    "Your Value": val,
                                    "Status": "❌ ข้อมูลผิดพลาด (กรุณาใส่เวลาห้ามเป็นเที่ยงคืนเป๊ะ)"
                                })

            # --- Display ---
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ ข้อมูลถูกต้องทั้งหมด!")
            else:
                if f_errors: st.warning("⚠️ Mismatch F3/F5"); st.table(pd.DataFrame(f_errors))
                if missing_data: st.warning("⚠️ Missing Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data: st.error("🚫 Extra Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors: st.error("⏰ พบข้อผิดพลาดในคอลัมน์ D หรือ K (ใส่เวลาไม่ครบ)"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"Error: {e}")
