import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

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
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            # --- เทคนิคพิเศษ: ใช้ openpyxl อ่านค่าดิบ (Raw Value) เพื่อดักจำนวน Digit ---
            wb = load_workbook(uploaded_file, data_only=False) # data_only=False เพื่อดูว่า user พิมพ์อะไร
            ws = wb[TARGET_SHEET]
            
            # แปลงเป็น DataFrame สำหรับช่องอื่นๆ
            df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []

            # 1. Check F3 / F5
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # 2. Check ข้อมูลทั่วไป (ยกเว้น D, K)
            for r in range(76):
                for c in range(df_ref.shape[1]):
                    if r >= 12 and c in [3, 10]: continue
                    ref_v = str(df_ref.iloc[r, c]).strip()
                    user_v = str(df_user.iloc[r, c]).strip()
                    if ref_v != "" and (user_v == "" or user_v == "nan"):
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    elif ref_v == "" and (user_v != "" and user_v != "nan"):
                        if r+1 == 12: continue
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # 3. Check Date Digit Count (D13-76 และ K13-76) แบบแคะค่าดิบ
            for row_idx in range(13, 77): # แถว Excel 13-76
                for col_idx, col_label in zip([4, 11], ['D', 'K']): # openpyxl นับเริ่มที่ 1 (D=4, K=11)
                    raw_cell_value = ws.cell(row=row_idx, column=col_idx).value
                    
                    if raw_cell_value is not None:
                        # แปลงเป็น string เพื่อให้นับหลักได้
                        val_str = str(raw_cell_value).strip()
                        
                        # ถ้าเป็นคอลัมน์ D แล้วยาวไม่ถึง 15 หลัก (เช่นพิมพ์แค่ 4/5/2026) -> Alarm
                        # หรือถ้าเป็นคอลัมน์ K แล้วพิมพ์สั้น -> Alarm
                        if len(val_str) < 15:
                            date_errors.append({
                                "Column": col_label,
                                "Row": row_idx,
                                "Value": val_str,
                                "Issue": "ข้อมูลช่องนั้นผิดพลาด (กรุณาใส่เวลาให้ครบ)"
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
