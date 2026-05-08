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
            # --- เทคนิคแคะข้อความดิบ (Raw Values Only) ---
            # ใช้ openpyxl เพื่อดึงสิ่งที่ User พิมพ์จริงๆ ออกมา
            wb = load_workbook(uploaded_file, data_only=False) 
            ws = wb[TARGET_SHEET]
            
            # อ่านไฟล์ Reference และ User สำหรับส่วนอื่นๆ
            df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []

            # 1. Check F3 / F5
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # 2. Check ข้อมูลทั่วไป (ข้าม D, K)
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

            # 3. Check คอลัมน์ D และ K โดย "นับตัวอักษร" เท่านั้น
            # เกณฑ์: วันที่+เวลา (xx/xx/xx xx:xx:xx) ควรมีประมาณ 17-19 ตัวอักษร
            # ถ้าพิมพ์แค่ 4/5/2026 จะมีแค่ประมาณ 8-10 ตัวอักษร
            for row_idx in range(13, 77): # แถว 13 ถึง 76
                for col_idx, col_label in zip([4, 11], ['D', 'K']): # D=4, K=11 ใน openpyxl
                    cell = ws.cell(row=row_idx, column=col_idx)
                    
                    # ดึงค่าที่เป็น Internal Value ออกมา (ไม่เอา Format)
                    raw_val = cell.value
                    
                    if raw_val is not None:
                        # บังคับให้เป็น String แล้วลบช่องว่างหัวท้าย
                        val_str = str(raw_val).strip()
                        
                        # กรองคำว่า 00:00:00 ออกถ้ามันแอบติดมาจากการแปลงประเภทข้อมูลของ Library
                        # แต่ถ้าคุณพิมพ์มาเอง มันจะมีช่องว่างคั่น เช่น "2026-05-04 10:30:00"
                        
                        # ถ้าความยาวน้อยกว่า 15 (แปลว่าไม่มีเวลาพ่วงมาด้วย) -> Alarm
                        if len(val_str) < 15:
                            date_errors.append({
                                "Column": col_label,
                                "Row": row_idx,
                                "Value": val_str,
                                "Status": "ข้อมูลช่องนั้นผิดพลาด (กรุณาใส่เวลาให้ครบ)"
                            })

            # --- สรุปผล ---
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ ข้อมูลถูกต้องทั้งหมด!")
            else:
                if f_errors: st.warning("⚠️ หัวตารางไม่ตรง"); st.table(pd.DataFrame(f_errors))
                if missing_data: st.warning("⚠️ ข้อมูลขาดหาย"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data: st.error("🚫 มีข้อมูลเกิน"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors: st.error("⏰ รูปแบบวันที่/เวลาผิดพลาด (D, K)"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
