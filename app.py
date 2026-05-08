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
        
        # 1. อ่านไฟล์แบบปกติ (ไม่ระบุ dtype=str เพื่อให้ช่องอื่นไม่พัง)
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []
            
            # --- ส่วนที่ 1: Check F3 / F5 ---
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # --- ส่วนที่ 2: Check ข้อมูลทั่วไป (A, B, C, E, ...) ---
            for r in range(76):
                for c in range(df_ref.shape[1]):
                    # ข้ามคอลัมน์ D(3) และ K(10) ในช่วงข้อมูล
                    if r >= 12 and c in [3, 10]: continue
                    
                    ref_v = str(df_ref.iloc[r, c]).strip() if r < df_ref.shape[0] else "nan"
                    user_v = str(df_user.iloc[r, c]).strip() if r < df_user.shape[0] else "nan"
                    
                    if ref_v != "nan" and user_v == "nan":
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    elif ref_v == "nan" and user_v != "nan":
                        if r+1 == 12: continue
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # --- ส่วนที่ 3: จัดการ D และ K (นับ Digit แบบ Force Format) ---
            for row_idx in range(12, 76): 
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < df_user.shape[0]:
                        raw_val = df_user.iloc[row_idx, col_idx]
                        
                        if pd.notna(raw_val) and str(raw_val).strip() != "":
                            # บังคับแปลงเป็น String Format "Y-m-d H:M:S" เพื่อให้นับหลักได้คงที่
                            try:
                                # ถ้าเป็นวันที่ (DateTime) มันจะกลายเป็น "2026-05-04 00:00:00" (19 หลัก)
                                formatted_val = pd.to_datetime(raw_val).strftime('%Y-%m-%d %H:%M:%S')
                            except:
                                # ถ้าไม่ใช่เครื่องหมายวันที่ ก็ให้อ่านเป็น string ตรงๆ
                                formatted_val = str(raw_val).strip()

                            # เช็ค Digit: ถ้าไม่มีข้อมูลเวลา (มักจะมาในรูป 00:00:00) 
                            # หรือถ้าเรานับจาก String ที่ "ควรจะเป็น" ใน Excel
                            # ผมเปลี่ยนมาเช็คว่า "ถ้าคุณใส่แค่ 4/5/2026" ใน Excel จริงๆ
                            # มันจะไม่มี Space ในตัวแปรดั้งเดิม
                            
                            original_str = str(raw_val).strip()
                            
                            # ถ้าไม่มีช่องว่าง (ไม่มีเวลาพ่วง) -> Alarm
                            # (เช่น '4/5/2026' ไม่มีช่องว่าง แต่ '11/4/2026 0:00:00' มีช่องว่าง)
                            if " " not in original_str:
                                date_errors.append({
                                    "Column": col_label,
                                    "Row": row_idx + 1,
                                    "Value": original_str,
                                    "Issue": "ข้อมูลช่องนั้นผิดพลาด (ใส่เวลาไม่ครบ)"
                                })

            # --- แสดงผล ---
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ ข้อมูลถูกต้องทั้งหมด!")
            else:
                if f_errors: st.warning("⚠️ Mismatch F3/F5"); st.table(pd.DataFrame(f_errors))
                if missing_data: st.warning("⚠️ Data Missing"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data: st.error("🚫 Extra Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors: st.error("⏰ พบข้อผิดพลาดในคอลัมน์ D หรือ K"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"Error: {e}")
