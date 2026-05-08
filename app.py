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
        
        # อ่านไฟล์ Reference และ User ให้เป็น String ทั้งหมดเพื่อความเสถียร
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl').fillna("")
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl').fillna("")
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []
            
            # --- ส่วนที่ 1: Check F3 / F5 (หัวตารางต้องตรง) ---
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = df_ref.iloc[r, c].strip()
                user_v = df_user.iloc[r, c].strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # --- ส่วนที่ 2: Check Missing/Extra (เน้นแค่ "ความว่าง") ---
            # วนลูปเช็คถึงแถว 76
            for r in range(76):
                for c in range(df_ref.shape[1]):
                    ref_v = df_ref.iloc[r, c].strip() if r < df_ref.shape[0] else ""
                    user_v = df_user.iloc[r, c].strip() if r < df_user.shape[0] else ""
                    
                    # ข้ามการเช็ค Missing/Extra สำหรับคอลัมน์ D(3) และ K(10) ในแถว 13-76 
                    # เพราะเราจะไปเช็คละเอียดที่ส่วนที่ 3 แทน
                    if r >= 12 and c in [3, 10]:
                        continue

                    # Missing: ต้นฉบับมี แต่ไฟล์ที่อัพโหลดไม่มี
                    if ref_v != "" and user_v == "":
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    
                    # Extra: ต้นฉบับไม่มี แต่ไฟล์ที่อัพโหลดดันมี
                    elif ref_v == "" and user_v != "":
                        if r+1 == 12: continue # ข้ามหัวตารางแถว 12
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # --- ส่วนที่ 3: Targeted Date Check (D13-76 และ K13-76) ---
            # ใช้ Logic การนับ Digit ตามที่คุณต้องการ
            for row_idx in range(12, 76): 
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < df_user.shape[0]:
                        val = df_user.iloc[row_idx, col_idx].strip()
                        
                        # เช็คเฉพาะถ้ามีข้อมูล (ถ้าว่าง ส่วน Missing Data จะจัดการเอง)
                        if val != "":
                            # ถ้าน้อยกว่า 15 หลัก (เช่น มีแค่ 4/5/2026) ให้ Alarm
                            if len(val) < 15:
                                date_errors.append({
                                    "Column": col_label,
                                    "Row": row_idx + 1,
                                    "Value": val,
                                    "Issue": "ข้อมูลช่องนั้นผิดพลาด (กรุณาใส่เวลาให้ครบ)"
                                })

            # --- แสดงผลลัพธ์ ---
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons()
                st.success("✅ ข้อมูลถูกต้องทั้งหมด!")
            else:
                if f_errors: 
                    st.warning("⚠️ ข้อมูลส่วนหัวไม่ตรง (F3/F5)"); st.table(pd.DataFrame(f_errors))
                if missing_data: 
                    st.warning("⚠️ ข้อมูลหายไป"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data: 
                    st.error("🚫 มีข้อมูลเกินมาจากต้นฉบับ"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors: 
                    st.error("⏰ พบข้อผิดพลาดในคอลัมน์ D หรือ K"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"Error: {e}")
