import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="Excel Validator", layout="wide")
st.title("📊 โปรแกรมตรวจสอบไฟล์ตามรายโมเดล (RAMP v1.3)")

TARGET_SHEET = "RAMP v1.3"

# ฟังก์ชันแปลง Index เป็นชื่อคอลัมน์ Excel
def get_column_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

# --- 1. สแกนหาไฟล์โมเดลในโฟลเดอร์ ---
def get_available_models():
    files = [f for f in os.listdir('.') if f.startswith('reference_') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
    # ตัดคำว่า reference_ ออก และตัดนามสกุลไฟล์ออกเพื่อเอาชื่อโมเดล
    models = {f.replace('reference_', '').split('.')[0]: f for f in files}
    return models

try:
    available_models = get_available_models()
    
    if not available_models:
        st.error("❌ ไม่พบไฟล์ Reference ในระบบ (ต้องตั้งชื่อแบบ reference_ModelName.xlsx)")
        st.stop()

    # 2. เลือกโมเดลจากชื่อไฟล์ที่มี
    model_names = sorted(list(available_models.keys()))
    selected_model_name = st.selectbox("1. กรุณาเลือก Model:", ["เลือก Model"] + model_names)

    if selected_model_name != "เลือก Model":
        ref_filename = available_models[selected_model_name]
        
        # โหลดไฟล์ Reference ของโมเดลนั้นๆ
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        # ดึงค่า Part No (F3) และ Dwg No (F5) จากไฟล์ต้นฉบับเพื่อใช้เป็นเกณฑ์
        correct_part_no = str(df_ref.iloc[2, 5]).strip()
        correct_dwg_no = str(df_ref.iloc[4, 5]).strip()

        uploaded_file = st.file_uploader(f"2. อัปโหลดไฟล์ที่ต้องการตรวจสอบ (เทียบกับต้นฉบับ {selected_model_name})", type=["xlsx", "xlsm"])

        if uploaded_file:
            try:
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            except ValueError:
                st.error(f"❌ ไม่พบชีทชื่อ '{TARGET_SHEET}' ในไฟล์ที่อัปโหลด")
                st.stop()

            # --- ส่วนการเก็บข้อมูล Error ---
            f_errors = [] 
            missing_data = {} 
            date_errors = [] 

            max_row = min(len(df_ref), len(df_user), 76)
            max_col = min(df_ref.shape[1], df_user.shape[1])

            # 1. ตรวจสอบ F3 และ F5 เทียบกับต้นฉบับรายไฟล์
            user_part_no = str(df_user.iloc[2, 5]).strip()
            user_dwg_no = str(df_user.iloc[4, 5]).strip()

            if user_part_no != correct_part_no:
                f_errors.append({"ตำแหน่ง": "F3 (Part No.)", "พบ": user_part_no, "ต้นฉบับระบุ": correct_part_no})
            if user_dwg_no != correct_dwg_no:
                f_errors.append({"ตำแหน่ง": "F5 (Dwg No.)", "พบ": user_dwg_no, "ต้นฉบับระบุ": correct_dwg_no})

            # 2. ตรวจสอบความครบถ้วน (จัดกลุ่มตาม Column)
            for r in range(max_row):
                for c in range(max_col):
                    ref_val = df_ref.iloc[r, c]
                    # เช็คเฉพาะช่องที่ต้นฉบับมีข้อมูล
                    if pd.notna(ref_val) and str(ref_val).strip() != "":
                        user_val = df_user.iloc[r, c]
                        # ถ้าผู้ใช้ไม่ได้กรอกข้อมูล
                        if pd.isna(user_val) or str(user_val).strip() == "":
                            col_name = get_column_letter(c)
                            if col_name not in missing_data:
                                missing_data[col_name] = []
                            missing_data[col_name].append(str(r + 1))

            # 3. ตรวจสอบรูปแบบวันที่ D และ K (แถว 13-76)
            date_format = "%m/%d/%Y %H:%M:%S"
            for row_idx in range(12, 76):
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < len(df_user):
                        val = df_user.iloc[row_idx, col_idx]
                        is_empty = pd.isna(val) or str(val).strip() == ""
                        
                        if is_empty:
                            date_errors.append({"คอลัมน์": col_label, "แถว": row_idx + 1, "สถานะ": "ข้อมูลว่างเปล่า"})
                        elif not isinstance(val, datetime):
                            check_str = str(val).strip()
                            if check_str != "00:00:00":
                                try:
                                    datetime.strptime(check_str, date_format)
                                except ValueError:
                                    date_errors.append({"คอลัมน์": col_label, "แถว": row_idx + 1, "สถานะ": "รูปแบบวันที่ผิด"})

            # --- การแสดงผลผลลัพธ์ ---
            st.divider()
            
            if not f_errors and not missing_data and not date_errors:
                st.balloons()
                st.success(f"✅ ไฟล์สมบูรณ์! ข้อมูลตรงตามต้นฉบับ {selected_model_name}")
            else:
                # 1. ตารางแสดงผล Part No / Dwg No ไม่ตรง
                if f_errors:
                    st.subheader("❌ ข้อมูลหลักไม่ตรงกับต้นฉบับ")
                    st.table(pd.DataFrame(f_errors))

                # 2. ตารางแสดงข้อมูลไม่ครบ
                if missing_data:
                    st.subheader("⚠️ รายการช่องที่ข้อมูลไม่ครบ")
                    table_data = []
                    for col, rows in missing_data.items():
                        table_data.append({
                            "คอลัมน์": col,
                            "จำนวนที่หายไป": len(rows),
                            "แถวที่ต้องเติมข้อมูล": ", ".join(rows)
                        })
                    st.table(pd.DataFrame(table_data))

                # 3. ตารางแสดง Error วันที่
                if date_errors:
                    with st.expander("🔍 คลิกเพื่อดูรายละเอียดความผิดพลาดของวันที่ (D, K)"):
                        st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
