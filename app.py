import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Excel Validator", layout="wide")
st.title("📊 โปรแกรมตรวจสอบไฟล์ Excel (Sheet: RAMP v1.3)")

TARGET_SHEET = "RAMP v1.3"

# ฟังก์ชันแปลง Index เป็นชื่อคอลัมน์ Excel (A, B, C... AA, AB...)
def get_column_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

try:
    df_ref = pd.read_excel("reference.xlsx", sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
    df_models = pd.read_excel("model_list.xlsx", engine='openpyxl')
    
    model_list = df_models['model'].unique().tolist()
    selected_model = st.selectbox("1. กรุณาเลือก Model:", ["เลือก Model"] + model_list)

    if selected_model != "เลือก Model":
        search_result = df_models[df_models['model'] == selected_model]
        if search_result.empty:
            st.error(f"❌ ไม่พบข้อมูลสำหรับ Model: {selected_model}")
            st.stop()

        model_info = search_result.iloc[0]
        correct_part_no = str(model_info['part_no']).strip()
        correct_dwg_no = str(model_info['dwg_no']).strip()

        uploaded_file = st.file_uploader(f"2. อัปโหลดไฟล์ที่ต้องการตรวจสอบ", type=["xlsx", "xlsm"])

        if uploaded_file:
            try:
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            except ValueError:
                st.error(f"❌ ไม่พบชีทชื่อ '{TARGET_SHEET}'")
                st.stop()

            errors = []
            
            # --- ป้องกัน Index Error ---
            max_row = min(len(df_ref), len(df_user), 76)
            max_col = min(df_ref.shape[1], df_user.shape[1])

            st.info(f"กำลังตรวจสอบข้อมูลสำหรับ Model: {selected_model}...")

            # --- 1. ตรวจสอบ F3 และ F5 (ผ่านแล้ว คงไว้เหมือนเดิม) ---
            user_part_no = str(df_user.iloc[2, 5]).strip()
            user_dwg_no = str(df_user.iloc[4, 5]).strip()

            if user_part_no != correct_part_no:
                errors.append(f"❌ **F3 (Part No.) ไม่ตรง:** ในไฟล์คือ '{user_part_no}' (ต้องเป็น '{correct_part_no}')")
            if user_dwg_no != correct_dwg_no:
                errors.append(f"❌ **F5 (Dwg No.) ไม่ตรง:** ในไฟล์คือ '{user_dwg_no}' (ต้องเป็น '{correct_dwg_no}')")

            # --- 2. ตรวจสอบความครบถ้วน (แก้ปัญหา Error ในช่องว่าง) ---
            # เราจะเช็คเฉพาะแถวที่ 1-76 และเฉพาะช่องที่ใน Reference 'ไม่ว่าง' เท่านั้น
            for r in range(max_row):
                for c in range(max_col):
                    ref_val = df_ref.iloc[r, c]
                    # ถ้าในไฟล์ Reference มีข้อมูล แต่ในไฟล์ที่อัปโหลด 'ว่าง' ถึงจะแจ้งเตือน
                    if pd.notna(ref_val) and str(ref_val).strip() != "":
                        user_val = df_user.iloc[r, c]
                        if pd.isna(user_val) or str(user_val).strip() == "":
                            col_letter = get_column_letter(c)
                            errors.append(f"⚠️ **ข้อมูลไม่ครบ:** ช่อง {col_letter}{r+1} ต้องกรอกข้อมูลตามไฟล์ต้นแบบ")

            # --- 3. ตรวจสอบ Column D และ K (แถว 13-76) ---
            date_format = "%m/%d/%Y %H:%M:%S"
            for row_idx in range(12, 76):
                for col_idx, col_name in zip([3, 10], ['D', 'K']):
                    if row_idx < len(df_user):
                        val = df_user.iloc[row_idx, col_idx]
                        if pd.isna(val) or str(val).strip() == "":
                            errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** ห้ามเป็นค่าว่าง")
                            continue
                        
                        if isinstance(val, datetime):
                            continue
                        else:
                            check_str = str(val).strip()
                            if check_str == "00:00:00":
                                continue
                            try:
                                datetime.strptime(check_str, date_format)
                            except ValueError:
                                errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** รูปแบบวันที่ผิด")

            # --- แสดงผล ---
            st.divider()
            if not errors:
                st.balloons()
                st.success(f"✅ ตรวจสอบเรียบร้อย! ข้อมูลถูกต้องตามต้นแบบ")
            else:
                st.error(f"พบจุดที่ต้องแก้ไข {len(errors)} จุด:")
                for err in errors:
                    st.write(err)

except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
