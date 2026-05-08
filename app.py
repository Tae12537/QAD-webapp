import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Excel Validator", layout="wide")
st.title("📊 ตรวจสอบความสมบูรณ์ของไฟล์ (RAMP v1.3)")

TARGET_SHEET = "RAMP v1.3"

def get_column_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

try:
    # อ่านไฟล์ Reference และ Model
    df_ref = pd.read_excel("reference.xlsx", sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
    df_models = pd.read_excel("model_list.xlsx", engine='openpyxl')
    
    model_list = df_models['model'].unique().tolist()
    selected_model = st.selectbox("1. เลือก Model:", ["เลือก Model"] + model_list)

    if selected_model != "เลือก Model":
        search_result = df_models[df_models['model'] == selected_model]
        if search_result.empty:
            st.error(f"❌ ไม่พบข้อมูล Model: {selected_model}")
            st.stop()

        model_info = search_result.iloc[0]
        correct_part_no = str(model_info['part_no']).strip()
        correct_dwg_no = str(model_info['dwg_no']).strip()

        uploaded_file = st.file_uploader("2. อัปโหลดไฟล์ RAMP v1.3 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

        if uploaded_file:
            try:
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            except ValueError:
                st.error(f"❌ ไม่พบชีทชื่อ '{TARGET_SHEET}'")
                st.stop()

            # --- ส่วนการเก็บข้อมูล Error ---
            f_errors = [] # สำหรับ F3, F5
            missing_data = {} # สำหรับจัดกลุ่มข้อมูลไม่ครบ { 'Column': [Rows] }
            date_errors = [] # สำหรับ format วันที่ D, K

            max_row = min(len(df_ref), len(df_user), 76)
            max_col = min(df_ref.shape[1], df_user.shape[1])

            # 1. ตรวจสอบ F3 และ F5
            user_part_no = str(df_user.iloc[2, 5]).strip()
            user_dwg_no = str(df_user.iloc[4, 5]).strip()

            if user_part_no != correct_part_no:
                f_errors.append({"ตำแหน่ง": "F3 (Part No.)", "พบ": user_part_no, "ที่ถูกต้อง": correct_part_no})
            if user_dwg_no != correct_dwg_no:
                f_errors.append({"ตำแหน่ง": "F5 (Dwg No.)", "พบ": user_dwg_no, "ที่ถูกต้อง": correct_dwg_no})

            # 2. ตรวจสอบความครบถ้วน (จัดกลุ่มตาม Column)
            for r in range(max_row):
                for c in range(max_col):
                    ref_val = df_ref.iloc[r, c]
                    if pd.notna(ref_val) and str(ref_val).strip() != "":
                        user_val = df_user.iloc[r, c]
                        if pd.isna(user_val) or str(user_val).strip() == "":
                            col_name = get_column_letter(c)
                            if col_name not in missing_data:
                                missing_data[col_name] = []
                            missing_data[col_name].append(str(r + 1))

            # 3. ตรวจสอบรูปแบบวันที่ D และ K
            date_format = "%m/%d/%Y %H:%M:%S"
            for row_idx in range(12, 76):
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < len(df_user):
                        val = df_user.iloc[row_idx, col_idx]
                        is_empty = pd.isna(val) or str(val).strip() == ""
                        
                        if is_empty:
                            date_errors.append({"คอลัมน์": col_label, "แถว": row_idx + 1, "ปัญหา": "ข้อมูลว่างเปล่า"})
                        elif not isinstance(val, datetime):
                            check_str = str(val).strip()
                            if check_str != "00:00:00":
                                try:
                                    datetime.strptime(check_str, date_format)
                                except ValueError:
                                    date_errors.append({"คอลัมน์": col_label, "แถว": row_idx + 1, "ปัญหา": "รูปแบบวันที่ผิด"})

            # --- การแสดงผลแบบใหม่ (ตารางและ Expander) ---
            st.divider()
            
            if not f_errors and not missing_data and not date_errors:
                st.balloons()
                st.success("✅ ไฟล์ถูกต้องสมบูรณ์ 100%")
            else:
                # แสดง Error F3/F5 ถ้ามี
                if f_errors:
                    st.subheader("⚠️ ตรวจสอบ Part No. / Dwg No.")
                    st.table(pd.DataFrame(f_errors))

                # แสดงตารางข้อมูลไม่ครบ (จัดกลุ่ม)
                if missing_data:
                    st.subheader("⚠️ รายการช่องที่ข้อมูลไม่ครบ")
                    # แปลง Dictionary เป็น List ของตาราง
                    table_data = []
                    for col, rows in missing_data.items():
                        table_data.append({
                            "คอลัมน์": col,
                            "จำนวนที่หายไป": len(rows),
                            "แถวที่ต้องตรวจสอบ": ", ".join(rows)
                        })
                    st.table(pd.DataFrame(table_data))

                # แสดง Error วันที่
                if date_errors:
                    with st.expander("❌ คลิกเพื่อดูรายละเอียดรูปแบบวันที่ผิด (D, K)"):
                        st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
