import streamlit as st
import pandas as pd
from openpyxl import load_workbook

st.set_page_config(page_title="File Validator", layout="centered")

def reset_app():
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

st.title("📁 File Validator (Format Inspector)")
if st.button("🔄 Reset Application"):
    reset_app()

st.divider()
TARGET_SHEET = "RAMP v1.3"

uploaded_file = st.file_uploader("Upload File", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # ใช้ openpyxl เพื่อเข้าถึงการตั้งค่า Format หลังบ้านของ Excel
        wb = load_workbook(uploaded_file, data_only=False)
        if TARGET_SHEET not in wb.sheetnames:
            st.error(f"ไม่พบ Sheet ชื่อ '{TARGET_SHEET}'")
        else:
            ws = wb[TARGET_SHEET]
            date_errors = []

            # ตรวจสอบ D13-D76 (Col 4) และ K13-K76 (Col 11)
            for row_idx in range(13, 77):
                for col_idx, col_label in zip([4, 11], ['D', 'K']):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    val = cell.value
                    fmt = str(cell.number_format) # ดึงค่า Format เช่น 'm/d/yyyy' หรือ 'dd/mm/yyyy hh:mm:ss'
                    
                    if val is not None:
                        # Logic: ถ้าใน Format ไม่มีตัวอักษร 'h' (Hour) หรือ 'm' (Minute ที่อยู่หลังชั่วโมง)
                        # เราจะถือว่าเป็น Format Date ธรรมดาที่ไม่มีการโชว์เวลา
                        # *หมายเหตุ: m ใน Excel format เป็นได้ทั้ง Month และ Minute แต่ปกติถ้ามี h จะตามด้วย m
                        
                        is_format_valid = 'h' in fmt.lower()
                        
                        if not is_format_valid:
                            date_errors.append({
                                "Column": col_label,
                                "Row": row_idx,
                                "Current Format": fmt,
                                "Status": "❌ ข้อมูลไม่ครบ/Format ผิด (ขาดเวลา)"
                            })

            # --- แสดงผล ---
            if not date_errors:
                st.balloons()
                st.success("✅ ข้อมูลและ Format ถูกต้องทั้งหมด!")
            else:
                st.error("⏰ พบข้อผิดพลาดเรื่อง Format หรือข้อมูลเวลาไม่ครบ")
                st.write("กรุณาแก้ Format ให้เป็น 'm/d/yyyy hh:mm:ss' หรือใส่เวลาให้ครบ")
                st.table(pd.DataFrame(date_errors))

    except Exception as e:
        st.error(f"System Error: {e}")
