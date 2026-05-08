import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime

st.set_page_config(page_title="File Validator", layout="centered")

def get_column_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

st.title("📁 File Validator (Visual Check)")

uploaded_file = st.file_uploader("2️⃣ Upload File", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # ใช้ openpyxl โหลดเพื่อเช็ค Format การแสดงผล
        wb = load_workbook(uploaded_file, data_only=True)
        ws = wb["RAMP v1.3"]
        
        date_errors = []

        # ตรวจสอบ D13-D76 และ K13-K76
        for row_idx in range(13, 77):
            for col_idx, col_label in zip([4, 11], ['D', 'K']):
                cell = ws.cell(row=row_idx, column=col_idx)
                val = cell.value
                fmt = cell.number_format # ดูว่า Excel ตั้งค่าให้โชว์แบบไหน
                
                if val is not None:
                    # --- Logic: แปลงค่าให้เหมือนที่ตาเห็นใน Excel ---
                    if isinstance(val, (datetime.datetime, datetime.date)):
                        # ถ้า Format ใน Excel ไม่มีตัว 'h' หรือ 'm' (ที่เป็นเวลา) อยู่เลย
                        # แสดงว่า Excel โชว์แค่ "วันที่" เราต้องตัดเวลาทิ้งให้เหลือแค่ข้อความสั้นๆ
                        if 'h' not in fmt.lower():
                            display_text = val.strftime('%d/%m/%Y') # โชว์แค่วันที่
                        else:
                            display_text = val.strftime('%d/%m/%Y %H:%M:%S') # โชว์พร้อมเวลา
                    else:
                        display_text = str(val).strip()

                    # --- นับหลักจากสิ่งที่ "โชว์บนหน้าจอ" ---
                    # ถ้าคุณพิมพ์แค่ 4/5/2026 (และ Excel โชว์แค่นั้น) display_text จะยาวแค่ ~10 หลัก
                    # ถ้ามีเวลาโชว์ด้วย จะยาว 15-19 หลัก
                    if len(display_text) < 15:
                        date_errors.append({
                            "Column": col_label,
                            "Row": row_idx,
                            "Visual Value": display_text,
                            "Status": "❌ ข้อมูลช่องนั้นผิดพลาด (ใน Excel ไม่โชว์เวลา)"
                        })

        if not date_errors:
            st.balloons()
            st.success("✅ ข้อมูลถูกต้อง! ทุกช่องมีเวลาโชว์ครบถ้วน")
        else:
            st.error("⏰ พบช่องที่ไม่มีเวลาโชว์ (นับจากสิ่งที่เห็นใน Excel)")
            st.table(pd.DataFrame(date_errors))

    except Exception as e:
        st.error(f"Error: {e}")
