import streamlit as st
import pandas as pd
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Excel Validator", layout="wide")

st.title("📊 โปรแกรมตรวจสอบความถูกต้องไฟล์ Excel")

# --- ส่วนการโหลดไฟล์จากโฟลเดอร์เดียวกันบน GitHub ---
try:
    # อ่านไฟล์ Reference และ Model list (ดึงจากไฟล์ที่อยู่ใน GitHub เดียวกับ app.py)
    df_ref = pd.read_excel("reference.xlsx")
    df_models = pd.read_excel("model_list.xlsx")
    
    # ดึงรายชื่อ Model มาใส่ใน Dropdown
    model_list = df_models['model'].unique().tolist()
    selected_model = st.selectbox("1. กรุณาเลือก Model:", ["เลือก Model"] + model_list)

    if selected_model != "เลือก Model":
        # ดึงข้อมูล part_no และ dwg_no ที่ถูกต้องของ model นั้น
        model_info = df_models[df_models['model'] == selected_model].iloc[0]
        correct_part_no = str(model_info['part_no']).strip()
        correct_dwg_no = str(model_info['dwg_no']).strip()

        # 2. ปุ่มอัปโหลดไฟล์
        uploaded_file = st.file_uploader("2. อัปโหลดไฟล์ Excel ที่ต้องการตรวจสอบ", type=["xlsx"])

        if uploaded_file:
            # อ่านไฟล์ที่ผู้ใช้อัปโหลด (ไม่อ่าน header เพื่อระบุตำแหน่ง cell ได้แม่นยำ)
            df_user = pd.read_excel(uploaded_file, header=None)
            errors = []

            st.info(f"กำลังตรวจสอบไฟล์สำหรับ Model: {selected_model}...")

            # --- เงื่อนไขที่ 1: ตรวจสอบ F3 (Part No.) และ F5 (Dwg No.) ---
            # Index ใน Python: แถวที่ 3 คือ index 2, แถวที่ 5 คือ index 4 | Column F คือ index 5
            try:
                user_part_no = str(df_user.iloc[2, 5]).strip() # F3
                user_dwg_no = str(df_user.iloc[4, 5]).strip()  # F5

                if user_part_no != correct_part_no:
                    errors.append(f"❌ **F3 (Part No.) ไม่ตรง:** ในไฟล์เป็น '{user_part_no}' แต่ Model {selected_model} ต้องเป็น '{correct_part_no}'")
                
                if user_dwg_no != correct_dwg_no:
                    errors.append(f"❌ **F5 (Dwg No.) ไม่ตรง:** ในไฟล์เป็น '{user_dwg_no}' แต่ Model {selected_model} ต้องเป็น '{correct_dwg_no}'")
            except IndexError:
                errors.append("❌ ไม่พบข้อมูลในตำแหน่ง F3 หรือ F5")

            # --- เงื่อนไขที่ 2: ตรวจสอบความครบถ้วนของข้อมูลเทียบกับไฟล์ Reference ---
            # เราจะเช็คทุกช่องที่ไฟล์ Reference มีข้อมูล (ไม่เป็นค่าว่าง)
            for r in range(len(df_ref)):
                for c in range(len(df_ref.columns)):
                    if pd.notna(df_ref.iloc[r, c]): # ถ้าต้นแบบมีข้อมูล
                        try:
                            # ตรวจสอบว่าไฟล์ที่อัปโหลดมีข้อมูลในช่องเดียวกันไหม
                            if pd.isna(df_user.iloc[r, c]):
                                col_letter = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
                                errors.append(f"⚠️ **ข้อมูลไม่ครบ:** ช่อง {col_letter}{r+1} ว่างเปล่า (เทียบตามไฟล์ Reference)")
                        except IndexError:
                            col_letter = chr(65 + c)
                            errors.append(f"⚠️ **ข้อมูลขาด:** ไม่พบแถว/คอลัมน์ที่ {col_letter}{r+1}")

            # --- เงื่อนไขที่ 3: ตรวจสอบ Column D12 และ K12 (ข้อมูลแถว 13-76) ---
            # D = index 3, K = index 10 | แถว 13-76 = index 12 ถึง 75
            date_format = "%m/%d/%Y %H:%M:%S"
            
            for row_idx in range(12, 76):
                for col_idx, col_name in zip([3, 10], ['D', 'K']):
                    try:
                        val = df_user.iloc[row_idx, col_idx]
                        if pd.isna(val):
                            errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** ข้อมูลว่างเปล่า")
                            continue
                        
                        # พยายามแปลงค่าเป็น datetime ตาม format ที่กำหนด
                        if isinstance(val, datetime):
                            check_val = val.strftime(date_format)
                        else:
                            check_val = str(val).strip()
                            datetime.strptime(check_val, date_format)
                            
                    except (ValueError, TypeError):
                        # ถ้าแปลงไม่ได้ และไม่ใช่ 00:00:00
                        if "00:00:00" not in str(val):
                            errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** รูปแบบวันที่ผิด (ต้องเป็น M/D/YYYY HH:MM:SS เช่น 4/13/2026 00:00:00)")

            # --- แสดงผลลัพธ์ ---
            st.divider()
            if not errors:
                st.balloons()
                st.success("✅ ตรวจสอบผ่าน! ไฟล์ข้อมูลครบถ้วนและถูกต้องตามเงื่อนไขทั้งหมด สามารถใช้งานได้")
            else:
                st.error(f"พบข้อผิดพลาดทั้งหมด {len(errors)} จุด:")
                for err in errors:
                    st.write(err)

except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์พื้นฐาน: {e}")
    st.info("ตรวจสอบว่าไฟล์ 'reference.xlsx' และ 'model_list.xlsx' อยู่ใน GitHub และพิมพ์ชื่อถูกต้อง")
