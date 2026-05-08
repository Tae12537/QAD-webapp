import streamlit as st
import pandas as pd
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Excel Validator", layout="wide")

st.title("📊 โปรแกรมตรวจสอบความถูกต้องไฟล์ Excel (.xlsx, .xlsm)")

# --- 1. โหลดไฟล์พื้นฐานจาก GitHub ---
try:
    # อ่านไฟล์ Reference และ Model list จากโฟลเดอร์เดียวกัน
    df_ref = pd.read_excel("reference.xlsx", engine='openpyxl')
    df_models = pd.read_excel("model_list.xlsx", engine='openpyxl')
    
    # ดึงรายชื่อ Model
    model_list = df_models['model'].unique().tolist()
    selected_model = st.selectbox("1. กรุณาเลือก Model:", ["เลือก Model"] + model_list)

    if selected_model != "เลือก Model":
        # ตรวจสอบว่ามีข้อมูล Model นี้จริงไหม เพื่อป้องกัน Index Out of Bounds
        search_result = df_models[df_models['model'] == selected_model]
        
        if search_result.empty:
            st.error(f"❌ ไม่พบข้อมูลสำหรับ Model: {selected_model} ในระบบ")
            st.stop()

        model_info = search_result.iloc[0]
        correct_part_no = str(model_info['part_no']).strip()
        correct_dwg_no = str(model_info['dwg_no']).strip()

        # --- 2. ส่วนการอัปโหลดไฟล์ (รองรับ .xlsm) ---
        uploaded_file = st.file_uploader("2. อัปโหลดไฟล์ที่ต้องการตรวจสอบ (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

        if uploaded_file:
            # อ่านไฟล์ผู้ใช้ (header=None เพื่อใช้พิกัด index แบบ 0, 1, 2...)
            df_user = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
            errors = []

            # --- ตัวป้องกัน: เช็คขนาดไฟล์เบื้องต้น ---
            if df_user.shape[0] < 76:
                st.error(f"❌ ไฟล์ที่อัปโหลดมีจำนวนแถวน้อยเกินไป (มี {df_user.shape[0]} แถว) ข้อมูลต้องมีอย่างน้อย 76 แถว")
                st.stop()

            st.info(f"กำลังตรวจสอบไฟล์สำหรับ Model: {selected_model}...")

            # --- เงื่อนไขที่ 1: ตรวจสอบ F3 (Part No.) และ F5 (Dwg No.) ---
            # Index Python: F3 = [2, 5], F5 = [4, 5]
            user_part_no = str(df_user.iloc[2, 5]).strip()
            user_dwg_no = str(df_user.iloc[4, 5]).strip()

            if user_part_no != correct_part_no:
                errors.append(f"❌ **F3 (Part No.) ไม่ตรง:** ในไฟล์คือ '{user_part_no}' (ต้องเป็น '{correct_part_no}')")
            
            if user_dwg_no != correct_dwg_no:
                errors.append(f"❌ **F5 (Dwg No.) ไม่ตรง:** ในไฟล์คือ '{user_dwg_no}' (ต้องเป็น '{correct_dwg_no}')")

            # --- เงื่อนไขที่ 2: ตรวจสอบการเติมข้อมูลครบถ้วนตาม Reference ---
            # เช็คเฉพาะ cell ที่ใน Reference ไม่เป็นค่าว่าง
            for r in range(len(df_ref)):
                for c in range(len(df_ref.columns)):
                    if pd.notna(df_ref.iloc[r, c]): 
                        try:
                            if pd.isna(df_user.iloc[r, c]):
                                # แปลง index เป็นชื่อ Column Excel (A, B, C...)
                                col_letter = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
                                errors.append(f"⚠️ **ข้อมูลไม่ครบ:** ช่อง {col_letter}{r+1} ว่างเปล่า")
                        except IndexError:
                            pass

            # --- เงื่อนไขที่ 3: ตรวจสอบ Column D และ K (แถว 13-76) ---
            # D = index 3, K = index 10 | แถว 13-76 = index 12 ถึง 75
            date_format = "%m/%d/%Y %H:%M:%S"
            
            for row_idx in range(12, 76):
                for col_idx, col_name in zip([3, 10], ['D', 'K']):
                    val = df_user.iloc[row_idx, col_idx]
                    
                    if pd.isna(val):
                        errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** ห้ามเป็นค่าว่าง")
                        continue
                    
                    # เช็ค Format วันที่
                    if isinstance(val, datetime):
                        # ถ้าเป็น datetime อยู่แล้ว (Excel แปลงให้) ถือว่าผ่าน
                        continue
                    else:
                        check_str = str(val).strip()
                        # อนุโลมกรณี 00:00:00
                        if check_str == "00:00:00":
                            continue
                        try:
                            datetime.strptime(check_str, date_format)
                        except ValueError:
                            errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** รูปแบบวันที่ผิด (ต้องเป็น M/D/YYYY HH:MM:SS)")

            # --- สรุปผล ---
            st.divider()
            if not errors:
                st.balloons()
                st.success("✅ ตรวจสอบเรียบร้อย! ไฟล์ข้อมูลถูกต้องและครบถ้วนตามเงื่อนไข")
            else:
                st.error(f"พบข้อผิดพลาดทั้งหมด {len(errors)} จุด:")
                for err in errors:
                    st.write(err)

except FileNotFoundError:
    st.error("❌ หาไฟล์ไม่เจอ: ตรวจสอบว่ามีไฟล์ 'reference.xlsx' และ 'model_list.xlsx' อยู่ใน GitHub")
except Exception as e:
    st.error(f"เกิดข้อผิดพลาดไม่คาดคิด: {e}")
