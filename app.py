import streamlit as st
import pandas as pd
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Excel Validator", layout="wide")

st.title("📊 โปรแกรมตรวจสอบไฟล์ Excel (Sheet: RAMP v1.3)")

# ชื่อชีทที่ต้องการให้อ่าน
TARGET_SHEET = "RAMP v1.3"

# --- 1. โหลดไฟล์พื้นฐานจาก GitHub ---
try:
    # อ่านไฟล์ Reference และ Model list
    # สำหรับ reference.xlsx ให้อ่านเฉพาะชีท RAMP v1.3
    df_ref = pd.read_excel("reference.xlsx", sheet_name=TARGET_SHEET, engine='openpyxl')
    df_models = pd.read_excel("model_list.xlsx", engine='openpyxl') # model_list ปกติจะมีชีทเดียว
    
    # ดึงรายชื่อ Model
    model_list = df_models['model'].unique().tolist()
    selected_model = st.selectbox("1. กรุณาเลือก Model:", ["เลือก Model"] + model_list)

    if selected_model != "เลือก Model":
        search_result = df_models[df_models['model'] == selected_model]
        
        if search_result.empty:
            st.error(f"❌ ไม่พบข้อมูลสำหรับ Model: {selected_model} ในระบบ")
            st.stop()

        model_info = search_result.iloc[0]
        correct_part_no = str(model_info['part_no']).strip()
        correct_dwg_no = str(model_info['dwg_no']).strip()

        # --- 2. ส่วนการอัปโหลดไฟล์ ---
        uploaded_file = st.file_uploader(f"2. อัปโหลดไฟล์ที่ต้องการตรวจสอบ (ระบบจะอ่านเฉพาะชีท '{TARGET_SHEET}')", type=["xlsx", "xlsm"])

        if uploaded_file:
            try:
                # อ่านไฟล์ผู้ใช้ เฉพาะชีท RAMP v1.3
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
            except ValueError:
                # กรณีหาชีทชื่อ RAMP v1.3 ไม่เจอ
                st.error(f"❌ ไม่พบชีทชื่อ '{TARGET_SHEET}' ในไฟล์ที่อัปโหลด กรุณาตรวจสอบชื่อชีท")
                st.stop()

            errors = []

            # --- ตัวป้องกัน: เช็คจำนวนแถว ---
            if df_user.shape[0] < 76:
                st.error(f"❌ ข้อมูลในชีท {TARGET_SHEET} มีจำนวนแถวน้อยเกินไป (มี {df_user.shape[0]} แถว) ต้องมีอย่างน้อย 76 แถว")
                st.stop()

            st.info(f"กำลังตรวจสอบชีท {TARGET_SHEET} สำหรับ Model: {selected_model}...")

            # --- เงื่อนไขที่ 1: ตรวจสอบ F3 (Part No.) และ F5 (Dwg No.) ---
            # Index: F3 = [2, 5], F5 = [4, 5]
            user_part_no = str(df_user.iloc[2, 5]).strip()
            user_dwg_no = str(df_user.iloc[4, 5]).strip()

            if user_part_no != correct_part_no:
                errors.append(f"❌ **F3 (Part No.) ไม่ตรง:** ในไฟล์คือ '{user_part_no}' (ต้องเป็น '{correct_part_no}')")
            
            if user_dwg_no != correct_dwg_no:
                errors.append(f"❌ **F5 (Dwg No.) ไม่ตรง:** ในไฟล์คือ '{user_dwg_no}' (ต้องเป็น '{correct_dwg_no}')")

            # --- เงื่อนไขที่ 2: ตรวจสอบการเติมข้อมูลครบถ้วนตาม Reference ---
            for r in range(len(df_ref)):
                for c in range(len(df_ref.columns)):
                    if pd.notna(df_ref.iloc[r, c]): 
                        try:
                            if pd.isna(df_user.iloc[r, c]):
                                col_letter = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
                                errors.append(f"⚠️ **ข้อมูลไม่ครบ:** ช่อง {col_letter}{r+1} ว่างเปล่า")
                        except:
                            pass

            # --- เงื่อนไขที่ 3: ตรวจสอบ Column D และ K (แถว 13-76) ---
            date_format = "%m/%d/%Y %H:%M:%S"
            
            for row_idx in range(12, 76): # แถว 13-76 (index 12-75)
                for col_idx, col_name in zip([3, 10], ['D', 'K']):
                    val = df_user.iloc[row_idx, col_idx]
                    
                    if pd.isna(val):
                        errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** ห้ามเป็นค่าว่าง")
                        continue
                    
                    # เช็ค Format วันที่
                    if isinstance(val, datetime):
                        continue
                    else:
                        check_str = str(val).strip()
                        if check_str == "00:00:00":
                            continue
                        try:
                            datetime.strptime(check_str, date_format)
                        except ValueError:
                            errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** รูปแบบวันที่ผิด (ต้องเป็น M/D/YYYY HH:MM:SS)")

            # --- แสดงผล ---
            st.divider()
            if not errors:
                st.balloons()
                st.success(f"✅ ตรวจสอบชีท {TARGET_SHEET} เรียบร้อย! ข้อมูลถูกต้องและครบถ้วน")
            else:
                st.error(f"พบข้อผิดพลาดทั้งหมด {len(errors)} จุด:")
                for err in errors:
                    st.write(err)

except FileNotFoundError:
    st.error("❌ หาไฟล์ไม่เจอ: ตรวจสอบว่ามีไฟล์ 'reference.xlsx' และ 'model_list.xlsx' อยู่ใน GitHub")
except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
