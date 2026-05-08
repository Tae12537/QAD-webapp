import streamlit as st
import pandas as pd
from datetime import datetime

# ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Excel Validator", layout="wide")

st.title("📊 โปรแกรมตรวจสอบความถูกต้องไฟล์ Excel (.xlsx, .xlsm)")

# --- ส่วนการโหลดไฟล์จาก GitHub ---
try:
    # อ่านไฟล์จาก GitHub (ใช้ engine='openpyxl' เพื่อความชัวร์กับไฟล์ที่มี Macro หรือ Format พิเศษ)
    df_ref = pd.read_excel("reference.xlsx", engine='openpyxl')
    df_models = pd.read_excel("model_list.xlsx", engine='openpyxl')
    
    model_list = df_models['model'].unique().tolist()
    selected_model = st.selectbox("1. กรุณาเลือก Model:", ["เลือก Model"] + model_list)

    if selected_model != "เลือก Model":
        model_info = df_models[df_models['model'] == selected_model].iloc[0]
        correct_part_no = str(model_info['part_no']).strip()
        correct_dwg_no = str(model_info['dwg_no']).strip()

        # --- แก้ไขตรงนี้: เพิ่ม xlsm ใน type ---
        uploaded_file = st.file_uploader("2. อัปโหลดไฟล์ Excel (.xlsx หรือ .xlsm)", type=["xlsx", "xlsm"])

        if uploaded_file:
            # อ่านไฟล์ที่ผู้ใช้อัปโหลด โดยระบุ engine='openpyxl'
            df_user = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
            errors = []

            st.info(f"กำลังตรวจสอบไฟล์สำหรับ Model: {selected_model}...")

            # --- 1. เช็ค F3 (Part No.) และ F5 (Dwg No.) ---
            try:
                user_part_no = str(df_user.iloc[2, 5]).strip() # F3 (Index 2, 5)
                user_dwg_no = str(df_user.iloc[4, 5]).strip()  # F5 (Index 4, 5)

                if user_part_no != correct_part_no:
                    errors.append(f"❌ **F3 (Part No.) ไม่ตรง:** พบ '{user_part_no}' (ต้องเป็น '{correct_part_no}')")
                
                if user_dwg_no != correct_dwg_no:
                    errors.append(f"❌ **F5 (Dwg No.) ไม่ตรง:** พบ '{user_dwg_no}' (ต้องเป็น '{correct_dwg_no}')")
            except Exception:
                errors.append("❌ ไม่สามารถอ่านตำแหน่ง F3 หรือ F5 ได้")

            # --- 2. เช็คความครบถ้วนเทียบกับ Reference ---
            # เช็คเฉพาะพื้นที่ที่มีข้อมูลในไฟล์ Reference
            for r in range(len(df_ref)):
                for c in range(len(df_ref.columns)):
                    if pd.notna(df_ref.iloc[r, c]): 
                        try:
                            if pd.isna(df_user.iloc[r, c]):
                                col_letter = chr(65 + c) if c < 26 else f"A{chr(65 + c - 26)}"
                                errors.append(f"⚠️ **ข้อมูลไม่ครบ:** ช่อง {col_letter}{r+1} ว่างเปล่า")
                        except:
                            pass

            # --- 3. เช็ค Column D12 และ K12 (แถว 13-76) ---
            date_format = "%m/%d/%Y %H:%M:%S"
            
            for row_idx in range(12, 76): # แถว 13-76
                for col_idx, col_name in zip([3, 10], ['D', 'K']): # D=3, K=10
                    try:
                        val = df_user.iloc[row_idx, col_idx]
                        if pd.isna(val):
                            errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** ข้อมูลว่าง")
                            continue
                        
                        # ตรวจสอบ Format
                        if isinstance(val, datetime):
                            pass # ถ้าเป็น datetime object จาก Excel อยู่แล้วถือว่าผ่าน
                        else:
                            check_str = str(val).strip()
                            # ถ้าเป็น 00:00:00 อย่างเดียวให้ผ่านตามเงื่อนไขอนุโลม
                            if "00:00:00" in check_str and len(check_str) <= 8:
                                pass
                            else:
                                datetime.strptime(check_str, date_format)
                            
                    except ValueError:
                        errors.append(f"❌ **คอลัมน์ {col_name} แถว {row_idx+1}:** รูปแบบวันที่ผิด (ต้องเป็น M/D/YYYY HH:MM:SS)")

            # --- แสดงผล ---
            st.divider()
            if not errors:
                st.balloons()
                st.success("✅ ตรวจสอบผ่าน! ไฟล์ `.xlsm` ถูกต้องตามเงื่อนไข")
            else:
                st.error(f"พบข้อผิดพลาด {len(errors)} จุด:")
                for err in errors:
                    st.write(err)

except Exception as e:
    st.error(f"เกิดข้อผิดพลาด: {e}")
