import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Excel Validator", layout="wide")

# --- ฟังก์ชันช่วยเหลือ ---
def validate_date_format(val):
    if pd.isna(val):
        return False
    try:
        # ตรวจสอบ format M/D/YYYY HH:MM:SS
        datetime.strptime(str(val), '%m/%d/%Y %H:%M:%S')
        return True
    except ValueError:
        return False

# --- ส่วนติดต่อผู้ใช้งาน ---
st.title("📊 Excel Data Validation Tool")

# 1. โหลดไฟล์ Reference และ Model จาก GitHub (ใส่ URL ของคุณ)
# หมายเหตุ: ต้องใช้ URL แบบ 'raw' จาก GitHub
GITHUB_RAW_URL = "https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/"

try:
    df_ref = pd.read_excel(f"{GITHUB_RAW_URL}reference.xlsx")
    df_models = pd.read_excel(f"{GITHUB_RAW_URL}model_list.xlsx")
    
    # 2. เลือก Model
    model_list = df_models['model'].unique().tolist()
    selected_model = st.selectbox("เลือก Model ที่ต้องการตรวจสอบ:", model_list)
    
    # ดึงค่า part_no และ dwg_no ที่ถูกต้องของ model นั้นๆ
    model_info = df_models[df_models['model'] == selected_model].iloc[0]
    correct_part_no = str(model_info['part_no'])
    correct_dwg_no = str(model_info['dwg_no'])

    # 3. อัปโหลดไฟล์
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xlsx) เพื่อเริ่มการตรวจสอบ", type=["xlsx"])

    if uploaded_file:
        # อ่านไฟล์ที่อัปโหลด (อ่านแบบไม่มี Header ก่อนเพื่อเข้าถึง Cell เจาะจง)
        df_user_raw = pd.read_excel(uploaded_file, header=None)
        
        errors = []
        is_valid = True

        # --- Check 1: ตรวจสอบ F3 (Part No.) และ F5 (Dwg No.) ---
        # Excel F3 = index [2, 5], F5 = index [4, 5] (ใน pandas เริ่มนับจาก 0)
        user_part_no = str(df_user_raw.iloc[2, 5])
        user_dwg_no = str(df_user_raw.iloc[4, 5])

        if user_part_no != correct_part_no:
            errors.append(f"❌ ช่อง F3 (Part No.) ไม่ตรงกับ Model: คาดหวัง {correct_part_no} แต่พบ {user_part_no}")
        if user_dwg_no != correct_dwg_no:
            errors.append(f"❌ ช่อง F5 (Dwg No.) ไม่ตรงกับ Model: คาดหวัง {correct_dwg_no} แต่พบ {user_dwg_no}")

        # --- Check 2: ตรวจสอบการกรอกข้อมูลครบถ้วนเทียบกับ Reference ---
        # (สมมติว่าเทียบขนาดและตำแหน่งข้อมูลที่ควรจะมี)
        if df_user_raw.shape[0] < df_ref.shape[0] or df_user_raw.shape[1] < df_ref.shape[1]:
            errors.append("❌ จำนวนแถวหรือคอลัมน์ในไฟล์น้อยกว่าไฟล์ Reference")
        else:
            for r in range(len(df_ref)):
                for c in range(len(df_ref.columns)):
                    if pd.notna(df_ref.iloc[r, c]) and pd.isna(df_user_raw.iloc[r, c]):
                        # แจ้งตำแหน่งแบบ Excel (เช่น A1, B2)
                        col_letter = chr(65 + c)
                        errors.append(f"❌ ข้อมูลไม่ครบที่ตำแหน่ง {col_letter}{r+1}")

        # --- Check 3: ตรวจสอบ Column D12 และ K12 (แถว 13-76) ---
        # Index สำหรับ pandas: Column D=3, K=10 | แถว 13-76 = index 12 ถึง 75
        for row_idx in range(12, 76):
            # ตรวจสอบคอลัมน์ D
            val_d = df_user_raw.iloc[row_idx, 3]
            if not validate_date_format(val_d):
                errors.append(f"❌ รูปแบบวันที่คอลัมน์ D แถว {row_idx+1} ไม่ถูกต้อง (ต้องเป็น M/D/YYYY HH:MM:SS)")
            
            # ตรวจสอบคอลัมน์ K
            val_k = df_user_raw.iloc[row_idx, 10]
            if not validate_date_format(val_k):
                errors.append(f"❌ รูปแบบวันที่คอลัมน์ K แถว {row_idx+1} ไม่ถูกต้อง (ต้องเป็น M/D/YYYY HH:MM:SS)")

        # --- สรุปผลการตรวจสอบ ---
        st.divider()
        if not errors:
            st.success("✅ ไฟล์สมบูรณ์! สามารถใช้งานได้")
        else:
            st.error("พบข้อผิดพลาดในการตรวจสอบ:")
            for err in errors:
                st.write(err)

except Exception as e:
    st.warning("กรุณาตรวจสอบการเชื่อมต่อไฟล์ Reference บน GitHub")
    st.error(f"Error: {e}")
