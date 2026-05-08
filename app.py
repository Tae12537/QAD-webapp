import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="File Validator", layout="centered")

def reset_app():
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

def get_column_letter(n):
    result = ""
    while n >= 0:
        result = chr(n % 26 + 65) + result
        n = n // 26 - 1
    return result

def get_available_models():
    files = [f for f in os.listdir('.') if f.startswith('reference_') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
    return {f.replace('reference_', '').split('.')[0]: f for f in files}

st.title("📁 File Validator")
if st.button("🔄 Reset Application"):
    reset_app()

st.divider()
TARGET_SHEET = "RAMP v1.3"

try:
    available_models = get_available_models()
    selected_model_name = st.selectbox("1️⃣ Select Model:", ["-- Please Select --"] + sorted(list(available_models.keys())))

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        
        # อ่าน Reference
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl').fillna("")
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            # *** จุดสำคัญ: บังคับอ่านแบบ RAW โดยไม่แปลงวันที่ (keep_default_na=False) ***
            df_user = pd.read_excel(
                uploaded_file, 
                sheet_name=TARGET_SHEET, 
                header=None, 
                dtype=str, 
                keep_default_na=False, # ห้ามเปลี่ยนค่าว่างเป็น NaN
                engine='openpyxl'
            )
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []

            # 1. Check F3 / F5
            for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

            # 2. Check Missing/Extra (ยกเว้น D, K ในช่วงข้อมูล)
            for r in range(76):
                for c in range(df_ref.shape[1]):
                    if r >= 12 and c in [3, 10]: continue
                    
                    ref_v = str(df_ref.iloc[r, c]).strip() if r < df_ref.shape[0] else ""
                    user_v = str(df_user.iloc[r, c]).strip() if r < df_user.shape[0] else ""
                    
                    if ref_v != "" and user_v == "":
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    elif ref_v == "" and user_v != "":
                        if r+1 == 12: continue
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # 3. Check Date Digit Count (D13-76 และ K13-76)
            for row_idx in range(12, 76): 
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < df_user.shape[0]:
                        val = str(df_user.iloc[row_idx, col_idx]).strip()
                        
                        if val != "":
                            # ถ้าน้อยกว่า 15 หลัก (เช่น '4/5/2026' ยาว 8-10 หลัก) -> Alarm
                            # ถ้ามีเวลาพ่วง '4/5/2026 10:00:00' จะยาว 19 หลัก -> ผ่าน
                            if len(val) < 15:
                                date_errors.append({
                                    "Column": col_label,
                                    "Row": row_idx + 1,
                                    "Value": val,
                                    "Status": "❌ ข้อมูลผิดพลาด (ใส่เวลาไม่ครบ)"
                                })

            # --- Display ---
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ ข้อมูลถูกต้องทั้งหมด!")
            else:
                if f_errors: st.warning("⚠️ Mismatch F3/F5"); st.table(pd.DataFrame(f_errors))
                if missing_data: st.warning("⚠️ Missing Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data: st.error("🚫 Extra Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors: st.error("⏰ พบข้อผิดพลาดในคอลัมน์ D หรือ K"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"Error: {e}")
