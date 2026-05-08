import streamlit as st
import pandas as pd
import os

# ตั้งค่าหน้าจอให้อยู่ตรงกลาง (Centered Layout)
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

# --- UI Header ---
st.title("📁 File Validator Before Upload")
if st.button("🔄 Reset Application"):
    reset_app()

st.divider()
TARGET_SHEET = "RAMP v1.3"

try:
    available_models = get_available_models()
    selected_model_name = st.selectbox("1️⃣ Select Model:", ["-- Please Select --"] + sorted(list(available_models.keys())))

    if selected_model_name != "-- Please Select --":
        ref_filename = available_models[selected_model_name]
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File (Target: {selected_model_name})", type=["xlsx", "xlsm"])

        if uploaded_file:
            # อ่านไฟล์โดยไม่ให้ Pandas แอบแปลงวันที่เอง
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl').fillna("")
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []
            user_rows, user_cols = df_user.shape
            ref_rows, ref_cols = df_ref.shape

            # --- 1. Check F3 & F5 (ต้องตรงเป๊ะ) ---
            correct_f3 = str(df_ref.iloc[2, 5]).strip()
            correct_f5 = str(df_ref.iloc[4, 5]).strip()
            user_f3 = str(df_user.iloc[2, 5]).strip() if user_rows > 2 else ""
            user_f5 = str(df_user.iloc[4, 5]).strip() if user_rows > 4 else ""

            if user_f3 != correct_f3: f_errors.append({"Position": "F3", "Found": user_f3, "Target": correct_f3})
            if user_f5 != correct_f5: f_errors.append({"Position": "F5", "Found": user_f5, "Target": correct_f5})

            # --- 2. Check All Cells (Missing & Extra) ---
            # ตรวจสอบขอบเขตแถว 1-76
            check_until_row = 76
            for r in range(max(check_until_row, user_rows)):
                for c in range(max(ref_cols, user_cols)):
                    ref_val = str(df_ref.iloc[r, c]).strip() if (r < ref_rows and c < ref_cols) else ""
                    user_val = str(df_user.iloc[r, c]).strip() if (r < user_rows and c < user_cols) else ""
                    
                    is_ref_empty = ref_val == "" or ref_val.lower() == "nan"
                    is_user_empty = user_val == "" or user_val.lower() == "nan"
                    col_name = get_column_letter(c)

                    # Missing: ต้นฉบับมี แต่เราไม่มี
                    if not is_ref_empty and is_user_empty:
                        if col_name not in missing_data: missing_data[col_name] = []
                        missing_data[col_name].append(str(r+1))
                    
                    # Extra: ต้นฉบับไม่มี แต่เราดันมี
                    elif is_ref_empty and not is_user_empty:
                        if r+1 == 12: continue # ข้าม Header แถว 12
                        if col_name not in extra_data: extra_data[col_name] = []
                        extra_data[col_name].append(str(r+1))

            # --- 3. Check Date/Time (D13-76, K13-76) ---
            for row_idx in range(12, 76): # แถว 13 ถึง 76
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < user_rows:
                        val = str(df_user.iloc[row_idx, col_idx]).strip()
                        
                        if val != "" and val.lower() != "nan":
                            # Logic ใหม่: 
                            # 1. ถ้าความยาวน้อยกว่า 11 (มีแค่วันที่) -> Alarm
                            # 2. ถ้าไม่มีเครื่องหมาย ":" (ไม่มีเวลา) -> Alarm
                            # 3. ถ้าเวลาเป็น 00:00:00 (Excel เติมให้เอง) -> Alarm (เพื่อบังคับให้ผู้ใช้ใส่เวลาจริง)
                            
                            has_time = ":" in val
                            is_midnight = "00:00:00" in val
                            
                            if len(val) < 11 or not has_time or is_midnight:
                                date_errors.append({
                                    "Column": col_label,
                                    "Row": row_idx + 1,
                                    "Found Value": val,
                                    "Status": "❌ ข้อมูลผิดพลาด (กรุณาใส่เวลาให้ถูกต้อง ห้ามเป็น 00:00:00)"
                                })

            # --- 4. Display Results ---
            st.divider()
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons()
                st.success("✅ Perfect! File is ready to upload.")
            else:
                if f_errors:
                    st.warning("⚠️ Critical Info Mismatch (F3/F5)"); st.table(pd.DataFrame(f_errors))
                
                if missing_data:
                    st.warning("⚠️ Missing Data (กรุณากรอกข้อมูลให้ครบ)"); 
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                
                if extra_data:
                    st.error("🚫 Unexpected Extra Data (Reference ว่างแต่มีข้อมูลโผล่มา)"); 
                    st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                
                if date_errors:
                    st.error("⏰ Date/Time Format Error (D, K)"); 
                    st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"System Error: {e}")
