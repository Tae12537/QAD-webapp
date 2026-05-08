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
        df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None, engine='openpyxl')
        
        uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"])

        if uploaded_file:
            # อ่านเป็น String ทั้งหมดและห้ามเติมค่าอัตโนมัติ
            df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None, dtype=str, engine='openpyxl').fillna("")
            
            f_errors, missing_data, extra_data, date_errors = [], {}, {}, []
            user_rows, user_cols = df_user.shape
            ref_rows, ref_cols = df_ref.shape

            # 1. Check F3 / F5
            for r, c, target_name in [(2, 5, "F3"), (4, 5, "F5")]:
                ref_v = str(df_ref.iloc[r, c]).strip()
                user_v = str(df_user.iloc[r, c]).strip()
                if user_v != ref_v:
                    f_errors.append({"Position": target_name, "Found": user_v, "Target": ref_v})

            # 2. Check Missing & Extra (แถว 1-76)
            for r in range(max(76, user_rows)):
                for c in range(max(ref_cols, user_cols)):
                    ref_v = str(df_ref.iloc[r, c]).strip() if (r < ref_rows and c < ref_cols) else ""
                    user_v = str(df_user.iloc[r, c]).strip() if (r < user_rows and c < user_cols) else ""
                    
                    if ref_v != "" and user_v == "":
                        col = get_column_letter(c)
                        if col not in missing_data: missing_data[col] = []
                        missing_data[col].append(str(r+1))
                    elif ref_v == "" and user_v != "":
                        if r+1 == 12: continue # ข้ามหัวตาราง
                        col = get_column_letter(c)
                        if col not in extra_data: extra_data[col] = []
                        extra_data[col].append(str(r+1))

            # 3. Check Date Digit Length (D13-76, K13-76)
            # ถ้าพิมพ์แค่ 4/5/2026 จะยาวประมาณ 8-10 ตัว -> Alarm
            # ถ้าพิมพ์ 11/4/2026 0:00:00 จะยาวประมาณ 17-18 ตัว -> ผ่าน
            for row_idx in range(12, 76): 
                for col_idx, col_label in zip([3, 10], ['D', 'K']):
                    if row_idx < user_rows:
                        val = str(df_user.iloc[row_idx, col_idx]).strip()
                        if val != "" and val.lower() != "nan":
                            # ถ้าน้อยกว่า 15 หลัก (ซึ่งคือระดับที่มีแค่วันที่) ให้ Alarm
                            if len(val) < 15:
                                date_errors.append({
                                    "Column": col_label,
                                    "Row": row_idx + 1,
                                    "Found Value": val,
                                    "Status": "❌ ข้อมูลผิดพลาด (ใส่เวลาไม่ครบ)"
                                })

            # --- Display ---
            if not (f_errors or missing_data or extra_data or date_errors):
                st.balloons(); st.success("✅ Perfect match!")
            else:
                if f_errors: st.warning("⚠️ Mismatch"); st.table(pd.DataFrame(f_errors))
                if missing_data: st.warning("⚠️ Missing"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                if extra_data: st.error("🚫 Extra"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                if date_errors: st.error("⏰ Digit Count Error (D, K)"); st.table(pd.DataFrame(date_errors))

except Exception as e:
    st.error(f"Error: {e}")
