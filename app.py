import streamlit as st
import pandas as pd
import os
import io
import re
from openpyxl import load_workbook
import datetime

# 1. Config Layout ให้เล็กลงอยู่ตรงกลาง
st.set_page_config(page_title="Internal Solution Hub", layout="centered")

st.markdown("""
    <style>
    div.stButton > button:first-child {
        width: 100%;
        height: 60px;
        font-size: 18px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

if "current_app" not in st.session_state:
    st.session_state.current_app = "Main Menu"

def go_to_menu():
    for key in list(st.session_state.keys()):
        if key != "reset_counter":
            del st.session_state[key]
    st.session_state.current_app = "Main Menu"
    st.rerun()

# ==========================================
# APP 1: FILE VALIDATOR
# ==========================================
def app_file_validator():
    st.title("📁 File Validator")
    if st.sidebar.button("⬅️ Back to Main Menu"):
        go_to_menu()

    if 'reset_counter' not in st.session_state:
        st.session_state.reset_counter = 0

    def reset_validator():
        st.session_state.reset_counter += 1
        st.rerun()

    if st.button("🔄 Reset This App"):
        reset_validator()

    st.divider()
    TARGET_SHEET = "RAMP v1.3"

    def get_column_letter(n):
        result = ""
        while n >= 0:
            result = chr(n % 26 + 65) + result
            n = n // 26 - 1
        return result

    def get_available_models():
        files = [f for f in os.listdir('.') if f.startswith('reference_') and (f.endswith('.xlsx') or f.endswith('.xlsm'))]
        return {f.replace('reference_', '').split('.')[0]: f for f in files}

    try:
        available_models = get_available_models()
        model_list = ["-- Please Select --"] + sorted(list(available_models.keys()))
        selected_model_name = st.selectbox("1️⃣ Select Model:", model_list, index=0, key=f"v_model_{st.session_state.reset_counter}")

        if selected_model_name != "-- Please Select --":
            ref_filename = available_models[selected_model_name]
            uploaded_file = st.file_uploader("2️⃣ Upload File to Check", type=["xlsx", "xlsm"], key=f"v_file_{st.session_state.reset_counter}")

            if uploaded_file:
                wb = load_workbook(uploaded_file, data_only=False)
                ws = wb[TARGET_SHEET]
                df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
                
                f_errors, missing_data, extra_data = [], {}, {}
                d_errors, k_errors = [], []

                # Check F3/F5
                for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                    if str(df_user.iloc[r, c]).strip() != str(df_ref.iloc[r, c]).strip():
                        f_errors.append({"Position": label, "Found": df_user.iloc[r, c], "Target": df_ref.iloc[r, c]})

                # Check Missing/Extra
                for r in range(76):
                    for c in range(df_ref.shape[1]):
                        if r >= 12 and c in [3, 10]: continue
                        ref_v = str(df_ref.iloc[r, c]).strip()
                        user_v = str(df_user.iloc[r, c]).strip()
                        if ref_v != "" and (user_v == "" or user_v == "nan"):
                            missing_data.setdefault(get_column_letter(c), []).append(str(r+1))
                        elif ref_v == "" and (user_v != "" and user_v != "nan") and r+1 != 12:
                            extra_data.setdefault(get_column_letter(c), []).append(str(r+1))

                # Strict Format Check
                for row_idx in range(13, 77): 
                    for col_idx, (lbl, err_list) in zip([4, 11], [('D', d_errors), ('K', k_errors)]):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if cell.value is not None:
                            fmt = str(cell.number_format).lower()
                            if not (('y' in fmt or ('d' in fmt and 'm' in fmt)) and 'h' in fmt):
                                err_list.append({"Row": row_idx, "Format": fmt, "Status": "❌ ขาดเวลา"})

                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons(); st.success("✅ ข้อมูลและรูปแบบถูกต้องทั้งหมด!")
                else:
                    if f_errors: st.warning("⚠️ F3/F5 ไม่ตรง"); st.table(pd.DataFrame(f_errors))
                    if missing_data: st.warning("⚠️ Missing Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    if extra_data: st.error("🚫 Extra Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                    st.subheader("⏰ รูปแบบวันที่ (D & K)")
                    if d_errors: st.error("❌ คอลัมน์ D ผิด"); st.table(pd.DataFrame(d_errors))
                    if k_errors: st.error("❌ คอลัมน์ K ผิด"); st.table(pd.DataFrame(k_errors))
    except Exception as e:
        st.error(f"Error: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR
# ==========================================
def app_washing_processor():
    st.title("📊 Washing Date Processor")
    if st.sidebar.button("⬅️ Back to Main Menu"):
        go_to_menu()

    if "p_uploader_key" not in st.session_state:
        st.session_state.p_uploader_key = 0

    if st.button("🔄 Reset This App"):
        st.session_state.p_output = None
        st.session_state.p_summary = None
        st.session_state.p_uploader_key += 1
        st.rerun()

    file1 = st.file_uploader("📂 Upload File 1 (Lot/Serial)", type=["xlsx", "csv"], key=f"p1_{st.session_state.p_uploader_key}")
    file2 = st.file_uploader("📂 Upload File 2 (Runcard)", type=["xlsx", "csv"], key=f"p2_{st.session_state.p_uploader_key}")

    if st.button("🚀 Process"):
        if file1 and file2:
            try:
                # Read File 1
                df1_raw = pd.read_excel(file1, header=None)
                df1 = pd.DataFrame({"Lot": [str(x).strip() for x in df1_raw.iloc[16:, 5] if pd.notna(x) and str(x).strip() != ""]})
                
                # Read File 2
                df2_raw = pd.read_excel(file2, header=None)
                h_row = next(i for i in range(20) if "runcard" in str(df2_raw.iloc[i]).lower() and "barcode" in str(df2_raw.iloc[i]).lower())
                df2 = df2_raw.iloc[h_row+1:].copy()
                df2.columns = [str(c).strip().lower() for c in df2_raw.iloc[h_row]]
                
                l_col = [c for c in df2.columns if "runcard" in c][0]
                b_col = [c for c in df2.columns if "barcode" in c][0]
                p_col = [c for c in df2.columns if "packed" in c and "date" in c][0]
                
                df2_clean = df2[[l_col, b_col, p_col]].copy()
                df2_clean.columns = ["Lot", "Barcode No", "Packed Date"]
                df2_clean["Lot"] = df2_clean["Lot"].astype(str).str.strip()
                df2_clean["Packed Date"] = pd.to_datetime(df2_clean["Packed Date"], errors="coerce")
                
                # Merge & Extract
                merged = pd.merge(df1, df2_clean, on="Lot", how="left").drop_duplicates(subset=["Lot"])
                def ext(b):
                    m = re.search('[A-Za-z]', str(b))
                    if m:
                        code = str(b)[m.start()+3:m.start()+6]
                        if code.isdigit(): return int(code[:2]), int(code[2])
                    return None, None
                merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(ext(x)))
                
                # Database & Best Date
                db = pd.read_csv("database.txt")
                db["Date"] = pd.to_datetime(db["Date"], format="%d-%b-%Y", errors="coerce")
                def find_d(r, db):
                    c = db[(db["WW"] == r["WW"]) & (db["Day"] == r["Day"])].copy()
                    if c.empty or pd.isna(r["Packed Date"]): return None
                    c["diff"] = (c["Date"] - r["Packed Date"]).abs()
                    return c.sort_values("diff").iloc[0]["Date"]
                
                merged["Washing Date"] = merged.apply(lambda r: find_d(r, db), axis=1)
                
                # Final Output
                res = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date"]].copy()
                res["Washing Date"] = pd.to_datetime(res["Washing Date"]).dt.strftime("%d-%b-%Y")
                
                # ✅ SUMMARY FUNCTION กลับมาแล้ว
                summary = res.groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})
                
                # Create Excel Buffer
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    res.to_excel(writer, index=False, sheet_name="Result")
                    summary.to_excel(writer, index=False, sheet_name="Summary")
                
                st.session_state.p_output = res
                st.session_state.p_summary = summary
                st.session_state.p_file = buf.getvalue()
                st.rerun()
            except Exception as e:
                st.error(f"Error during process: {e}")

    if st.session_state.get("p_output") is not None:
        st.success("✅ Process สำเร็จ")
        st.subheader("📋 Result")
        st.dataframe(st.session_state.p_output)
        
        # ✅ แสดงตาราง Summary บนหน้าเว็บ
        st.subheader("📊 Summary")
        st.dataframe(st.session_state.p_summary)
        
        st.download_button("📥 Download Excel (Result + Summary)", st.session_state.p_file, "washing_date_report.xlsx")

# ==========================================
# MAIN MENU
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<h1 style='text-align: center;'>🏭 Internal System Hub</h1>", unsafe_allow_html=True)
    st.write("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📁 File Validator\n(ตรวจ Format D/K)"):
            st.session_state.current_app = "Validator"; st.rerun()
    with c2:
        if st.button("📊 Washing Date\nProcessor"):
            st.session_state.current_app = "Processor"; st.rerun()

elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
