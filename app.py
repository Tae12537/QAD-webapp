import streamlit as st
import pandas as pd
import os
import io
import re
from openpyxl import load_workbook
import datetime

# ==========================================
# CONFIG & STYLE
# ==========================================
st.set_page_config(page_title="Washing Data Solution Hub", layout="centered")

# CSS ตกแต่งปุ่มเมนูให้ดูน่ากด
st.markdown("""
    <style>
    div.stButton > button:first-child {
        width: 100%;
        height: 80px;
        font-size: 20px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# SESSION STATE MANAGEMENT
# ==========================================
if "current_app" not in st.session_state:
    st.session_state.current_app = "Main Menu"

def go_to_menu():
    # ล้างค่า session เฉพาะที่จำเป็นแต่เก็บ reset_counter ไว้
    temp_counter = st.session_state.get('reset_counter', 0)
    st.session_state.clear()
    st.session_state.reset_counter = temp_counter
    st.session_state.current_app = "Main Menu"
    st.rerun()

# ==========================================
# APP 1: FILE VALIDATOR (Strict Format Check)
# ==========================================
def app_file_validator():
    st.title("📁 File Validator (Strict Format)")
    
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
            uploaded_file = st.file_uploader(f"2️⃣ Upload File to Check", type=["xlsx", "xlsm"], key=f"v_file_{st.session_state.reset_counter}")

            if uploaded_file:
                wb = load_workbook(uploaded_file, data_only=False)
                ws = wb[TARGET_SHEET]
                df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
                
                f_errors, missing_data, extra_data = [], {}, {}
                d_errors, k_errors = [], []

                for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                    ref_v = str(df_ref.iloc[r, c]).strip()
                    user_v = str(df_user.iloc[r, c]).strip()
                    if user_v != ref_v:
                        f_errors.append({"Position": label, "Found": user_v, "Target": ref_v})

                for r in range(76):
                    for c in range(df_ref.shape[1]):
                        if r >= 12 and c in [3, 10]: continue
                        ref_v = str(df_ref.iloc[r, c]).strip() if r < df_ref.shape[0] else ""
                        user_v = str(df_user.iloc[r, c]).strip() if r < df_user.shape[0] else ""
                        if ref_v != "" and (user_v == "" or user_v == "nan"):
                            col = get_column_letter(c); missing_data.setdefault(col, []).append(str(r+1))
                        elif ref_v == "" and (user_v != "" and user_v != "nan"):
                            if r+1 != 12: col = get_column_letter(c); extra_data.setdefault(col, []).append(str(r+1))

                for row_idx in range(13, 77): 
                    for col_idx, (label, errs) in zip([4, 11], [('D', d_errors), ('K', k_errors)]):
                        cell = ws.cell(row=row_idx, column=col_idx)
                        if cell.value is not None:
                            fmt = str(cell.number_format).lower()
                            has_date = ('y' in fmt) or ('d' in fmt and 'm' in fmt)
                            has_time = ('h' in fmt)
                            if not (has_date and has_time):
                                errs.append({"Row": row_idx, "Format": fmt, "Status": "❌ ขาดเวลา"})

                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons(); st.success("✅ ข้อมูลและรูปแบบถูกต้องทั้งหมด!")
                else:
                    if f_errors: st.warning("⚠️ F3/F5 ไม่ตรง"); st.table(pd.DataFrame(f_errors))
                    if missing_data: st.warning("⚠️ Missing Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    if extra_data: st.error("🚫 Extra Data"); st.table([{"Col": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                    st.subheader("⏰ ตรวจสอบรูปแบบวันที่ (D & K)")
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

    file1 = st.file_uploader("📂 Upload File 1 (Lot/Serial)", type=["xls", "xlsx", "csv"], key=f"p1_{st.session_state.p_uploader_key}")
    file2 = st.file_uploader("📂 Upload File 2 (Runcard / Barcode)", type=["xls", "xlsx", "csv"], key=f"p2_{st.session_state.p_uploader_key}")

    def read_excel(file):
        try: return pd.read_excel(file, engine="openpyxl", header=None)
        except: return pd.read_excel(file, engine="xlrd", header=None)

    def read_file1(file):
        df = read_excel(file)
        data = df.iloc[16:, 5]
        lot_list = []
        for val in data:
            if pd.isna(val) or str(val).strip() == "": break
            lot_list.append(str(val).strip())
        return pd.DataFrame({"Lot": lot_list})

    def read_file2(file):
        df = read_excel(file)
        header_row = None
        for i in range(20):
            row = df.iloc[i].astype(str).str.lower()
            if row.str.contains("runcard").any() and row.str.contains("barcode").any():
                header_row = i; break
        if header_row is None: return pd.DataFrame()
        df.columns = df.iloc[header_row]
        df = df[header_row + 1:]
        df.columns = df.columns.astype(str).str.strip().str.lower()
        lot_col = [c for c in df.columns if "runcard" in str(c).lower()][0]
        barcode_col = [c for c in df.columns if "barcode" in str(c).lower()][0]
        packed_col = [c for c in df.columns if "packed" in str(c).lower() and "date" in str(c).lower()][0]
        df_out = df[[lot_col, barcode_col, packed_col]].copy()
        df_out.columns = ["Lot", "Barcode No", "Packed Date"]
        df_out["Packed Date"] = pd.to_datetime(df_out["Packed Date"], errors="coerce")
        return df_out

    if st.button("🚀 Process"):
        if file1 and file2:
            df1 = read_file1(file1)
            df2 = read_file2(file2)
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
            
            def extract(b):
                try:
                    s = str(b)
                    match = re.search('[A-Za-z]', s)
                    if not match: return None, None
                    start = match.start()
                    code = s[start+3:start+6]
                    return (int(code[:2]), int(code[2])) if code.isdigit() else (None, None)
                except: return None, None
            
            merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(extract(x)))
            date_db = pd.read_csv("database.txt")
            date_db["Date"] = pd.to_datetime(date_db["Date"], format="%d-%b-%Y", errors="coerce")
            
            def find_date(row, db):
                cands = db[(db["WW"] == row["WW"]) & (db["Day"] == row["Day"])].copy()
                if cands.empty or pd.isna(row["Packed Date"]): return None
                cands["diff"] = (cands["Date"] - row["Packed Date"]).abs()
                return cands.sort_values("diff").iloc[0]["Date"]

            merged["Washing Date"] = merged.apply(lambda r: find_date(r, date_db), axis=1)
            output = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date"]].copy()
            output["Washing Date"] = pd.to_datetime(output["Washing Date"]).dt.strftime("%d-%b-%Y")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                output.to_excel(writer, index=False, sheet_name="Result")
            
            st.session_state.p_output = output
            st.session_state.p_file = buffer.getvalue()
            st.rerun()

    if st.session_state.get("p_output") is not None:
        st.success("✅ Process สำเร็จ")
        st.dataframe(st.session_state.p_output)
        st.download_button("📥 Download Excel", st.session_state.p_file, "result.xlsx")

# ==========================================
# MAIN ROUTING (MENU)
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<h1 style='text-align: center;'>🏭 Internal System Hub</h1>", unsafe_allow_html=True)
    st.write("---")
    st.subheader("กรุณาเลือกแอปที่ต้องการใช้งาน:")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📁 File Validator\n(ตรวจ Format D/K)"):
            st.session_state.current_app = "Validator"
            st.rerun()
            
    with col2:
        if st.button("📊 Washing Date\nProcessor"):
            st.session_state.current_app = "Processor"
            st.rerun()

elif st.session_state.current_app == "Validator":
    app_file_validator()

elif st.session_state.current_app == "Processor":
    app_washing_processor()
