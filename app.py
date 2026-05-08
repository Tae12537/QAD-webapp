import streamlit as st
import pandas as pd
import io
import re
import os
from openpyxl import load_workbook

# ==========================================
# CONFIG & STYLE
# ==========================================
st.set_page_config(page_title="Production Tools Hub", layout="centered")

# ปุ่มเมนูหน้าแรก
st.markdown("""
    <style>
    div.stButton > button:first-child {
        width: 100%;
        height: 70px;
        font-size: 18px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

if "current_app" not in st.session_state:
    st.session_state.current_app = "Main Menu"

def go_to_menu():
    # ล้าง Session ยกเว้นตัวที่คุม Navigation
    for key in list(st.session_state.keys()):
        if key not in ["current_app", "reset_counter", "p_uploader_key"]:
            del st.session_state[key]
    st.session_state.current_app = "Main Menu"
    st.rerun()

# ==========================================
# APP 1: FILE VALIDATOR (โค้ดตัวล่าสุดของคุณ)
# ==========================================
def app_file_validator():
    st.title("📁 File Validator")
    if st.sidebar.button("⬅️ Exit App (Back to Menu)"):
        go_to_menu()

    if 'reset_counter' not in st.session_state:
        st.session_state.reset_counter = 0

    def reset_app():
        for key in list(st.session_state.keys()):
            if key != "current_app": del st.session_state[key]
        st.session_state.reset_counter = st.session_state.get('reset_counter', 0) + 1
        st.rerun()

    if st.button("🔄 Reset Application"):
        reset_app()

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
        selected_model_name = st.selectbox("1️⃣ Select Model:", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}")

        if selected_model_name != "-- Please Select --":
            ref_filename = available_models[selected_model_name]
            uploaded_file = st.file_uploader(f"2️⃣ Upload File", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}")

            if uploaded_file:
                wb = load_workbook(uploaded_file, data_only=False)
                ws = wb[TARGET_SHEET]
                df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
                
                f_errors, missing_data, extra_data = [], {}, {}
                d_errors, k_errors = [], []

                for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                    if str(df_user.iloc[r, c]).strip() != str(df_ref.iloc[r, c]).strip():
                        f_errors.append({"Position": label, "Found": df_user.iloc[r, c], "Target": df_ref.iloc[r, c]})

                for r in range(76):
                    for c in range(df_ref.shape[1]):
                        if r >= 12 and c in [3, 10]: continue
                        ref_v = str(df_ref.iloc[r, c]).strip()
                        user_v = str(df_user.iloc[r, c]).strip()
                        if ref_v != "" and (user_v == "" or user_v == "nan"):
                            missing_data.setdefault(get_column_letter(c), []).append(str(r+1))
                        elif ref_v == "" and (user_v != "" and user_v != "nan") and r+1 != 12:
                            extra_data.setdefault(get_column_letter(c), []).append(str(r+1))

                for row_idx in range(13, 77): 
                    for col_idx, (col_label, error_list) in enumerate(zip(['D', 'K'], [d_errors, k_errors])):
                        target_col = 4 if col_label == 'D' else 11
                        cell = ws.cell(row=row_idx, column=target_col)
                        if cell.value is not None:
                            fmt = str(cell.number_format).lower()
                            has_date = ('y' in fmt) or ('d' in fmt and 'm' in fmt)
                            has_time = ('h' in fmt)
                            if not (has_date and has_time):
                                error_list.append({"Row": row_idx, "Format": fmt, "Status": "❌ รูปแบบผิด (ต้องมีทั้งวันที่และเวลา)"})

                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons(); st.success("✅ ข้อมูลและรูปแบบถูกต้องทั้งหมด!")
                else:
                    if f_errors: st.warning("⚠️ F3/F5 ไม่ตรง"); st.table(pd.DataFrame(f_errors))
                    if missing_data: st.warning("⚠️ Missing Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    if extra_data: st.error("🚫 Extra Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                    st.subheader("⏰ รูปแบบวันที่และเวลา")
                    if d_errors: st.error("❌ คอลัมน์ D (Washing Date) ผิด"); st.table(pd.DataFrame(d_errors))
                    if k_errors: st.error("❌ คอลัมน์ K (Finish Date) ผิด"); st.table(pd.DataFrame(k_errors))
    except Exception as e: st.error(f"Error: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR (โค้ดตัวที่คุณส่งมาเป๊ะๆ)
# ==========================================
def app_washing_processor():
    st.title("📊 Washing Date Processor")
    if st.sidebar.button("⬅️ Exit App (Back to Menu)"):
        go_to_menu()

    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0

    if st.button("🔄 Reset"):
        st.session_state.output = None
        st.session_state.summary = None
        st.session_state.file = None
        st.session_state.uploader_key += 1
        st.rerun()

    file1 = st.file_uploader("📂 Upload File 1 (Lot/Serial)", type=["xls", "xlsx", "csv"], key=f"file1_{st.session_state.uploader_key}")
    file2 = st.file_uploader("📂 Upload File 2 (Runcard / Barcode)", type=["xls", "xlsx", "csv"], key=f"file2_{st.session_state.uploader_key}")

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
        if header_row is None:
            st.error("❌ หา header ไม่เจอ (Runcard / Barcode)"); return pd.DataFrame()
        df.columns = df.iloc[header_row]
        df = df[header_row + 1:]
        df.columns = df.columns.astype(str).str.strip().str.lower()
        lot_cols = [c for c in df.columns if "runcard" in str(c).lower()]
        barcode_cols = [c for c in df.columns if "barcode" in str(c).lower()]
        if len(lot_cols) == 0 or len(barcode_cols) == 0: return pd.DataFrame()
        packed_col = None
        for c in df.columns:
            if "packed" in str(c).lower() and "date" in str(c).lower(): packed_col = c; break
        if packed_col is None: return pd.DataFrame()
        df_out = df[[lot_cols[0], barcode_cols[0], packed_col]].copy()
        df_out.columns = ["Lot", "Barcode No", "Packed Date"]
        df_out = df_out.dropna(subset=["Lot"])
        df_out["Lot"] = df_out["Lot"].astype(str).str.strip()
        df_out["Packed Date"] = pd.to_datetime(df_out["Packed Date"], errors="coerce")
        return df_out

    def extract_ww_day(barcode):
        try:
            s = str(barcode)
            match = re.search('[A-Za-z]', s)
            if not match: return None, None
            start = match.start()
            code = s[start+3:start+6]
            if len(code) != 3 or not code.isdigit(): return None, None
            return int(code[:2]), int(code[2])
        except: return None, None

    if st.button("🚀 Process"):
        if file1 is None or file2 is None:
            st.warning("⚠️ กรุณาอัพโหลดไฟล์ให้ครบ")
        else:
            df1 = read_file1(file1)
            df2 = read_file2(file2)
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
            merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(extract_ww_day(x)))
            date_db = pd.read_csv("database.txt")
            date_db["Date"] = pd.to_datetime(date_db["Date"], format="%d-%b-%Y", errors="coerce")
            
            def find_best_date(row, db):
                if pd.isna(row["WW"]) or pd.isna(row["Day"]) or pd.isna(row["Packed Date"]): return None
                cands = db[(db["WW"] == row["WW"]) & (db["Day"] == row["Day"])].copy()
                if cands.empty: return None
                cands["diff"] = (cands["Date"] - row["Packed Date"]).abs()
                return cands.sort_values("diff").iloc[0]["Date"]

            merged["Washing Date"] = merged.apply(lambda r: find_best_date(r, date_db), axis=1)
            output = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date"]].copy()
            output["Washing Date"] = pd.to_datetime(output["Washing Date"]).dt.strftime("%d-%b-%Y")
            output = output[output["Lot"].astype(str).str.lower() != "lot/serial"].reset_index(drop=True)
            summary = output.groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                output.to_excel(writer, index=False, sheet_name="Result")
                summary.to_excel(writer, index=False, sheet_name="Summary")
            
            st.session_state.output = output
            st.session_state.summary = summary
            st.session_state.file = buffer.getvalue()

    if "output" in st.session_state and st.session_state.output is not None:
        st.success("✅ Process สำเร็จ")
        st.subheader("📋 Result"); st.dataframe(st.session_state.output)
        st.subheader("📊 Summary"); st.dataframe(st.session_state.summary)
        st.download_button("📥 Download Excel", st.session_state.file, "washing_date_result.xlsx")

# ==========================================
# MAIN ROUTING
# ==========================================
if st.session_state.current_app == "Main Menu":
    st.markdown("<h1 style='text-align: center;'>🏭 QAD System Hub</h1>", unsafe_allow_html=True)
    st.write("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📁 File Validator\n(ตรวจ Format)"):
            st.session_state.current_app = "Validator"; st.rerun()
    with c2:
        if st.button("📊 Washing Date\nProcessor"):
            st.session_state.current_app = "Processor"; st.rerun()
elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
