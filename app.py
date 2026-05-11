import streamlit as st
import pandas as pd
import io
import re
import os
from openpyxl import load_workbook

# ==========================================
# 🎨 UI CUSTOMIZATION (Electric Blue Theme)
# ==========================================
st.set_page_config(page_title="Production Tools Hub", layout="centered")

st.markdown("""
    <style>
    /* พื้นหลังสีฟ้าอ่อน */
    .stApp {
        background-color: #eef7ff;
    }
    
    /* แต่งปุ่ม Main Menu ให้เป็นสีฟ้าตะโกน */
    div.stButton > button:first-child {
        border-radius: 15px;
        border: 2px solid #007bff;
        background-color: #ffffff;
        color: #007bff;
        height: 100px;
        font-size: 22px;
        font-weight: bold;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(0, 123, 255, 0.1);
    }
    
    /* Hover Effect */
    div.stButton > button:hover {
        border: 2px solid #0056b3;
        background-color: #007bff;
        color: #ffffff;
        transform: translateY(-3px);
    }

    /* แต่งกรอบ File Uploader */
    .stFileUploader {
        border: 2px dashed #007bff;
        border-radius: 12px;
        background-color: #ffffff;
    }

    /* Sidebar สีฟ้า */
    [data-testid="stSidebar"] {
        background-color: #d1e9ff;
        border-right: 2px solid #007bff;
    }

    /* หัวข้อสีฟ้าเข้ม */
    h1, h2, h3 {
        color: #004085;
    }

    /* ตารางขอบฟ้า */
    .stDataFrame {
        border: 1px solid #cce5ff;
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

if "current_app" not in st.session_state:
    st.session_state.current_app = "Main Menu"

def go_to_menu():
    for key in list(st.session_state.keys()):
        if key not in ["current_app", "reset_counter", "uploader_key"]:
            del st.session_state[key]
    st.session_state.current_app = "Main Menu"
    st.rerun()

# ==========================================
# APP 1: FILE VALIDATOR
# ==========================================
def app_file_validator():
    st.markdown("<h2 style='text-align: center;'>📁 File Validator</h2>", unsafe_allow_html=True)
    st.sidebar.markdown("### Menu")
    if st.sidebar.button("⬅️ Exit App"):
        go_to_menu()

    if 'reset_counter' not in st.session_state:
        st.session_state.reset_counter = 0

    def reset_app():
        for key in list(st.session_state.keys()):
            if key != "current_app": del st.session_state[key]
        st.session_state.reset_counter = st.session_state.get('reset_counter', 0) + 1
        st.rerun()

    st.sidebar.write("---")
    if st.sidebar.button("🔄 Reset This App"):
        reset_app()

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
        st.markdown("#### 1️⃣ Select Model")
        selected_model_name = st.selectbox("Select Model", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}", label_visibility="collapsed")

        if selected_model_name != "-- Please Select --":
            ref_filename = available_models[selected_model_name]
            st.markdown("#### 2️⃣ Upload File")
            uploaded_file = st.file_uploader("Upload", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}")

            if uploaded_file:
                wb = load_workbook(uploaded_file, data_only=False)
                ws = wb[TARGET_SHEET]
                df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
                
                f_errors, missing_data, extra_data, d_errors, k_errors = [], {}, {}, [], []

                for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                    if str(df_user.iloc[r, c]).strip() != str(df_ref.iloc[r, c]).strip():
                        f_errors.append({"Position": label, "Found": df_user.iloc[r, c], "Target": df_ref.iloc[r, c]})

                for r in range(76):
                    for c in range(df_ref.shape[1]):
                        if r >= 12 and c in [3, 10]: continue
                        ref_v, user_v = str(df_ref.iloc[r, c]).strip(), str(df_user.iloc[r, c]).strip()
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
                            if not (('y' in fmt or ('d' in fmt and 'm' in fmt)) and 'h' in fmt):
                                error_list.append({"Row": row_idx, "Format": fmt, "Status": "❌ รูปแบบผิด"})

                st.markdown("### 📋 Result")
                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons(); st.success("✅ ข้อมูลถูกต้องทั้งหมด!")
                else:
                    if f_errors: st.warning("⚠️ F3/F5 ไม่ตรง"); st.table(pd.DataFrame(f_errors))
                    if missing_data: st.info("⚠️ Missing Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    if extra_data: st.error("🚫 Extra Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**Column D**"); st.dataframe(pd.DataFrame(d_errors)) if d_errors else st.write("✅ ปกติ")
                    with c2:
                        st.markdown("**Column K**"); st.dataframe(pd.DataFrame(k_errors)) if k_errors else st.write("✅ ปกติ")
    except Exception as e: st.error(f"Error: {e}")

# ==========================================
# APP 2: WASHING DATE PROCESSOR
# ==========================================
def app_washing_processor():
    st.markdown("<h2 style='text-align: center;'>📊 Washing Date Processor</h2>", unsafe_allow_html=True)
    if st.sidebar.button("⬅️ Exit App"):
        go_to_menu()
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = 0
    if st.sidebar.button("🔄 Reset"):
        st.session_state.output, st.session_state.summary, st.session_state.file = None, None, None
        st.session_state.uploader_key += 1
        st.rerun()

    st.markdown("#### 📂 Upload Files")
    file1 = st.file_uploader("Upload File 1", type=["xls", "xlsx", "csv"], key=f"p1_{st.session_state.uploader_key}")
    file2 = st.file_uploader("Upload File 2", type=["xls", "xlsx", "csv"], key=f"p2_{st.session_state.uploader_key}")

    def read_excel(file):
        try: return pd.read_excel(file, engine="openpyxl", header=None)
        except: return pd.read_excel(file, engine="xlrd", header=None)

    def read_file1(file):
        df = read_excel(file)
        data = df.iloc[16:, 5]
        return pd.DataFrame({"Lot": [str(v).strip() for v in data if not pd.isna(v) and str(v).strip() != ""]})

    def read_file2(file):
        df = read_excel(file)
        header_row = next((i for i in range(20) if df.iloc[i].astype(str).str.lower().str.contains("runcard").any() and df.iloc[i].astype(str).str.lower().str.contains("barcode").any()), None)
        if header_row is None: return pd.DataFrame()
        df.columns = df.iloc[header_row]
        df = df[header_row + 1:]
        df.columns = df.columns.astype(str).str.strip().str.lower()
        lot_cols, barcode_cols = [c for c in df.columns if "runcard" in str(c)], [c for c in df.columns if "barcode" in str(c)]
        p_col = next((c for c in df.columns if "packed" in str(c) and "date" in str(c)), None)
        if not lot_cols or not barcode_cols or not p_col: return pd.DataFrame()
        df_out = df[[lot_cols[0], barcode_cols[0], p_col]].copy()
        df_out.columns = ["Lot", "Barcode No", "Packed Date"]
        df_out["Packed Date"] = pd.to_datetime(df_out["Packed Date"], errors="coerce")
        return df_out.dropna(subset=["Lot"])

    if st.button("🚀 Process"):
        if not file1 or not file2:
            st.warning("⚠️ อัปโหลดไฟล์ไม่ครบ")
        else:
            df1, df2 = read_file1(file1), read_file2(file2)
            merged = pd.merge(df1, df2, on="Lot", how="left").drop_duplicates(subset=["Lot"])
            def ext_ww(b):
                m = re.search('[A-Za-z]', str(b))
                if not m: return None, None
                c = str(b)[m.start()+3:m.start()+6]
                return (int(c[:2]), int(c[2])) if (len(c)==3 and c.isdigit()) else (None, None)
            merged[['WW', 'Day']] = merged['Barcode No'].apply(lambda x: pd.Series(ext_ww(x)))
            db = pd.read_csv("database.txt")
            db["Date"] = pd.to_datetime(db["Date"], format="%d-%b-%Y", errors="coerce")
            def find_d(r, d_db):
                if pd.isna(r["WW"]) or pd.isna(r["Day"]) or pd.isna(r["Packed Date"]): return None
                c = d_db[(d_db["WW"] == r["WW"]) & (d_db["Day"] == r["Day"])].copy()
                if c.empty: return None
                c["diff"] = (c["Date"] - r["Packed Date"]).abs()
                return c.sort_values("diff").iloc[0]["Date"]
            merged["Washing Date"] = merged.apply(lambda r: find_d(r, db), axis=1)
            out = merged[["Lot", "Barcode No", "WW", "Day", "Washing Date"]].copy()
            out["Washing Date"] = pd.to_datetime(out["Washing Date"]).dt.strftime("%d-%b-%Y")
            out = out[out["Lot"].astype(str).str.lower() != "lot/serial"].reset_index(drop=True)
            sum_df = out.groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                out.to_excel(writer, index=False, sheet_name="Result")
                sum_df.to_excel(writer, index=False, sheet_name="Summary")
            st.session_state.output, st.session_state.summary, st.session_state.file = out, sum_df, buf.getvalue()

    if st.session_state.get("output") is not None:
        st.success("✅ สำเร็จ!"); st.dataframe(st.session_state.output)
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
        if st.button("📁 File Validator\n(ตรวจ Format ไฟล์)"):
            st.session_state.current_app = "Validator"; st.rerun()
    with c2:
        if st.button("📊 Washing Date\nProcessor"):
            st.session_state.current_app = "Processor"; st.rerun()
elif st.session_state.current_app == "Validator": app_file_validator()
elif st.session_state.current_app == "Processor": app_washing_processor()
