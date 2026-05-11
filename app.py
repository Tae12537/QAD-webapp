import streamlit as st
import pandas as pd
import io
import re
import os
from openpyxl import load_workbook

# ==========================================
# 🎨 UI CUSTOMIZATION (Electric Blue Theme)
# ==========================================
st.set_page_config(page_title="QAD System Hub", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #eef7ff; }
    div.stButton > button:first-child {
        border-radius: 15px;
        border: 2px solid #007bff;
        background-color: #ffffff;
        color: #007bff;
        height: 80px;
        font-size: 20px;
        font-weight: bold;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(0, 123, 255, 0.1);
    }
    div.stButton > button:hover {
        border: 2px solid #0056b3;
        background-color: #007bff;
        color: #ffffff;
        transform: translateY(-3px);
    }
    .stFileUploader {
        border: 2px dashed #007bff;
        border-radius: 12px;
        background-color: #ffffff;
    }
    [data-testid="stSidebar"] {
        background-color: #d1e9ff;
        border-right: 2px solid #007bff;
    }
    h1, h2, h3 { color: #004085; }
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
# APP 1: FILE VALIDATOR (Updated Logic)
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
        selected_model_name = st.selectbox("1️⃣ Select Model", model_list, index=0, key=f"v_sel_{st.session_state.reset_counter}")

        if selected_model_name != "-- Please Select --":
            ref_filename = available_models[selected_model_name]
            uploaded_file = st.file_uploader("2️⃣ Upload File", type=["xlsx", "xlsm"], key=f"v_up_{st.session_state.reset_counter}")

            if uploaded_file:
                wb = load_workbook(uploaded_file, data_only=False)
                ws = wb[TARGET_SHEET]
                df_ref = pd.read_excel(ref_filename, sheet_name=TARGET_SHEET, header=None).fillna("")
                df_user = pd.read_excel(uploaded_file, sheet_name=TARGET_SHEET, header=None).fillna("")
                
                f_errors, missing_data, extra_data, d_errors, k_errors = [], {}, {}, [], []

                # ตรวจ F3, F5
                for r, c, label in [(2, 5, "F3"), (4, 5, "F5")]:
                    if str(df_user.iloc[r, c]).strip() != str(df_ref.iloc[r, c]).strip():
                        f_errors.append({"Position": label, "Found": df_user.iloc[r, c], "Target": df_ref.iloc[r, c]})

                # ตรวจ Missing/Extra Data
                for r in range(76):
                    for c in range(df_ref.shape[1]):
                        if r >= 12 and c in [3, 10]: continue
                        ref_v, user_v = str(df_ref.iloc[r, c]).strip(), str(df_user.iloc[r, c]).strip()
                        if ref_v != "" and (user_v == "" or user_v == "nan"):
                            missing_data.setdefault(get_column_letter(c), []).append(str(r+1))
                        elif ref_v == "" and (user_v != "" and user_v != "nan") and r+1 != 12:
                            extra_data.setdefault(get_column_letter(c), []).append(str(r+1))

                # ตรวจ Column D และ K (Logic แก้ไขใหม่)
                for row_idx in range(13, 77): 
                    for col_idx, (col_label, error_list) in enumerate(zip(['D', 'K'], [d_errors, k_errors])):
                        target_col = 4 if col_label == 'D' else 11
                        cell = ws.cell(row=row_idx, column=target_col)
                        
                        if cell.value is not None:
                            # 1. เช็ค Format เบื้องต้น
                            fmt = str(cell.number_format).lower()
                            has_format = (('y' in fmt or ('d' in fmt and 'm' in fmt)) and 'h' in fmt)
                            
                            # 2. เช็ค "ไส้ใน" ข้อมูลว่ามีเวลาติดมาจริงหรือไม่
                            val = cell.value
                            has_time_data = False
                            try:
                                dt_val = pd.to_datetime(val)
                                # ถ้า hour, minute, second เป็น 0 หมด แสดงว่าไม่มีเวลาติดมา
                                if dt_val.hour == 0 and dt_val.minute == 0 and dt_val.second == 0:
                                    has_time_data = False
                                else:
                                    has_time_data = True
                            except:
                                has_time_data = False

                            if not has_format or not has_time_data:
                                reason = "❌ รูปแบบผิด" if not has_format else "❌ ไม่มีข้อมูลเวลา (Time is 00:00:00)"
                                error_list.append({"Row": row_idx, "Value": str(val), "Status": reason})

                st.markdown("### 📋 Result")
                if not (f_errors or missing_data or extra_data or d_errors or k_errors):
                    st.balloons(); st.success("✅ ข้อมูลและรูปแบบถูกต้องทั้งหมด!")
                else:
                    if f_errors: st.warning("⚠️ F3/F5 ไม่ตรง"); st.table(pd.DataFrame(f_errors))
                    if missing_data: st.info("⚠️ Missing Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in missing_data.items()])
                    if extra_data: st.error("🚫 Extra Data"); st.table([{"Column": k, "Rows": ", ".join(v)} for k, v in extra_data.items()])
                    
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("**Column D (Washing Date)**")
                        if d_errors: st.dataframe(pd.DataFrame(d_errors))
                        else: st.write("✅ ปกติ")
                    with c2:
                        st.markdown("**Column K (Finish Date)**")
                        if k_errors: st.dataframe(pd.DataFrame(k_errors))
                        else: st.write("✅ ปกติ")
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

    file1 = st.file_uploader("📂 Upload File 1", type=["xls", "xlsx", "csv"], key=f"p1_{st.session_state.uploader_key}")
    file2 = st.file_uploader("📂 Upload File 2", type=["xls", "xlsx", "csv"], key=f"p2_{st.session_state.uploader_key}")

    def read_excel(file):
        try: return pd.read_excel(file, engine="openpyxl", header=None)
        except: return pd.read_excel(file, engine="xlrd", header=None)

    if st.button("🚀 Process"):
        if not file1 or not file2:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ให้ครบ")
        else:
            # Logic การอ่านไฟล์และ Process (เหมือนเดิม)
            df1_raw = read_excel(file1)
            df1 = pd.DataFrame({"Lot": [str(v).strip() for v in df1_raw.iloc[16:, 5] if not pd.isna(v) and str(v).strip() != ""]})
            
            df2_raw = read_excel(file2)
            h_row = next((i for i in range(20) if df2_raw.iloc[i].astype(str).str.lower().str.contains("runcard").any()), None)
            if h_row is not None:
                df2_raw.columns = df2_raw.iloc[h_row]
                df2 = df2_raw[h_row + 1:].copy()
                df2.columns = df2.columns.astype(str).str.strip().str.lower()
                lot_c = [c for c in df2.columns if "runcard" in c][0]
                bar_c = [c for c in df2.columns if "barcode" in c][0]
                pak_c = [c for c in df2.columns if "packed" in c and "date" in c][0]
                
                df2_final = df2[[lot_c, bar_c, pak_c]].dropna(subset=[lot_c])
                df2_final.columns = ["Lot", "Barcode No", "Packed Date"]
                df2_final["Packed Date"] = pd.to_datetime(df2_final["Packed Date"], errors="coerce")
                
                merged = pd.merge(df1, df2_final, on="Lot", how="left").drop_duplicates(subset=["Lot"])
                
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
                sum_df = out.groupby("Washing Date")["Lot"].count().reset_index().rename(columns={"Lot": "Total Lot"})
                
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    out.to_excel(writer, index=False, sheet_name="Result")
                    sum_df.to_excel(writer, index=False, sheet_name="Summary")
                st.session_state.output, st.session_state.summary, st.session_state.file = out, sum_df, buf.getvalue()

    if st.session_state.get("output") is not None:
        st.success("✅ Process สำเร็จ")
        st.dataframe(st.session_state.output)
        st.subheader("📊 Summary")
        st.dataframe(st.session_state.summary)
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
