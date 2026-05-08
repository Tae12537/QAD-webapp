import streamlit as st
import pandas as pd

st.set_page_config(page_title="Debug Data Format", layout="wide")

st.title("🔍 Data Inspector (D & K Column)")
st.write("ลองอัพโหลดไฟล์เพื่อดูว่าระบบ 'เห็น' ข้อมูลของคุณเป็นแบบไหน")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # อ่านไฟล์ 2 แบบเพื่อเปรียบเทียบ
        df_normal = pd.read_excel(uploaded_file, sheet_name="RAMP v1.3", header=None)
        df_string = pd.read_excel(uploaded_file, sheet_name="RAMP v1.3", header=None, dtype=str)

        st.subheader("📊 ผลการวิเคราะห์ข้อมูลแถวที่ 13-15")
        
        debug_list = []
        # ตรวจสอบแถวที่ 13, 14, 15 (Index 12, 13, 14)
        for r in [12, 13, 14]:
            for c, label in zip([3, 10], ['D', 'K']):
                # ข้อมูลแบบปกติ (ที่มักจะโดนแอบแก้)
                val_normal = df_normal.iloc[r, c]
                # ข้อมูลแบบ String (ที่พยายามจะอ่านตรงๆ)
                val_str = str(df_string.iloc[r, c])
                
                debug_list.append({
                    "Row": r + 1,
                    "Column": label,
                    "Raw Value (String Read)": f"'{val_str}'",
                    "Length": len(val_str),
                    "Has Space?": "Yes" if " " in val_str else "No",
                    "Data Type": type(val_normal).__name__
                })
        
        st.table(pd.DataFrame(debug_list))
        
        st.info("""
        **วิธีสังเกต:**
        - ถ้าช่อง D ของคุณ (เช่น 4/5/2026) ขึ้นว่า **Has Space?: No** และความยาวน้อยกว่า 15 -> แสดงว่านับหลักได้
        - แต่ถ้ามันขึ้นว่า **Has Space?: Yes** ทั้งที่คุณไม่ได้พิมพ์ -> แสดงว่า Excel/Pandas แอบเติมเวลาให้ตอนอ่านไฟล์ครับ
        """)

    except Exception as e:
        st.error(f"Error: {e}")
