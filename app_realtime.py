import pandas as pd
import streamlit as st

st.subheader("📁 อัปโหลดไฟล์เพื่ออัปเดต Window Time")

# 1. ช่องทางอัปโหลดไฟล์
uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel หรือ CSV (รองรับ .xlsx, .xls, .csv)",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
  try:
    # 2. อ่านไฟล์ตามนามสกุล
    if uploaded_file.name.endswith(".csv"):
      df_update = pd.read_csv(uploaded_file)
    else:
      df_update = pd.read_excel(uploaded_file)

    st.write("ตัวอย่างข้อมูลที่อัปโหลด:")
    st.dataframe(df_update.head())

    # ตรวจสอบคอลัมน์ที่จำเป็น
    required_columns = ["Vendor ID", "Order Group", "Window Time"]
    if all(col in df_update.columns for col in required_columns):
      if st.button("ยืนยันการอัปเดต Window Time"):
        # จำลองการนำข้อมูลไปอัปเดตเข้าฐานข้อมูลหรือ Session State หลัก
        # ตัวอย่างเช่น การ Merge กับ DataFrame หลักของระบบ
        # st.session_state['main_data'] = ...

        st.success(
            f"อัปเดตข้อมูล Window Time สำเร็จจำนวน {len(df_update)} รายการ!"
        )
    else:
      st.error(
          "รูปแบบไฟล์ไม่ถูกต้อง กรุณาตรวจสอบชื่อคอลัมน์ให้มี: Vendor ID,"
          " Order Group, Window Time"
      )

  except Exception as e:
    st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์: {e}")
