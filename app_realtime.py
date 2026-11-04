import pandas as pd
import streamlit as st

st.subheader("📁 อัปโหลดไฟล์เพื่ออัปเดต Window Time")

# 1. กำหนดสถานะเริ่มต้นใน session_state เพื่อป้องกันข้อมูลหาย
if "main_data" not in st.session_state:
  # สมมติว่านี่คือข้อมูลหลักที่มีอยู่เดิม (หรือโหลดจากฐานข้อมูล/ไฟล์เริ่มต้นของคุณ)
  st.session_state["main_data"] = pd.DataFrame(
      {
          "Vendor ID": ["V001", "V002", "V003"],
          "Order Group": ["Group A", "Group B", "Group C"],
          "Window Time": ["08:00", "10:00", "13:00"],
      }
  )

# 2. ช่องทางอัปโหลดไฟล์
uploaded_file = st.file_uploader(
    "เลือกไฟล์ Excel หรือ CSV (รองรับ .xlsx, .xls, .csv)",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is not None:
  try:
    # อ่านไฟล์ตามนามสกุล
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
        # ทำการ Merge หรืออัปเดตข้อมูลทับข้อมูลเดิมใน session_state
        # เพื่อให้ข้อมูลถูกบันทึกและไม่สูญหายเมื่อมีการใช้งานต่อ
        main_df = st.session_state["main_data"]

        # ตัวอย่างการอัปเดตข้อมูล (Merge ตาม Vendor ID และ Order Group)
        for _, row in df_update.iterrows():
          mask = (main_df["Vendor ID"] == row["Vendor ID"]) & (
              main_df["Order Group"] == row["Order Group"]
          )
          if mask.any():
            main_df.loc[mask, "Window Time"] = row["Window Time"]
          else:
            # ถ้ายังไม่มี ให้เพิ่มเข้าไปใหม่
            new_row = pd.DataFrame([row])
            main_df = pd.concat([main_df, new_row], ignore_index=True)

        st.session_state["main_data"] = main_df
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

# 3. แสดงตารางข้อมูลปัจจุบันในระบบ
st.markdown("---")
st.subheader("📊 ข้อมูล Window Time ปัจจุบันในระบบ")
st.dataframe(st.session_state["main_data"])
