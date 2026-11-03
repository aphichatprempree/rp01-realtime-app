import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="ระบบรายงานและจัดการคิว Real-time", layout="wide")

st.markdown("<h2 style='text-align: center; background-color: #2c3e50; color: white; padding: 10px;'>ระบบบริหารจัดการคิวและรายงานสินค้า (Real-time Sync)</h2>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01 - Schedule Appointment List Report)", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.success("อัปโหลดและซิงค์ข้อมูลสำเร็จ!")
    st.markdown("### 📊 ตารางแสดงข้อมูลนัดหมายสินค้า")
    st.dataframe(df, use_container_width=True)
    
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col2:
        st.markdown("#### สถานะคิวปัจจุบัน")
        current_queue = st.selectbox("เลือกคิวที่กำลังเรียก:", ["คิวที่ 1: BEE", "คิวที่ 2: DKSH", "คิวที่ 3: CROSSDOCK", "คิวที่ 4: PEPSI"], key="queue_status")
        st.info(f"กำลังเรียกใช้งาน: {current_queue}")
else:
    st.warning("⚠️ กรุณาอัปโหลดไฟล์ Excel (RP01) เพื่อเริ่มใช้งานระบบ Dashboard")
