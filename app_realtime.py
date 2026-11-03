import streamlit as st
import pandas as pd
import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide และใช้ Dark Theme
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS ตกแต่งดีไซน์ใหม่ให้ดูน่าเชื่อถือ เป็นทางการ และพรีเมียมยิ่งขึ้น
st.markdown("""
    <style>
    /* ธีมสีหลักและฟอนต์ */
    .main { 
        background-color: #0b0f19; 
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    /* กล่องหัวข้อหลัก (Header Box) ธีมน้ำเงินกรมท่าพรีเมียม */
    .header-box {
        background: linear-gradient(135deg, #1e3a8a 0%, #172554 100%);
        padding: 20px; 
        border-radius: 12px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 800; 
        font-size: 24px; 
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(30, 58, 138, 0.4);
        border: 1px solid #3b82f6;
        letter-spacing: 0.5px;
    }
    
    /* หัวข้อรอง (Sub-title) สีฟ้าสว่างตัดขอบ */
    .sub-title {
        color: #60a5fa; 
        font-weight: 700; 
        font-size: 18px;
        border-left: 5px solid #3b82f6; 
        padding-left: 10px; 
        margin-top: 25px; 
        margin-bottom: 12px;
        letter-spacing: 0.3px;
    }
    
    /* ตกแต่งการ์ดตัวเลข (Metric Cards) ให้ดูมีมิติ น่าเชื่อถือ */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #1e293b, #0f172a);
        border: 1px solid #334155;
        padding: 18px 22px;
        border-radius: 12px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.4);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(59, 130, 246, 0.25);
        border-color: #60a5fa;
    }
    div[data-testid="stMetric"] label {
        font-size: 15px !important;
        font-weight: 600 !important;
        color: #94a3b8 !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 28px !important;
        font-weight: 800 !important;
        color: #f8fafc !important;
    }
    </style>
""", unsafe_allow_html=True)

# ช่องอัปโหลดไฟล์ Excel
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01 - Schedule Appointment List Report)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # เปิดไฟล์ Excel และอ่านข้อมูลจากแถวหัวตารางจริง (header=3 สำหรับไฟล์รูปแบบใหม่)
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = xls.sheet_names[0]
        
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=3)
        
        # ดึงวันที่จากค
