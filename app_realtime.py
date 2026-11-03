import streamlit as st
import pandas as pd
import datetime
import os

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# กำหนดชื่อไฟล์สำหรับเซฟเก็บไว้บน Server เพื่อให้ผู้ใช้คนอื่นเปิดมาเจอข้อมูลเดิมเสมอ
SERVER_SAVE_FILE = "latest_uploaded_rp01.xlsx"

# --- กำหนดค่าเริ่มต้นใน Session State ---
if 'font_size' not in st.session_state:
    st.session_state['font_size'] = 36  # ปรับขนาดเริ่มต้นให้พอดีตาและประหยัดพื้นที่
if 'font_weight' not in st.session_state:
    st.session_state['font_weight'] = "900 (หนาพิเศษ)"
if 'enable_time_filter' not in st.session_state:
    st.session_state['enable_time_filter'] = False
if 'filter_status' not in st.session_state:
    st.session_state['filter_status'] = "ทั้งหมด"
if 'postpone_pallets' not in st.session_state:
    st.session_state['postpone_pallets'] = 0.0

# --- Sidebar: แผงควบคุมตัวกรองข้อมูล คลังข้อมูล และการปรับแต่งการแสดงผล ---
st.sidebar.markdown("<h2>🎛️ ควบคุมการแสดงผล</h2>", unsafe_allow_html=True)

# 1. ส่วนควบคุมขนาดและฟอนต์ (Font Controls)
st.sidebar.markdown("### 🔠 ปรับแต่งตัวอักษรหน้าจอ")
st.session_state['font_size'] = st.sidebar.slider("ขนาดฟอนต์ตารางหลัก (Pixel)", min_value=14, max_value=45, value=st.session_state['font_size'], step=2)
st.session_state['font_weight'] = st.sidebar.selectbox("ความหนาของตัวอักษร", ["900 (หนาพิเศษ)", "800 (หนามาก)", "700 (หนาปานกลาง)"], index=["900 (หนาพิเศษ)", "800 (หนามาก)", "700 (หนาปานกลาง)"].index(st.session_state['font_weight']))
weight_val = st.session_state['font_weight'].split()[0]

# 2. ส่วนอัปโหลดไฟล์ Excel
st.sidebar.markdown("---")
st.sidebar.markdown("### 📂 คลังข้อมูล & ตัวกรอง")
uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel รายงาน (RP01)", type=["xlsx", "xls"])

# หากมีการอัปโหลดไฟล์ใหม่ ให้บันทึกทับลงบน Server ทันที
if uploaded_file is not None:
    try:
        with open(SERVER_SAVE_FILE, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success("✅ อัปโหลดและบันทึกไฟล์ลง Server เรียบร้อย!")
    except Exception as e:
        st.sidebar.error(f"ไม่สามารถบันทึกไฟล์บน Server ได้: {e}")

# ตรวจสอบว่ามีไฟล์เซฟไว้บน Server หรือไม่ (เพื่อให้คนอื่นเข้ามาแล้วเจอข้อมูลเดิม)
active_source = None
if os.path.exists(SERVER_SAVE_FILE):
    active_source = SERVER_SAVE_FILE

# นำค่าที่ผู้ใช้ปรับแต่งมาใส่ลงใน Dynamic CSS ทันที (ลด Padding ให้กระชับขึ้น)
st.markdown(f"""
    <style>
    /* ลดระยะห่างขอบบน-ล่างของหน้าเว็บหลักเพื่อให้เนื้อหาแน่นขึ้น */
    .block-container {{
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
    }}
    
    /* พื้นหลังหลักของหน้าเว็บเป็นสีขาว */
    .stApp {{ 
        background-color: #ffffff !important;
        color: #000000 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }}
    .main {{ 
        background-color: #ffffff !important;
    }}
    
    /* กล่องหัวข้อหลัก (Header Box) ให้กระชับขึ้น */
    .header-box {{
        background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #172554 100%);
        padding: 18px; 
        border-radius: 10px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 900; 
        font-size: 32px; 
        margin-bottom: 12px;
        box-shadow: 0 6px 15px rgba(30, 58, 138, 0.3);
        border: 2px solid #3b82f6;
    }}
    
    /* หัวข้อรอง (Sub-title) */
    .sub-title {{
        color: #1e40af; 
        font-weight: 900; 
        font-size: 26px;
        border-left: 8px solid #1e40af; 
        padding-left: 12px; 
        margin-top: 15px; 
        margin-bottom: 10px;
        height: 38px;
        display: flex;
        align-items: center;
    }}
    
    /* แถบ TOTAL PALLETS */
    .total-box {{
        background: linear-gradient(135deg, #172554 0%, #1e3a8a 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 12px; 
        border-radius: 8px; 
        font-size: 24px; 
        box-shadow: 0 4px 10px rgba(23, 37, 84, 0.25); 
        border: 2px solid #3b82f6;
        margin-top: 10px;
    }}

    /* กล่อง DCSBx และ DCSB */
    .dcsbx-box {{
        background: linear-gradient(135deg, #1e40af 0%, #2563eb 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 10px; 
        border-radius: 8px 8px 0 0; 
        font-size: 22px;
    }}
    .dcsb-box {{
        background: linear-gradient(135deg, #0369a1 0%, #0284c7 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 10px; 
        border-radius: 8px 8px 0 0; 
        font-size: 22px;
    }}
    
    /* === ปรับขนาดฟอนต์เฉพาะตาราง "ภาพรวมสรุปกะการทำงาน" ตามที่ผู้ใช้ตั้งค่าไว้ === */
    div.element-container:has(div.sub-title) + div[data-testid="stDataFrame"] th {{
        font-size: {int(st.session_state['font_size'] * 0.8)}px !important;
        font-weight: {weight_val} !important;
        color: #000000 !important;
        background-color: #cbd5e1 !important;
        padding: 12px !important;
    }}
    div.element-container:has(div.sub-title) + div[data-testid="stDataFrame"] td {{
        font-size: {st.session_state['font_size']}px !important;
        font-weight: {weight_val} !important;
        color: #000000 !important;
        padding: 12px !important;
    }}

    /* ตารางอื่นๆ ให้ขนาดได้สัดส่วนและกระชับขึ้น */
    div[data-testid="stDataFrame"] th {{
        font-size: 18px !important;
        font-weight: 900 !important;
        color: #000000 !important;
        background-color: #e2e8f0 !important;
        padding: 10px !important;
    }}
    div[data-testid="stDataFrame"] td {{
        font-size: 20px !important;
        font-weight: 800 !important;
        color: #000000 !important;
        padding: 10px !important;
    }}
    
    /* ขนาดตัวอักษรในช่องกรอกตัวเลข (Number Input) */
    label[data-baseweb="label"], div[data-baseweb="input"] input {{
        font-size: 20px !important;
        font-weight: 900 !important;
    }}
    div[data-baseweb="input"] input {{
        font-size: 24px !important;
        font-weight: 900 !important;
        height: 48px !important;
    }}
    
    /* ขนาดตัวอักษรในกล่องข้อความแจ้งเตือน */
    div.stAlert p {{
        font-size: 18px !important;
        font-weight: 900 !important;
    }}

    /* ตกแต่งการ์ดตัวเลข (Metric Cards) ให้กระชับ */
    div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #f1f5f9, #e2e8f0);
        border: 2px solid #94a3b8;
        padding: 12px 18px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }}
    div[data-testid="stMetric"] label {{
        font-size: 18px !important;
        font-weight: 900 !important;
        color: #1e293b !important;
    }}
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        font-size: 34px !important;
        font-weight: 900 !important;
        color: #000000 !important;
    }}
    
    /* แถบ Sidebar ด้านข้าง */
    section[data-testid="stSidebar"] {{
        background-color: #f8fafc;
        border-right: 2px solid #e2e8f0;
    }}
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {{
        font-size: 18px !important;
        font-weight: 900 !important;
    }}
    </style>
""", unsafe_allow_html=True)

if active_source is not None:
    try:
        xls = pd.ExcelFile(active_source)
        target_sheet = xls.sheet_names[0]
        df_raw = pd.read_excel(active_source, sheet_name=target_sheet, header=3)
        
        df_raw = df_raw.dropna(how='all')
        if 'Item Name' in df_raw.columns:
            df_raw = df_raw[df_raw['Item Name'].notna()]

        # ตรวจสอบและแปลงวันที่ทั้งหมดที่มีในไฟล์ เพื่อจัดเก็บลงใน Session State (ประวัติย้อนหลัง)
        if 'Delivery Date' in df_raw.columns:
            df_raw['Parsed_Date'] = pd.to_datetime(df_raw['Delivery Date'], errors='coerce', day
