import streamlit as st
import pandas as pd
import datetime
import os

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide พร้อมซ่อนแถบเมนูด้านบนของ Streamlit (Header & Toolbar)
st.set_page_config(
    page_title="รายงานประมาณการรับสินค้า", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# กำหนดชื่อไฟล์สำหรับเซฟเก็บไว้บน Server เพื่อให้ผู้ใช้คนอื่นเปิดมาเจอข้อมูลเดิมเสมอ
SERVER_SAVE_FILE = "latest_uploaded_rp01.xlsx"

# --- ฟังก์ชันอ่านไฟล์ Excel พร้อมระบบ Caching (@st.cache_data) เพื่อความเร็วสูงสุด ---
@st.cache_data(show_spinner=False)
def load_excel_data(file_path):
    xls = pd.ExcelFile(file_path)
    target_sheet = xls.sheet_names[0]
    df_raw = pd.read_excel(file_path, sheet_name=target_sheet, header=3)
    
    df_raw = df_raw.dropna(how='all')
    if 'Item Name' in df_raw.columns:
        df_raw = df_raw[df_raw['Item Name'].notna()]
    return df_raw

# --- กำหนดค่าเริ่มต้นใน Session State สำหรับแยกส่วนควบคุม ---
if 'dcs_font_size' not in st.session_state:
    st.session_state['dcs_font_size'] = 22  # ขนาดฟอนต์ตารางบน (DCSBx / DCSB)
if 'vendor_font_size' not in st.session_state:
    st.session_state['vendor_font_size'] = 20  # ขนาดฟอนต์ตารางล่าง (Vendors CARD & COINS)
if 'enable_time_filter' not in st.session_state:
    st.session_state['enable_time_filter'] = False
if 'filter_status' not in st.session_state:
    st.session_state['filter_status'] = "ทั้งหมด"
if 'postpone_in_pallets' not in st.session_state:
  st.session_state['postpone_in_pallets'] = 0.0
if 'postpone_out_pallets' not in st.session_state:
  st.session_state['postpone_out_pallets'] = 0.0

dcs_size = st.session_state['dcs_font_size']
vendor_size = st.session_state['vendor_font_size']

# นำค่าฟอนต์ที่ปรับแต่งมาใส่ใน Dynamic CSS พร้อมปรับแต่งสีตัวหนังสือใน Sidebar ให้สว่างและชัดเจน
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700;800&display=swap');

<style>
@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;500;600;700;800&display=swap');

footer {{
    visibility: hidden !important;
}}

    footer {{
        visibility: hidden !important;
    }}

    .block-container {{
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }}
    .stApp {{ 
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: 'Prompt', sans-serif !important;
    }}
    .main {{ 
        background-color: #f8fafc !important;
    }}
    
    /* กรอบหัวข้อหลัก (Header Box) */
    .header-box {{
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #1e40af 100%);
        padding: 22px 24px 20px 24px; 
        border-radius: 12px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 800; 
        font-size: 30px; 
        margin-bottom: 14px;
        box-shadow: 0 6px 15px -4px rgba(37, 99, 235, 0.4);
        border: 2px solid rgba(255, 255, 255, 0.2);
        line-height: 1.5;
    }}
    
    /* หัวข้อรอง (Sub-title) */
    .sub-title {{
        color: #1e40af; 
        font-weight: 800; 
        font-size: 24px;
        border-left: 8px solid #2563eb; 
        padding-left: 12px; 
        margin-top: 10px; 
        margin-bottom: 8px;
        height: 34px;
        display: flex;
        align-items: center;
        background: linear-gradient(90deg, rgba(37, 99, 235, 0.08) 0%, rgba(255, 255, 255, 0) 100%);
        border-radius: 0 8px 8px 0;
    }}

    /* กล่อง DCSBx และ DCSB */
    .dcsbx-box {{
        background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
        color: white; 
        text-align: center; 
        font-weight: 800; 
        padding: 10px; 
        border-radius: 8px 8px 0 0; 
        font-size: 22px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }}
    .dcsb-box {{
        background: linear-gradient(135deg, #0284c7 0%, #38bdf8 100%);
        color: white; 
        text-align: center; 
        font-weight: 800; 
        padding: 10px; 
        border-radius: 8px 8px 0 0; 
        font-size: 22px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }}

    /* === ขยายขนาดตัวหนังสือยอด PO (Delta) ในการ์ดให้ใหญ่สะใจ === */
    div[data-testid="stMetric"] div[data-testid="stMetricDelta"] {{
        font-size: 18px !important;
        font-weight: 900 !important;
    }}

    /* === ปรับขนาดฟอนต์ตารางด้านบน (DCSBx / DCSB) ตาม Sidebar === */
    div.element-container:has(iframe) dataframe, div[data-testid="stDataFrame"] {{
        font-size: {dcs_size}px !important;
    }}
    
    [data-testid="stDataFrame"] div[data-baseweb="tab"] {{
        font-size: {dcs_size}px !important;
    }}

    .stDataFrame [data-testid="StyledMarkupTable"] td, .stDataFrame [data-testid="StyledMarkupTable"] th {{
        font-size: {dcs_size}px !important;
    }}

    /* === เอฟเฟกต์การ์ดกะดึกและกะเช้าแบบมีมิติสูง (3D Shadow) === */
    @keyframes pulse-night {{
        0% {{ box-shadow: 0 0 0 0 rgba(245, 158, 11, 0.4); }}
        70% {{ box-shadow: 0 0 0 10px rgba(245, 158, 11, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(245, 158, 11, 0); }}
    }}
    @keyframes pulse-morning {{
        0% {{ box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.4); }}
        70% {{ box-shadow: 0 0 0 10px rgba(34, 197, 94, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(34, 197, 94, 0); }}
    }}

    /* การ์ดกะดึก */
    div.element-container:has(div.metric-night) div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #fffbeb, #fef3c7) !important;
        border: 2px solid #f59e0b !important;
        box-shadow: 0 6px 15px rgba(245, 158, 11, 0.25) !important;
        animation: pulse-night 2.5s infinite;
        border-radius: 10px !important;
        padding: 10px 14px !important;
    }}
    div.element-container:has(div.metric-night) div[data-testid="stMetric"] label {{
        color: #92400e !important;
    }}
    div.element-container:has(div.metric-night) div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        color: #b45309 !important;
    }}

    /* การ์ดกะเช้า */
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #f0fdf4, #dcfce7) !important;
        border: 2px solid #22c55e !important;
        box-shadow: 0 6px 15px rgba(34, 197, 94, 0.25) !important;
        animation: pulse-morning 2.5s infinite;
        border-radius: 10px !important;
        padding: 10px 14px !important;
    }}
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] label {{
        color: #166534 !important;
    }}
    div.element-container:has(div.metric-morning) div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        color: #15803d !important;
    }}
    
    /* การ์ด Metric ทั่วไป */
    div[data-testid="stMetric"] {{
        background: linear-gradient(145deg, #ffffff, #f1f5f9);
        border: 2px solid #cbd5e1;
        padding: 10px 14px;
        border-radius: 10px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.05);
    }}
    div[data-testid="stMetric"] label {{
        font-size: 16px !important;
        font-weight: 800 !important;
        color: #334155 !important;
    }}
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        font-size: 28px !important;
        font-weight: 900 !important;
        color: #0f172a !important;
    }}
    
    /* === ปรับแต่งแถบ Sidebar ด้านข้างให้สมดุลและคมชัด === */
    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
        border-right: 3px solid #38bdf8;
        padding: 0.8rem 0.6rem;
    }}
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] span {{
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #f8fafc !important;
        text-align: center !important;
    }}
    section[data-testid="stSidebar"] h3 {{
        font-size: 15px !important;
        font-weight: 700 !important;
        color: #38bdf8 !important;
        margin-top: 10px !important;
        margin-bottom: 4px !important;
        letter-spacing: 0.5px;
        border-bottom: 1px solid rgba(56, 189, 248, 0.3);
        padding-bottom: 4px;
        text-align: center !important;
    }}
    
    /* บังคับให้ข้อความหัวข้อย่อยและป้ายกำกับใน Sidebar เป็นสีขาวสว่างทั้งหมด */
    section[data-testid="stSidebar"] .stFileUploader label, 
    section[data-testid="stSidebar"] p {{
        color: #f8fafc !important;
        text-align: center !important;
    }}
    
    /* กรอบบล็อกเครื่องมือใน Sidebar พร้อมเส้นขอบคมชัด */
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"] {{
        background: rgba(255, 255, 255, 0.06);
        padding: 12px 14px;
        border-radius: 8px;
        border: 2px solid #475569;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3);
        margin-bottom: 10px;
        text-align: center;
    }}

    /* ตกแต่งกล่องอัปโหลดไฟล์ให้ตัวหนังสือชัดเจนและจัดกึ่งกลาง */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] {{
        text-align: center !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] small, 
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] span {{
        color: #334155 !important;
        font-weight: 700 !important;
    }}
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {{
        background-color: #ffffff !important;
        border: 2px dashed #0284c7 !important;
        border-radius: 8px;
        padding: 10px !important;
    }}

    @media print {{
        section[data-testid="stSidebar"], button, .stDownloadButton {{
            display: none !important;
        }}
        .stApp {{
            background-color: #ffffff !important;
        }}
    }}
    </style>
""", unsafe_allow_html=True)

# --- Sidebar: แผงควบคุมตัวกรองข้อมูล คลังข้อมูล และการปรับแต่งการแสดงผล ---
st.sidebar.markdown("""
    <div style="text-align: center; padding: 10px 0;">
        <h2 style='color: #ffffff; font-size: 18px; font-weight: 800; margin: 0;'>DATA ENTRY (DCSB)</h2>
        <p style='color: #38bdf8; font-size: 12px; font-weight: 600; margin-top: 4px;'>📊 ระบบรายงานประมาณการรับสินค้า</p>
    </div>
    <hr style="border-color: #38bdf8; margin: 5px 0 15px 0;">
""", unsafe_allow_html=True)

st.sidebar.markdown("<h2 style='color: #ffffff; font-size: 18px; font-weight: 800; border-bottom: 3px solid #38bdf8; padding-bottom: 6px; margin-bottom: 12px; text-align: center;'>🎛️ แผงควบคุมระบบ</h2>", unsafe_allow_html=True)

# 1. ส่วนคลังข้อมูล & อัปโหลดไฟล์
st.sidebar.markdown("### 📂 1. จัดการข้อมูลไฟล์")
uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        with open(SERVER_SAVE_FILE, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.cache_data.clear()
        st.sidebar.success("✅ อัปโหลดสำเร็จ!")
    except Exception as e:
        st.sidebar.error(f"ไม่สามารถบันทึก: {e}")

active_source = None
if os.path.exists(SERVER_SAVE_FILE):
    active_source = SERVER_SAVE_FILE

st.sidebar.markdown("---")

# 2. ส่วนเลือกวันที่ดูข้อมูลย้อนหลัง
st.sidebar.markdown("### 📅 2. เลือกวันที่ดูย้อนหลัง")

if active_source is not None:
    try:
        df_raw = load_excel_data(active_source)
        
        if 'Delivery Date' in df_raw.columns:
            df_raw['Parsed_Date'] = pd.to_datetime(df_raw['Delivery Date'], errors='coerce', dayfirst=True)
            
            if 'historical_data' not in st.session_state:
                st.session_state['historical_data'] = {}
                
            for dt_val, group_df in df_raw.groupby(df_raw['Parsed_Date'].dt.date):
                if pd.notna(dt_val):
                    st.session_state['historical_data'][str(dt_val)] = group_df

        if 'historical_data' in st.session_state and st.session_state['historical_data']:
            available_dates = sorted(list(st.session_state['historical_data'].keys()))
            selected_date_str = st.sidebar.selectbox("เลือกวันที่", available_dates, index=len(available_dates)-1, label_visibility="collapsed")
            df = st.session_state['historical_data'][selected_date_str].copy()
        else:
            df = df_raw.copy()
            selected_date_str = str(datetime.date.today())

    except Exception as e:
        st.sidebar.error(f"อ่านไฟล์ไม่สำเร็จ: {e}")
        df = pd.DataFrame()
        selected_date_str = str(datetime.date.today())
else:
    df = pd.DataFrame()
    selected_date_str = str(datetime.date.today())

st.sidebar.markdown("---")

# 3. ส่วนกรองประเภท Supplier & ช่วงเวลา
st.sidebar.markdown("### 🏷️ 3. ตัวกรองเพิ่มเติม")
status_options = ["ทั้งหมด", "STAMP", "CARD", "COINS"]
st.session_state['filter_status'] = st.sidebar.selectbox("ประเภท Supplier", status_options, index=status_options.index(st.session_state['filter_status']), label_visibility="collapsed")
filter_status = st.session_state['filter_status']

st.session_state['enable_time_filter'] = st.sidebar.checkbox("⏰ กรองตามช่วงเวลา", value=st.session_state['enable_time_filter'])
if st.session_state['enable_time_filter']:
    col_t1, col_t2 = st.sidebar.columns(2)
    with col_t1:
        start_t = st.time_input("ตั้งแต่", datetime.time(2, 0), label_visibility="collapsed")
    with col_t2:
        end_t = st.time_input("ถึง", datetime.time(5, 0), label_visibility="collapsed")
    
    time_col_filter = next((col for col in df.columns if 'time' in str(col).lower() or 'slot' in str(col).lower()), None)
    if time_col_filter and not df.empty:
        df['Parsed_Time'] = pd.to_datetime(df[time_col_filter], errors='coerce').dt.time
        df = df[(df['Parsed_Time'] >= start_t) & (df['Parsed_Time'] <= end_t)]

st.sidebar.markdown("---")

# 4. ส่วนจัดการงานเลื่อนส่ง (แยกอิสระ: เลื่อนเข้าบวกเพิ่ม / เลื่อนออกหักลบ พร้อมบันทึกความจำ 7 วัน)
st.sidebar.markdown("### ⏳ 4. จัดการงานเลื่อนส่ง (เข้า/ออก)")

# สร้างตัวแปรเก็บประวัติแยกตามวันที่ใน Session State (ถ้ายังไม่มี)
if 'postpone_records' not in st.session_state:
  st.session_state['postpone_records'] = {}

# ฟอร์มบันทึกยอดงานเลื่อนส่งประจำวัน
with st.sidebar.form(key='save_postpone_form', clear_on_submit=True):
  p_target_date = st.date_input(
      "เลือกวันที่ต้องการบันทึก", value=datetime.date.today()
  )
  p_in_val = st.number_input(
      "📥 ยอดเลื่อนเข้า (บวกเพิ่ม +)", min_value=0.0, step=1.0, value=0.0
  )
  p_out_val = st.number_input(
      "📤 ยอดเลื่อนออก (หักลบ -)", min_value=0.0, step=1.0, value=0.0
  )
  p_note = st.text_input("หมายเหตุ (เช่น เลื่อนเข้า/ออก)", value="")
  submit_btn = st.form_submit_button("💾 บันทึกเก็บไว้")

if submit_btn and (p_in_val > 0 or p_out_val > 0):
  date_key = str(p_target_date)
  st.session_state['postpone_records'][date_key] = {
      'in_pallets': p_in_val,
      'out_pallets': p_out_val,
      'note': p_note,
      'timestamp': datetime.datetime.now(),
  }
  st.sidebar.success(
      f'บันทึกสำเร็จ! (เข้า: +{p_in_val:,.0f} | ออก: -{p_out_val:,.0f})'
  )

# --- ระบบเคลียร์ข้อมูลอัตโนมัติ (เก็บรักษาไว้ไม่เกิน 7 วัน) ---
current_time = datetime.datetime.now()
keys_to_delete = []
for d_str, info in st.session_state['postpone_records'].items():
  record_date = datetime.datetime.strptime(d_str, '%Y-%m-%d')
  days_diff = (current_time - record_date).days
  if days_diff > 7 or days_diff < -1:
    keys_to_delete.append(d_str)

for k in keys_to_delete:
  del st.session_state['postpone_records'][k]

# --- ดึงข้อมูลงานเลื่อนส่งของวันที่กำลังเลือกดูมาใช้งานอัตโนมัติ ---
selected_key = str(selected_date_str)
if selected_key in st.session_state['postpone_records']:
  saved_data = st.session_state['postpone_records'][selected_key]
  postpone_in_pallets = saved_data.get('in_pallets', 0.0)
  postpone_out_pallets = saved_data.get('out_pallets', 0.0)
  postpone_from_str = saved_data['note']
  st.sidebar.info(
      f"📌 ข้อมูลวันที่ {selected_key}: 📥 +{postpone_in_pallets:,.0f} | 📤"
      f" -{postpone_out_pallets:,.0f} ({saved_data['note']})"
  )
else:
  postpone_in_pallets = 0.0
  postpone_out_pallets = 0.0
  postpone_from_str = selected_date_str

st.sidebar.markdown("---")

# 5. ส่วนปรับขนาดตัวอักษร
st.sidebar.markdown("### 🔠 5. ปรับขนาดฟอนต์ตาราง")
col_fs1, col_fs2 = st.sidebar.columns(2)
with col_fs1:
    st.session_state['dcs_font_size'] = st.slider("ตารางบน", min_value=16, max_value=40, value=st.session_state['dcs_font_size'], step=2)
with col_fs2:
    st.session_state['vendor_font_size'] = st.slider("ตารางล่าง", min_value=16, max_value=35, value=st.session_state['vendor_font_size'], step=2)

# --- ส่วนประมวลผลและแสดงผลหน้าจอหลัก ---
if active_source is not None and not df.empty:
    try:
        try:
            dt_obj = datetime.datetime.strptime(selected_date_str, "%Y-%m-%d")
            months_th = {1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'เม.ย.', 5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.', 9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'}
            current_date = f"{dt_obj.day:02d} {months_th[dt_obj.month]} {str(dt_obj.year + 543)[-2:]}"
        except:
            current_date = selected_date_str
        
        st.markdown(f'<div class="header-box">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</div>', unsafe_allow_html=True)
        if st.session_state['enable_time_filter']:
            st.info(f"🔍 กำลังแสดงผลข้อมูลเฉพาะช่วงเวลา: **{start_t.strftime('%H:%M')} - {end_t.strftime('%H:%M')}**")

        st.markdown("---")
        
        if postpone_in_pallets > 0 or postpone_out_pallets > 0:
          st.warning(
              f"⚠️ มีการจัดการงานเลื่อนส่ง (เข้า: +{postpone_in_pallets:,.0f} พาเลท |"
              f" ออก: -{postpone_out_pallets:,.0f} พาเลท) ({postpone_from_str})"
          )
        else:
          st.info("ℹ️ กำลังแสดงยอดพาเลทตามแผนปกติ (ไม่มีการเลื่อนส่งในวันนี้)")

        sup_col = 'Vendor Name' if 'Vendor Name' in df.columns else (next((col for col in df.columns if 'vendor' in str(col).lower() or 'supplier' in str(col).lower()), df.columns[4]))
        total_sup = df[sup_col].dropna().nunique() if sup_col in df.columns else len(df)

        item_col = 'Item Name' if 'Item Name' in df.columns else (next((col for col in df.columns if 'item' in str(col).lower()), df.columns[12]))
        total_item = df[item_col].dropna().nunique() if item_col in df.columns else len(df)

        po_col_raw = next((c for c in df.columns if 'p/o' in str(c).lower() or 'po' in str(c).lower()), None)
        if not po_col_raw:
            po_col_raw = 'P/O' if 'P/O' in df.columns else df.columns[1]

        if po_col_raw in df.columns:
            df['PO_clean'] = pd.to_numeric(df[po_col_raw], errors='coerce').fillna(df[po_col_raw]).astype(str).str.split('.').str[0].str.strip()
            total_po = df['PO_clean'].dropna().nunique()
        else:
            total_po = len(df)

        uom_col = 'UOM Qty' if 'UOM Qty' in df.columns else df.columns[17]
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[21]

        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        raw_pallets = pd.to_numeric(
            df[pallets_col], errors="coerce"
        ).fillna(0).sum()
        total_pallets = max(
            0.0,
            raw_pallets + postpone_in_pallets - postpone_out_pallets,
        )
        
        target_night = 3879
        target_morning = 1663
        total_target = target_night + target_morning
        
        ratio_night = target_night / total_target
        ratio_morning = target_morning / total_target

        night_pallets = round(total_pallets * ratio_night, 2)
        morning_pallets = round(total_pallets - night_pallets, 2)

        night_po = round(total_po * ratio_night)
        morning_po = total_po - night_po
        
        night_pct = (night_pallets / total_pallets) * 100 if total_pallets > 0 else 0
        morning_pct = (morning_pallets / total_pallets) * 100 if total_pallets > 0 else 0

        asrs_col = 'ASRS_ITEM' if 'ASRS_ITEM' in df.columns else (next((c for c in df.columns if 'asrs' in str(c).lower()), 'AE'))
        
        if asrs_col in df.columns:
            df['ASRS_match'] = df[asrs_col].astype(str).str.strip().str.upper()
            dcsbx_df_filtered = df[df['ASRS_match'] == 'ASRS_ITEM']
            dcsb_df_filtered = df[df['ASRS_match'] != 'ASRS_ITEM']
        else:
            dcsbx_df_filtered = df.iloc[:int(len(df)*0.4)]
            dcsb_df_filtered = df.iloc[int(len(df)*0.4):]

        dcsbx_item = dcsbx_df_filtered[item_col].dropna().nunique() if item_col in dcsbx_df_filtered.columns else len(dcsbx_df_filtered)
        dcsbx_pallets = pd.to_numeric(dcsbx_df_filtered[pallets_col], errors='coerce').fillna(0).sum()
        dcsbx_po = dcsbx_df_filtered['PO_clean'].dropna().nunique() if 'PO_clean' in dcsbx_df_filtered.columns else len(dcsbx_df_filtered)
        dcsbx_pct = (dcsbx_pallets / total_pallets) * 100 if total_pallets > 0 else 0

        dcsb_item = dcsb_df_filtered[item_col].dropna().nunique() if item_col in dcsb_df_filtered.columns else len(dcsb_df_filtered)
        dcsb_pallets = pd.to_numeric(dcsb_df_filtered[pallets_col], errors='coerce').fillna(0).sum()
        dcsb_po = dcsb_df_filtered['PO_clean'].dropna().nunique() if 'PO_clean' in dcsb_df_filtered.columns else len(dcsb_df_filtered)
        dcsb_pct = (dcsb_pallets / total_pallets) * 100 if total_pallets > 0 else 0

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            st.metric("Total Sup.", f"{total_sup:,}")
        with c2:
            st.metric("Item", f"{total_item:,}")
        with c3:
            st.metric("UOM Qty", f"{total_uom:,.0f}")
        with c4:
          if postpone_in_pallets > 0:
            delta_desc = f"+{postpone_in_pallets:,.0f} ปรับเข้า"
          elif postpone_out_pallets > 0:
            delta_desc = f"-{postpone_out_pallets:,.0f} ปรับออก"
          else:
            delta_desc = None

          st.metric(
              "Net Pallets", f"{total_pallets:,.0f}", delta=delta_desc, delta_color="inverse")
            
        with c5:
            st.markdown('<div class="metric-night"></div>', unsafe_allow_html=True)
            st.metric("🌙 กะดึก (05:00-17:00)", f"{night_pallets:,.0f}", f"{night_po:,} PO ({night_pct:.1f}%)")
        with c6:
            st.markdown('<div class="metric-morning"></div>', unsafe_allow_html=True)
            st.metric("☀️ กะเช้า (17:00-05:00)", f"{morning_pallets:,.0f}", f"{morning_po:,} PO ({morning_pct:.1f}%)")

        st.markdown('<div class="sub-title">(DCSBx / DCSB)</div>', unsafe_allow_html=True)
        sub_c1, sub_c2 = st.columns(2)
        with sub_c1:
            st.markdown("<div class='dcsbx-box'>DCSBx (ASRS_ITEM)</div>", unsafe_allow_html=True)
            dcsbx_table_df = pd.DataFrame({
                "Date": [current_date], 
                "Item": [dcsbx_item], 
                "Palletes": [dcsbx_pallets], 
                "PO": [dcsbx_po], 
                "%": [f"{dcsbx_pct:.2f}%"]
            })
            st.dataframe(dcsbx_table_df, use_container_width=True, hide_index=True)
        with sub_c2:
            st.markdown("<div class='dcsb-box'>DCSB (Normal)</div>", unsafe_allow_html=True)
            dcsb_table_df = pd.DataFrame({
                "Date": [current_date], 
                "Item": [dcsb_item], 
                "Palletes": [dcsb_pallets], 
                "PO": [dcsb_po], 
                "%": [f"{dcsb_pct:.2f}%"]
            })
            st.dataframe(dcsb_table_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.markdown('<div class="sub-title">Vendors STAMP, CARD & COINS (Special Suppliers)</div>', unsafe_allow_html=True)
        
        stamp_vendors = ["70979", "21189" , "2055418"]
        card_vendors = [
            "592", "3751", "3924", "12457", "12790", "12810", "12922", "12990", "13725", "17157",
            "21265", "21541", "21653", "21738", "40334", "40352", "40678", "41068", "70082"
        ]
        coins_vendors = [
            "70335", "70341", "70585", "70605", "70609", "70757", "71007", "71026", "71316", "71338",
            "71573", "71758", "72204", "72215", "72472", "72715", "72736", "72747", "72852", "72853",
            "73204", "73208", "1101530", "FS", "73005", "1102038", "2052496", "2053616",
        ]
        
        v_col = next((c for c in df.columns if 'vendor' in str(c).lower() and 'name' not in str(c).lower()), None)
        if not v_col:
            v_col = 'Vendor' if 'Vendor' in df.columns else df.columns[3]

        if v_col in df.columns and po_col_raw in df.columns:
            df['Vendor_str'] = df[v_col].astype(str).str.split('.').str[0].str.strip()
            
            special_df = df[df['Vendor_str'].isin(stamp_vendors + card_vendors + coins_vendors)].copy()
            
            if not special_df.empty:
                def get_status(v_id):
                    if v_id in stamp_vendors:
                        return 'STAMP'
                    elif v_id in card_vendors:
                        return 'CARD'
                    else:
                        return 'COINS'
                
                special_df['Status_val'] = special_df['Vendor_str'].apply(get_status)
                
                if filter_status != "ทั้งหมด":
                    special_df = special_df[special_df['Status_val'] == filter_status]
                
                unique_special = special_df.drop_duplicates(subset=['PO_clean'])
                vendor_table = unique_special[['Vendor_str', sup_col, 'PO_clean', pallets_col, uom_col, 'Status_val']].copy()
                vendor_table.columns = ["Vendor ID", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty", "Status"]
                
                st.dataframe(vendor_table, use_container_width=True, hide_index=True)
            else:
                st.info("ℹ️ ไม่พบรายการ Supplier พิเศษ (STAMP/CARD/COINS) ในเงื่อนไขตัวกรองหรือวันที่เลือกนี้")
        else:
            st.warning("⚠️ ไม่พบคอลัมน์ Vendor หรือ P/O ในไฟล์รีพอร์ต")
        
        postpone_display_str = (
            f"ปรับเข้า: {postpone_in_pallets:,.0f} |"
            f" ปรับออก: {postpone_out_pallets:,.0f} ({postpone_from_str})"
        )

        st.markdown(f"""
            <div style='background: linear-gradient(135deg, #047857 0%, #065f46 100%); color: white; padding: 14px 22px; border-radius: 8px; font-weight: 800; display: flex; justify-content: space-between; margin-top: 10px; font-size: 24px; box-shadow: 0 4px 12px rgba(4, 120, 87, 0.3); border: 2px solid #10b981;'>
                <span>Grand Total (Net Pallets)</span>
                <span>Pallets: {total_pallets:,.0f} ({postpone_display_str})</span>
                <span>UOM Qty: {total_uom:,.0f}</span>
            </div>
        """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 📥 บันทึกและพิมพ์รายงาน")
        
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        
        with col_dl1:
            summary_df_export = pd.DataFrame({
                "Date": [current_date],
                "Total Sup.": [total_sup],
                "Item": [total_item],
                "UOM Qty": [f"{total_uom:,.0f}"],
                "Net Pallets": [f"{total_pallets:,.0f}"]
            })
            csv_data = summary_df_export.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label=f"📥 ดาวน์โหลดตารางสรุป (CSV)",
                data=csv_data,
                file_name=f"report_summary_{current_date.replace(' ', '_')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        with col_dl2:
            html_filename = f"report_dashboard_{current_date.replace(' ', '_')}.html"
            html_content = f"""<html>
<head>
    <meta charset='utf-8'>
    <title>Report {current_date}</title>
    <style>
        body {{ background-color: #ffffff; color: #000000; font-family: 'Prompt', sans-serif; padding: 20px; }}
        h2 {{ color: #1e40af; text-align: center; font-size: 28px; font-weight: 800; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th, td {{ border: 2px solid #94a3b8; padding: 10px; text-align: center; font-size: 18px; font-weight: 800; }}
        th {{ background-color: #1e3a8a; color: white; }}
        td {{ background-color: #f1f5f9; color: #000000; }}
    </style>
</head>
<body>
    <h2>📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</h2>
    <p style="font-size: 18px; font-weight: bold; text-align: center;"><b>Total Sup:</b> {total_sup:,} | <b>Item:</b> {total_item:,} | <b>UOM Qty:</b> {total_uom:,.0f} | <b>Net Pallets:</b> {total_pallets:,.0f} ({postpone_display_str})</p>
    <br>
    <div style="text-align: center;">
        <button onclick="window.print()" style="background-color: #2563eb; color: white; padding: 12px 24px; font-size: 18px; font-weight: bold; border: none; border-radius: 8px; cursor: pointer;">🖨️ พิมพ์หน้านี้ / บันทึกเป็น PDF</button>
    </div>
</body>
</html>"""
            st.download_button(
                label=f"📄 ดาวน์โหลดไฟล์หน้าเว็บ (HTML)",
                data=html_content.encode('utf-8-sig'),
                file_name=html_filename,
                mime="text/html",
                use_container_width=True
            )

        with col_dl3:
            st.markdown(f"""
                <div style="display: flex; align-items: center; height: 100%;">
                    <button onclick="window.print()" style="width: 100%; background: linear-gradient(135deg, #2563eb, #1e40af); color: white; padding: 10px 16px; font-size: 18px; font-weight: 800; border: none; border-radius: 8px; cursor: pointer; box-shadow: 0 4px 10px rgba(37,99,235,0.3);">
                        🖨️ พิมพ์/บันทึก PDF ทันที
                    </button>
                </div>
            """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("📂 กรุณาอัปโหลดไฟล์ Excel (RP01) ที่แถบด้านข้าง (Sidebar) ระบบจะทำการบันทึกข้อมูลไว้บน Server เพื่อให้ทุกคนเปิดมาเจอหน้าจอและข้อมูลเดิมเสมอครับ")
    # --- ส่วนที่เพิ่มเข้ามา: แสดงตารางข้อมูล Vendor, PO, Item, Ti, Hi และระบบแจ้งเตือน ---
st.markdown("---")
st.subheader("📋 ตรวจสอบเงื่อนไข Ti และ Hi (Ti = 1 และ Hi = 1)")

# สมมติว่าดึง DataFrame หลักของหน้าจอคุณมาใช้ (เช่น st.session_state['main_data'] หรือตัวแปรตารางของคุณ)
# หากในโค้ดเดิมของคุณมีตัวแปรตารางอยู่แล้ว สามารถเปลี่ยนชื่อตัวแปร df ให้ตรงกันได้เลยครับ
if "main_data" in st.session_state:
  df_check = st.session_state["main_data"]
else:
  # หากต้องการดึงข้อมูลจากตารางที่มีอยู่จริง ให้แทนที่ด้วยตัวแปร DataFrame ของระบบคุณ
  df_check = df  # (ตัวอย่างอ้างอิงจากตัวแปร df ในระบบ)

# ตรวจสอบว่ามีคอลัมน์ Vendor, PO, Item, Ti, Hi ครบถ้วนหรือไม่
required_cols = ["Vendor", "PO", "Item", "Ti", "Hi"]
if all(col in df_check.columns for col in required_cols):

  # 1. แสดงผลตารางข้อมูลทั้ง 5 คอลัมน์
  st.dataframe(df_check[required_cols], use_container_width=True)

  # 2. ตรวจสอบเงื่อนไขแจ้งเตือน เมื่อ Ti = 1 และ Hi = 1 พร้อมกัน
  alert_items = df_check[(df_check["Ti"] == 1) & (df_check["Hi"] == 1)]

  if not alert_items.empty:
    st.warning(
        "🚨 **แจ้งเตือนระบบ:** ตรวจพบรายการที่มีค่า **Ti = 1** และ **Hi ="
        " 1** พร้อมกัน จำนวน "
        f"{len(alert_items)} รายการ!"
    )

    # แสดงรายละเอียดเฉพาะรายการที่เข้าเงื่อนไขเตือน เพื่อให้ตรวจสอบง่าย
    with st.expander(
        "🔍 คลิกเพื่อดูรายละเอียดรายการที่ต้องตรวจสอบ (Ti=1 และ Hi=1)"
    ):
      st.dataframe(alert_items[required_cols], use_container_width=True)
  else:
    st.success("✅ ไม่พบรายการที่ค่า Ti = 1 และ Hi = 1 พร้อมกันในขณะนี้")
else:
  st.info(
      "ℹ️ กรุณาตรวจสอบชื่อคอลัมน์ในฐานข้อมูลว่ามี 'Vendor', 'PO', 'Item',"
      " 'Ti', 'Hi' ครบถ้วน"
  )
import streamlit as st
import streamlit.components.v1 as components

# --- ส่วนเพิ่มปุ่มพิมพ์/บันทึก PDF ---
st.markdown("---")
import io
import pandas as pd
import streamlit as st

# --- ส่วนเพิ่มพิเศษ: ตรวจสอบงานนอกรอบ (เปรียบเทียบ PO แบบ VLOOKUP และรวมยอดพาเลท) ---
st.markdown("---")
st.markdown(
    '<div class="sub-title">🔍 ตรวจสอบงานนอกรอบ (เปรียบเทียบ PO แบบ VLOOKUP'
    " และรวมยอดพาเลท)</div>",
    unsafe_allow_html=True,
)

rp03_file = st.file_uploader(
    "📂 อัปโหลดไฟล์ RP03 (Goods Receive Note Summary Report)",
    type=["xlsx", "xls", "csv"],
    key="unique_rp03_vlookup_uploader_final",
)

if rp03_file is not None:
  try:
    # 1. อ่านไฟล์ RP03 ตามนามสกุล
    if rp03_file.name.endswith(".csv"):
      df_rp03 = pd.read_csv(rp03_file)
    else:
      df_rp03 = pd.read_excel(rp03_file)

    st.success("อัปโหลดไฟล์ RP03 สำเร็จ!")

    # 2. ดึงข้อมูลตารางหลัก (RP01) จากระบบ หรือ session_state
    if "df" in st.session_state and st.session_state["df"] is not None:
      df_main = st.session_state["df"]
    elif "df" in locals() and df is not None and not df.empty:
      df_main = df
    else:
      df_main = None

    if df_main is not None and not df_main.empty:
      # ค้นหาชื่อคอลัมน์ Purchase Order และ Total Case Qty ในไฟล์ RP03
      po_col_rp03 = next(
          (
              c
              for c in df_rp03.columns
              if any(
                  k in c.lower() for k in ["purchase order", "po", "order"]
              )
          ),
          None,
      )

      pallet_col = next(
          (
              c
              for c in df_rp03.columns
              if any(
                  k in c.lower()
                  for k in ["total case", "ilpn", "pallet", "qty"]
              )
          ),
          None,
      )

      # ค้นหาคอลัมน์ PO ในตารางหลัก (RP01)
      main_po_col = next(
          (
              c
              for c in df_main.columns
              if any(k in c.lower() for k in ["po", "order"])
          ),
          None,
      )

      # ตรวจสอบความถูกต้องว่าพบคอลัมน์ครบถ้วนหรือไม่
      if po_col_rp03 and pallet_col and main_po_col:
        df_rp03[po_col_rp03] = df_rp03[po_col_rp03].astype(str).str.strip()
        existing_pos = set(df_main[main_po_col].astype(str).str.strip())

        # VLOOKUP เปรียบเทียบหาตัวที่ไม่มีใน RP01 (เทียบเท่า #N/A หรือ งานนอกรอบ)
        df_mismatch = df_rp03[~df_rp03[po_col_rp03].isin(existing_pos)].copy()

        if not df_mismatch.empty:
          df_mismatch[pallet_col] = pd.to_numeric(
              df_mismatch[pallet_col], errors="coerce"
          ).fillna(0)
          total_sum_pallets = df_mismatch[pallet_col].sum()

          # แสดงผลสรุปยอดพาเลทงานนอกรอบ
          st.metric(
              label=(
                  "📦 ยอดรวม Total Case Qty (iLPNs) ของงานนอกรอบ (PO ไม่ตรงกับ"
                  " RP01)"
              ),
              value=(
                  f"{total_sum_pallets:,.2f} พาเลท (จากทั้งหมด"
                  f" {len(df_mismatch)} รายการ)"
              ),
          )

          st.subheader("📋 รายละเอียดรายการงานนอกรอบทั้งหมด (เทียบเท่า #N/A)")
          st.dataframe(df_mismatch, use_container_width=True)

          # สร้างไฟล์ Excel สำหรับกดดาวน์โหลด
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_mismatch.to_excel(
                writer, index=False, sheet_name="RP03_NA_Summary"
            )
          processed_data = output.getvalue()

          st.download_button(
              label="📥 ดาวน์โหลดรายงานงานนอกรอบ (Excel)",
              data=processed_data,
              file_name="RP03_Mismatch_N_A_Report.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              key="unique_download_mismatch_btn_final_v2",
          )
          st.success("✅ ไม่พบงานนอกรอบ ทุก PO ใน RP03 ตรงกับ RP01 ทั้งหมด")
        st.warning(
            "⚠️ ไม่พบคอลัมน์ Purchase Order หรือคอลัมน์ Total Case Qty (iLPNs)"
            f" ในไฟล์ RP03 (พบ PO: {po_col_rp03}, Qty: {pallet_col})"
        )
      st.warning("⚠️ กรุณามีข้อมูลหลัก (RP01) ในระบบก่อนทำการอัปโหลดเพื่อเทียบข้อมูล")
  except Exception as e:
     st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")
# --- เพิ่มปุ่มอัปโหลดไฟล์ RP01 ไว้ด้านล่างสุดของหน้าจอหลักด้วย ---
st.markdown("---")
st.markdown(
    '<div class="sub-title">📂 เพิ่มเติม: อัปโหลดไฟล์ Excel (RP01)'
    " ข้อมูลหน้าจอหลัก</div>",
    unsafe_allow_html=True,
)

uploaded_file_bottom = st.file_uploader(
    "📂 เลือกไฟล์ Excel (RP01) ด้านล่าง",
    type=["xlsx", "xls"],
    key="unique_rp01_extra_bottom_main_key",
)

if uploaded_file_bottom is not None:
  try:
    with open(SERVER_SAVE_FILE, "wb") as f:
      f.write(uploaded_file_bottom.getbuffer())
    st.cache_data.clear()
    st.success("✅ อัปโหลดไฟล์สำเร็จ! กำลังรีเฟรชระบบ...")
    st.rerun()
  except Exception as e:
    st.error(f"ไม่สามารถบันทึกไฟล์ได้: {e}")
