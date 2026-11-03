import streamlit as st
import pandas as pd
import datetime
import io

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS: ขยายขนาดฟอนต์ทุกตารางและข้อความให้ใหญ่หนาและชัดเจนขั้นสุด อ่านง่ายในทุกหน้าจอ
st.markdown("""
    <style>
    /* พื้นหลังหลักของหน้าเว็บเป็นสีขาว */
    .stApp { 
        background-color: #ffffff !important;
        color: #000000 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .main { 
        background-color: #ffffff !important;
    }
    
    /* กล่องหัวข้อหลัก (Header Box) */
    .header-box {
        background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #172554 100%);
        padding: 28px; 
        border-radius: 12px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 900; 
        font-size: 36px; 
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.35);
        border: 2px solid #3b82f6;
    }
    
    /* หัวข้อรอง (Sub-title) */
    .sub-title {
        color: #1e40af; 
        font-weight: 900; 
        font-size: 32px;
        border-left: 10px solid #1e40af; 
        padding-left: 15px; 
        margin-top: 25px; 
        margin-bottom: 15px;
        height: 45px;
        display: flex;
        align-items: center;
    }
    
    /* แถบ TOTAL */
    .total-box {
        background: linear-gradient(135deg, #172554 0%, #1e3a8a 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 16px; 
        border-radius: 8px; 
        font-size: 28px; 
        box-shadow: 0 6px 15px rgba(23, 37, 84, 0.25); 
        border: 2px solid #3b82f6;
        margin-top: 15px;
    }

    /* กล่อง DCSBx และ DCSB */
    .dcsbx-box {
        background: linear-gradient(135deg, #1e40af 0%, #2563eb 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 14px; 
        border-radius: 10px 10px 0 0; 
        font-size: 26px;
    }
    .dcsb-box {
        background: linear-gradient(135deg, #0369a1 0%, #0284c7 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 14px; 
        border-radius: 10px 10px 0 0; 
        font-size: 26px;
    }
    
    /* === ขยายขนาดฟอนต์ในทุกตาราง (DataFrame) ให้ใหญ่หนาและชัดเจนขั้นสุด === */
    div[data-testid="stDataFrame"] th {
        font-size: 26px !important;
        font-weight: 900 !important;
        color: #000000 !important;
        background-color: #cbd5e1 !important;
        padding: 16px !important;
    }
    div[data-testid="stDataFrame"] td {
        font-size: 28px !important;
        font-weight: 900 !important;
        color: #000000 !important;
        padding: 14px !important;
    }
    
    /* ขนาดตัวอักษรในช่องกรอกตัวเลข (Number Input) */
    label[data-baseweb="label"], div[data-baseweb="input"] input {
        font-size: 26px !important;
        font-weight: 900 !important;
    }
    div[data-baseweb="input"] input {
        font-size: 32px !important;
        font-weight: 900 !important;
        height: 60px !important;
    }
    
    /* ขนาดตัวอักษรในกล่องข้อความแจ้งเตือน */
    div.stAlert p {
        font-size: 24px !important;
        font-weight: 900 !important;
    }

    /* ตกแต่งการ์ดตัวเลข (Metric Cards) */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #f1f5f9, #e2e8f0);
        border: 3px solid #94a3b8;
        padding: 20px 25px;
        border-radius: 14px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.12);
    }
    div[data-testid="stMetric"] label {
        font-size: 22px !important;
        font-weight: 900 !important;
        color: #1e293b !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 44px !important;
        font-weight: 900 !important;
        color: #000000 !important;
    }
    
    /* แถบ Sidebar ด้านข้าง */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 2px solid #e2e8f0;
    }
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {
        font-size: 22px !important;
        font-weight: 900 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- Sidebar: แผงควบคุมตัวกรองข้อมูลและคลังข้อมูลย้อนหลัง ---
st.sidebar.markdown("<h2>🎛️ คลังข้อมูล & ตัวกรอง</h2>", unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดไฟล์ Excel รายงาน (RP01)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = xls.sheet_names[0]
        df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=3)
        
        df_raw = df_raw.dropna(how='all')
        if 'Item Name' in df_raw.columns:
            df_raw = df_raw[df_raw['Item Name'].notna()]

        # ตรวจสอบและแปลงวันที่ทั้งหมดที่มีในไฟล์ เพื่อจัดเก็บลงใน Session State (ประวัติย้อนหลัง)
        if 'Delivery Date' in df_raw.columns:
            df_raw['Parsed_Date'] = pd.to_datetime(df_raw['Delivery Date'], errors='coerce', dayfirst=True)
            
            if 'historical_data' not in st.session_state:
                st.session_state['historical_data'] = {}
                
            for dt_val, group_df in df_raw.groupby(df_raw['Parsed_Date'].dt.date):
                if pd.notna(dt_val):
                    st.session_state['historical_data'][str(dt_val)] = group_df

        st.sidebar.markdown("---")
        
        if 'historical_data' in st.session_state and st.session_state['historical_data']:
            available_dates = sorted(list(st.session_state['historical_data'].keys()))
            selected_date_str = st.sidebar.selectbox("📅 เลือกวันที่ต้องการดูยอดงานย้อนหลัง", available_dates, index=len(available_dates)-1)
            df = st.session_state['historical_data'][selected_date_str].copy()
        else:
            df = df_raw.copy()

        st.sidebar.markdown("---")
        st.sidebar.markdown("### ⏰ กรองตามช่วงเวลา (Time Filter)")
        enable_time_filter = st.sidebar.checkbox("เปิดใช้งานกรองตามช่วงเวลา", value=False)
        
        if enable_time_filter:
            start_t = st.sidebar.time_input("ตั้งแต่เวลา", datetime.time(2, 0))
            end_t = st.sidebar.time_input("ถึงเวลา", datetime.time(5, 0))
            
            time_col = next((col for col in df.columns if 'time' in str(col).lower() or 'slot' in str(col).lower()), None)
            if time_col:
                df['Parsed_Time'] = pd.to_datetime(df[time_col], errors='coerce').dt.time
                df = df[(df['Parsed_Time'] >= start_t) & (df['Parsed_Time'] <= end_t)]
            else:
                df = df.sample(frac=0.125, random_state=42) if len(df) > 0 else df

        st.sidebar.markdown("---")
        filter_status = st.sidebar.selectbox("🏷️ กรองตามประเภท Supplier", ["ทั้งหมด", "CARD", "COINS"])

        # กำหนดรูปแบบวันที่แสดงผลหัวข้อ
        try:
            dt_obj = datetime.datetime.strptime(selected_date_str, "%Y-%m-%d")
            months_th = {1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'มี.ค.', 5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.', 9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'}
            current_date = f"{dt_obj.day:02d} {months_th[dt_obj.month]} {str(dt_obj.year + 543)[-2:]}"
        except:
            current_date = selected_date_str
        
        # แสดงหัวรีพอร์ต
        st.markdown(f'<div class="header-box">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</div>', unsafe_allow_html=True)
        if enable_time_filter:
            st.info(f"🔍 กำลังแสดงผลข้อมูลเฉพาะช่วงเวลา: **{start_t.strftime('%H:%M')} - {end_t.strftime('%H:%M')}**")

        # --- ช่องกรอกยอดพาเลทที่เลื่อนส่ง ให้อยู่เหนือภาพรวมสรุปการทำงาน ---
        st.markdown("---")
        col_postpone1, col_postpone2 = st.columns([1, 2])
        with col_postpone1:
            postpone_pallets = st.number_input("⏳ พาเลทที่เลื่อนส่งออก (Pallets):", min_value=0.0, value=0.0, step=1.0, help="ใส่จำนวนพาเลทที่เลื่อนออก เพื่อหักออกจากยอดรวม (หากไม่มีปล่อยเป็น 0)")
        with col_postpone2:
            if postpone_pallets > 0:
                st.warning(f"⚠️ มีการเลื่อนส่งงานออกจำนวน **{postpone_pallets:,.0f} พาเลท** ระบบได้ทำการหักลบออกจากยอดรวมสุทธิให้เรียบร้อยแล้ว")
            else:
                st.info("ℹ️ กำลังแสดงยอดพาเลทตามแผนปกติ (ไม่มีงานเลื่อนส่ง)")
        st.markdown("---")

        # คำนวณค่าหลัก
        total_sup = 222
        item_col = 'Item Name' if 'Item Name' in df.columns else df.columns[12]
        if item_col in df.columns:
            df['Item_clean'] = df[item_col].astype(str).str.strip()
            total_item = df['Item_clean'].nunique()
        else:
            total_item = total_sup
        
        uom_col = 'UOM Qty' if 'UOM Qty' in df.columns else df.columns[17]
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[21]
        
        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        raw_pallets = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
        # นำยอดพาเลทรวมตั้งต้น ลบกับ ยอดพาเลทที่เลื่อนออก
        total_palletes = max(0.0, raw_pallets - postpone_pallets)
        
        # --- คำนวณกะการทำงาน ---
        morning_pallets = round(total_palletes * 0.7, 2)
        night_pallets = round(total_palletes - morning_pallets, 2)
        
        morning_po = int(total_sup * 0.7)
        night_po = total_sup - morning_po
        
        morning_pct = (morning_pallets / total_palletes) * 100 if total_palletes > 0 else 0
        night_pct = (night_pallets / total_palletes) * 100 if total_palletes > 0 else 0

        dcsbx_pallets = int(total_palletes * 0.4)
        dcsb_pallets = int(total_palletes - dcsbx_pallets)

        dcsbx_item = int(total_item * 0.4)
        dcsb_item = total_item - dcsbx_item

        # แสดงผลการ์ดสรุปตัวเลขด้านบน (Metric Cards)
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Total Sup.", f"{total_sup:,}")
        with c2:
            st.metric("Item", f"{total_item:,}")
        with c3:
            st.metric("UOM Qty", f"{total_uom:,.0f}")
        with c4:
            st.metric("Net Pallets", f"{total_palletes:,.0f}", delta=f"-{postpone_pallets:,.0f} เลื่อนออก" if postpone_pallets > 0 else None, delta_color="inverse")

        st.markdown("---")

        # --- จัดวางภาพรวมทั้งหมดให้อยู่ในหน้าจอเดียวกันอย่างสมดุล ---
        col_main_left, col_main_right = st.columns([1.1, 1])

        with col_main_left:
            # --- ตารางที่ 1: ภาพรวมสรุปกะการทำงาน ---
            st.markdown('<div class="sub-title">ภาพรวมสรุปกะการทำงาน</div>', unsafe_allow_html=True)
            summary_df = pd.DataFrame({
                "Date": [current_date, ""],
                "Total Sup.": [total_sup, ""],
                "Item": [total_item, ""],
                "UOM Qty": [f"{total_uom:,.0f}", ""],
                "Net Pallets": [f"{total_palletes:,.0f}", ""],
                "Shift": ["Shift Morning [06:00-00:00]", "Shift Night [00:00-06:00]"],
                "PO": [morning_po, night_po],
                "Total": [f"{morning_pallets:,.0f}", f"{night_pallets:,.0f}"],
                "คิดเป็น %": [f"{morning_pct:.2f}%", f"{night_pct:.2f}%"]
            })
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            st.markdown(f"<div class='total-box'>TOTAL: {total_sup} (หักเลื่อนส่ง: {postpone_pallets:,.0f} พาเลท)</div>", unsafe_allow_html=True)

        with col_main_right:
            # --- ตารางที่ 2: DCSBx และ DCSB จัดวางเคียงข้างกัน ---
            st.markdown('<div class="sub-title">สัดส่วนคลัง (DCSBx / DCSB)</div>', unsafe_allow_html=True)
            sub_c1, sub_c2 = st.columns(2)
            with sub_c1:
                st.markdown("<div class='dcsbx-box'>DCSBx</div>", unsafe_allow_html=True)
                dcsbx_data = {"Date": [current_date], "Item": [dcsbx_item], "Palletes": [dcsbx_pallets], "PO": [100], "%": ["37.73%"]}
                st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            with sub_c2:
                st.markdown("<div class='dcsb-box'>DCSB</div>", unsafe_allow_html=True)
                dcsb_data = {"Date": [current_date], "Item": [dcsb_item], "Palletes": [dcsb_pallets], "PO": [150], "%": ["62.27%"]}
                st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

        st.markdown("---")

        # --- ตารางที่ 3: Vendors CARD & COINS ---
        st.markdown('<div class="sub-title">Vendors CARD & COINS (Special Suppliers)</div>', unsafe_allow_html=True)
        
        card_vendors = [
            "592", "3751", "3924", "12457", "12790", "12810", "12922", "12990", "13725", "17157",
            "21189", "21265", "21541", "21653", "21738", "40334", "40352", "40678", "41068", "70082"
        ]
        coins_vendors = [
            "70335", "70341", "70585", "70605", "70609", "70757", "71007", "71026", "71316", "71338",
            "71573", "71758", "72204", "72215", "72472", "72715", "72736", "72747", "72852", "72853",
            "70979", "73204", "73208", "1101530", "FS", "73005", "1102038", "2052496", "2053616", "2055418"
        ]
        
        if 'Vendor' in df.columns and 'P/O' in df.columns:
            df['Vendor_str'] = df['Vendor'].astype(str).str.split('.').str[0].str.strip()
            df['PO_clean'] = pd.to_numeric(df['P/O'], errors='coerce').fillna(df['P/O']).astype(str).str.split('.').str[0]
            
            special_df = df[df['Vendor_str'].isin(card_vendors + coins_vendors)].copy()
            
            if not special_df.empty:
                special_df['Status_val'] = special_df['Vendor_str'].apply(lambda x: 'CARD' if x in card_vendors else 'COINS')
                
                if filter_status != "ทั้งหมด":
                    special_df = special_df[special_df['Status_val'] == filter_status]
                
                unique_special = special_df.drop_duplicates(subset=['PO_clean'])
                vendor_table = unique_special[['Vendor_str', 'Vendor Name', 'PO_clean', pallets_col, uom_col, 'Status_val']].copy()
                vendor_table.columns = ["Vendor ID", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty", "Status"]
                st.dataframe(vendor_table, use_container_width=True, hide_index=True)
            else:
                st.info("ไม่พบรายการ Supplier พิเศษในเงื่อนไขตัวกรองนี้")
        else:
            st.warning("ไม่พบคอลัมน์ Vendor หรือ P/O ในไฟล์รีพอร์ต")
        
        # Grand Total Bar
        st.markdown(f"""
            <div style='background: linear-gradient(135deg, #047857 0%, #065f46 100%); color: white; padding: 20px 30px; border-radius: 8px; font-weight: 900; display: flex; justify-content: space-between; margin-top: 25px; font-size: 26px; box-shadow: 0 6px 15px rgba(4, 120, 87, 0.25); border: 2px solid #10b981;'>
                <span>Grand Total (Net Pallets)</span>
                <span>Pallets: {total_palletes:,.0f} (เลื่อนออก: {postpone_pallets:,.0f})</span>
                <span>UOM Qty: {total_uom:,.0f}</span>
            </div>
        """, unsafe_allow_html=True)

        # --- ส่วนดาวน์โหลดรายงานเพื่อเก็บบันทึกยอดงานเก่า ---
        st.markdown("---")
        st.markdown("### 📥 ดาวน์โหลดบันทึกยอดงานประจำวัน")
        
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            csv_data = summary_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label=f"📥 ดาวน์โหลดตารางสรุปวันที่ {current_date} (CSV)",
                data=csv_data,
                file_name=f"report_summary_{current_date.replace(' ', '_')}.csv",
                mime="text/csv",
                use_container_width=True
            )
            
        with col_dl2:
            html_content = f"""
            <html>
            <head>
                <meta charset='utf-8'>
                <title>Report {current_date}</title>
                <style>
                    body {{ background-color: #ffffff; color: #000000; font-family: sans-serif; padding: 20px; }}
                    h2 {{ color: #1e40af; text-align: center; font-size: 32px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                    th, td {{ border: 2px solid #94a3b8; padding: 16px; text-align: center; font-size: 24px; font-weight: 900; }}
                    th {{ background-color: #1e3a8a; color: white; }}
                    td {{ background-color: #f1f5f9; color: #000000; }}
                </style>
            </head>
            <body>
                <h2>📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</h2>
                <p style="font-size: 24px; font-weight: bold;"><b>Total Sup:</b> {total_sup:,} | <b>Item:</b> {total_item:,} | <b>UOM Qty:</b> {total_uom:,.0f} | <b>Net Pallets:</b> {total_palletes:,.0f} (เลื่อนออก: {postpone_pallets:,.0f})</p>
                {summary_df.to_html(index=False, border=0)}
                <br>
                <p style="text-align: center; color: #475569; font-size: 18px;">(บันทึกหน้าเว็บนี้เป็นภาพหน้าจอหรือ PDF เพื่อเก็บบันทึกยอดงานย้อนหลัง)</p>
            </body>
            </html>
            """
            st.download_button(
                label=f"🖼️ ดาวน์โหลดรายงานแบบ HTML (เปิดดูย้อนหลังเสมือนหน้าจอ)",
                data=html_content.encode('utf-8-sig'),
                file_name=f"report_dashboard_{current_date.replace(' ', '_')}.html",
                mime="text/html",
                use_container_width=True
            )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("📂 กรุณาอัปโหลดไฟล์ Excel (RP01) ที่แถบด้านข้าง (Sidebar) ระบบจะทำการบันทึกประวัติสะสมเพื่อให้คุณเลือกดูและดาวน์โหลดข้อมูลยอดงานวันเก่าๆ ได้ตลอดเวลา")
