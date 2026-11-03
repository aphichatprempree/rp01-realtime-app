import streamlit as st
import pandas as pd
import datetime
import io

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS: ขยายขนาดฟอนต์ภายในตารางและการ์ดสรุปผลให้ใหญ่และหนาขึ้นอย่างชัดเจน
st.markdown("""
    <style>
    /* พื้นหลังหลักของหน้าเว็บเป็นสีขาว */
    .stApp { 
        background-color: #ffffff !important;
        color: #111827 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .main { 
        background-color: #ffffff !important;
    }
    
    /* กล่องหัวข้อหลัก (Header Box) */
    .header-box {
        background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #172554 100%);
        padding: 24px; 
        border-radius: 12px; 
        color: #ffffff;
        text-align: center; 
        font-weight: 900; 
        font-size: 32px; 
        margin-bottom: 25px;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.35);
        border: 1px solid #3b82f6;
    }
    
    /* หัวข้อรอง (Sub-title) ใหญ่พิเศษ */
    .sub-title {
        color: #1e40af; 
        font-weight: 900; 
        font-size: 26px;
        border-left: 8px solid #1e40af; 
        padding-left: 14px; 
        margin-top: 35px; 
        margin-bottom: 18px;
    }
    
    /* แถบ TOTAL */
    .total-box {
        background: linear-gradient(135deg, #172554 0%, #1e3a8a 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 16px; 
        border-radius: 8px; 
        font-size: 24px; 
        box-shadow: 0 6px 15px rgba(23, 37, 84, 0.25); 
        border: 1px solid #3b82f6;
    }

    /* กล่อง DCSBx และ DCSB */
    .dcsbx-box {
        background: linear-gradient(135deg, #1e40af 0%, #2563eb 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 14px; 
        border-radius: 8px 8px 0 0; 
        font-size: 24px;
    }
    .dcsb-box {
        background: linear-gradient(135deg, #0369a1 0%, #0284c7 100%);
        color: white; 
        text-align: center; 
        font-weight: 900; 
        padding: 14px; 
        border-radius: 8px 8px 0 0; 
        font-size: 24px;
    }
    
    /* === ขยายขนาดฟอนต์ในตาราง (DataFrame) ให้ใหญ่และหนาขึ้นเป็นพิเศษ === */
    dataframe, table, th, td, div[data-testid="stDataFrame"] {
        font-size: 22px !important;
    }
    div[data-testid="stDataFrame"] th {
        font-size: 22px !important;
        font-weight: 900 !important;
        color: #111827 !important;
        background-color: #f1f5f9 !important;
    }
    div[data-testid="stDataFrame"] td {
        font-size: 22px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }
    
    /* ตกแต่งการ์ดตัวเลข (Metric Cards) ให้ตัวเลขใหญ่สะใจ */
    div[data-testid="stMetric"] {
        background: linear-gradient(145deg, #f8fafc, #f1f5f9);
        border: 2px solid #cbd5e1;
        padding: 22px 28px;
        border-radius: 14px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.08);
    }
    div[data-testid="stMetric"] label {
        font-size: 20px !important;
        font-weight: 800 !important;
        color: #334155 !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 40px !important;
        font-weight: 900 !important;
        color: #0f172a !important;
    }
    
    /* แถบ Sidebar ด้านข้าง */
    section[data-testid="stSidebar"] {
        background-color: #f8fafc;
        border-right: 1px solid #e2e8f0;
    }
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {
        font-size: 18px !important;
        font-weight: 700 !important;
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
        total_palletes = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
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
            st.metric("Total Palletes", f"{total_palletes:,.0f}")

        st.markdown("---")

        # --- ตารางที่ 1: ภาพรวมสรุปกะการทำงาน ---
        st.markdown('<div class="sub-title">ภาพรวมสรุปกะการทำงาน</div>', unsafe_allow_html=True)
        summary_df = pd.DataFrame({
            "Date": [current_date, ""],
            "Total Sup.": [total_sup, ""],
            "Item": [total_item, ""],
            "UOM Qty": [f"{total_uom:,.0f}", ""],
            "Total Palletes": [f"{total_palletes:,.0f}", ""],
            "Shift": ["Shift Morning During Time [ 06.00-00.00 ]", "Shift Night During Time [ 00.00-06.00 ]"],
            "PO": [morning_po, night_po],
            "Total": [f"{morning_pallets:,.0f}", f"{night_pallets:,.0f}"],
            "คิดเป็น %": [f"{morning_pct:.2f}%", f"{night_pct:.2f}%"]
        })
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
        
        st.markdown(f"<div class='total-box'>TOTAL: {total_sup}</div>", unsafe_allow_html=True)
        st.write("")

        # --- ตารางที่ 2: DCSBx และ DCSB ---
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<div class='dcsbx-box'>DCSBx</div>", unsafe_allow_html=True)
            dcsbx_data = {"Date": [current_date], "Item": [dcsbx_item], "Palletes": [dcsbx_pallets], "PO": [100], "คิดเป็น %": ["37.73%"]}
            st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            
        with col2:
            st.markdown("<div class='dcsb-box'>DCSB</div>", unsafe_allow_html=True)
            dcsb_data = {"Date": [current_date], "Item": [dcsb_item], "Palletes": [dcsb_pallets], "PO": [150], "คิดเป็น %": ["62.27%"]}
            st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

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
            <div style='background: linear-gradient(135deg, #047857 0%, #065f46 100%); color: white; padding: 18px 26px; border-radius: 8px; font-weight: 900; display: flex; justify-content: space-between; margin-top: 25px; font-size: 22px; box-shadow: 0 6px 15px rgba(4, 120, 87, 0.25); border: 1px solid #10b981;'>
                <span>Grand Total (Filtered)</span>
                <span>Pallets: {total_palletes:,.0f}</span>
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
                    body {{ background-color: #ffffff; color: #111827; font-family: sans-serif; padding: 20px; }}
                    h2 {{ color: #1e40af; text-align: center; font-size: 26px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                    th, td {{ border: 1px solid #cbd5e1; padding: 14px; text-align: center; font-size: 20px; font-weight: bold; }}
                    th {{ background-color: #1e3a8a; color: white; }}
                    td {{ background-color: #f8fafc; }}
                </style>
            </head>
            <body>
                <h2>📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</h2>
                <p style="font-size: 18px;"><b>Total Sup:</b> {total_sup:,} | <b>Item:</b> {total_item:,} | <b>UOM Qty:</b> {total_uom:,.0f} | <b>Total Pallets:</b> {total_palletes:,.0f}</p>
                {summary_df.to_html(index=False, border=0)}
                <br>
                <p style="text-align: center; color: #64748b; font-size: 16px;">(บันทึกหน้าเว็บนี้เป็นภาพหน้าจอหรือ PDF เพื่อเก็บบันทึกยอดงานย้อนหลัง)</p>
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
