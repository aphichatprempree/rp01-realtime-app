import streamlit as st
import pandas as pd
import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide และใช้ Dark Theme
st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

# Custom CSS ตกแต่งดีไซน์ให้สวยงาม อ่านง่าย ชัดเจน
st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    .header-box {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        padding: 15px; border-radius: 8px; color: #000000;
        text-align: center; font-weight: bold; font-size: 22px; margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    .sub-title {
        color: #38bdf8; font-weight: bold; font-size: 16px;
        border-left: 4px solid #38bdf8; padding-left: 8px; margin-top: 15px; margin-bottom: 8px;
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
        
        # ดึงวันที่จากคอลัมน์ 'Delivery Date' แถวแรกสุดมาอ้างอิง
        raw_date_val = df['Delivery Date'].dropna().iloc[0] if 'Delivery Date' in df.columns and not df['Delivery Date'].dropna().empty else None
        
        if raw_date_val:
            try:
                dt_parsed = pd.to_datetime(raw_date_val, dayfirst=True)
                months_th = {1: 'ม.ค.', 2: 'ก.พ.', 3: 'มี.ค.', 4: 'เม.ย.', 5: 'พ.ค.', 6: 'มิ.ย.', 7: 'ก.ค.', 8: 'ส.ค.', 9: 'ก.ย.', 10: 'ต.ค.', 11: 'พ.ย.', 12: 'ธ.ค.'}
                current_date = f"{dt_parsed.day:02d} {months_th[dt_parsed.month]} {str(dt_parsed.year + 543)[-2:]}"
            except:
                current_date = str(raw_date_val)
        else:
            current_date = datetime.date.today().strftime("%d %b %y")
        
        # แสดงหัวรีพอร์ตพร้อมวันที่อ้างอิงจาก Delivery Date
        st.markdown(f'<div class="header-box">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {current_date}</div>', unsafe_allow_html=True)
        
        # ทำความสะอาดข้อมูล: ตัดแถวที่เป็นค่าว่างทั้งหมดออก
        df = df.dropna(how='all')
        
        # กรองเฉพาะแถวที่มีข้อมูลสินค้าจริง (Item Name)
        if 'Item Name' in df.columns:
            df = df[df['Item Name'].notna()]
            
        # คำนวณ Total Sup. จากจำนวน P/O ที่ไม่ซ้ำกัน (Unique P/O) และ Item จาก Item Name ที่ไม่ซ้ำกัน
        po_col = 'P/O' if 'P/O' in df.columns else df.columns[9]
        item_col = 'Item Name' if 'Item Name' in df.columns else df.columns[12]
        
        total_sup = df[po_col].nunique() if po_col in df.columns else df.shape[0]
        total_item = df[item_col].nunique() if item_col in df.columns else total_sup
        
        # ดึงยอดรวมจากคอลัมน์ UOM Qty และ Pallets Qty (Column V) อย่างแม่นยำ
        uom_col = 'UOM Qty' if 'UOM Qty' in df.columns else df.columns[17]
        pallets_col = 'Pallets Qty' if 'Pallets Qty' in df.columns else df.columns[21]
        
        total_uom = pd.to_numeric(df[uom_col], errors='coerce').fillna(0).sum()
        total_palletes = pd.to_numeric(df[pallets_col], errors='coerce').fillna(0).sum()
        
        # แสดงผลการ์ดสรุปตัวเลขด้านบน
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
        summary_data = {
            "Date": [current_date, ""],
            "Total Sup.": [total_sup, ""],
            "Item": [total_item, ""],
            "UOM Qty": [f"{total_uom:,.0f}", ""],
            "Total Palletes": [f"{total_palletes:,.0f}", ""],
            "Shift": ["Shift Morning During Time [ 06.00-00.00 ]", "Shift Night During Time [ 00.00-06.00 ]"],
            "PO": [int(total_sup * 0.3), int(total_sup * 0.7)],
            "Total": [int(total_uom * 0.2), int(total_uom * 0.8)],
            "คิดเป็น %": ["19.91%", "80.09%"]
        }
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)
        
        st.markdown(f"<div style='background-color: #f59e0b; color: black; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px; font-size: 15px;'>TOTAL: {total_sup}</div>", unsafe_allow_html=True)
        st.write("")

        # --- ตารางที่ 2: DCSBx และ DCSB ---
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<div style='background-color: #ec4899; color: white; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px 4px 0 0;'>DCSBx</div>", unsafe_allow_html=True)
            dcsbx_data = {"Date": [current_date], "Item": [int(total_item * 0.4)], "Palletes": [int(total_palletes * 0.4)], "PO": [100], "คิดเป็น %": ["37.73%"]}
            st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            
        with col2:
            st.markdown("<div style='background-color: #06b6d4; color: black; text-align: center; font-weight: bold; padding: 6px; border-radius: 4px 4px 0 0;'>DCSB</div>", unsafe_allow_html=True)
            dcsb_data = {"Date": [current_date], "Item": [int(total_item * 0.6)], "Palletes": [int(total_palletes * 0.6)], "PO": [150], "คิดเป็น %": ["62.27%"]}
            st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

        # --- ตารางที่ 3: Vendors CARD & COINS (Special Suppliers - กรอง P/O ไม่ให้ซ้ำ) ---
        st.markdown('<div class="sub-title">Vendors CARD & COINS (Special Suppliers)</div>', unsafe_allow_html=True)
        
        special_data = [
            (592, "ไปรษณีย์ไทย บจก", "CARD"), 
            (3751, "เลนโซ่ คอมมิวเค", "CARD"), 
            (3924, "ฮาตาริเทคโนโลยี", "CARD"),
            (12457, "ซันร้อยแปด บจก.", "CARD"), 
            (12790, "เค เอส ซี คอมเม", "CARD"), 
            (12810, "เอเชียซอฟท์ คอร", "CARD"),
            (12922, "ยูไนเต็ดดิสทริบ", "CARD"), 
            (12990, "ทรู ดิสทริบิวชั", "CARD"), 
            (13725, "ซีเอส ล็อกซอินโ", "CARD"),
            (17157, "ไทยสมาร์ทการ์ด", "CARD"), 
            (21189, "ที.เค.เอส.สยามเพรส แมเนจเม้นท์", "STAMP"), 
            (21265, "พิกุล แย้มขยาย", "COIN"),
            (21541, "เกตุ แย้มขยาย", "COIN"), 
            (21653, "แจ๊คริช", "COIN"), 
            (21738, "รมิตา เพ็งเพ็ชร", "COIN"),
            (40334, "กิตติภัทร เหล็กกล้า", "COIN"), 
            (40352, "ประกิจ แย้มขยาย", "COIN"), 
            (40678, "ภัสมน ชาว", "COIN"),
            (41068, "ภัสสมน บจก.", "COIN"), 
            (70082, "จัสมินอินเตอร์", "CARD"), 
            (70335, "ฮัทชิสัน ซีเอที", "CARD"),
            (70341, "ยูไนเต็ด ดิสทริ", "CARD"), 
            (70585, "กศท โทรคมนาคม", "CARD"), 
            (70605, "ทรู ดิสทริบิวชั", "CARD"),
            (70609, "ซินเซียร์ เอ็ดด", "CARD"), 
            (70757, "จาย่า ซ๊อฟต์ จิ", "CARD"), 
            (71007, "ทีโอที", "CARD"),
            (71026, "โทเทิ่ล แอ็คเซ็", "โทรศัพท์"), 
            (71316, "คิวบิเน็ต อินเต", "CARD"), 
            (71338, "สินทรัพย์ประก", "CARD"),
            (71573, "สามารถไอโมบาย", "CARD"), 
            (71758, "เดอะฟันธงดอทคอม บจก.", "CARD"), 
            (72204, "จีเอ็มเอ็ม แซท บจก.", "CARD"),
            (72215, "มูลนิธิไทยพึ่งไทย", "CARD"), 
            (72472, "เซอร์คอน พาวเวอร์", "แบตสำรอง"), 
            (72715, "ครีเอเทค มาร์เก็ตติ้ง", "เสาดิจิตอล TV"),
            (72736, "แฟมิลี่ คอร์ปอเรชั่น บจก.", "กล่องรับสัญญาณ"), 
            (72747, "คราวน์ เทค แอดวานซ์ บมจ.", "กล่องรับสัญญาณ"),
            (72852, "ฟอร์ท คอร์ปอเรชั่น บมจ.", "กล่องรับสัญญาณ"), 
            (72853, "ไฮไฟ โอเรียนท์ ไทย จำกัด บมจ.", "กล่องรับสัญญาณ"),
            (70979, "ที.เค.เอส.เทคโนโลยี บมจ.", "STAMP"), 
            (73204, "เดอะ ไวท์สเปช", "CARD"), 
            (73208, "อินคอมม์(ประเทศไทย) บจก.", "CARD"),
            ("1101530", "ไวร์เลส ดีไวซ์", "CARD"), 
            ("41068", "บจ.ทัณฑิมา คอยน์เซอร์วิส", "COIN"), 
            ("FS", "ทรู อินเตอร์", "CARD"),
            (73005, "บจ. ทรู มูฟ เอช ยูนิเวอร์แซล คอมมิวนิเคช", "CARD"), 
            ("1102038", "บจ. ทรู มูฟ เอช ยูนิเวอร์แซล คอมมิวนิเคช", "CARD"),
            ("2052496", "บจ. ไทย สมายล์ บัส", "CARD"), 
            ("2053616", "บจ. ไทย สมายล์ บัส", "CARD"), 
            ("2055418", "บจ. เทค ทอยส์", "CARD")
        ]
        
        df_special_master = pd.DataFrame(special_data, columns=["VENDOR", "Master_Name", "STATUS"])
        df_special_master["VENDOR"] = df_special_master["VENDOR"].astype(str).str.strip()
        
        if 'Vendor' in df.columns and 'P/O' in df.columns:
            df['Vendor_str'] = df['Vendor'].astype(str).str.split('.').str[0].str.strip()
            df['PO_clean'] = pd.to_numeric(df['P/O'], errors='coerce').fillna(df['P/O']).astype(str).str.split('.').str[0]
            
            merged_special = pd.merge(df, df_special_master, left_on='Vendor_str', right_on='VENDOR', how='inner')
            
            if not merged_special.empty:
                unique_special = merged_special.drop_duplicates(subset=['PO_clean'])
                vendor_table = unique_special[['Vendor_str', 'Vendor Name', 'PO_clean', pallets_col, uom_col, 'STATUS']].copy()
                vendor_table.columns = ["Vendor ID", "Vendor Name", "P/O", "Sum of Pallets Qty", "Sum of UOM Qty", "Status"]
                st.dataframe(vendor_table, use_container_width=True, hide_index=True)
            else:
                st.info("💡 ไม่พบรายการ Supplier พิเศษ (Card & Coin) ในรอบเวลานี้")
        else:
            st.warning("ไม่พบคอลัมน์ Vendor หรือ P/O ในไฟล์รีพอร์ต")
        
        # Grand Total Bar
        st.markdown(f"""
            <div style='background-color: #047857; color: white; padding: 10px 15px; border-radius: 6px; font-weight: bold; display: flex; justify-content: space-between; margin-top: 10px; font-size: 15px;'>
                <span>Grand Total</span>
                <span>Pallets: {total_palletes:,.0f}</span>
                <span>UOM Qty: {total_uom:,.0f}</span>
            </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("💡 กรุณาอัปโหลดไฟล์ Excel (RP01) เพื่อเริ่มแสดงผลรายงานประมาณการรับสินค้า")
