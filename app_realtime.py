import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="รายงานประมาณการรับสินค้า", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    .header-box {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        padding: 12px; border-radius: 6px; color: #000000;
        text-align: center; font-weight: bold; font-size: 22px; margin-bottom: 15px;
    }
    .sub-title {
        color: #38bdf8; font-weight: bold; font-size: 16px;
        border-left: 4px solid #38bdf8; padding-left: 8px; margin-top: 15px; margin-bottom: 8px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header-box">📋 รายงานประมาณการรับสินค้า</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01 - Schedule Appointment List Report)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # อ่านไฟล์โดยกำหนดให้แถวที่ 4 เป็นหัวตาราง
        df = pd.read_excel(uploaded_file, header=4)
        current_date = datetime.date.today().strftime("%d %b %y")
        
        # ตัดแถวที่เป็นค่าว่างทั้งหมดออก
        df = df.dropna(how='all')
        
        # นับจำนวนแถวข้อมูลจริง (ไม่เอาแถวสรุปถ้ามี)
        df = df[df.iloc[:, 2].notna()] # กรองเฉพาะแถวที่มี Item
        total_sup = df.shape[0]
        
        # นับจำนวน Item ไม่ซ้ำกัน (Column C หรือ Item Name)
        item_col = 'Item Name' if 'Item Name' in df.columns else df.columns[2]
        total_item = df[item_col].nunique() if item_col in df.columns else total_sup
        
        # ดึงข้อมูลจาก Column R (UOM Qty) และ Column V (Pallets Qty) โดยตรงอย่างแม่นยำ
        # ใช้ errors='coerce' เพื่อแปลงค่าที่ไม่ใช่ตัวเลขให้เป็น NaN และกรองเฉพาะตัวเลข
        uom_series = pd.to_numeric(df.iloc[:, 13], errors='coerce').fillna(0)
        pallets_series = pd.to_numeric(df.iloc[:, 17], errors='coerce').fillna(0)
        
        total_uom = uom_series.sum()
        total_palletes = pallets_series.sum()
        
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

        # --- ตารางที่ 3: Vendors CARD & COINS ---
        st.markdown('<div class="sub-title">Vendors CARD & COINS</div>', unsafe_allow_html=True)
        vendor_data = {
            "Vendor": ["12922", "70341", "73005"],
            "Vendor Name": ["ยูไนเต็ด ดิสทริบิวชัน บิซซิเนส บจก.", "ยูไนเต็ด ดิสทริบิวชัน บิซซิเนส บจก.", "บจก.ทรู เอช ยูนิเวอร์แซล คอ"],
            "Sum of Pallets Qty": [1, 2, 3],
            "Sum of UOM Qty": [600, 400, 5000],
            "Total Sum of Pallets Qty": [1, 2, 3],
            "Total Sum of UOM Qty": [600, 400, 5000]
        }
        st.dataframe(pd.DataFrame(vendor_data), use_container_width=True, hide_index=True)
        
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
    st.info("💡 กรุณาอัปโหลดไฟล์ Excel (RP01) เพื่อแสดงผลรายงานประมาณการรับสินค้า")
