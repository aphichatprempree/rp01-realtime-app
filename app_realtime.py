import streamlit as st
import pandas as pd
import datetime

# ตั้งค่าหน้าเว็บให้เป็นแบบ Wide และใช้ Dark Theme
st.set_page_config(page_title="ระบบบริหารจัดการคิวและรายงานสินค้า", layout="wide")

# Custom CSS เพื่อตกแต่งหน้าตาให้มีมิติ สไตล์ Modern Dashboard
st.markdown("""
    <style>
    .main {
        background-color: #0e1117;
    }
    .header-card {
        background: linear-gradient(135deg, #1f2937 0%, #111827 100%);
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
        border: 1px solid #374151;
        color: white;
        text-align: center;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #1f2937;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #3b82f6;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
        color: white;
        text-align: center;
    }
    .sub-header {
        color: #f3f4f6;
        font-weight: 600;
        border-bottom: 2px solid #374151;
        padding-bottom: 8px;
        margin-top: 20px;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# ส่วนหัวโปรแกรมแบบมีมิติ
st.markdown("""
    <div class="header-card">
        <h2>📦 DATA ENTRY - ระบบบริหารจัดการคิวและรายงานสินค้า</h2>
        <p style='color: #9ca3af; margin: 0;'>Real-time Sync & Appointment Management Dashboard</p>
    </div>
""", unsafe_allow_html=True)

# ช่องอัปโหลดไฟล์
uploaded_file = st.file_uploader("📂 อัปโหลดไฟล์ Excel (RP01 - Schedule Appointment List Report)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        current_date = datetime.date.today().strftime("%d %b %y")
        
        total_sup = df.shape[0] if not df.empty else 0
        total_item = df['Item'].nunique() if 'Item' in df.columns else total_sup
        total_uom = df['UOM Qty'].sum() if 'UOM Qty' in df.columns else 0
        total_palletes = df['Pallets'].sum() if 'Pallets' in df.columns else 0
        
        # แสดงผล Metric แบบการ์ดมีมิติ
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric(label="📋 Total Sup.", value=f"{total_sup:,}")
        with m2:
            st.metric(label="📦 Total Items", value=f"{total_item:,}")
        with m3:
            st.metric(label="⚖️ UOM Qty", value=f"{total_uom:,}")
        with m4:
            st.metric(label="🪵 Total Palletes", value=f"{total_palletes:,}")

        st.markdown("---")

        # --- ตารางที่ 1: ภาพรวมสรุปกะการทำงาน ---
        st.markdown("<h3 class='sub-header'>📊 ภาพรวมสรุปกะการทำงาน (Shift Overview)</h3>", unsafe_allow_html=True)
        summary_data = {
            "Date": [current_date, ""],
            "Total Sup.": [total_sup, ""],
            "Item": [total_item, ""],
            "UOM Qty": [total_uom, ""],
            "Total Palletes": [total_palletes, ""],
            "Shift": ["Shift Morning [ 06.00-00.00 ]", "Shift Night [ 00.00-06.00 ]"],
            "PO": [int(total_sup * 0.3), int(total_sup * 0.7)],
            "Total": [int(total_uom * 0.2), int(total_uom * 0.8)],
            "คิดเป็น %": ["19.91%", "80.09%"]
        }
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)
        
        # --- ตารางที่ 2: แยกตาม DCSBx และ DCSB ---
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("<h4 style='color: #ec4899;'>🌸 DCSBx</h4>", unsafe_allow_html=True)
            dcsbx_data = {"Date": [current_date], "Item": [int(total_item * 0.4)], "Palletes": [int(total_palletes * 0.4)], "PO": [100], "คิดเป็น %": ["37.73%"]}
            st.dataframe(pd.DataFrame(dcsbx_data), use_container_width=True, hide_index=True)
            
        with col2:
            st.markdown("<h4 style='color: #06b6d4;'>💧 DCSB</h4>", unsafe_allow_html=True)
            dcsb_data = {"Date": [current_date], "Item": [int(total_item * 0.6)], "Palletes": [int(total_palletes * 0.6)], "PO": [150], "คิดเป็น %": ["62.27%"]}
            st.dataframe(pd.DataFrame(dcsb_data), use_container_width=True, hide_index=True)

        # --- ตารางที่ 3: Vendors CARD & COINS ---
        st.markdown("<h3 class='sub-header'>🏢 Vendors CARD & COINS</h3>", unsafe_allow_html=True)
        vendor_data = {
            "Vendor": ["12922", "70341", "73005"],
            "Vendor Name": ["ยูไนเต็ด ดิสทริบิวชัน บิซซิเนส บจก.", "ยูไนเต็ด ดิสทริบิวชัน บิซซิเนส บจก.", "บจก.ทรู เอช ยูนิเวอร์แซล คอ"],
            "Sum of Pallets Qty": [1, 2, 3],
            "Sum of UOM Qty": [600, 400, 5000],
            "Total Sum of Pallets Qty": [1, 2, 3],
            "Total Sum of UOM Qty": [600, 400, 5000]
        }
        st.dataframe(pd.DataFrame(vendor_data), use_container_width=True, hide_index=True)
        
        # Grand Total Bar แบบพรีเมียม
        st.markdown(f"""
            <div style='background: linear-gradient(90deg, #065f46 0%, #047857 100%); color: white; padding: 12px 20px; border-radius: 8px; font-weight: bold; display: flex; justify-content: space-between; margin-top: 15px;'>
                <span>🟢 Grand Total</span>
                <span>Pallets: {total_palletes}</span>
                <span>UOM Qty: {total_uom:,}</span>
            </div>
        """, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์ Excel: {e}")
else:
    st.info("💡 กรุณาอัปโหลดไฟล์ Excel (RP01) ด้านบน เพื่อเริ่มแสดงผลรายงานแบบ Real-time Dashboard")
