import pandas as pd
import streamlit as st

# ตั้งค่าหน้าเว็บ
st.set_page_config(
    page_title="รายงานประมาณการรับสินค้า", page_icon="📋", layout="wide"
)

# --- โค้ดสำหรับขยายขนาดตัวอักษรในตาราง ---
st.markdown(
    """
    <style>
    .stDataFrame td {
        font-size: 16px !important;
    }
    .stDataFrame th {
        font-size: 16px !important;
        font-weight: bold !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.subheader("📋 รายงานข้อมูล Vendor, PO, Item, Ti, Hi")

# ดึงข้อมูลจาก st.session_state หรือใช้ข้อมูลจำลอง
if "main_data" in st.session_state:
  df = st.session_state["main_data"]
else:
  df = pd.DataFrame({
      "Vendor": ["บริษัท เอ บี ซี จำกัด", "บริษัท เอ็กซ์ วาย ซี จำกัด", "บริษัท เทส จำกัด"],
      "PO": ["5011101837", "5011101838", "5011101893"],
      "Item": ["Item A", "Item B", "Item C"],
      "Ti": [1, 2, 1],
      "Hi": [1, 3, 1],
  })

# 1. แสดงผลตารางข้อมูล
required_cols = ["Vendor", "PO", "Item", "Ti", "Hi"]
if all(col in df.columns for col in required_cols):
  st.dataframe(df[required_cols], use_container_width=True)

  # 2. ตรวจสอบเงื่อนไขแจ้งเตือน เมื่อ Ti = 1 และ Hi = 1 เหมือนกัน
  alert_items = df[(df["Ti"] == 1) & (df["Hi"] == 1)]

  if not alert_items.empty:
    st.warning(
        "🚨 **แจ้งเตือนระบบ:** ตรวจพบรายการที่มีค่า **Ti = 1** และ **Hi ="
        " 1** พร้อมกัน จำนวน "
        f"{len(alert_items)} รายการ!"
    )

    with st.expander(
        "🔍 คลิกเพื่อดูรายละเอียดรายการที่ต้องตรวจสอบ (Ti=1 และ Hi=1)"
    ):
      st.dataframe(alert_items[required_cols], use_container_width=True)
  else:
    st.success("✅ ไม่พบรายการที่ค่า Ti = 1 และ Hi = 1 พร้อมกันในขณะนี้")
else:
  st.dataframe(df, use_container_width=True)

# --- ส่วนปุ่มพิมพ์/บันทึก PDF ---
st.markdown("---")
if st.button("🖨️ พิมพ์/บันทึก PDF ทันที"):
  import streamlit.components.v1 as components

  js_code = """
        <script>
            window.print();
        </script>
    """
  components.html(js_code, height=0)
