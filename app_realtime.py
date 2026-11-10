import streamlit as st
import pandas as pd
import datetime as dt
import sqlite3
import io
import os
import re
import html

# =========================================================
# DATA ENTRY (DCSB) | Receiving Forecast Dashboard
# =========================================================
st.set_page_config(
    page_title="รายงานประมาณการรับสินค้า | DATA ENTRY (DCSB)",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

RP01_PATH = "latest_uploaded_rp01.xlsx"
DB_PATH = "receiving_dashboard.db"

# ------------------------- CSS --------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700;800&display=swap');
:root { --navy:#082d62; --blue:#0759c7; --cyan:#08a7e5; --ink:#10234b; }
html, body, [class*="css"], .stApp { font-family:'Prompt',sans-serif !important; }
.stApp { background:linear-gradient(135deg,#eaf4ff 0%,#f7fbff 55%,#e6f2ff 100%); color:#10234b; }
.block-container { padding-top:1rem !important; padding-bottom:1.5rem !important; max-width:1800px; }
header[data-testid="stHeader"] { background:transparent; }
footer { visibility:hidden; }
section[data-testid="stSidebar"] {
  background:linear-gradient(180deg,#062958 0%,#0b3b78 100%);
  border-right:3px solid #48baff;
}
section[data-testid="stSidebar"] * { color:#f8fbff !important; }
section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {
  color:#fff !important; font-weight:800 !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {
  background:#f4f9ff !important; border:2px dashed #42b8ff !important; border-radius:12px;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section * { color:#174477 !important; }
.hero {
  background:linear-gradient(105deg,#062e72 0%,#075ed0 62%,#8bc8f6 100%);
  border:1px solid #b7d9ff; border-radius:16px; padding:19px 25px;
  color:white; box-shadow:0 7px 18px #164b8a35; margin-bottom:15px;
  display:flex; align-items:center; justify-content:space-between; gap:16px;
}
.hero-title {font-size:clamp(20px,2.2vw,32px);font-weight:800;line-height:1.35;}
.hero-sub {font-size:13px;opacity:.9;margin-top:4px;}
.hero-logo {font-size:27px;font-weight:800;background:#ffffffeb;color:#164b91;border-radius:13px;padding:10px 17px;}
.kpi {
  background:linear-gradient(145deg,#fff,#eef6ff); border:1px solid #c4dcfa;
  border-radius:14px; padding:13px 14px; min-height:112px;
  box-shadow:0 4px 12px #174b8520; text-align:center;
}
.kpi-label {font-size:14px;font-weight:700;color:#23416c;}
.kpi-value {font-size:clamp(23px,2vw,32px);font-weight:800;color:#071d4a;line-height:1.35;}
.kpi-note {font-size:12px;font-weight:700;color:#07843d;}

.kpi.night {
    background:
        radial-gradient(circle at 15% 20%, rgba(129,140,248,.35), transparent 40%),
        radial-gradient(circle at 85% 80%, rgba(168,85,247,.35), transparent 45%),
        linear-gradient(135deg,#172554 0%,#312e81 50%,#581c87 100%);
    border: 2px solid #818cf8;
    box-shadow: 0 0 15px rgba(124,58,237,.45);
}
.kpi.night .kpi-label {
    color: #e0e7ff;
}
.kpi.night .kpi-value {
    color: #ffffff;
    text-shadow: 0 0 12px #a78bfa;
}
.kpi.night .kpi-note {
    color: #c4b5fd;
}


/* === SUNRISE GOLDEN DAWN : MORNING SHIFT === */
.kpi.morning {
    background:
        radial-gradient(
            circle at 15% 15%,
            rgba(255,255,255,0.95),
            transparent 38%
        ),
        radial-gradient(
            circle at 85% 85%,
            rgba(251,146,60,0.25),
            transparent 45%
        ),
        linear-gradient(
            135deg,
            #fff7ed 0%,
            #ffedd5 40%,
            #fed7aa 75%,
            #fef3c7 100%
        );

    border: 2px solid #f59e0b;

    box-shadow:
        0 5px 15px rgba(245,158,11,0.22),
        inset 0 1px 5px rgba(255,255,255,0.9);
}

/* ชื่อกะเช้า */
.kpi.morning .kpi-label {
    color: #9a3412;
    font-weight: 800;
}

/* ตัวเลขยอดพาเลท */
.kpi.morning .kpi-value {
    color: #9a3412;
    text-shadow: 0 1px 3px rgba(255,255,255,0.8);
    font-weight: 900;
}

/* จำนวน PO และเปอร์เซ็นต์ */
.kpi.morning .kpi-note {
    color: #c2410c;
    font-weight: 800;
}

.panel {
  background:#fff;border:1px solid #c5dcf7;border-radius:13px;
  box-shadow:0 4px 12px #174b851c;overflow:hidden;margin:7px 0 14px 0;
}
.panel-head {padding:9px 15px;color:#fff;font-size:19px;font-weight:800;
  background:linear-gradient(90deg,#073b91,#0879df);}
.panel-head.cyan {background:linear-gradient(90deg,#078acb,#08b5e9);}
.panel-head.green {background:linear-gradient(90deg,#087b4c,#13b76e);}
.panel-head.purple {background:linear-gradient(90deg,#6840bb,#a46de9);}
.panel-body {padding:12px;}
.notice {
 border-radius:12px;padding:15px;text-align:center;font-weight:700;
 border:2px solid #ff7777;background:linear-gradient(145deg,#fff1f1,#ffdede);color:#a51f29;
}
.notice.ok {border-color:#27b76a;background:linear-gradient(145deg,#effff4,#d7ffe4);color:#12663a;}
.small-muted {color:#64748b;font-size:12px;}
div[data-testid="stMetric"] {background:#fff;border:1px solid #c8dcf4;border-radius:12px;padding:10px 13px;box-shadow:0 3px 9px #174b8515;}
div[data-testid="stMetric"] label {font-weight:700;color:#28456e !important;}
div[data-testid="stMetricValue"] {font-weight:800;color:#08275b;}
.stButton button,.stDownloadButton button {
 border-radius:10px !important;font-weight:700 !important;min-height:42px;
}
div[data-testid="stDataFrame"] {border:1px solid #c4d9f1;border-radius:9px;overflow:hidden;}
hr {border-color:#c7ddf5;}
@media print {
 section[data-testid="stSidebar"],button,.stDownloadButton {display:none!important;}
 .stApp {background:#fff!important;}
}
section[data-testid="stSidebar"] input {color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    font-weight: 700 !important;
}

/* ช่อง Dropdown ที่เลือกข้อมูลแล้ว */
section[data-testid="stSidebar"] [data-baseweb="select"] div {
    color: #0f172a !important;
}

/* ข้อความใน Dropdown */
section[data-testid="stSidebar"] [data-baseweb="select"] span {
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    font-weight: 700 !important;
}

/* ไอคอนลูกศร Dropdown */
section[data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: #0f172a !important;
}

/* ข้อความ Placeholder */
section[data-testid="stSidebar"] input::placeholder {
    color: #64748b !important;
    -webkit-text-fill-color: #64748b !important;
}

/* ช่องกรอกข้อมูลพื้นหลังสีขาว */
section[data-testid="stSidebar"] [data-baseweb="input"] {
    background-color: #ffffff !important;
}

/* ข้อความตัวเลือกใน Dropdown */
div[role="option"] {
    color: #0f172a !important;
    background-color: #ffffff !important;
}



/* ผลลัพธ์หลังแก้ไข
- ช่องเลือกวันที่: ตัวเลขสีน้ำเงินเข้ม
- ช่อง Supplier: ข้อความสีดำ อ่านง่าย
*/

/* === แก้สีตัวอักษรช่องเลือกวันที่ === */
section[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    font-weight: 700 !important;
    background-color: #ffffff !important;
    opacity: 1 !important;
}

/* แก้สีไอคอนปฏิทิน */
section[data-testid="stSidebar"] [data-testid="stDateInput"] svg {
    fill: #334155 !important;
    color: #334155 !important;
}


/* === แก้สีช่องวันที่งานเลื่อนส่ง === */
section[data-testid="stSidebar"] [data-baseweb="datepicker"] input {
    color: #0f172a !important;
    -webkit-text-fill-color: #0f172a !important;
    font-weight: 800 !important;
    background-color: #ffffff !important;
    opacity: 1 !important;
}

section[data-testid="stSidebar"] [data-baseweb="datepicker"] svg {
    color: #1e3a8a !important;
    fill: #1e3a8a !important;
}

/* แก้สีข้อความวันที่ในช่อง */
section[data-testid="stSidebar"] [data-testid="stDateInput"] div {
    color: #0f172a;
}


/* =====================================
   SIDEBAR - HIGH CONTRAST
   ===================================== */

/* ข้อความหัวข้อ */
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p {
    color: #ffffff !important;
    font-weight: 700 !important;
}

/* ช่องกรอกข้อมูล */
section[data-testid="stSidebar"] input {
    background: #ffffff !important;
    color: #10234b !important;
    -webkit-text-fill-color: #10234b !important;
    font-weight: 700 !important;
    border-radius: 9px !important;
}

/* ช่องเลือกวันที่ */
section[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    background: #ffffff !important;
    color: #10234b !important;
    -webkit-text-fill-color: #10234b !important;
    font-weight: 800 !important;
    opacity: 1 !important;
}

/* ไอคอนปฏิทิน */
section[data-testid="stSidebar"] [data-testid="stDateInput"] svg {
    color: #174477 !important;
    fill: #174477 !important;
}

/* Dropdown */
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background: #ffffff !important;
    border-radius: 9px !important;
}

section[data-testid="stSidebar"] [data-baseweb="select"] span {
    color: #10234b !important;
    -webkit-text-fill-color: #10234b !important;
    font-weight: 800 !important;
}

/* ลูกศร Dropdown */
section[data-testid="stSidebar"] [data-baseweb="select"] svg {
    fill: #174477 !important;
}

/* Placeholder */
section[data-testid="stSidebar"] input::placeholder {
    color: #64748b !important;
    -webkit-text-fill-color: #64748b !important;
}

/* ช่องเวลา */
section[data-testid="stSidebar"] [data-testid="stTimeInput"] input {
    background: #ffffff !important;
    color: #10234b !important;
    -webkit-text-fill-color: #10234b !important;
    font-weight: 800 !important;
}

/* ช่องหมายเหตุ */
section[data-testid="stSidebar"] [data-testid="stTextInput"] input {
    background: #ffffff !important;
    color: #10234b !important;
    -webkit-text-fill-color: #10234b !important;
}

/* ปุ่ม */
section[data-testid="stSidebar"] button {
    font-weight: 800 !important;
    border-radius: 9px !important;
}

/* Checkbox */
section[data-testid="stSidebar"] [data-testid="stCheckbox"] label {
    color: #ffffff !important;
    font-weight: 700 !important;
}

/* ขอบแบ่งเมนู */
section[data-testid="stSidebar"] hr {
    border-color: #5b9bd5 !important;
}

</style>
""", unsafe_allow_html=True)


# ---------------------- Utilities ----------------------
def norm_name(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())

def find_col(df, candidates=(), contains=()):
    if df is None or df.empty:
        return None
    lookup = {norm_name(c): c for c in df.columns}
    for candidate in candidates:
        key = norm_name(candidate)
        if key in lookup:
            return lookup[key]
    for c in df.columns:
        key = norm_name(c)
        if any(norm_name(term) in key for term in contains):
            return c
    return None

def clean_po(series):
    # Normalize numeric-looking PO values without converting alphanumeric IDs.
    def one(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        if s.lower() in {"nan", "none", "nat"}:
            return ""
        if re.fullmatch(r"\d+\.0+", s):
            s = s.split(".")[0]
        return s.upper()
    return series.map(one)

def to_num(series):
    return pd.to_numeric(series.astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)

@st.cache_data(show_spinner=False)
def read_rp01(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = xls.sheet_names[0]
    # RP01 template typically has its column headers on Excel row 4.
    data = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=3)
    data = data.dropna(how="all")
    item = find_col(data, ["Item Name"], contains=["itemname"])
    if item:
        data = data[data[item].notna()].copy()
    return data

def read_uploaded_table(upload):
    raw = upload.getvalue()
    if upload.name.lower().endswith(".csv"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(io.BytesIO(raw), encoding="cp874")
    return pd.read_excel(io.BytesIO(raw))

def init_db():
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS postpone (
            report_date TEXT PRIMARY KEY,
            in_pallets REAL NOT NULL DEFAULT 0,
            out_pallets REAL NOT NULL DEFAULT 0,
            note TEXT DEFAULT '',
            updated_at TEXT NOT NULL
        )
        """)

def save_postpone(date_key, incoming, outgoing, note):
    with sqlite3.connect(DB_PATH) as con:
        con.execute("""
        INSERT INTO postpone(report_date,in_pallets,out_pallets,note,updated_at)
        VALUES(?,?,?,?,?)
        ON CONFLICT(report_date) DO UPDATE SET
          in_pallets=excluded.in_pallets,
          out_pallets=excluded.out_pallets,
          note=excluded.note,
          updated_at=excluded.updated_at
        """, (date_key, incoming, outgoing, note, dt.datetime.now().isoformat(timespec="seconds")))

def get_postpone(date_key):
    with sqlite3.connect(DB_PATH) as con:
        row = con.execute(
            "SELECT in_pallets,out_pallets,note FROM postpone WHERE report_date=?",
            (date_key,)
        ).fetchone()
    return row if row else (0.0, 0.0, "")

def fmt_date(date_value):
    try:
        d = pd.to_datetime(date_value).date()
        months = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."]
        return f"{d.day:02d} {months[d.month-1]} {str(d.year+543)[-2:]}"
    except Exception:
        return str(date_value)

def section(title, color=""):
    st.markdown(f'<div class="panel"><div class="panel-head {color}">{title}</div><div class="panel-body">', unsafe_allow_html=True)

def section_end():
    st.markdown("</div></div>", unsafe_allow_html=True)


init_db()

# ------------------------- Sidebar ----------------------
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:8px 0 14px">
      <div style="font-size:28px">📊</div>
      <div style="font-size:20px;font-weight:800">DATA ENTRY (DCSB)</div>
      <div style="font-size:12px;color:#7dd3fc!important">ระบบรายงานประมาณการรับสินค้า</div>
    </div><hr>
    """, unsafe_allow_html=True)

    st.markdown("### 📂 1. จัดการข้อมูลไฟล์")
    rp01_upload = st.file_uploader("อัปโหลดไฟล์ Excel (RP01)", type=["xlsx","xls"], key="rp01_side")
    if rp01_upload is not None:
        try:
            with open(RP01_PATH, "wb") as f:
                f.write(rp01_upload.getvalue())
            st.success("อัปโหลด RP01 สำเร็จ")
            st.cache_data.clear()
        except Exception as e:
            st.error(f"บันทึกไฟล์ไม่สำเร็จ: {e}")

    st.markdown("---")
    st.markdown("### 📅 2. เลือกวันที่ย้อนหลัง")
    raw_df = pd.DataFrame()
    df = pd.DataFrame()
    selected_date = dt.date.today()
    if os.path.exists(RP01_PATH):
        try:
            with open(RP01_PATH, "rb") as f:
                raw_bytes = f.read()
            raw_df = read_rp01(raw_bytes)
            delivery_col = find_col(raw_df, ["Delivery Date"], contains=["deliverydate"])
            if delivery_col:
                parsed = pd.to_datetime(raw_df[delivery_col], errors="coerce", dayfirst=True)
                date_values = sorted(parsed.dropna().dt.date.unique())
                if date_values:
                    selected_date = st.selectbox(
                        "เลือกวันที่", date_values, index=len(date_values)-1,
                        format_func=fmt_date
                    )
                    df = raw_df.loc[parsed.dt.date == selected_date].copy()
                else:
                    df = raw_df.copy()
            else:
                st.caption("ไม่พบ Delivery Date จึงแสดงข้อมูลทั้งหมด")
                df = raw_df.copy()
        except Exception as e:
            st.error(f"อ่าน RP01 ไม่สำเร็จ: {e}")
    else:
        st.info("ยังไม่มีไฟล์ RP01")

    st.markdown("---")
    st.markdown("### 🏷️ 3. ตัวกรองเพิ่มเติม")
    status_filter = st.selectbox("ประเภท Supplier", ["ทั้งหมด","STAMP","CARD","COINS"])
    use_time = st.checkbox("⏰ กรองตามช่วงเวลา", value=False)
    start_time, end_time = dt.time(2,0), dt.time(5,0)
    if use_time:
        tc1, tc2 = st.columns(2)
        with tc1:
            start_time = st.time_input("ตั้งแต่", value=dt.time(2,0))
        with tc2:
            end_time = st.time_input("ถึง", value=dt.time(5,0))
        time_col = find_col(df, contains=["time","slot"])
        if time_col:
            parsed_time = pd.to_datetime(df[time_col].astype(str), errors="coerce").dt.time
            if start_time <= end_time:
                df = df.loc[parsed_time.between(start_time,end_time)].copy()
            else:
                df = df.loc[(parsed_time >= start_time) | (parsed_time <= end_time)].copy()
        else:
            st.warning("ไม่พบคอลัมน์เวลา/Slot ใน RP01")

    st.markdown("---")
    st.markdown("### ⏳ 4. จัดการงานเลื่อนส่ง (เข้า/ออก)")
    with st.form("postpone_form", clear_on_submit=False):
        postpone_date = st.date_input("วันที่บันทึก", value=selected_date)
        p_in = st.number_input("📥 เลื่อนเข้า (+)", min_value=0.0, step=1.0, value=0.0)
        p_out = st.number_input("📤 เลื่อนออก (-)", min_value=0.0, step=1.0, value=0.0)
        p_note = st.text_input("หมายเหตุ", placeholder="เช่น เลื่อนเข้า/เลื่อนออก")
        save_btn = st.form_submit_button("💾 บันทึกเก็บไว้", use_container_width=True)
    if save_btn:
        if p_in == 0 and p_out == 0:
            st.warning("กรุณาระบุยอดเลื่อนเข้า หรือเลื่อนออก")
        else:
            save_postpone(str(postpone_date), p_in, p_out, p_note)
            st.success("บันทึกข้อมูลแล้ว")
    postpone_in, postpone_out, postpone_note = get_postpone(str(selected_date))
    st.info(f"วันที่ {fmt_date(selected_date)} | เข้า +{postpone_in:,.0f} | ออก -{postpone_out:,.0f}")

    st.markdown("---")
    st.markdown("### 🔠 5. ปรับขนาดฟอนต์ตาราง")
    dcs_font = st.slider("ตาราง DCSBx / DCSB", 12, 30, 18, 1)
    vendor_font = st.slider("ตาราง Supplier", 12, 28, 16, 1)

# ---------------------- Main Dashboard -------------------
st.markdown(f"""
<div class="hero">
 <div><div class="hero-title">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {fmt_date(selected_date)}</div>
 <div class="hero-sub">DATA ENTRY (DCSB) • RECEIVING FORECAST DASHBOARD</div></div>
 <div class="hero-logo">CP ALL</div>
</div>
""", unsafe_allow_html=True)

if df.empty:
    st.info("📂 กรุณาอัปโหลดไฟล์ RP01 ที่แถบด้านซ้าย เพื่อเริ่มแสดงรายงาน")
    st.stop()

# Identify core columns safely
vendor_name_col = find_col(df, ["Vendor Name"], contains=["vendorname","suppliername"])
vendor_id_col = find_col(df, ["Vendor","Vendor ID","Supplier","Supplier ID"], contains=["vendor","supplier"])
item_col = find_col(df, ["Item Name","Item"], contains=["itemname"])
po_col = find_col(df, ["P/O","PO","Purchase Order"], contains=["purchaseorder","pono"])
uom_col = find_col(df, ["UOM Qty"], contains=["uomqty"])
pallet_col = find_col(df, ["Pallets Qty","Pallet Qty","Palletes"], contains=["pallet"])
asrs_col = find_col(df, ["ASRS_ITEM"], contains=["asrs"])
ti_col = find_col(df, ["Ti"], contains=["ti"])
hi_col = find_col(df, ["Hi"], contains=["hi"])

if po_col is None:
    st.error("ไม่พบคอลัมน์ PO / P/O ในไฟล์ RP01")
    st.stop()
if pallet_col is None:
    st.error("ไม่พบคอลัมน์ Pallets Qty ในไฟล์ RP01")
    st.stop()

df["_PO"] = clean_po(df[po_col])
df["_PALLETS"] = to_num(df[pallet_col])
df["_UOM"] = to_num(df[uom_col]) if uom_col else 0

total_sup = df[vendor_id_col].nunique() if vendor_id_col else (df[vendor_name_col].nunique() if vendor_name_col else 0)
total_item = df[item_col].nunique() if item_col else 0
total_uom = float(df["_UOM"].sum())
raw_pallets = float(df["_PALLETS"].sum())
total_pallets = max(0.0, raw_pallets + float(postpone_in) - float(postpone_out))

# Keep original allocation ratio from the supplied application.
night_target, morning_target = 3879, 1663
night_ratio = night_target / (night_target + morning_target)
night_pallets = round(total_pallets * night_ratio, 2)
morning_pallets = round(total_pallets - night_pallets, 2)
night_po = round(df["_PO"].replace("", pd.NA).nunique() * night_ratio)
total_po = int(df["_PO"].replace("", pd.NA).nunique())
morning_po = total_po - night_po
night_pct = night_pallets / total_pallets * 100 if total_pallets else 0
morning_pct = morning_pallets / total_pallets * 100 if total_pallets else 0

if asrs_col:
    asrs_flag = df[asrs_col].astype(str).str.strip().str.upper()
    dcsbx_df = df.loc[asrs_flag.eq("ASRS_ITEM")].copy()
    dcsb_df = df.loc[~asrs_flag.eq("ASRS_ITEM")].copy()
else:
    dcsbx_df, dcsb_df = df.iloc[0:0].copy(), df.copy()
    st.warning("ไม่พบคอลัมน์ ASRS_ITEM ระบบจึงจัดรายการทั้งหมดเป็น DCSB (Normal)")

def unique_item_count(frame):
    return frame[item_col].nunique() if item_col and item_col in frame else 0

def unique_po_count(frame):
    return frame["_PO"].replace("", pd.NA).nunique()

def pallet_sum(frame):
    return float(frame["_PALLETS"].sum())

def share(frame):
    return pallet_sum(frame) / total_pallets * 100 if total_pallets else 0

# KPI cards
kpis = [
    ("👥","Total Sup.",f"{total_sup:,}",""),
    ("📦","Item",f"{total_item:,}",""),
    ("🛢️","UOM Qty",f"{total_uom:,.0f}",""),
    ("🪵","Net Pallets",f"{total_pallets:,.0f}",f"ปรับเข้า +{postpone_in:,.0f} / ออก -{postpone_out:,.0f}"),
    ("🌙","กะดึก (05:00-17:00)",f"{night_pallets:,.0f}",f"{night_po:,} PO ({night_pct:.1f}%)"),
    ("☀️","กะเช้า (17:00-05:00)",f"{morning_pallets:,.0f}",f"{morning_po:,} PO ({morning_pct:.1f}%)"),
]
cols = st.columns(6)
for i,(icon,label,value,note) in enumerate(kpis):
    cls = "night" if i == 4 else ("morning" if i == 5 else "")
    cols[i].markdown(
        f'<div class="kpi {cls}"><div class="kpi-label">{icon} {label}</div>'
        f'<div class="kpi-value">{value}</div><div class="kpi-note">{html.escape(note)}</div></div>',
        unsafe_allow_html=True
    )

if postpone_in or postpone_out:
    st.warning(f"⚠️ มีการปรับแผนเลื่อนส่ง: เข้า +{postpone_in:,.0f} พาเลท | ออก -{postpone_out:,.0f} พาเลท"
               + (f" | หมายเหตุ: {postpone_note}" if postpone_note else ""))
else:
    st.caption("ℹ️ แสดงยอดพาเลทตามแผนปกติ ไม่มีการปรับเลื่อนส่งในวันที่เลือก")

# DCSBx and DCSB tables
st.markdown('<div class="panel"><div class="panel-head">📊 DCSBx / DCSB — สรุปแยกประเภทพื้นที่จัดเก็บ</div><div class="panel-body">', unsafe_allow_html=True)
dc1, dc2 = st.columns(2)
with dc1:
    st.markdown('<div class="panel-head" style="border-radius:8px">DCSBx (ASRS_ITEM)</div>', unsafe_allow_html=True)
    dcsbx_table = pd.DataFrame([{
        "วันที่":fmt_date(selected_date),"Item":unique_item_count(dcsbx_df),
        "Pallets":f"{pallet_sum(dcsbx_df):,.0f}","PO":unique_po_count(dcsbx_df),
        "%":f"{share(dcsbx_df):.1f}%"
    }])
    st.dataframe(dcsbx_table, use_container_width=True, hide_index=True)
with dc2:
    st.markdown('<div class="panel-head cyan" style="border-radius:8px">DCSB (Normal)</div>', unsafe_allow_html=True)
    dcsb_table = pd.DataFrame([{
        "วันที่":fmt_date(selected_date),"Item":unique_item_count(dcsb_df),
        "Pallets":f"{pallet_sum(dcsb_df):,.0f}","PO":unique_po_count(dcsb_df),
        "%":f"{share(dcsb_df):.1f}%"
    }])
    st.dataframe(dcsb_table, use_container_width=True, hide_index=True)
st.markdown("</div></div>", unsafe_allow_html=True)

# Supplier special categories
stamp_vendors = {"70979","21189","2055418"}
card_vendors = {"592","3751","3924","12457","12790","12810","12922","12990","13725","17157",
"21265","21541","21653","21738","40334","40352","40678","41068","70082","70335","70341","70585","70605","70609","70757","71007","71316","71338","71573","71758",
"72204","72215","73204","73208","1101530","FS","73005","1102038","2052496","2053616",}
coins_vendors = {"21265","21541","21653","21738","40334","40352","40678","41068",}

st.markdown('<div class="panel"><div class="panel-head">🏷️ Vendors STAMP, CARD & COINS (Special Suppliers)</div><div class="panel-body">', unsafe_allow_html=True)
if vendor_id_col:
    special = df.copy()
    special["_VENDOR_ID"] = clean_po(special[vendor_id_col])
    def supplier_type(v):
        if v in stamp_vendors: return "STAMP"
        if v in card_vendors: return "CARD"
        if v in coins_vendors: return "COINS"
        return ""
    special["Status"] = special["_VENDOR_ID"].map(supplier_type)
    special = special[special["Status"] != ""]
    if status_filter != "ทั้งหมด":
        special = special[special["Status"] == status_filter]
    name_col = vendor_name_col or vendor_id_col
    if not special.empty:
        special = special.drop_duplicates(subset=["_PO"])
        view = pd.DataFrame({
            "Vendor ID":special["_VENDOR_ID"],
            "Vendor Name":special[name_col].astype(str),
            "P/O":special["_PO"],
            "Sum of Pallets Qty":special["_PALLETS"],
            "Sum of UOM Qty":special["_UOM"],
            "Status":special["Status"],
        })
        st.dataframe(view, use_container_width=True, hide_index=True)
        st.metric("รวมยอด Supplier พิเศษ", f"{view['Sum of Pallets Qty'].sum():,.0f} พาเลท",
                  f"{len(view):,} รายการ")
    else:
        st.info("ไม่พบ Supplier พิเศษตามตัวกรองหรือวันที่ที่เลือก")
else:
    st.warning("ไม่พบคอลัมน์ Vendor ID ใน RP01")
st.markdown("</div></div>", unsafe_allow_html=True)

# Ti / Hi validation
st.markdown('<div class="panel"><div class="panel-head green">✅ ตรวจสอบเงื่อนไข Ti และ Hi (Ti = 1 และ Hi = 1)</div><div class="panel-body">', unsafe_allow_html=True)
if ti_col and hi_col:
    ti_num, hi_num = to_num(df[ti_col]), to_num(df[hi_col])
    alert = df.loc[(ti_num == 1) & (hi_num == 1)].copy()
    display_cols = [c for c in [vendor_id_col, po_col, item_col, ti_col, hi_col] if c]
    left, right = st.columns([1.25, .85])
    with left:
        st.dataframe(alert[display_cols] if not alert.empty else df[display_cols].head(0),
                     use_container_width=True, hide_index=True)
    with right:
        if not alert.empty:
            st.markdown(f'<div class="notice">🚨 แจ้งเตือนระบบ!<br>พบรายการที่ Ti = 1 และ Hi = 1 พร้อมกัน<br><span style="font-size:30px">{len(alert):,} รายการ</span></div>', unsafe_allow_html=True)
            with st.expander("🔍 ดูรายละเอียดรายการที่ต้องตรวจสอบ"):
                st.dataframe(alert[display_cols], use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="notice ok">✅ ไม่พบรายการที่เข้าเงื่อนไข Ti = 1 และ Hi = 1</div>', unsafe_allow_html=True)
else:
    st.info("ไม่พบคอลัมน์ Ti และ/หรือ Hi ใน RP01 กรุณาตรวจสอบชื่อหัวคอลัมน์")
st.markdown("</div></div>", unsafe_allow_html=True)

# RP03 out-of-cycle comparison
st.markdown('<div class="panel"><div class="panel-head purple">🔎 ตรวจสอบงานนอกรอบ (RP03) — เปรียบเทียบ PO กับ RP01</div><div class="panel-body">', unsafe_allow_html=True)
rp03_upload = st.file_uploader("อัปโหลดไฟล์ RP03 (Goods Receive Note Summary Report)", type=["xlsx","xls","csv"], key="rp03_upload")
if rp03_upload is not None:
    try:
        rp03 = read_uploaded_table(rp03_upload)
        rp03_po_col = find_col(rp03, ["Purchase Order","PO","P/O"], contains=["purchaseorder","pono"])
        qty_col = find_col(rp03, ["Total Case Qty","Total Case Qty (iLPNs)","iLPNs","Pallets Qty"], contains=["totalcase","ilpn","pallet"])
        if rp03_po_col is None or qty_col is None:
            st.error(f"ไม่พบคอลัมน์ PO หรือ Total Case Qty ใน RP03 (PO: {rp03_po_col}, Qty: {qty_col})")
        else:
            rp03["_PO"] = clean_po(rp03[rp03_po_col])
            rp03["_QTY"] = to_num(rp03[qty_col])
            main_pos = set(df["_PO"].loc[df["_PO"] != ""])
            mismatch = rp03.loc[(rp03["_PO"] != "") & (~rp03["_PO"].isin(main_pos))].copy()
            total_mismatch = float(mismatch["_QTY"].sum())
            if mismatch.empty:
                st.markdown('<div class="notice ok">✅ ไม่พบงานนอกรอบ — ทุก PO ใน RP03 ตรงกับ RP01</div>', unsafe_allow_html=True)
            else:
                st.metric("ยอดรวม Total Case Qty (iLPNs) ของงานนอกรอบ", f"{total_mismatch:,.2f}", f"{len(mismatch):,} รายการ")
                st.dataframe(mismatch.drop(columns=["_PO","_QTY"], errors="ignore").assign(
                    **{"PO ที่ตรวจสอบ":mismatch["_PO"],"จำนวนที่รวม":mismatch["_QTY"]}
                ), use_container_width=True, hide_index=True)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    mismatch.to_excel(writer, index=False, sheet_name="RP03_Mismatch")
                st.download_button("📥 ดาวน์โหลดรายงานงานนอกรอบ (Excel)", output.getvalue(),
                                   "RP03_Mismatch_Report.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
    except Exception as e:
        st.error(f"ประมวลผล RP03 ไม่สำเร็จ: {e}")
else:
    st.caption("อัปโหลด RP03 เพื่อเปรียบเทียบ PO และตรวจหารายการที่ไม่มีใน RP01")
st.markdown("</div></div>", unsafe_allow_html=True)

# Grand total and export controls
st.markdown(f"""
<div style="background:linear-gradient(100deg,#08764c,#079b65);color:white;border:2px solid #42d69a;
border-radius:12px;padding:14px 20px;display:flex;justify-content:space-between;gap:10px;flex-wrap:wrap;
font-size:18px;font-weight:800;box-shadow:0 4px 12px #08764c30">
<span>GRAND TOTAL (NET PALLETS)</span>
<span>📦 Pallets: {total_pallets:,.0f}</span>
<span>UOM Qty: {total_uom:,.0f}</span>
</div>
""", unsafe_allow_html=True)

st.markdown("### 📥 ดาวน์โหลดและพิมพ์รายงาน")
summary = pd.DataFrame([{
    "Date":fmt_date(selected_date),"Total Sup.":total_sup,"Item":total_item,
    "UOM Qty":total_uom,"Net Pallets":total_pallets,
    "Postpone In":postpone_in,"Postpone Out":postpone_out
}])
vendor_export = pd.DataFrame()
if vendor_id_col:
    vendor_export = df[[c for c in [vendor_id_col,vendor_name_col,po_col,item_col,pallet_col,uom_col] if c]].copy()
c1,c2,c3 = st.columns(3)
with c1:
    st.download_button("📗 ดาวน์โหลดตารางสรุป (CSV)", summary.to_csv(index=False).encode("utf-8-sig"),
                       f"RP01_Summary_{selected_date}.csv","text/csv",use_container_width=True)
with c2:
    xout = io.BytesIO()
    with pd.ExcelWriter(xout, engine="openpyxl") as writer:
        summary.to_excel(writer,index=False,sheet_name="Summary")
        dcsbx_table.to_excel(writer,index=False,sheet_name="DCSBx")
        dcsb_table.to_excel(writer,index=False,sheet_name="DCSB")
        vendor_export.to_excel(writer,index=False,sheet_name="RP01_Detail")
    st.download_button("📘 ดาวน์โหลดรายงาน Excel", xout.getvalue(),
                       f"RP01_Report_{selected_date}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
with c3:
    st.button("🖨️ พิมพ์ / บันทึก PDF", on_click=None, use_container_width=True,
              help="กด Ctrl+P หรือเลือก Print จากเบราว์เซอร์ แล้วเลือก Save as PDF")
    st.caption("ใช้ Ctrl + P แล้วเลือก Save as PDF")
st.caption("หมายเหตุ: การจัดสรรยอดกะดึก/กะเช้ายังคงใช้อัตราส่วนเป้าหมาย 3,879 : 1,663 ตามโค้ดเดิม")
