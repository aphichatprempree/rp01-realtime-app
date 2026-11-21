import streamlit as st
import pandas as pd
import datetime as dt
import sqlite3
import io
import os
import re
import html
import hashlib
import pickle

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
# Column skeleton used for a day that has never been imported (every number = 0)
DEFAULT_COLS = ["Delivery Date", "Vendor", "Vendor Name", "Item Name", "P/O", "UOM Qty",
                "Pallets Qty", "ASRS_ITEM", "Ti", "Hi", "Window Time", "PMA"]

# PMA TYP lookup from the supplied "TYP PMA" workbook
PMA_TYPE_MAP = {1: 'Fresh RTE Food', 2: 'Fresh RTE Food', 3: 'Fresh RTE Food', 5: 'Counter Drink', 6: 'Fresh RTE Food', 7: 'Fresh RTE Food', 8: 'Fresh RTE Food', 10: 'Package Bev.', 11: 'Package Bev.', 12: 'Non Food', 14: 'Non Food', 16: 'Fresh RTE Food', 18: 'Non Food', 19: 'Fresh RTE Food', 20: 'Fresh RTE Food', 21: 'Fresh RTE Food', 22: 'Fresh RTE Food', 23: 'Fresh RTE Food', 24: 'Fresh RTE Food', 25: 'Fresh RTE Food', 27: 'Non Food', 28: 'Non Food', 29: 'Package Bev.', 32: 'Process Food', 33: 'Process Food', 34: 'Process Food', 35: 'Non Food', 36: 'Fresh RTE Food', 38: 'Package Bev.', 39: 'Non Food', 40: 'Process Food', 41: 'Process Food', 42: 'Package Bev.', 43: 'Non Food', 44: 'Package Bev.', 45: 'Package Bev.', 46: 'Non Food', 49: 'Package Bev.', 50: 'Non Food', 51: 'Non Food', 52: 'Non Food', 53: 'Non Food', 54: 'Non Food', 55: 'Non Food', 56: 'Non Food', 58: 'Non Food', 59: 'Non Food', 60: 'Fresh RTE Food', 61: 'Fresh RTE Food', 63: 'Process Food', 64: 'Non Food', 66: 'Non Food', 68: 'Non Food', 69: 'Non Food', 70: 'Non Food', 71: 'Non Food', 72: 'Non Food', 74: 'Non Food', 78: 'Fresh RTE Food', 80: 'Fresh RTE Food', 82: 'Fresh RTE Food', 83: 'Fresh RTE Food', 85: 'Fresh RTE Food', 86: 'Non Food', 30: 'Counter Drink', 31: 'Counter Drink', 75: 'Counter Drink', 76: 'Counter Drink', 62: 'Counter Drink', 79: 'Counter Drink', 87: 'Counter Drink', 95: 'Counter Drink', 73: 'Counter Drink'}

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
/* Sidebar: light blue-grey theme, navy text on white cards, collapsible expanders */
section[data-testid="stSidebar"] {
  background:linear-gradient(180deg,#f3f7fc 0%,#e4edf8 100%);
  border-right:1px solid #bcd0e8;
  box-shadow:4px 0 16px rgba(16,35,75,.10);
  color:#10234b;
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] *,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
section[data-testid="stSidebar"] [data-testid="stCheckbox"] *,
section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 {
  color:#10234b !important; -webkit-text-fill-color:#10234b !important; font-weight:600;
}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] { opacity:1 !important; }
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * { color:#35507a !important; -webkit-text-fill-color:#35507a !important; }
section[data-testid="stSidebar"] hr { border-color:#b9cfe8; }
section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] button,
section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] svg { color:#10234b !important; fill:#10234b !important; }
/* Expanders (kept collapsed by default) */
section[data-testid="stSidebar"] [data-testid="stExpander"] {
  margin:.35rem 0; background:#fff !important; border:1px solid #b9d0ea !important; border-left:4px solid #0759c7 !important;
  border-radius:12px !important; box-shadow:0 2px 8px rgba(16,35,75,.10); overflow:hidden;
}
section[data-testid="stSidebar"] [data-testid="stExpander"]:hover { border-color:#0759c7 !important; box-shadow:0 3px 12px rgba(7,89,199,.20); }
section[data-testid="stSidebar"] [data-testid="stExpander"] details,
section[data-testid="stSidebar"] [data-testid="stExpander"] summary { background:#fff !important; }
section[data-testid="stSidebar"] [data-testid="stExpander"] summary { padding:.62rem .75rem !important; }
section[data-testid="stSidebar"] [data-testid="stExpander"] summary,
section[data-testid="stSidebar"] [data-testid="stExpander"] summary * { color:#082d62 !important; -webkit-text-fill-color:#082d62 !important; font-weight:700 !important; }
section[data-testid="stSidebar"] [data-testid="stExpander"] summary svg { fill:#082d62 !important; color:#082d62 !important; }
section[data-testid="stSidebar"] [data-testid="stExpander"] [data-testid="stExpanderDetails"] { padding:.3rem .75rem .8rem !important; background:#fff; }
/* Inputs */
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea {
  color:#10234b !important; -webkit-text-fill-color:#10234b !important; font-weight:600 !important; background:#fff !important;
}
section[data-testid="stSidebar"] input::placeholder { color:#6b7f9e !important; -webkit-text-fill-color:#6b7f9e !important; }
section[data-testid="stSidebar"] [data-baseweb="input"],
section[data-testid="stSidebar"] [data-baseweb="base-input"],
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
  background:#fff !important; border:1px solid #9db8d8 !important; border-radius:9px !important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] span,
section[data-testid="stSidebar"] [data-baseweb="select"] div { color:#10234b !important; -webkit-text-fill-color:#10234b !important; font-weight:600 !important; }
section[data-testid="stSidebar"] [data-baseweb="select"] svg,
section[data-testid="stSidebar"] [data-testid="stDateInput"] svg,
section[data-testid="stSidebar"] [data-testid="stTimeInput"] svg { fill:#174477 !important; color:#174477 !important; }
section[data-testid="stSidebar"] [data-testid="stNumberInput"] button { background:#e6f0fb !important; color:#10234b !important; }
section[data-testid="stSidebar"] [data-testid="stNumberInput"] button svg { fill:#10234b !important; }
/* Forms */
section[data-testid="stSidebar"] [data-testid="stForm"] { padding:.75rem; border:1px solid #c4d9f1; border-radius:12px; background:#f4f8fd; box-shadow:none; }
/* Buttons: solid blue, white text */
section[data-testid="stSidebar"] .stButton button,
section[data-testid="stSidebar"] .stDownloadButton button,
section[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
  background:linear-gradient(135deg,#0759c7,#0a7be0) !important; border-color:#064aa8 !important;
}
section[data-testid="stSidebar"] .stButton button *,
section[data-testid="stSidebar"] .stDownloadButton button *,
section[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button *,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button *,
section[data-testid="stSidebar"] .stButton button,
section[data-testid="stSidebar"] .stDownloadButton button,
section[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button { color:#fff !important; -webkit-text-fill-color:#fff !important; font-weight:700 !important; }
/* File uploader */
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"],
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section { background:#f2f8ff !important; border:2px dashed #4aa3e8 !important; border-radius:12px; }
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] *:not(button):not(button *),
section[data-testid="stSidebar"] [data-testid="stFileUploader"] small,
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section *:not(button):not(button *) { color:#174477 !important; -webkit-text-fill-color:#174477 !important; }
section[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] * { color:#10234b !important; -webkit-text-fill-color:#10234b !important; }
/* Alerts / metrics / tables: light card, navy text */
section[data-testid="stSidebar"] [data-testid="stAlert"] { background:#fff !important; border:1px solid #b9d0ea; border-radius:10px; }
section[data-testid="stSidebar"] [data-testid="stAlert"] * { color:#10234b !important; -webkit-text-fill-color:#10234b !important; }
section[data-testid="stSidebar"] [data-testid="stMetric"] { background:#fff!important; }
section[data-testid="stSidebar"] [data-testid="stMetric"] *,
section[data-testid="stSidebar"] [data-testid="stMetricLabel"],
section[data-testid="stSidebar"] [data-testid="stMetricValue"],
section[data-testid="stSidebar"] [data-testid="stMetricDelta"] { color:#10234b!important; -webkit-text-fill-color:#10234b!important; }
section[data-testid="stSidebar"] [data-testid="stDataFrame"] { background:#fff!important; }
.hero {
  background:linear-gradient(105deg,#062e72 0%,#075ed0 62%,#8bc8f6 100%);
  border:1px solid #b7d9ff; border-radius:16px; padding:19px 25px;
  color:white; box-shadow:0 7px 18px #164b8a35; margin-bottom:15px;
  display:flex; align-items:center; justify-content:space-between; gap:16px;
}
.hero-title {font-size:clamp(20px,2.2vw,32px);font-weight:800;line-height:1.35;}
.hero-sub {font-size:13px;opacity:.9;margin-top:4px;}
.hero-logo {font-size:27px;font-weight:800;background:#ffffffeb;color:#164b91;border-radius:13px;padding:10px 17px;}
.kpi { background:linear-gradient(145deg,#fff,#eef6ff); border:1px solid #c4dcfa; border-radius:14px; padding:13px 14px; min-height:112px; box-shadow:0 4px 12px #174b8520; text-align:center; }
.kpi-label {font-size:14px;font-weight:700;color:#23416c;}
.kpi-value {font-size:clamp(23px,2vw,32px);font-weight:800;color:#071d4a;line-height:1.35;}
.kpi-note {font-size:12px;font-weight:700;color:#07843d;}
.kpi.night { background:radial-gradient(circle at 15% 20%, rgba(129,140,248,.35), transparent 40%),radial-gradient(circle at 85% 80%, rgba(168,85,247,.35), transparent 45%),linear-gradient(135deg,#172554 0%,#312e81 50%,#581c87 100%); border:2px solid #818cf8; box-shadow:0 0 15px rgba(124,58,237,.45); }
.kpi.night .kpi-label {color:#e0e7ff;}.kpi.night .kpi-value {color:#fff;text-shadow:0 0 12px #a78bfa;}.kpi.night .kpi-note {color:#c4b5fd;}
.kpi.morning { background:radial-gradient(circle at 15% 15%,rgba(255,255,255,.95),transparent 38%),radial-gradient(circle at 85% 85%,rgba(251,146,60,.25),transparent 45%),linear-gradient(135deg,#fff7ed 0%,#ffedd5 40%,#fed7aa 75%,#fef3c7 100%); border:2px solid #f59e0b; box-shadow:0 5px 15px rgba(245,158,11,.22),inset 0 1px 5px rgba(255,255,255,.9); }
.kpi.morning .kpi-label,.kpi.morning .kpi-value {color:#9a3412;font-weight:900;}.kpi.morning .kpi-note {color:#c2410c;font-weight:800;}
.panel { background:#fff;border:1px solid #c5dcf7;border-radius:13px;box-shadow:0 4px 12px #174b851c;overflow:hidden;margin:7px 0 14px; }
.panel-head {padding:9px 15px;color:#fff;font-size:19px;font-weight:800;background:linear-gradient(90deg,#073b91,#0879df);}
.panel-head.cyan {background:linear-gradient(90deg,#078acb,#08b5e9);}.panel-head.green {background:linear-gradient(90deg,#087b4c,#13b76e);}.panel-head.purple {background:linear-gradient(90deg,#6840bb,#a46de9);}
.panel-body {padding:12px;}
.notice {border-radius:12px;padding:15px;text-align:center;font-weight:700;border:2px solid #ff7777;background:linear-gradient(145deg,#fff1f1,#ffdede);color:#a51f29;}
.notice.ok {border-color:#27b76a;background:linear-gradient(145deg,#effff4,#d7ffe4);color:#12663a;}
.small-muted {color:#64748b;font-size:12px;}
div[data-testid="stMetric"] {background:#fff;border:1px solid #c8dcf4;border-radius:12px;padding:10px 13px;box-shadow:0 3px 9px #174b8515;}
div[data-testid="stMetric"] label {font-weight:700;color:#28456e!important;}div[data-testid="stMetricValue"] {font-weight:800;color:#08275b;}
.stButton button,.stDownloadButton button {border-radius:10px!important;font-weight:700!important;min-height:42px;}
div[data-testid="stDataFrame"] {border:1px solid #c4d9f1;border-radius:9px;overflow:hidden;}
hr {border-color:#c7ddf5;}
@media print { section[data-testid="stSidebar"],button,.stDownloadButton {display:none!important;} .stApp {background:#fff!important;} }
</style>
""", unsafe_allow_html=True)

# ---------------------- Utilities ----------------------
def norm_name(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())

def find_col(df, candidates=(), contains=()):
    if df is None:
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

def parse_rp01_time(value):
    """Convert RP01 Excel/text time values to minutes after midnight."""
    if pd.isna(value):
        return float("nan")
    if isinstance(value, dt.datetime):
        value = value.time()
    if isinstance(value, dt.time):
        return value.hour * 60 + value.minute + value.second / 60
    if isinstance(value, (int, float)) and 0 <= float(value) < 1:
        return float(value) * 24 * 60
    parsed = pd.to_datetime(str(value).strip(), errors="coerce")
    if pd.isna(parsed):
        return float("nan")
    return parsed.hour * 60 + parsed.minute + parsed.second / 60

def shift_label(minutes):
    """Night 00:00-05:00; morning = everything else (05:00 onward)."""
    if pd.isna(minutes):
        return "ไม่ระบุเวลา"
    if minutes < 5 * 60:
        return "🌙 กะดึก"
    return "☀️ กะเช้า"

def minutes_to_hms(minutes):
    if pd.isna(minutes):
        return "N/A"
    total = int(round(minutes * 60))
    return f"{total // 3600 % 24:02d}:{total % 3600 // 60:02d}:{total % 60:02d}"

def night_first_order(minutes):
    """Sort key: plain time of day, so night 00:00-05:00 comes first."""
    return minutes

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
        con.execute("""
        CREATE TABLE IF NOT EXISTS vendor_window_time (
            vendor_id TEXT PRIMARY KEY,
            vendor_name TEXT NOT NULL DEFAULT '',
            window_time TEXT NOT NULL DEFAULT '00:00:00',
            updated_at TEXT NOT NULL
        )
        """)
        # One stored snapshot of the RP01 rows per delivery date (keeps each day's numbers)
        con.execute("""
        CREATE TABLE IF NOT EXISTS daily_snapshot (
            report_date TEXT PRIMARY KEY,
            payload BLOB NOT NULL,
            row_count INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL
        )
        """)

def get_window_time(vendor_id):
    """Return the stored supplier window or a manual setup prompt."""
    v_id = str(vendor_id).strip()
    with sqlite3.connect(DB_PATH) as con:
        row = con.execute(
            "SELECT vendor_name, window_time FROM vendor_window_time WHERE vendor_id = ?",
            (v_id,),
        ).fetchone()
    if row:
        return {"vendor_name": row[0], "window_time": row[1]}
    return {"vendor_name": "New Supplier (กรุณาตั้งค่า Manual ด้านล่าง)", "window_time": "N/A"}
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


def save_rp01_snapshots(raw_df):
    """Split RP01 by Delivery Date and store every day in SQLite (re-import overwrites that day)."""
    delivery_col = find_col(raw_df, ["Delivery Date"], contains=["deliverydate"])
    if delivery_col:
        parsed = pd.to_datetime(raw_df[delivery_col], errors="coerce", dayfirst=True)
        day_keys = parsed.dt.date
        groups = {d: raw_df.loc[day_keys == d] for d in sorted(parsed.dropna().dt.date.unique())}
    else:
        groups = {dt.date.today(): raw_df}
    now = dt.datetime.now().isoformat(timespec="seconds")
    with sqlite3.connect(DB_PATH) as con:
        for d, g in groups.items():
            con.execute(
                "INSERT OR REPLACE INTO daily_snapshot(report_date,payload,row_count,updated_at) VALUES(?,?,?,?)",
                (str(d), pickle.dumps(g.reset_index(drop=True), protocol=4), int(len(g)), now),
            )
    return list(groups)

def list_snapshot_dates():
    with sqlite3.connect(DB_PATH) as con:
        rows = con.execute("SELECT report_date FROM daily_snapshot ORDER BY report_date").fetchall()
    return [dt.date.fromisoformat(r[0]) for r in rows]

def load_snapshot(date_value):
    """Return the stored RP01 rows for the day, or None if that day was never imported."""
    with sqlite3.connect(DB_PATH) as con:
        row = con.execute("SELECT payload FROM daily_snapshot WHERE report_date=?", (str(date_value),)).fetchone()
    if not row:
        return None
    try:
        return pickle.loads(row[0])
    except Exception:
        return None


init_db()

# One-time migration: if an RP01 file exists from the old version but nothing is stored yet.
if not list_snapshot_dates() and os.path.exists(RP01_PATH):
    try:
        with open(RP01_PATH, "rb") as f:
            save_rp01_snapshots(read_rp01(f.read()))
    except Exception:
        pass

# ------------------------- Sidebar ----------------------
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:8px 0 14px">
      <div style="display:inline-flex;align-items:center;justify-content:center;width:54px;height:54px;border-radius:16px;background:linear-gradient(145deg,#38bdf8,#0b62c4);font-size:27px;box-shadow:0 5px 14px #0b3d7a40">📦</div>
      <div style="font-size:19px;font-weight:800;margin-top:9px;letter-spacing:.3px;color:#082d62!important">DATA ENTRY <span style="color:#0759c7!important">(DCSB)</span></div>
      <div style="font-size:12px;color:#35507a!important;font-weight:600;margin-top:3px">RECEIVING FORECAST</div>
    </div><hr>
    """, unsafe_allow_html=True)

    with st.expander("📂 IMPORT RP01", expanded=False):
        rp01_upload = st.file_uploader(
            "ไฟล์ประมาณการหลัก · RP01", type=["xlsx", "xls"], key="rp01_side",
            help="รองรับไฟล์ Excel .xlsx และ .xls"
        )
        if rp01_upload is not None:
            up_bytes = rp01_upload.getvalue()
            up_sig = hashlib.md5(up_bytes).hexdigest()
            # Process each uploaded file once, so later reruns don't reset the chosen date.
            if st.session_state.get("rp01_sig") != up_sig:
                try:
                    with open(RP01_PATH, "wb") as f:
                        f.write(up_bytes)
                    saved_days = save_rp01_snapshots(read_rp01(up_bytes))
                    st.session_state["rp01_sig"] = up_sig
                    if saved_days:
                        st.session_state["sel_date"] = saved_days[-1]
                    st.success(f"นำเข้า RP01 สำเร็จ · บันทึก {len(saved_days)} วัน")
                except Exception as e:
                    st.error(f"นำเข้าไฟล์ไม่สำเร็จ: {e}")
            else:
                st.caption("✅ ไฟล์นี้นำเข้าและบันทึกแล้ว")

    # Expander contents run on every Streamlit rerun; data preparation stays available when collapsed.
    with st.expander("📅 Date", expanded=False):
        saved_dates = list_snapshot_dates()
        if "sel_date" not in st.session_state:
            st.session_state["sel_date"] = saved_dates[-1] if saved_dates else dt.date.today()
        selected_date = st.date_input("เลือกวันที่", key="sel_date", format="DD/MM/YYYY")
        if selected_date is None:
            selected_date = saved_dates[-1] if saved_dates else dt.date.today()

        # Each day is read back from the database; a day never imported has no data (all zeros).
        snapshot = load_snapshot(selected_date)
        has_data = snapshot is not None
        df = snapshot.copy() if has_data else pd.DataFrame(columns=DEFAULT_COLS)
        if has_data:
            st.caption(f"✅ มีข้อมูลที่บันทึกไว้ · {len(df):,} แถว")
        else:
            st.caption("⚪ วันนี้ยังไม่เคย Import · ทุกตัวเลข = 0")
        st.caption(f"วันที่ที่บันทึกไว้แล้ว: {len(saved_dates):,} วัน")

    with st.expander("🏷️ Fillter time", expanded=False):
        status_filter = st.selectbox(
            "ประเภท Supplier", ["ทั้งหมด", "STAMP", "CARD", "COINS"],
            help="เลือกดู Supplier ตามประเภท"
        )
        use_time = st.checkbox("⏰ กรองตามช่วงเวลา", value=False)
        start_time, end_time = dt.time(2, 0), dt.time(5, 0)
        if use_time:
            tc1, tc2 = st.columns(2)
            with tc1:
                start_time = st.time_input("ตั้งแต่", value=dt.time(2, 0))
            with tc2:
                end_time = st.time_input("ถึง", value=dt.time(5, 0))
            time_col = find_col(df, ["Window Time"], contains=["windowtime", "slot"])
            if df.empty:
                st.caption("ไม่มีข้อมูลในวันที่เลือก")
            elif time_col:
                parsed_time = pd.to_datetime(df[time_col].astype(str), errors="coerce").dt.time
                if start_time <= end_time:
                    df = df.loc[parsed_time.between(start_time, end_time)].copy()
                else:
                    df = df.loc[(parsed_time >= start_time) | (parsed_time <= end_time)].copy()
            else:
                st.warning("ไม่พบคอลัมน์เวลา/Slot ใน RP01")

    with st.expander("⏳ งานเลื่อนส่ง", expanded=False):
        st.markdown("""
        <div style="background:#fff8e1;border:1px solid #f0b429;border-left:5px solid #e08a00;border-radius:10px;padding:9px 11px;margin-bottom:10px;font-size:12.5px;line-height:1.55;color:#4a3000">
          <div style="font-weight:800;font-size:13.5px;margin-bottom:3px">⚠️ วิธีบันทึกงานเลื่อนส่ง (อ่านก่อนกด)</div>
          <div>1) <b>"วันที่บันทึก" = วันที่ของรายงานที่ยอดเปลี่ยน</b> (วันที่ที่จะรับสินค้า) ไม่ใช่วันที่ที่ทำรายการ และต้องบันทึก <b>แยกทีละวันที่</b></div>
          <div>2) <b>📥 เลื่อนเข้า (+)</b> = งานจากวันอื่นเลื่อนมารับ<b>ในวันนี้</b> → ยอดเพิ่ม</div>
          <div>3) <b>📤 เลื่อนออก (-)</b> = งานของ<b>วันนี้</b>ถูกเลื่อนไปรับวันอื่น → ยอดลด</div>
          <div style="margin-top:4px"><b>ตัวอย่าง:</b> วันที่ 21 ทำรายงานของวันที่ 22 และมีงาน 10 พาเลท เลื่อนจากวันที่ 23 มารับวันที่ 22</div>
          <div>&nbsp;• วันที่บันทึก <b>22</b> → เลื่อนเข้า <b>10</b></div>
          <div>&nbsp;• วันที่บันทึก <b>23</b> → เลื่อนออก <b>10</b></div>
          <div style="margin-top:4px">4) ยอดสุทธิ = พาเลทใน RP01 + เลื่อนเข้า − เลื่อนออก</div>
          <div>5) ถ้าวันเดียวกันมีทั้งเข้าและออก ให้ใส่<b>ทั้งสองช่องแล้วกดบันทึกครั้งเดียว</b> (บันทึกซ้ำจะ<b>ทับ</b>ค่าเดิม ไม่ได้บวกเพิ่ม)</div>
          <div>6) บันทึกวันอื่นแล้ว ต้องเลือกวันนั้นที่ช่อง "Date" เพื่อดูผล</div>
        </div>
        """, unsafe_allow_html=True)
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
        st.caption(f"ค่าที่บันทึกไว้ของวันที่ {fmt_date(selected_date)}: เข้า +{postpone_in:,.0f} / ออก -{postpone_out:,.0f}")
# ---------------------- Main Dashboard -------------------
st.markdown(f"""
<div class="hero">
 <div><div class="hero-title">📋 รายงานประมาณการรับสินค้า ประจำวันที่ {fmt_date(selected_date)}</div>
 <div class="hero-sub">DATA ENTRY (DCSB) • RECEIVING FORECAST DASHBOARD</div></div>
 <div class="hero-logo">CP ALL</div>
</div>
""", unsafe_allow_html=True)

if not has_data:
    st.info("📂 วันที่นี้ยังไม่เคย Import ข้อมูล RP01 — ทุกตัวเลขแสดงเป็น 0 (Import ไฟล์ RP01 ที่แถบด้านซ้ายเพื่อบันทึกข้อมูลวันนี้)")

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

# Assign actual RP01 records to shifts using the time/slot column.
# Night: 00:00-05:00; morning: everything else (including unreadable times).
shift_time_col = find_col(df, ["Window Time"], contains=["windowtime", "slot"])
valid_po = df["_PO"].replace("", pd.NA)
total_po = int(valid_po.nunique())
if shift_time_col:
    df["_TIME_MINUTES"] = df[shift_time_col].map(parse_rp01_time).astype(float)
    shift_minutes = df["_TIME_MINUTES"]
    night_mask = shift_minutes < 5 * 60
    morning_mask = ~night_mask
    night_pallets = round(float(df.loc[night_mask, "_PALLETS"].sum()), 2)
    morning_pallets = round(float(df.loc[morning_mask, "_PALLETS"].sum()), 2)
    night_po = int(df.loc[night_mask, "_PO"].replace("", pd.NA).nunique())
    morning_po = int(df.loc[morning_mask, "_PO"].replace("", pd.NA).nunique())
else:
    night_pallets = morning_pallets = 0.0
    night_po = morning_po = 0
    st.warning("ไม่พบคอลัมน์เวลา/Slot ใน RP01 จึงคำนวณยอดแยกกะไม่ได้")
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
    ("📋","PO รวมทั้งหมด",f"{total_po:,}",f"กะดึก {night_po:,} + กะเช้า {morning_po:,} PO"),
    ("🌙","กะดึก (17:00-05:00)",f"{night_pallets:,.0f}",f"{night_po:,} PO ({night_pct:.1f}%)"),
    ("☀️","กะเช้า (05:00-17:00)",f"{morning_pallets:,.0f}",f"{morning_po:,} PO ({morning_pct:.1f}%)"),
]
cols = st.columns(7)
for i,(icon,label,value,note) in enumerate(kpis):
    cls = "night" if i == 5 else ("morning" if i == 6 else "")
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
"21265","21541","21653","21738","70082","70335","70341","70585","70605","70609","70757","71007","71316","71338","71573","71758",
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

# PMA pallet totals summary (no chart)
st.markdown('<div class="panel"><div class="panel-head cyan">📦 สรุปยอด Pallets ตาม PMA TYP</div><div class="panel-body">', unsafe_allow_html=True)
pma_col = find_col(df, ["PMA"], contains=["pma"])
if pma_col is None:
    st.info("ไม่พบคอลัมน์ PMA ในไฟล์ RP01 จึงยังสรุปยอดตาม PMA TYP ไม่ได้")
else:
    pma_view = df[[pma_col, "_PALLETS"]].copy()
    pma_view["PMA"] = pd.to_numeric(pma_view[pma_col], errors="coerce")
    pma_view = pma_view.dropna(subset=["PMA"])
    pma_view["PMA"] = pma_view["PMA"].astype(int)
    pma_view["PMA TYP"] = pma_view["PMA"].map(PMA_TYPE_MAP).fillna("ไม่พบประเภท")
    pma_summary = (pma_view.groupby("PMA TYP", as_index=False)
                   .agg(**{"จำนวน PMA": ("PMA", "nunique"), "ยอด Pallets": ("_PALLETS", "sum")})
                   .sort_values("ยอด Pallets", ascending=False))
    total_pma_pallets = float(pma_summary["ยอด Pallets"].sum()) if not pma_summary.empty else 0.0
    total_pma_count = int(pma_view["PMA"].nunique())
    total_types = int(pma_summary["PMA TYP"].nunique()) if not pma_summary.empty else 0
    k1, k2, k3 = st.columns(3)
    k1.metric("ยอด Pallets รวม", f"{total_pma_pallets:,.0f} พาเลท")
    k2.metric("จำนวน PMA", f"{total_pma_count:,} หมายเลข")
    k3.metric("จำนวน PMA TYP", f"{total_types:,} ประเภท")
    st.dataframe(pma_summary, use_container_width=True, hide_index=True)
st.markdown("</div></div>", unsafe_allow_html=True)

# Ti / Hi validation
with st.sidebar.expander("✅  Ti & Hi", expanded=False):
    if ti_col and hi_col:
        ti_num, hi_num = to_num(df[ti_col]), to_num(df[hi_col])
        alert = df.loc[(ti_num == 1) & (hi_num == 1)].copy()
        display_cols = [c for c in [vendor_id_col, po_col, item_col, ti_col, hi_col] if c]
        if alert.empty:
            st.success("ไม่พบรายการที่ Ti = 1 และ Hi = 1 พร้อมกัน")
        else:
            st.warning(f"พบ {len(alert):,} รายการที่ Ti = 1 และ Hi = 1")
            st.dataframe(alert[display_cols], use_container_width=True, hide_index=True)
    else:
        st.info("ไม่พบคอลัมน์ Ti และ/หรือ Hi ใน RP01")
# RP03 out-of-cycle comparison
with st.sidebar.expander("🔎 IMPORT RP03(งานรับนอกรอบ)", expanded=False):
    rp03_upload = st.file_uploader(
        "อัปโหลด RP03 (Goods Receive Note Summary Report)",
        type=["xlsx", "xls", "csv"], key="rp03_upload"
    )
    if rp03_upload is not None:
        try:
            rp03 = read_uploaded_table(rp03_upload)
            rp03_po_col = find_col(rp03, ["Purchase Order", "PO", "P/O"], contains=["purchaseorder", "pono"])
            qty_col = find_col(rp03, ["Total Case Qty", "Total Case Qty (iLPNs)", "iLPNs", "Pallets Qty"], contains=["totalcase", "ilpn", "pallet"])
            if rp03_po_col is None or qty_col is None:
                st.error("ไม่พบคอลัมน์ PO หรือ Total Case Qty ใน RP03")
            else:
                rp03["_PO"] = clean_po(rp03[rp03_po_col])
                rp03["_QTY"] = to_num(rp03[qty_col])
                main_pos = set(df["_PO"].loc[df["_PO"] != ""])
                mismatch = rp03.loc[(rp03["_PO"] != "") & (~rp03["_PO"].isin(main_pos))].copy()
                if mismatch.empty:
                    st.success("ไม่พบงานนอกรอบ — ทุก PO ใน RP03 ตรงกับ RP01")
                else:
                    st.metric("ยอด Total Case Qty ของงานนอกรอบ", f"{mismatch['_QTY'].sum():,.2f}", f"{len(mismatch):,} รายการ")
                    st.dataframe(
                        mismatch.drop(columns=["_PO", "_QTY"], errors="ignore").assign(
                            **{"PO ที่ตรวจสอบ": mismatch["_PO"], "จำนวนที่รวม": mismatch["_QTY"]}
                        ), use_container_width=True, hide_index=True
                    )
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        mismatch.to_excel(writer, index=False, sheet_name="RP03_Mismatch")
                    st.download_button(
                        "📥 ดาวน์โหลดรายงาน RP03 (Excel)", output.getvalue(),
                        "RP03_Mismatch_Report.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"ประมวลผล RP03 ไม่สำเร็จ: {e}")
    else:
        st.caption("อัปโหลด RP03 เพื่อเปรียบเทียบ PO และตรวจหารายการที่ไม่มีใน RP01")
with st.sidebar.expander("⚙️ Checking Window Time", expanded=False):
    # Window time comes from the RP01 time/slot column (per vendor), split by shift.
    # Night 00:00-05:00 | Morning = the rest
    if vendor_id_col and shift_time_col:
        wt = df[[vendor_id_col, "_TIME_MINUTES"]].copy()
        wt["Vendor ID"] = clean_po(wt[vendor_id_col])
        wt = wt[(wt["Vendor ID"] != "") & wt["_TIME_MINUTES"].notna()]
        if wt.empty:
            st.info("ไม่พบเวลา Window ใน RP01 ของวันที่เลือก")
        else:
            wt["_ORDER"] = wt["_TIME_MINUTES"].map(night_first_order)
            first = wt.sort_values("_ORDER").groupby("Vendor ID", as_index=False).first()
            last = wt.sort_values("_ORDER").groupby("Vendor ID", as_index=False).last()
            win = first[["Vendor ID", "_TIME_MINUTES", "_ORDER"]].rename(columns={"_TIME_MINUTES": "_FIRST"})
            win["_LAST"] = last["_TIME_MINUTES"].values
            if vendor_name_col:
                names = (df.assign(**{"Vendor ID": clean_po(df[vendor_id_col])})
                         .drop_duplicates("Vendor ID").set_index("Vendor ID")[vendor_name_col].astype(str))
                win["Vendor Name"] = win["Vendor ID"].map(names).fillna("")
            else:
                win["Vendor Name"] = ""
            win["Window Time (RP01)"] = win["_FIRST"].map(minutes_to_hms)
            win["ถึงเวลา"] = win["_LAST"].map(minutes_to_hms)
            win["กะ"] = win["_FIRST"].map(shift_label)
            win = win.sort_values("_ORDER")
            shift_pick = st.selectbox("แสดงกะ", ["ทั้งหมด", "🌙 กะดึก", "☀️ กะเช้า"], key="wt_shift")
            if shift_pick != "ทั้งหมด":
                win = win[win["กะ"] == shift_pick]
            st.dataframe(
                win[["Vendor ID", "Vendor Name", "Window Time (RP01)", "ถึงเวลา", "กะ"]],
                use_container_width=True, hide_index=True
            )
            st.caption("กะดึก 17:00-05:00 | กะเช้า 05:00-17:00 · เวลาอ้างอิงจากคอลัมน์ Time/Slot ใน RP01")
    else:
        st.info("ไม่พบคอลัมน์ Vendor หรือเวลา/Slot ใน RP01")
    st.markdown("**ตั้งค่า Manual (สำรองกรณี RP01 ไม่มีเวลา)**")
    with st.form("manual_window_form", clear_on_submit=False):
        new_v_id = st.text_input("รหัส Supplier ID")
        new_v_name = st.text_input("ชื่อ Supplier")
        new_w_time = st.text_input("กำหนด Window Time (เช่น 09:00:00)", value="00:00:00")
        submitted = st.form_submit_button("บันทึก / อัปเดตข้อมูล", use_container_width=True)
    if submitted:
        if not new_v_id.strip():
            st.warning("กรุณากรอกรหัส Supplier ID")
        else:
            with sqlite3.connect(DB_PATH) as con:
                con.execute(
                    """
                    INSERT INTO vendor_window_time (vendor_id, vendor_name, window_time, updated_at)
                    VALUES (?, ?, ?, datetime('now'))
                    ON CONFLICT(vendor_id) DO UPDATE SET
                        vendor_name = excluded.vendor_name,
                        window_time = excluded.window_time,
                        updated_at = datetime('now')
                    """,
                    (new_v_id.strip(), new_v_name.strip(), new_w_time.strip()),
                )
            st.success(f"บันทึก Window Time สำหรับรหัส {new_v_id} เรียบร้อยแล้ว")
            st.rerun()
