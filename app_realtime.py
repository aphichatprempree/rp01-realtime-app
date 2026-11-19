import json
import hashlib
import math
import re
from pathlib import Path
from datetime import date, datetime, time

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

def _norm_key(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    return value[:-2] if value.endswith(".0") else value


def _read_report(upload, required_headers):
    """Find the report header row from its known fields, then read its detail rows."""
    if upload is None:
        return None
    raw = pd.read_excel(upload, header=None, dtype=object)
    required = [str(x).strip().casefold() for x in required_headers]
    header_idx = None
    for i, row in raw.head(30).iterrows():
        labels = {str(v).strip().casefold() for v in row if pd.notna(v)}
        if all(any(req in label for label in labels) for req in required):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("หารายการหัวคอลัมน์ที่ต้องใช้ไม่พบ กรุณาตรวจว่าเลือกไฟล์รายงานถูกประเภท")
    upload.seek(0)
    frame = pd.read_excel(upload, header=header_idx, dtype=object)
    frame.columns = [str(c).strip() for c in frame.columns]
    frame = frame.loc[:, ~frame.columns.str.startswith("Unnamed:")]
    return frame.dropna(how="all").reset_index(drop=True)


def _column(frame, name):
    matches = [c for c in frame.columns if c.strip().casefold() == name.casefold()]
    if not matches:
        raise ValueError(f"ไม่พบคอลัมน์ {name}")
    return matches[0]


def _sum_numeric(frame, name, mask=None):
    col = _column(frame, name)
    values = pd.to_numeric(frame[col], errors="coerce").fillna(0)
    total = values[mask].sum() if mask is not None else values.sum()
    total = float(total)
    return int(total) if total.is_integer() else total


def _format_quantity(value):
    number = float(value)
    return f"{int(number):,}" if number.is_integer() else f"{number:,.2f}".rstrip("0").rstrip(".")


def _minutes_since_midnight(value):
    if pd.isna(value):
        return None
    if isinstance(value, (datetime, pd.Timestamp, time)):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)) and 0 <= float(value) < 1:
        return int(round(float(value) * 24 * 60))
    # RP exports may store Window Time as either 17:00 or 17.00 (sometimes
    # followed by the end of the range, e.g. 17.00-05.00).
    match = re.search(r"(?<!\d)(\d{1,2})\s*[:.]\s*(\d{2})(?!\d)", str(value).strip())
    if not match:
        return None
    hour, minute = int(match.group(1)), int(match.group(2))
    valid_time = minute < 60 and (hour < 24 or (hour == 24 and minute == 0))
    return hour * 60 + minute if valid_time else None


def _window_mask(frame, shift):
    col = _column(frame, "Window Time")
    minutes = frame[col].map(_minutes_since_midnight)
    # Shift windows are half-open intervals: night 18:00–07:00,
    # morning 07:00–18:00. This keeps boundary rows in exactly one shift.
    if shift == "night":
        return minutes.notna() & ((minutes >= 18 * 60) | (minutes < 7 * 60))
    return minutes.notna() & (minutes >= 7 * 60) & (minutes < 18 * 60)


def _load_morning_metrics(files):
    result = {}
    try:
        rp01_file = files["rp01_1610"] if files["rp01_1610"] is not None else files["rp01_2200"]
        if rp01_file is not None:
            rp01 = _read_report(rp01_file, ["Pallets Qty", "P/O", "Window Time"])
            result["report_date"] = _report_date_from_frame(rp01)
            result["inv_all"] = _sum_numeric(rp01, "Pallets Qty")
            # Divide RP01 Pallets Qty by its Window Time: night 18:00–06:59,
            # morning 07:00–17:59. Unreadable times go to morning so totals
            # still reconcile to the full RP01 Pallets Qty.
            night_mask = _window_mask(rp01, "night")
            morning_mask = _window_mask(rp01, "morning")
            result["night_received"] = _sum_numeric(rp01, "Pallets Qty", night_mask)
            morning_window_total = _sum_numeric(rp01, "Pallets Qty", morning_mask)
            unclassified = max(0, result["inv_all"] - result["night_received"] - morning_window_total)
            result["morning_received"] = morning_window_total + unclassified
        if files["rp42"] is not None:
            rp42 = _read_report(files["rp42"], ["Pallets Qty", "ASN_STATUS_DESC", "Window Time", "CHECK_IN_USER"])
            status = rp42[_column(rp42, "ASN_STATUS_DESC")].astype(str).str.strip().str.casefold()
            verified = status.isin({"verified", "receiving verified"})
            started = status.eq("receiving started")
            in_transit = status.eq("intransit")
            user = rp42[_column(rp42, "CHECK_IN_USER")]
            checked_in = user.notna() & user.astype(str).str.strip().ne("")
            result["complete_night"] = _sum_numeric(rp42, "Pallets Qty", verified & _window_mask(rp42, "night"))
            result["complete_morning"] = _sum_numeric(rp42, "Pallets Qty", verified & _window_mask(rp42, "morning"))
            result["recv_start"] = _sum_numeric(rp42, "Pallets Qty", started)
            result["in_transit"] = _sum_numeric(rp42, "Pallets Qty", in_transit & checked_in)
            result["sup_night_missing"] = _sum_numeric(
                rp42, "Pallets Qty", in_transit & ~checked_in & _window_mask(rp42, "night")
            )
            result["sup_morning_missing"] = _sum_numeric(
                rp42, "Pallets Qty", in_transit & ~checked_in & _window_mask(rp42, "morning")
            )
        if files["rp03"] is not None:
            rp03 = _read_report(files["rp03"], ["Total Case Qty (iLPNs)", "Schedule Date"])
            schedule = pd.to_datetime(rp03[_column(rp03, "Schedule Date")], dayfirst=True, errors="coerce").dt.date
            today = pd.Timestamp.now(tz="Asia/Bangkok").date()
            result["extra_work"] = _sum_numeric(
                rp03, "Total Case Qty (iLPNs)", schedule.notna() & schedule.ne(today)
            )
        if all(k in result for k in ("recv_start", "in_transit", "sup_night_missing", "sup_morning_missing")):
            result["total_remaining"] = (
                result["recv_start"] + result["in_transit"]
                + result["sup_night_missing"] + result["sup_morning_missing"]
            )
        return result, None
    except Exception as exc:
        return result, str(exc)


def _report_date_from_frame(frame):
    """Use the most common delivery date in RP01 as the report date."""
    try:
        col = _column(frame, "Delivery Date")
    except ValueError:
        return None
    dates = pd.to_datetime(frame[col], dayfirst=True, errors="coerce").dropna()
    if dates.empty:
        return None
    return dates.dt.date.value_counts().index[0]


def _date_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    patterns = (
        r"\d{4}-\d{1,2}-\d{1,2}",
        r"\d{1,2}/\d{1,2}/\d{2,4}",
        r"\d{1,2}/[A-Za-z]{3}/\d{4}",
        r"[A-Za-z]{3}\s+\d{1,2},\s*\d{4}",
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            parsed = pd.to_datetime(match.group(0), dayfirst=True, errors="coerce")
            if not pd.isna(parsed):
                return parsed.date()
    return None


def _report_metadata_date(upload):
    """Read the export/report date from the preamble above a report's table."""
    if upload is None:
        return None
    upload.seek(0)
    raw = pd.read_excel(upload, header=None, dtype=object, nrows=12)
    upload.seek(0)
    for _, row in raw.iterrows():
        values = list(row)
        for index, value in enumerate(values):
            if "report date" not in str(value).strip().casefold():
                continue
            parsed = _date_value(value)
            if parsed:
                return parsed
            for next_value in values[index + 1:]:
                parsed = _date_value(next_value)
                if parsed:
                    return parsed
    return None


def _report_date_warnings(files, rp01_date):
    if rp01_date is None:
        if files.get("rp01_1610") is not None or files.get("rp01_2200") is not None:
            return ["อ่านวันที่จาก Delivery Date ใน RP01 ไม่ได้ จึงยังตรวจสอบวันที่ของ RP03 และ RP42 ไม่ได้"]
        return []
    warnings = []
    for key, label in (("rp03", "RP03"), ("rp42", "RP42")):
        upload = files.get(key)
        if upload is None:
            continue
        try:
            source_date = _report_metadata_date(upload)
        except Exception:
            source_date = None
        if source_date is None:
            warnings.append(f"ไม่พบ Report Date ใน {label} ({upload.name}) จึงตรวจสอบวันเทียบกับ RP01 ไม่ได้")
        elif source_date != rp01_date:
            warnings.append(
                f"วันที่รายงานไม่ตรงกัน: RP01 {rp01_date:%d/%m/%Y} แต่ {label} {source_date:%d/%m/%Y} ({upload.name})"
            )
    return warnings


def _extra_complete(rp01_22, rp03):
    """Sum only RP03 case quantities absent from RP01 Pallets Qty (VLOOKUP #N/A)."""
    rp01_col = _column(rp01_22, "Pallets Qty")
    rp03_col = _column(rp03, "Total Case Qty (iLPNs)")
    rp01_qty = pd.to_numeric(rp01_22[rp01_col], errors="coerce").dropna()
    rp03_qty = pd.to_numeric(rp03[rp03_col], errors="coerce")
    unmatched = rp03_qty.notna() & ~rp03_qty.isin(set(rp01_qty.tolist()))
    missing_count = int(unmatched.sum())
    total = float(rp03_qty[unmatched].sum())
    total = int(total) if total.is_integer() else total
    return total, rp01_col, rp03_col, missing_count


def _po_values(text_value):
    values = re.split(r"[\s,;]+", str(text_value or "").strip())
    return list(dict.fromkeys(key for value in values if (key := _norm_key(value))))


def _item_column(frame, preferred):
    preferred_match = next((c for c in frame.columns if c.strip().casefold() == preferred.casefold()), None)
    if preferred_match is not None:
        return preferred_match
    aliases = {
        "item code", "item name", "item", "item number", "item no", "item no.",
        "sku", "product code",
    }
    return next((c for c in frame.columns if c.strip().casefold() in aliases), None)


def _lookup_delayed_pallets(rp01_file, po_text):
    frame = _read_report(rp01_file, ["P/O", "Pallets Qty"])
    requested = set(_po_values(po_text))
    po_col = _column(frame, "P/O")
    qty_col = _column(frame, "Pallets Qty")
    matched = frame[po_col].map(_norm_key).isin(requested)
    found = set(frame.loc[matched, po_col].map(_norm_key))
    missing = sorted(requested - found)
    total = pd.to_numeric(frame.loc[matched, qty_col], errors="coerce").fillna(0).sum()
    total = float(total)
    return (int(total) if total.is_integer() else total), missing


def _lookup_reject_pallets(rp01_file, rp03_file, item_text):
    requested = {value.casefold() for value in _po_values(item_text)}
    if not requested:
        return 0, [], []
    rp01 = _read_report(rp01_file, ["Pack Qty", "TI", "HI"])
    rp03 = _read_report(rp03_file, ["Total Pack Qty"])
    item1, item3 = _item_column(rp01, "Item Name"), _item_column(rp03, "Item Code")
    if item1 is None:
        raise ValueError("ไม่พบคอลัมน์ Item Name ใน RP01")
    if item3 is None:
        raise ValueError("ไม่พบคอลัมน์ Item Code ใน RP03")
    rp01 = rp01.assign(
        _item=rp01[item1].map(lambda value: _norm_key(value).casefold()),
        _expected=pd.to_numeric(rp01[_column(rp01, "Pack Qty")], errors="coerce").fillna(0),
        _ti=pd.to_numeric(rp01[_column(rp01, "TI")], errors="coerce"),
        _hi=pd.to_numeric(rp01[_column(rp01, "HI")], errors="coerce"),
    )
    rp03 = rp03.assign(
        _item=rp03[item3].map(lambda value: _norm_key(value).casefold()),
        _received=pd.to_numeric(rp03[_column(rp03, "Total Pack Qty")], errors="coerce").fillna(0),
    )
    rp01 = rp01[rp01["_item"].isin(requested) & rp01["_item"].ne("")]
    rp03 = rp03[rp03["_item"].isin(requested) & rp03["_item"].ne("")]
    expected = rp01.groupby("_item", as_index=False).agg(
        _expected=("_expected", "sum"), _ti=("_ti", "first"), _hi=("_hi", "first")
    )
    received = rp03.groupby("_item", as_index=False)["_received"].sum()
    ti_hi_counts = (
        rp01[["_item", "_ti", "_hi"]]
        .dropna(subset=["_ti", "_hi"])
        .drop_duplicates()
        .groupby("_item")
        .size()
    )
    ambiguous_items = set(ti_hi_counts[ti_hi_counts.gt(1)].index)
    expected = expected[~expected["_item"].isin(ambiguous_items)]
    matched_items = set(expected["_item"])
    missing = sorted(requested - matched_items - ambiguous_items)
    joined = expected.merge(received, on="_item", how="left")
    joined["_received"] = joined["_received"].fillna(0)
    valid = (
        joined["_ti"].notna() & joined["_ti"].gt(0)
        & joined["_hi"].notna() & joined["_hi"].gt(0)
    )
    invalid_items = joined.loc[~valid, "_item"].drop_duplicates().tolist()
    joined = joined.loc[valid].copy()
    joined["_reject_pack"] = (joined["_expected"] - joined["_received"]).clip(lower=0)
    total = (joined["_reject_pack"] / joined["_ti"] / joined["_hi"]).sum()
    total = math.ceil(float(total))
    invalid_items.extend(f"Item {item} มี TI/HI หลายชุด" for item in sorted(ambiguous_items))
    return total, missing, invalid_items


def _load_imported_metrics(files):
    result = {}
    try:
        if files["rp01_1610"] is not None:
            d = _read_report(files["rp01_1610"], ["Pallets Qty", "P/O"])
            result["hist_total_recv"] = _sum_numeric(d, "Pallets Qty")
            result["report_date"] = _report_date_from_frame(d)
        if files["rp01_2200"] is not None:
            d22 = _read_report(files["rp01_2200"], ["Pallets Qty", "P/O"])
            result["hist_not_billed"] = _sum_numeric(d22, "Pallets Qty")
            if not result.get("report_date"):
                result["report_date"] = _report_date_from_frame(d22)
        else:
            d22 = None
        if files["rp42"] is not None:
            d42 = _read_report(files["rp42"], ["Pallets Qty", "ASN_STATUS_DESC"])
            status = _column(d42, "ASN_STATUS_DESC")
            result["hist_complete"] = _sum_numeric(
                d42, "Pallets Qty", d42[status].astype(str).str.strip().str.casefold().eq("receiving verified")
            )
        if files["rp03"] is not None and d22 is not None:
            d03 = _read_report(files["rp03"], ["Total Case Qty (iLPNs)", "Purchase Order"])
            extra, left, right, overlap = _extra_complete(d22, d03)
            result["hist_extra"] = extra
            result["extra_join"] = (
                f"VLOOKUP {left} (RP01 22.00) ↔ {right} (RP03): "
                f"พบค่า #N/A {overlap:,} แถว นับเฉพาะแถวที่ไม่ตรงกัน"
            )
        return result, None
    except Exception as exc:
        return result, str(exc)

REPORT_DEFAULTS = {
    "hist_inv": 6126,
    "hist_total_recv": 5904,
    "hist_complete": 5722,
    "hist_extra": 138,
    "hist_delayed": 0,
    "hist_not_po": 73,
    "hist_not_billed": 250,
}
HISTORY_PO_FIELDS = ("hist_delayed_po_input", "hist_reject_item_input", "hist_po_rp01_source")
MORNING_DEFAULTS = {
    "inv_all": 6058,
    "night_received": 4240,
    "morning_received": 1818,
    "complete_night": 4475,
    "complete_morning": 0,
    "extra_work": 0,
    "total_remaining": 2425,
    "recv_start": 473,
    "in_transit": 268,
    "sup_night_missing": 239,
    "sup_morning_missing": 0,
}


REPORT_STORE_PATH = Path(__file__).with_name("daily_reports.json")


def _load_report_store():
    try:
        with REPORT_STORE_PATH.open("r", encoding="utf-8") as stream:
            data = json.load(stream)
        return data if isinstance(data.get("records"), dict) else {"records": {}}
    except (OSError, ValueError, TypeError):
        return {"records": {}}


def _get_report_record(report_date, page="history"):
    records = REPORT_STORE.get("records", {})
    return records.get(f"{page}:{report_date}") or (records.get(report_date) if page == "history" else None)


def _save_report_record(report_date, values, page="history"):
    store = _load_report_store()
    fields = MORNING_DEFAULTS if page == "morning" else REPORT_DEFAULTS
    store.setdefault("records", {})[f"{page}:{report_date}"] = {
        **{field: values.get(field, fields[field]) for field in fields},
        **({field: values.get(field, "") for field in HISTORY_PO_FIELDS} if page == "history" else {}),
        "updated_at": pd.Timestamp.now().isoformat(timespec="seconds"),
    }
    store["morning_latest_date" if page == "morning" else "latest_date"] = report_date
    temp_path = REPORT_STORE_PATH.with_suffix(".json.tmp")
    with temp_path.open("w", encoding="utf-8") as stream:
        json.dump(store, stream, ensure_ascii=False, indent=2)
    temp_path.replace(REPORT_STORE_PATH)


def _autosave_current_report():
    report_date = st.session_state.get("h_date_input")
    if not report_date:
        return
    values = {
        field: st.session_state.get(field, REPORT_DEFAULTS[field])
        for field in REPORT_DEFAULTS
    }
    values.update({field: st.session_state.get(field, "") for field in HISTORY_PO_FIELDS})
    try:
        _save_report_record(report_date.strftime("%Y-%m-%d"), values)
        st.session_state["_autosaved_at"] = pd.Timestamp.now().strftime("%H:%M:%S")
    except OSError as exc:
        st.session_state["_autosave_error"] = str(exc)


def _autosave_morning_report():
    report_date = st.session_state.get("m_date_input")
    if not report_date:
        return
    values = {field: st.session_state.get(field, default) for field, default in MORNING_DEFAULTS.items()}
    values["total_remaining"] = sum(values[field] for field in ("recv_start", "in_transit", "sup_night_missing", "sup_morning_missing"))
    try:
        _save_report_record(report_date.strftime("%Y-%m-%d"), values, page="morning")
        st.session_state["_morning_autosaved_at"] = pd.Timestamp.now().strftime("%H:%M:%S")
    except OSError as exc:
        st.session_state["_morning_autosave_error"] = str(exc)


# ตั้งค่าหน้าเว็บแดชบอร์ดหลัก
st.set_page_config(
    page_title="DCSB Distribution Center Morning Report",
    page_icon="🟦",
    layout="wide",
    initial_sidebar_state="expanded",
)

REPORT_STORE = _load_report_store()

# ชุดสีโทนน้ำเงิน-เทาฟ้า ตามธีม Soft UI
RAINBOW = [
    "#3a3dff", "#2e8bff", "#7b6cff", "#8d9bb3",
    "#27b3d6", "#5b6b8c", "#9aa8ff", "#b4bfd0",
]
INK = "#2d3650"
MUTED = "#8590a6"
FONT = "IBM Plex Sans Thai, sans-serif"
CHART_TRACK = "rgba(200,209,222,0.55)"
CHART_INNER = "rgba(255,255,255,0.7)"
CHART_GAP = "#eef2f8"
CHART_LEGEND_BORDER = "#9aaccb"
BAR_TEXT = "white"
BUBBLE_TEXT = None

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap');
:root {
    --ink:#2d3650; --muted:#8590a6; --bg:#e4e9f1; --bg2:#eef2f8;
    --blue:#3a3dff; --blue2:#2e8bff; --rose:#e0536a;
    --sh: 8px 8px 18px rgba(163,177,198,.55), -8px -8px 18px rgba(255,255,255,.95);
    --sh-sm: 5px 5px 11px rgba(163,177,198,.5), -5px -5px 11px rgba(255,255,255,.95);
    --sh-in: inset 4px 4px 9px rgba(163,177,198,.45), inset -4px -4px 9px rgba(255,255,255,.9);
    --grad: linear-gradient(135deg,#3a3dff,#2e8bff);
    --line:#9aaccb; --line-strong:#5f77a6;
}
html, body, [class*="css"], [data-testid="stAppViewContainer"] { font-family:'IBM Plex Sans Thai',sans-serif; color:var(--ink); }
[data-testid="stAppViewContainer"] { background: linear-gradient(145deg,#e9eef5,#dde4ee); }
[data-testid="stHeader"] { background:transparent; height:2.2rem; }
[data-testid="stToolbar"], [data-testid="stDecoration"], footer { display:none !important; }
.block-container { padding:.6rem 1.4rem 1rem; max-width:1700px; }
h1,h2,h3 { color:var(--ink); letter-spacing:-.02em; }
[data-testid="stVerticalBlock"] { gap:.55rem; }

/* ---------- HERO ---------- */
.hero { background:linear-gradient(145deg,#eef2f8,#e1e7f0); border-radius:26px; padding:14px 26px 12px; text-align:center; box-shadow:var(--sh); border:2px solid var(--line-strong); margin:0 auto 8px; max-width:1000px; }
.hero-kicker { color:var(--blue2); font-size:.68rem; font-weight:600; letter-spacing:.2em; }
.hero-title { color:var(--ink); font-size:1.5rem; font-weight:700; margin:2px 0 0; }
.hero-subtitle { color:var(--muted); font-size:.82rem; }
.hero-bar { height:4px; border-radius:4px; margin:9px auto 0; max-width:200px; background:var(--grad); box-shadow:0 3px 8px rgba(58,61,255,.3); }
.page-sub { display:none; }
.section-title { display:flex; align-items:center; gap:10px; color:var(--ink); font-size:1rem; font-weight:700; margin:14px 0 4px; padding-bottom:6px; border-bottom:2px solid var(--line); }
.section-title::before { content:""; width:5px; height:20px; border-radius:5px; background:var(--grad); box-shadow:0 2px 7px rgba(58,61,255,.3); }

/* ---------- FOCUS CARDS ---------- */
.focus-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:18px; margin:6px 0 6px; }
.focus-card { background:linear-gradient(145deg,#eef2f8,#e1e7f0); border-radius:24px; padding:16px 24px 14px 30px; box-shadow:var(--sh); border:2px solid var(--line-strong); position:relative; overflow:hidden; }
.focus-card::before { content:""; position:absolute; left:0; top:16px; bottom:16px; width:5px; border-radius:0 5px 5px 0; background:var(--grad); }
.focus-card.alert::before { background:linear-gradient(180deg,#ffb35c,#e0536a); }
.focus-label { font-size:.82rem; font-weight:600; color:var(--muted); }
.focus-value { font-size:2.1rem; font-weight:700; color:var(--blue); line-height:1.25; margin:6px 0 4px; font-variant-numeric:tabular-nums; }
.focus-card.alert .focus-value { color:var(--rose); }
.focus-value span { font-size:.95rem; font-weight:600; color:var(--muted); }
.focus-unit { font-size:.74rem; color:var(--muted); font-weight:500; }

/* ---------- REPORT TABLE ---------- */
.report-wrap { background:linear-gradient(145deg,#eef2f8,#e1e7f0); border-radius:20px; overflow:hidden; box-shadow:var(--sh); border:2px solid var(--line-strong); padding:5px 7px 8px; }
.report-head { color:var(--ink); text-align:left; padding:5px 10px 6px; margin-bottom:2px; border-bottom:2px solid var(--line); font-size:.98rem; font-weight:700; }
.report-date { font-size:.82rem; font-weight:500; color:var(--muted); margin-left:5px; }
.report-grid { display:grid; gap:7px; padding:2px 3px; }
.report-group { grid-column:1 / -1; color:var(--ink); font-size:.9rem; line-height:1.2; font-weight:700; padding:3px 3px 2px; border-bottom:1px solid var(--line); letter-spacing:.01em; }
.report-cell { border-radius:12px; padding:6px 6px; text-align:center; background:#f1f4f9; box-shadow:var(--sh-sm); border:1.5px solid var(--line); min-height:68px; box-sizing:border-box; display:flex; flex-direction:column; justify-content:space-between; gap:4px; }
.report-label { color:#4d5a78; font-size:.78rem; line-height:1.25; font-weight:500; }
.report-value { color:var(--ink); font-size:1.22rem; font-weight:700; line-height:1.2; overflow-wrap:anywhere; font-variant-numeric:tabular-nums; }
.report-wrap.compact { padding:4px 6px 6px; border-radius:16px; }
.report-wrap.compact .report-head { padding:4px 8px 5px; margin-bottom:1px; }
.report-wrap.compact .report-grid { gap:5px; padding:2px; }
.report-wrap.compact .report-group { font-size:.82rem; padding:2px 2px 1px; }
.report-wrap.compact .report-cell { min-height:58px; padding:4px 5px; border-radius:10px; gap:3px; }
.report-wrap.compact .report-label { font-size:.72rem; line-height:1.16; }
.report-wrap.compact .report-value { font-size:1.1rem; }
.report-value.alert { color:var(--rose); }

/* ---------- CHART CARDS ---------- */
[class*="st-key-card"] { background:linear-gradient(145deg,#eef2f8,#e1e7f0); border-radius:26px; box-shadow:var(--sh); padding:.8rem .9rem .6rem; border:2px solid var(--line-strong); }
.chart-title { color:var(--ink); font-size:.92rem; font-weight:700; margin:0 0 4px; text-align:center; }
.chart-sub { color:var(--muted); font-size:.7rem; text-align:center; margin-bottom:6px; padding-bottom:6px; border-bottom:1.5px dashed var(--line); }
@media (max-width:1200px) { .report-grid:not(.compact-grid) { grid-template-columns:repeat(4,minmax(0,1fr)) !important; } .report-grid.compact-grid { grid-template-columns:repeat(6,minmax(0,1fr)) !important; } .focus-grid { grid-template-columns:1fr; } }
@media (max-width:900px) { .report-grid.compact-grid { grid-template-columns:repeat(3,minmax(0,1fr)) !important; } }
@media (max-width:700px) { .hero-title{font-size:1.25rem} .report-grid:not(.compact-grid){grid-template-columns:repeat(2,minmax(0,1fr)) !important} .report-grid.compact-grid{grid-template-columns:repeat(2,minmax(0,1fr)) !important} .focus-value{font-size:1.8rem} }

/* ---------- EXPANDER + INPUTS ---------- */
[data-testid="stExpander"] { background:linear-gradient(145deg,#eef2f8,#e1e7f0); border:2px solid var(--line-strong) !important; border-radius:22px; box-shadow:var(--sh-sm); }
[data-testid="stExpander"] summary p { font-weight:700; color:var(--ink); }
[data-testid="stNumberInput"] > label p { font-size:.8rem; font-weight:600; color:#56627d; }
[data-testid="stNumberInput"] [data-baseweb="input"], [data-testid="stNumberInput"] [data-baseweb="base-input"] { background:#e6ebf3 !important; border:1.5px solid var(--line) !important; border-radius:12px; box-shadow:var(--sh-in); }
[data-testid="stNumberInput"] input { color:var(--ink); background:transparent !important; }
[data-testid="stNumberInput"] button { background:transparent; color:var(--blue); }
.stButton > button { background:var(--grad); color:#fff; border:0; border-radius:14px; font-weight:600; padding:.55rem 1.1rem; box-shadow:0 8px 16px rgba(58,61,255,.28); }
.stButton > button:hover { color:#fff; filter:brightness(1.07); }

/* ---------- SIDEBAR (soft light) ---------- */
[data-testid="stSidebar"] { background:linear-gradient(185deg,#eaeff6,#dde4ee); border-right:2px solid var(--line-strong); box-shadow:8px 0 24px rgba(163,177,198,.35); }
[data-testid="stSidebarContent"] { padding:0 1rem 1rem; }
[data-testid="stSidebarHeader"] { position:absolute; top:.4rem; right:.4rem; z-index:5; height:auto; padding:0; width:auto; background:transparent; }
[data-testid="stSidebarHeader"] button { color:var(--muted); }
[data-testid="stSidebarUserContent"] { padding-top:1rem; }
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap:.45rem; }
[data-testid="stSidebar"] hr { display:none; }
.sb-brand { display:flex; align-items:center; gap:12px; padding:4px 6px 14px; margin-bottom:8px; border-bottom:2px solid var(--line); }
.sb-logo { width:46px; height:46px; border-radius:15px; display:grid; place-items:center; flex:none; background:linear-gradient(145deg,#4a4dff,#2e8bff); border:2px solid #2a2cc9; box-shadow:0 8px 16px rgba(58,61,255,.35), inset 2px 2px 4px rgba(255,255,255,.35), inset -3px -3px 6px rgba(20,30,160,.35); }
.sb-logo svg { width:27px; height:27px; display:block; filter:drop-shadow(0 2px 2px rgba(20,30,160,.4)); }
.sb-name { color:var(--ink); font-weight:700; font-size:1rem; line-height:1.2; }
.sb-sub { color:var(--muted); font-size:.68rem; letter-spacing:.08em; }
.sb-label { color:var(--muted); font-size:.72rem; font-weight:600; letter-spacing:.05em; text-transform:uppercase; margin:14px 4px 0; padding-bottom:.5rem; }
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] { gap:.5rem; }
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] > div { width:100%; }
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"],
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] { width:100%; margin:0; padding:.65rem .8rem; border-radius:16px; background:#f1f4f9; box-shadow:var(--sh-sm); border:1.5px solid var(--line); transition:all .16s ease; cursor:pointer; box-sizing:border-box; }
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"] > div > div:first-child,
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] > div:first-child { display:none; }
[data-testid="stSidebar"] [data-testid="stRadio"] label p { color:#56627d; font-weight:600; font-size:.78rem; line-height:1.4; margin:0; }
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"]:has(input:checked),
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) { background:var(--grad); border-color:#2a2cc9 !important; box-shadow:0 9px 18px rgba(58,61,255,.3); }
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) p { color:#fff; }
[data-testid="stSidebar"] label p, [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p { color:#56627d; font-size:.78rem; font-weight:500; }
[data-testid="stSidebar"] [data-baseweb="input"], [data-testid="stSidebar"] [data-baseweb="base-input"] { background:#e6ebf3 !important; border-radius:12px; border:0 !important; box-shadow:var(--sh-in); }
[data-testid="stSidebar"] [data-baseweb="input"] input { color:var(--ink) !important; -webkit-text-fill-color:var(--ink); background:transparent !important; }
[data-testid="stSidebar"] [data-testid="stDateInputField"] { background:#e6ebf3 !important; border:1.5px solid var(--line) !important; border-radius:12px; box-shadow:var(--sh-in); }
[data-testid="stSidebar"] [data-testid="stDateInputField"] * { color:var(--ink) !important; -webkit-text-fill-color:var(--ink); }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { background:#e6ebf3; border:1.5px dashed var(--line-strong); border-radius:14px; box-shadow:var(--sh-in); padding:.5rem .7rem; display:flex; flex-direction:column; align-items:flex-start; gap:.3rem; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * { color:#6a7690; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small, [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span { font-size:.66rem !important; line-height:1.3; }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button { font-size:.76rem; background:var(--grad); border:0; border-radius:10px; color:#fff; padding:.22rem .7rem; box-shadow:0 5px 10px rgba(58,61,255,.25); }
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * { color:#fff; }
[data-testid="stSidebar"] [data-testid="stFileUploaderFile"] * { color:var(--ink); }
.sb-foot { display:flex; gap:8px; flex-wrap:wrap; justify-content:center; margin-top:14px; padding-top:12px; border-top:2px solid var(--line); }
.sb-chip { font-size:.7rem; font-weight:600; padding:5px 13px; border-radius:999px; background:#f1f4f9; border:1.5px solid var(--line); box-shadow:var(--sh-sm); color:var(--muted); }
.sb-chip.on { background:var(--grad); color:#fff; box-shadow:0 5px 12px rgba(58,61,255,.3); }

/* Shared 3D finish for every report page */
.hero, .focus-card, .report-wrap, [class*="st-key-card"], [data-testid="stExpander"] {
    box-shadow:0 2px 4px rgba(43,57,78,.08),0 12px 24px rgba(87,103,129,.16),0 24px 42px rgba(87,103,129,.07),inset 0 2px 0 rgba(255,255,255,.72),inset 0 -2px 5px rgba(79,95,119,.06);
    transition:transform .18s ease,box-shadow .18s ease;
}
.focus-card:hover, [class*="st-key-card"]:hover { transform:translateY(-3px); box-shadow:0 4px 7px rgba(43,57,78,.1),0 17px 30px rgba(87,103,129,.2),0 30px 48px rgba(87,103,129,.09),inset 0 2px 0 rgba(255,255,255,.86); }
.report-cell { box-shadow:0 2px 4px rgba(43,57,78,.09),0 7px 13px rgba(87,103,129,.12),inset 0 2px 0 rgba(255,255,255,.86),inset 0 -2px 4px rgba(79,95,119,.05); transition:transform .16s ease,box-shadow .16s ease; }
.report-cell:hover { transform:translateY(-2px); box-shadow:0 4px 6px rgba(43,57,78,.1),0 11px 18px rgba(87,103,129,.17),inset 0 2px 0 rgba(255,255,255,.92); }
.hero { border-top-color:rgba(255,255,255,.95); }
[data-testid="stExpander"] { border-top-color:rgba(255,255,255,.92) !important; }
[data-testid="stSidebar"] [data-testid="stRadio"] label,
[data-testid="stSidebar"] [data-testid="stDateInputField"],
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"],
[data-testid="stNumberInput"] [data-baseweb="input"],
[data-testid="stNumberInput"] [data-baseweb="base-input"] { box-shadow:var(--sh-sm),inset 0 1px 0 rgba(255,255,255,.8); }
[data-testid="stSidebar"] [data-testid="stRadio"] label { transition:transform .16s ease,box-shadow .16s ease; }
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover { transform:translateY(-2px); box-shadow:0 7px 14px rgba(87,103,129,.16),inset 0 1px 0 rgba(255,255,255,.86); }
.stButton > button, [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button { box-shadow:0 4px 0 rgba(50,58,174,.5),0 8px 15px rgba(58,61,255,.22),inset 0 1px 0 rgba(255,255,255,.45); transition:transform .14s ease,box-shadow .14s ease; }
.stButton > button:hover, [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover { transform:translateY(-2px); box-shadow:0 6px 0 rgba(50,58,174,.48),0 12px 19px rgba(58,61,255,.26),inset 0 1px 0 rgba(255,255,255,.5); }
.sb-chip { box-shadow:0 3px 6px rgba(87,103,129,.12),inset 0 1px 0 rgba(255,255,255,.82); }
.sb-chip.on { box-shadow:0 5px 12px rgba(58,61,255,.28),inset 0 1px 0 rgba(255,255,255,.45); }
.report-cell.status-complete { background:linear-gradient(160deg,#fff,#f8fbf8) !important; border-color:#e0e8e1 !important; border-top-color:#5d9568 !important; }
.report-cell.status-complete .report-value { color:#477d52; }
.report-cell.status-progress { background:linear-gradient(160deg,#fff,#fcf9f3) !important; border-color:#ebe4d7 !important; border-top-color:#c4832d !important; }
.report-cell.status-progress .report-value { color:#9d6925; }
.report-cell.status-transit { background:linear-gradient(160deg,#fff,#f7f9fc) !important; border-color:#dde5ed !important; border-top-color:#5d8fba !important; }
.report-cell.status-transit .report-value { color:#4a779c; }
.report-cell.status-alert { background:linear-gradient(160deg,#fff,#fcf7f7) !important; border-color:#ecdede !important; border-top-color:#bc5a5c !important; }
.report-cell.status-alert .report-value { color:#a64d4f; }
.report-cell.status-special { background:linear-gradient(160deg,#fff,#f9f7fc) !important; border-color:#e5dfed !important; border-top-color:#896eb0 !important; }
.report-cell.status-special .report-value { color:#765d99; }
.report-cell.status-neutral { background:linear-gradient(160deg,#fff,#f8f9fa) !important; border-color:#e1e4e6 !important; border-top-color:#9ba2a7 !important; }
.report-cell.status-neutral .report-value { color:#626a70; }
.focus-card.status-complete { background:linear-gradient(145deg,#fff,#f4faf5) !important; border-color:#d8e5da !important; }
.focus-card.status-complete .focus-value { color:#477d52 !important; }
</style>

""",
    unsafe_allow_html=True,
)


# =========================================================================
# ฟังก์ชันวาดกราฟ
# =========================================================================
def _pct(value, base):
    return (value / base * 100) if base and base > 0 else 0.0


def radial_chart(labels, values, base=None, height=300, palette=None):
    """กราฟพัดครึ่งวงกลม + ฟองเปอร์เซ็นต์ครึ่งซ้าย"""
    n = len(labels)
    base = base if base else sum(values)
    maxv = max(values) if max(values) > 0 else 1
    width = 180 / n
    chart_palette = palette or RAINBOW
    colors = [chart_palette[i % len(chart_palette)] for i in range(n)]

    fig = go.Figure()

    fig.add_trace(
        go.Barpolar(
            r=[1.12], theta=[180], width=[180],
            marker=dict(color=CHART_TRACK, line=dict(width=0)),
            hoverinfo="skip", showlegend=False,
        )
    )
    fig.add_trace(
        go.Barpolar(
            r=[0.72], theta=[180], width=[180],
            marker=dict(color=CHART_INNER, line=dict(width=0)),
            hoverinfo="skip", showlegend=False,
        )
    )

    for i, (lb, v) in enumerate(zip(labels, values)):
        center = 90 - width * (i + 0.5)
        r = 0.6 + 0.4 * (v / maxv)
        fig.add_trace(
            go.Barpolar(
                r=[r], theta=[center], width=[width], name=lb,
                marker=dict(color=colors[i], line=dict(color=CHART_GAP, width=2)),
                hovertemplate=f"{lb}<br>{v:,} Pallets<extra></extra>",
            )
        )
        fig.add_trace(
            go.Scatterpolar(
                r=[r - 0.12], theta=[center], mode="text",
                text=[f"{v:,}"],
                textfont=dict(color=BAR_TEXT, size=13, family=FONT),
                hoverinfo="skip", showlegend=False,
            )
        )

    angles = [180] if n == 1 else [100 + (160 / (n - 1)) * i for i in range(n)]
    fig.add_trace(
        go.Scatterpolar(
            r=[0.9] * n, theta=angles, mode="markers+text",
            text=[f"{_pct(v, base):.0f}%" for v in values],
            marker=dict(size=40, color="white", line=dict(color=colors, width=3)),
            textfont=dict(color=BUBBLE_TEXT or colors, size=11, family=FONT),
            hovertext=labels, hoverinfo="text", showlegend=False,
        )
    )

    fig.update_layout(
        height=height,
        margin=dict(t=10, b=46, l=10, r=10),
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, color=INK),
        polar=dict(
            bgcolor="rgba(0,0,0,0)",
            hole=0.16,
            radialaxis=dict(visible=False, range=[0, 1.12]),
            angularaxis=dict(visible=False, rotation=0, direction="counterclockwise"),
        ),
        legend=dict(
            orientation="h", yanchor="top", y=-0.04, xanchor="center", x=0.5,
            font=dict(size=11), bgcolor="rgba(255,255,255,.6)", bordercolor=CHART_LEGEND_BORDER, borderwidth=1,
        ),
        annotations=[
            dict(
                x=0.5, y=0.5, xref="paper", yref="paper",
                text="<b>REPORT</b>", showarrow=False,
                font=dict(size=10, color=MUTED),
            )
        ],
    )
    return fig


def tube_chart(labels, values, base, height=320, palette=None):
    """กราฟหลอดแท่งเงา + เส้นประเปอร์เซ็นต์ด้านบน"""
    n = len(labels)
    chart_palette = palette or RAINBOW
    colors = [chart_palette[i % len(chart_palette)] for i in range(n)]
    xs = list(range(n))
    maxv = max(values) if max(values) > 0 else 1
    bg_max = maxv * 1.08
    pcts = [_pct(v, base) for v in values]

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=xs, y=[bg_max] * n, width=0.68,
            marker=dict(color="rgba(255,255,255,0.95)", line=dict(color="#d8dde1", width=3)),
            hoverinfo="skip", showlegend=False,
        )
    )
    fig.add_trace(
        go.Bar(
            x=xs, y=values, width=0.52,
            marker=dict(color=colors, line=dict(color="rgba(255,255,255,.7)", width=1)),
            text=[f"{v:,}" for v in values],
            textposition="inside", insidetextanchor="middle", textangle=-90,
            textfont=dict(color="white", size=13, family=FONT),
            customdata=[[lb, p] for lb, p in zip(labels, pcts)],
            hovertemplate="%{customdata[0]}<br>%{y:,} Pallets (%{customdata[1]:.1f}%)<extra></extra>",
            showlegend=False,
        )
    )
    fig.add_trace(
        go.Bar(
            x=[x - 0.17 for x in xs], y=[v * 0.95 for v in values], width=0.07,
            marker=dict(color="rgba(255,255,255,0.38)", line=dict(width=0)),
            hoverinfo="skip", showlegend=False,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=xs, y=pcts, yaxis="y2", mode="lines+markers",
            line=dict(color="#929ba3", width=2, dash="dash"),
            marker=dict(size=12, symbol="square", color=colors, line=dict(color="white", width=2)),
            hovertemplate="%{y:.1f}%<extra></extra>", showlegend=False,
        )
    )
    lo, hi = min(pcts), max(pcts)
    span = max(hi - lo, 10)
    annotations = [
        dict(
            x=x, y=p, xref="x", yref="y2", text=f"<b>{p:.0f}%</b>",
            showarrow=False, yshift=28, bgcolor="white",
            bordercolor="#c7cdd2", borderwidth=1.5, borderpad=3,
            font=dict(size=11, color=INK, family=FONT),
        )
        for x, p in zip(xs, pcts)
    ]

    fig.update_layout(
        height=height, barmode="overlay",
        margin=dict(t=24, b=14, l=14, r=14),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, color=INK), showlegend=False,
        annotations=annotations,
        xaxis=dict(
            tickmode="array", tickvals=xs,
            ticktext=[lb.replace(" ", "<br>", 1) if len(lb) > 12 else lb for lb in labels],
            tickfont=dict(size=11), range=[-0.6, n - 0.4],
            showgrid=False, zeroline=False,
        ),
        yaxis=dict(domain=[0, 0.62], range=[0, bg_max], showgrid=False, showticklabels=False, zeroline=False),
        yaxis2=dict(
            domain=[0.74, 0.96], range=[lo - span * 0.6, hi + span * 0.9],
            showgrid=False, showticklabels=False, zeroline=False,
        ),
    )
    try:
        fig.update_traces(marker_cornerradius=18, selector=dict(type="bar"))
    except Exception:
        pass
    return fig


def pending_donut_chart(hist_not_billed, hist_not_po, hist_not_recv, height=250):
    """Flat, readable donut chart highlighting the total work not received."""
    values = [max(0, float(hist_not_billed)), max(0, float(hist_not_po))]
    raw_labels = ["ยังไม่ยื่นบิล", "Reject / ส่งไม่ครบ PO"]
    colors = ["#fff9e8", "#ffd43b"]
    active = [(label, value, color) for label, value, color in zip(raw_labels, values, colors) if value > 0]
    if active:
        raw_labels, values, colors = map(list, zip(*active))
    else:
        raw_labels, values, colors = ["ไม่มีงานค้าง"], [1], ["#e6eaf0"]
    display_labels = [f"{label} · {int(value):,}" for label, value in zip(raw_labels, values)]

    fig = go.Figure(go.Pie(
        labels=display_labels, values=values, customdata=raw_labels,
        hole=.64, sort=False, direction="clockwise",
        textinfo="none", marker=dict(colors=colors, line=dict(color="#c62828", width=2.5)),
        hovertemplate="%{customdata}<br>%{value:,.0f} Pallets (%{percent})<extra></extra>",
        showlegend=bool(active),
    ))
    fig.add_annotation(
        x=.5, y=.68, xref="paper", yref="paper", text="⚠ งานไม่ได้รับทั้งหมด",
        showarrow=False, font=dict(size=13, color="#c62828", family=FONT),
    )
    fig.add_annotation(
        x=.5, y=.51, xref="paper", yref="paper", text=f"<b>{hist_not_recv:,}</b>",
        showarrow=False, font=dict(size=32, color="#c92f25", family=FONT),
    )
    fig.add_annotation(
        x=.5, y=.39, xref="paper", yref="paper", text="Pallets",
        showarrow=False, font=dict(size=13, color="#402c28", family=FONT),
    )
    fig.update_layout(
        height=height, margin=dict(t=2, b=38, l=12, r=12),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, color="#172c4b"),
        legend=dict(
            orientation="h", yanchor="top", y=-.02, xanchor="center", x=.5,
            font=dict(size=12, color="#34241f"), bgcolor="rgba(255,250,246,.92)",
            bordercolor="#c6d7ec", borderwidth=1,
        ),
        showlegend=bool(active),
    )
    return fig


def _shade_hex(color, factor=0.68):
    color = color.lstrip("#")
    channels = [int(color[i:i + 2], 16) for i in (0, 2, 4)]
    return "#" + "".join(f"{max(0, min(255, round(channel * factor))):02x}" for channel in channels)


def metric_bar_3d_chart(labels, values, palette, reference=None, height=240):
    """Raised horizontal bars with a shaded extrusion and a reference track."""
    values = [max(0, float(value)) for value in values]
    colors = [palette[i % len(palette)] for i in range(len(values))]
    denominator = float(reference) if reference and reference > 0 else sum(values)
    denominator = denominator if denominator > 0 else 1
    track_max = max(max(values, default=0), float(reference or 0), 1) * 1.08
    track_colors = ["#edf0f2"] * len(values)
    percentages = [value / denominator * 100 for value in values]
    value_text = [f"{int(v):,} · {p:.0f}%" if v.is_integer() else f"{v:,.1f} · {p:.0f}%" for v, p in zip(values, percentages)]
    shadow_colors = [_shade_hex(color, .48) for color in colors]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=[track_max] * len(values), y=labels, orientation="h", width=.72,
        marker=dict(color=track_colors, line=dict(color="#fff", width=2)),
        hoverinfo="skip", showlegend=False,
    ))
    fig.add_trace(go.Bar(
        x=values, y=labels, orientation="h", width=.66, offset=.08,
        marker=dict(color=shadow_colors, line=dict(color="rgba(40,45,50,.16)", width=1)),
        hoverinfo="skip", showlegend=False,
    ))
    # Stack narrow adjacent sections inside each bar to create a dark-to-light
    # gradient that follows the dashboard card accent color.
    segment_count = 32
    for segment in range(segment_count):
        progress = segment / max(segment_count - 1, 1)
        shade_factor = .13 + progress * 1.35
        segment_colors = [_shade_hex(color, shade_factor) for color in colors]
        fig.add_trace(go.Bar(
            x=[value / segment_count for value in values],
            base=[value * segment / segment_count for value in values],
            y=labels, orientation="h", width=.5, offset=0,
            marker=dict(color=segment_colors, line=dict(width=0)),
            hoverinfo="skip", showlegend=False,
        ))
    fig.add_trace(go.Scatter(
        x=[value + track_max * .025 for value in values], y=labels,
        mode="text", text=value_text,
        textfont=dict(color="#17212b", size=17, family=FONT),
        customdata=[[label, value, pct] for label, value, pct in zip(labels, values, percentages)],
        hovertemplate="%{customdata[0]}<br>%{customdata[1]:,.0f} Pallets · %{customdata[2]:.1f}%<extra></extra>",
        cliponaxis=False, showlegend=False,
    ))
    fig.update_layout(
        height=height, barmode="overlay", bargap=.28,
        margin=dict(t=10, b=8, l=138, r=148),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, color=INK), showlegend=False,
        xaxis=dict(range=[0, track_max * 1.48], visible=False, fixedrange=True),
        yaxis=dict(autorange="reversed", fixedrange=True, tickfont=dict(size=15, color="#17212b"),
                   showgrid=False, zeroline=False, ticks=""),
    )
    return fig


def report_table(title, date_str, time_str, items, cols, compact=False, accent_colors=None):
    rendered_cells = []
    metric_index = 0
    colors = accent_colors or RAINBOW
    for item in items:
        if len(item) == 2 and item[0] == "group":
            rendered_cells.append(f'<div class="report-group">{item[1]}</div>')
            continue
        label, value, alert = item
        label_key = label.casefold()
        if "นอกรอบ" in label_key or "rp03" in label_key:
            status_class = "status-special"
        elif "complete" in label_key:
            status_class = "status-complete"
        elif "receiving started" in label_key or "กำลังลงสินค้า" in label_key:
            status_class = "status-progress"
        elif "intransit" in label_key or "check-in แล้ว" in label_key:
            status_class = "status-transit"
        elif (
            "supplier" in label_key or "reject" in label_key or "ไม่ได้รับ" in label_key
            or "ยังไม่ยื่นบิล" in label_key or "ส่งไม่ครบ" in label_key
        ):
            status_class = "status-alert"
        elif "เลื่อนส่ง" in label_key:
            status_class = "status-progress"
        else:
            status_class = "status-neutral"
        classes = f"report-cell tone{metric_index % len(colors)} {status_class}".strip()
        rendered_cells.append(
            f'<div class="{classes}" style="border-top:4px solid {colors[metric_index % len(colors)]};">'
            f'<div class="report-label">{label.replace(chr(10), "<br>")}</div>'
            f'<div class="report-value{" alert" if alert else ""}">{value}</div></div>'
        )
        metric_index += 1
    cells = "".join(rendered_cells)
    return (
        f'<div class="report-wrap{" compact" if compact else ""}"><div class="report-head">{title} {date_str}'
        f' <span class="report-date">{time_str}</span></div>'
        f'<div class="report-grid{" compact-grid" if compact else ""}" style="grid-template-columns: repeat({cols}, minmax(0, 1fr));">'
        f"{cells}</div></div>"
    )


def chart_card_title(title, sub=""):
    sub_html = f'<div class="chart-sub">{sub}</div>' if sub else ""
    st.markdown(
        f'<div class="chart-title">{title}</div>{sub_html}',
        unsafe_allow_html=True,
    )


# --- Sidebar เมนูเลือกชีตรายงาน ---
NAV = [
    "สรุปงานรับและงานค้างส่ง(ย้อนหลัง)",
    "สถานการณ์การรับ(กะเช้าปัจจุบัน)",
]
NAV_ICON = {NAV[0]: "🌙", NAV[1]: "☀️"}

st.sidebar.markdown(
    '<div class="sb-brand"><div class="sb-logo"><svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2.8 20 7v10l-8 4.2L4 17V7z"/><path d="M4 7l8 4.2L20 7"/><path d="M12 11.2v10"/><path d="M8 4.9 16 9.1" stroke-opacity=".6"/></svg></div><div>'
    '<div class="sb-name">DCSB Distribution Center</div>'
    '<div class="sb-sub">REPORT SYSTEM</div></div></div>',
    unsafe_allow_html=True,
)
st.sidebar.markdown('<div class="sb-label">เมนูรายงาน</div>', unsafe_allow_html=True)
sheet_selection = st.sidebar.radio(
    "เลือกชีต:",
    NAV,
    format_func=lambda s: f"{NAV_ICON[s]}  {s}",
    label_visibility="collapsed",
)
if sheet_selection == NAV[1]:
    RAINBOW = ["#5d9568", "#c4832d", "#5d8fba", "#bc5a5c", "#896eb0", "#9ba2a7", "#75a982", "#a9666d"]
    CHART_TRACK = "rgba(236,239,240,0.9)"
    CHART_INNER = "rgba(255,255,255,0.95)"
    CHART_GAP = "#ffffff"
    CHART_LEGEND_BORDER = "#eceff0"
    BAR_TEXT = "#1f2527"
    BUBBLE_TEXT = "#1f2527"
    INK, MUTED = "#1f2527", "#8a9296"
    st.markdown(
        """
<style>
body:has(.morning-theme) {
    --ink:#1f2527; --muted:#8a9296; --blue:#5b8fc0; --blue2:#8ab4d8;
    --rose:#e0707f; --grad:linear-gradient(120deg,#cdeed7,#f7e6bd);
    --line:#eceff0; --line-strong:#e8ecee;
    --sh:0 1px 2px rgba(31,37,39,.03), 0 10px 28px rgba(31,37,39,.05);
    --sh-sm:0 1px 2px rgba(31,37,39,.03), 0 4px 14px rgba(31,37,39,.04);
    --sh-in:none;
}
body:has(.morning-theme) [data-testid="stAppViewContainer"] { background:#fcfcfc; }

/* hero: left-aligned greeting, no card */
body:has(.morning-theme) .hero { background:transparent; border:0; box-shadow:none; text-align:left; max-width:none; padding:6px 4px 2px; margin:0 0 4px; border-radius:0; }
body:has(.morning-theme) .hero-kicker { color:var(--muted); font-size:.62rem; font-weight:500; letter-spacing:.14em; }
body:has(.morning-theme) .hero-title { color:var(--ink); font-size:1.85rem; font-weight:700; letter-spacing:-.025em; margin:2px 0 0; }
body:has(.morning-theme) .hero-subtitle { color:var(--muted); font-size:.82rem; }
body:has(.morning-theme) .hero-bar { display:none; }
body:has(.morning-theme) .section-title { border-bottom:0; font-size:.95rem; margin:12px 0 4px; padding-bottom:0; }
body:has(.morning-theme) .section-title::before { width:8px; height:8px; border-radius:50%; background:#83c89d; box-shadow:none; }

/* focus cards: pastel gradient */
body:has(.morning-theme) .focus-card { background:linear-gradient(145deg,#fff 0%,#e4f5e9 48%,#d2ebda 100%); border:1px solid #c1dfca; border-radius:22px; box-shadow:0 2px 3px rgba(36,65,45,.08),0 12px 24px rgba(73,120,88,.17),0 22px 38px rgba(73,120,88,.08),inset 0 2px 0 rgba(255,255,255,.96),inset 0 -2px 4px rgba(83,136,98,.08); transition:transform .18s ease,box-shadow .18s ease; padding:12px 20px 11px 22px; }
body:has(.morning-theme) .focus-card:hover { transform:translateY(-3px); box-shadow:0 3px 5px rgba(36,65,45,.1),0 16px 28px rgba(73,120,88,.21),0 28px 42px rgba(73,120,88,.1),inset 0 2px 0 rgba(255,255,255,.98); }
body:has(.morning-theme) .focus-card::before { display:none; }
body:has(.morning-theme) .focus-card::after { content:""; position:absolute; right:-34px; top:-34px; width:140px; height:140px; border-radius:50%; background:radial-gradient(circle,rgba(255,255,255,.95),rgba(255,255,255,0) 70%); pointer-events:none; }
body:has(.morning-theme) .focus-card .focus-value { color:var(--ink); }
body:has(.morning-theme) .focus-card.alert { background:linear-gradient(145deg,#fff 0%,#fde8e9 48%,#f8d9dc 100%); border-color:#edc3cb; box-shadow:0 2px 3px rgba(83,42,48,.08),0 12px 24px rgba(166,88,98,.17),0 22px 38px rgba(166,88,98,.08),inset 0 2px 0 rgba(255,255,255,.96),inset 0 -2px 4px rgba(166,88,98,.07); }
body:has(.morning-theme) .focus-card.alert .focus-value { color:var(--rose); }

/* report table */
body:has(.morning-theme) .report-wrap { background:linear-gradient(150deg,#fff,#f8fbf9); border:1px solid #d8e3dc; border-radius:24px; box-shadow:0 2px 4px rgba(37,57,43,.06),0 12px 28px rgba(50,70,58,.14),0 24px 44px rgba(50,70,58,.06),inset 0 2px 0 #fff,inset 0 -2px 5px rgba(86,117,95,.04); }
body:has(.morning-theme) .report-head { border-bottom:0; color:var(--ink); }
body:has(.morning-theme) .report-group { border-bottom:1px solid #f0f2f3; color:var(--muted); font-size:.74rem; font-weight:600; letter-spacing:.04em; }
body:has(.morning-theme) .report-cell { border:1px solid rgba(90,112,99,.18) !important; box-shadow:0 2px 3px rgba(43,60,49,.07),0 6px 12px rgba(43,60,49,.1),inset 0 2px 0 rgba(255,255,255,.94),inset 0 -2px 3px rgba(70,94,76,.04); border-radius:16px; background:#fff; }
body:has(.morning-theme) .report-label { color:#6b7478; }
body:has(.morning-theme) .report-value { color:var(--ink); }
body:has(.morning-theme) .report-value.alert { color:var(--rose); }
body:has(.morning-theme) .report-cell.tone0 { background:linear-gradient(160deg,#dcf2e3,#f6fbf8); }
body:has(.morning-theme) .report-cell.tone1 { background:linear-gradient(160deg,#e1edf9,#f7fbff); }
body:has(.morning-theme) .report-cell.tone2 { background:linear-gradient(160deg,#ebe5f7,#faf8fe); }
body:has(.morning-theme) .report-cell.tone3 { background:linear-gradient(160deg,#fde8cd,#fff8ee); }
body:has(.morning-theme) .report-cell.tone4 { background:linear-gradient(160deg,#fbe2e8,#fff7f9); }
body:has(.morning-theme) .report-cell.tone5 { background:linear-gradient(160deg,#faf0c8,#fffcec); }
body:has(.morning-theme) .report-cell.tone6 { background:linear-gradient(160deg,#fadfd5,#fff6f2); }
body:has(.morning-theme) .report-cell.tone7 { background:linear-gradient(160deg,#d9efec,#f5fbfa); }

/* chart cards, expander, inputs, buttons */
body:has(.morning-theme) [class*="st-key-card"] { background:linear-gradient(150deg,#fff,#f7fafb); border:1px solid #d9e2e8; border-radius:22px; box-shadow:0 2px 4px rgba(38,56,68,.07),0 12px 26px rgba(45,63,78,.15),0 22px 38px rgba(45,63,78,.06),inset 0 2px 0 #fff,inset 0 -2px 5px rgba(60,87,105,.04); transition:transform .18s ease,box-shadow .18s ease; }
body:has(.morning-theme) [class*="st-key-card"]:hover { transform:translateY(-3px); box-shadow:0 4px 6px rgba(38,56,68,.08),0 16px 30px rgba(45,63,78,.2),0 28px 44px rgba(45,63,78,.09),inset 0 2px 0 #fff; }
body:has(.morning-theme) .chart-sub { border-bottom:1px dashed #e6eaeb; }
body:has(.morning-theme) [data-testid="stExpander"] { background:#fff; border:1px solid #eceff0 !important; border-radius:20px; box-shadow:var(--sh-sm); }
body:has(.morning-theme) [data-testid="stNumberInput"] [data-baseweb="input"],
body:has(.morning-theme) [data-testid="stNumberInput"] [data-baseweb="base-input"] { background:#fff !important; border:1px solid #e6eaeb !important; box-shadow:none; }
body:has(.morning-theme) [data-testid="stNumberInput"] button { color:var(--muted); }
body:has(.morning-theme) .stButton > button { background:#1f2527; box-shadow:none; border-radius:12px; }

/* sidebar: white, light-grey active pill */
body:has(.morning-theme) [data-testid="stSidebar"] { background:#fff; border-right:1px solid #eef0f1; box-shadow:none; }
body:has(.morning-theme) .sb-brand { background:transparent; border-bottom:1px solid #eef0f1; border-radius:0; padding:5px 6px 14px; margin-bottom:8px; }
body:has(.morning-theme) .sb-name { color:var(--ink); }
body:has(.morning-theme) .sb-sub, body:has(.morning-theme) .sb-label { color:var(--muted); }
body:has(.morning-theme) .sb-logo { background:linear-gradient(145deg,#21110d 0%,#791b12 26%,#e43a19 52%,#ff851f 76%,#ffd653 100%); border:1px solid #ffc15b; border-radius:14px; box-shadow:0 0 0 2px rgba(255,105,35,.18),0 0 12px rgba(239,65,25,.58),0 0 24px rgba(255,117,28,.38),0 0 32px rgba(255,193,66,.24),inset 0 2px 0 rgba(255,245,211,.6),inset 0 -3px 7px rgba(91,24,8,.38); }
body:has(.morning-theme) .sb-logo svg { stroke:#fff9e7; filter:drop-shadow(0 0 3px rgba(255,255,237,.98)) drop-shadow(0 0 7px rgba(255,158,48,.9)) drop-shadow(0 0 10px rgba(255,71,25,.62)); }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"],
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"] { background:transparent; border:1px solid transparent; box-shadow:none; border-radius:12px; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"]:hover,
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"]:hover { background:#f7f8f8; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label p { color:#62696d; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"]:has(input:checked),
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) { background:#f1f3f3; border-color:#f1f3f3 !important; box-shadow:none; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) p { color:var(--ink); }
body:has(.morning-theme) [data-testid="stSidebar"] [data-baseweb="input"],
body:has(.morning-theme) [data-testid="stSidebar"] [data-baseweb="base-input"],
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stDateInputField"] { background:#fff !important; border:1px solid #e6eaeb !important; box-shadow:none; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { background:#fafbfb; border:1px dashed #dde3e4; box-shadow:none; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button { background:#1f2527; box-shadow:none; }
body:has(.morning-theme) .sb-foot { border-top:1px solid #eef0f1; }
body:has(.morning-theme) .sb-chip { background:#fff; border:1px solid #e6eaeb; box-shadow:none; color:var(--muted); }
body:has(.morning-theme) .sb-chip.on { background:#e4f4e9; border-color:#d3ead9; color:#3f7a56; box-shadow:none; }

/* consistent soft 3D depth across the entire morning report page */
body:has(.morning-theme) .section-title { text-shadow:0 1px 0 #fff; }
body:has(.morning-theme) .section-title::before { box-shadow:0 2px 4px rgba(83,136,98,.22),inset 0 1px 0 rgba(255,255,255,.7); }
body:has(.morning-theme) .hero-title { text-shadow:0 1px 0 #fff,0 2px 5px rgba(31,37,39,.08); }
body:has(.morning-theme) .report-cell { transition:transform .16s ease,box-shadow .16s ease; }
body:has(.morning-theme) .report-cell:hover { transform:translateY(-2px); box-shadow:0 3px 5px rgba(43,60,49,.08),0 10px 17px rgba(43,60,49,.13),inset 0 2px 0 rgba(255,255,255,.98); }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stExpander"] { box-shadow:0 2px 4px rgba(37,57,43,.06),0 8px 18px rgba(50,70,58,.11),inset 0 1px 0 #fff; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stDateInputField"],
body:has(.morning-theme) [data-testid="stSidebar"] [data-baseweb="input"] { box-shadow:0 2px 4px rgba(37,57,43,.06),0 5px 10px rgba(50,70,58,.08),inset 0 1px 0 #fff; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { box-shadow:0 2px 4px rgba(37,57,43,.05),0 7px 14px rgba(50,70,58,.09),inset 0 1px 0 #fff; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button,
body:has(.morning-theme) .stButton > button { box-shadow:0 3px 0 #b8c0c2,0 6px 12px rgba(31,37,39,.16),inset 0 1px 0 rgba(255,255,255,.45); transition:transform .14s ease,box-shadow .14s ease; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover,
body:has(.morning-theme) .stButton > button:hover { transform:translateY(-2px); box-shadow:0 4px 0 #b8c0c2,0 9px 16px rgba(31,37,39,.2),inset 0 1px 0 rgba(255,255,255,.5); }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) { box-shadow:0 2px 4px rgba(37,57,43,.06),0 5px 10px rgba(50,70,58,.08),inset 0 1px 0 #fff !important; }
body:has(.morning-theme) .sb-chip { box-shadow:0 2px 4px rgba(37,57,43,.06),inset 0 1px 0 #fff; }
body:has(.morning-theme) .sb-chip.on { box-shadow:0 2px 5px rgba(73,120,88,.12),inset 0 1px 0 rgba(255,255,255,.9); }
body:has(.morning-theme) [data-testid="stNumberInput"] [data-baseweb="input"],
body:has(.morning-theme) [data-testid="stNumberInput"] [data-baseweb="base-input"] { box-shadow:0 2px 4px rgba(37,57,43,.05),0 5px 10px rgba(50,70,58,.08),inset 0 1px 0 #fff; }

/* warm magma palette matched to the morning page logo */
body:has(.morning-theme) {
    --ink:#302019; --muted:#8b7568; --blue:#c13a27; --blue2:#e86d25;
    --rose:#b92325; --grad:linear-gradient(120deg,#a91e24 0%,#e94f24 52%,#ffc642 100%);
    --line:#f0d9c8; --line-strong:#d99a65;
    --sh:0 2px 4px rgba(115,53,26,.08),0 12px 26px rgba(155,72,31,.15),0 24px 42px rgba(155,72,31,.07);
    --sh-sm:0 2px 4px rgba(115,53,26,.07),0 6px 13px rgba(155,72,31,.11);
    --sh-in:inset 0 2px 4px rgba(115,53,26,.08),inset 0 1px 0 rgba(255,255,255,.9);
}
body:has(.morning-theme) [data-testid="stAppViewContainer"] { background:linear-gradient(145deg,#fffdf9,#fff6ed 58%,#fffaf2); }
body:has(.morning-theme) .hero-kicker { color:#c9572d; }
body:has(.morning-theme) .hero-title { color:#302019; }
body:has(.morning-theme) .section-title { color:#51271e; }
body:has(.morning-theme) .section-title::before { background:linear-gradient(160deg,#b51e24,#ee5926 60%,#ffc642); box-shadow:0 2px 6px rgba(208,74,35,.28),inset 0 1px 0 rgba(255,255,255,.7); }
body:has(.morning-theme) .focus-card { background:linear-gradient(145deg,#fffdf9 0%,#fff0dc 52%,#ffe3bf 100%); border-color:#efc394; box-shadow:0 2px 3px rgba(100,43,22,.08),0 12px 24px rgba(177,84,35,.17),0 22px 38px rgba(177,84,35,.08),inset 0 2px 0 rgba(255,255,255,.97),inset 0 -2px 4px rgba(177,84,35,.08); }
body:has(.morning-theme) .focus-card:hover { box-shadow:0 3px 5px rgba(100,43,22,.1),0 16px 28px rgba(177,84,35,.22),0 28px 42px rgba(177,84,35,.11),inset 0 2px 0 rgba(255,255,255,.98); }
body:has(.morning-theme) .focus-card.alert { background:linear-gradient(145deg,#fff9f3 0%,#ffe3d7 48%,#ffc9b5 100%); border-color:#e9a185; box-shadow:0 2px 3px rgba(108,34,22,.08),0 12px 24px rgba(180,49,31,.18),0 22px 38px rgba(180,49,31,.08),inset 0 2px 0 rgba(255,255,255,.96),inset 0 -2px 4px rgba(180,49,31,.08); }
body:has(.morning-theme) .focus-card.alert .focus-value,
body:has(.morning-theme) .report-value.alert { color:#b92325; }
body:has(.morning-theme) .report-wrap { background:linear-gradient(150deg,#fffefa,#fff8ef); border-color:#e9c9ad; box-shadow:0 2px 4px rgba(91,52,27,.07),0 12px 28px rgba(133,75,36,.15),0 24px 44px rgba(133,75,36,.07),inset 0 2px 0 #fff,inset 0 -2px 5px rgba(133,75,36,.05); }
body:has(.morning-theme) .report-group { border-color:#efd8c3; color:#87523a; }
body:has(.morning-theme) .report-cell { border-color:rgba(164,91,48,.2) !important; box-shadow:0 2px 3px rgba(91,52,27,.07),0 6px 12px rgba(133,75,36,.1),inset 0 2px 0 rgba(255,255,255,.95),inset 0 -2px 3px rgba(133,75,36,.05); }
body:has(.morning-theme) .report-cell.tone0 { background:linear-gradient(160deg,#fff2dd,#fffaf1); }
body:has(.morning-theme) .report-cell.tone1 { background:linear-gradient(160deg,#ffe7d5,#fff8f0); }
body:has(.morning-theme) .report-cell.tone2 { background:linear-gradient(160deg,#ffe1d8,#fff7f2); }
body:has(.morning-theme) .report-cell.tone3 { background:linear-gradient(160deg,#ffebe0,#fff9f4); }
body:has(.morning-theme) .report-cell.tone4 { background:linear-gradient(160deg,#fff0d2,#fffaf0); }
body:has(.morning-theme) .report-cell.tone5 { background:linear-gradient(160deg,#fff4cb,#fffdf3); }
body:has(.morning-theme) .report-cell.tone6 { background:linear-gradient(160deg,#ffe5c8,#fff8ea); }
body:has(.morning-theme) .report-cell.tone7 { background:linear-gradient(160deg,#f8e5d4,#fff8f1); }
body:has(.morning-theme) [class*="st-key-card"] { background:linear-gradient(150deg,#fffefa,#fff7ed); border-color:#e8cbb2; box-shadow:0 2px 4px rgba(91,52,27,.07),0 12px 26px rgba(133,75,36,.16),0 22px 38px rgba(133,75,36,.07),inset 0 2px 0 #fff,inset 0 -2px 5px rgba(133,75,36,.05); }
body:has(.morning-theme) [class*="st-key-card"]:hover { box-shadow:0 4px 6px rgba(91,52,27,.09),0 16px 30px rgba(133,75,36,.22),0 28px 44px rgba(133,75,36,.1),inset 0 2px 0 #fff; }
body:has(.morning-theme) .chart-sub { border-color:#edd5bf; }
body:has(.morning-theme) [data-testid="stSidebar"] { background:linear-gradient(180deg,#fffaf4,#fff4e8); border-color:#efd6c0; }
body:has(.morning-theme) .sb-brand { border-color:#eed2b7; }
body:has(.morning-theme) .sb-name { color:#38231a; }
body:has(.morning-theme) .sb-sub, body:has(.morning-theme) .sb-label { color:#967463; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-testid="stRadioOption"]:has(input:checked),
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label[data-baseweb="radio"]:has(input:checked) { background:linear-gradient(120deg,#a91e24,#e94f24 58%,#f6a22c); border-color:#d6552d !important; box-shadow:0 5px 12px rgba(177,66,35,.22),inset 0 1px 0 rgba(255,255,255,.55); }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) p { color:#fffaf3; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stRadio"] label:hover { background:#fff1e4; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-baseweb="input"],
body:has(.morning-theme) [data-testid="stSidebar"] [data-baseweb="base-input"],
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stDateInputField"],
body:has(.morning-theme) [data-testid="stNumberInput"] [data-baseweb="input"],
body:has(.morning-theme) [data-testid="stNumberInput"] [data-baseweb="base-input"] { background:#fffdf9 !important; border-color:#e8c9ad !important; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] { background:linear-gradient(145deg,#fffaf3,#fff0df); border-color:#e6b98c; }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button,
body:has(.morning-theme) .stButton > button { background:linear-gradient(120deg,#a91e24 0%,#e94f24 55%,#f6a22c 100%); color:#fffaf1; box-shadow:0 3px 0 #a64228,0 7px 14px rgba(177,66,35,.24),inset 0 1px 0 rgba(255,255,255,.48); }
body:has(.morning-theme) [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover,
body:has(.morning-theme) .stButton > button:hover { background:linear-gradient(120deg,#b91f25 0%,#f05a24 55%,#ffb331 100%); box-shadow:0 4px 0 #a64228,0 10px 18px rgba(177,66,35,.29),inset 0 1px 0 rgba(255,255,255,.52); }
body:has(.morning-theme) .sb-foot { border-color:#eed2b7; }
body:has(.morning-theme) .sb-chip { background:#fffaf4; border-color:#ecd4be; color:#8a6350; }
body:has(.morning-theme) .sb-chip.on { background:linear-gradient(120deg,#a91e24,#e94f24 58%,#f6a22c); border-color:#d6552d; color:#fffaf3; }
body:has(.morning-theme) .report-cell.status-complete { background:linear-gradient(160deg,#fff,#f8fbf8) !important; border-color:#e0e8e1 !important; border-top-color:#5d9568 !important; }
body:has(.morning-theme) .report-cell.status-complete .report-value { color:#477d52; }
body:has(.morning-theme) .report-cell.status-progress { background:linear-gradient(160deg,#fff,#fcf9f3) !important; border-color:#ebe4d7 !important; border-top-color:#c4832d !important; }
body:has(.morning-theme) .report-cell.status-progress .report-value { color:#9d6925; }
body:has(.morning-theme) .report-cell.status-transit { background:linear-gradient(160deg,#fff,#f7f9fc) !important; border-color:#dde5ed !important; border-top-color:#5d8fba !important; }
body:has(.morning-theme) .report-cell.status-transit .report-value { color:#4a779c; }
body:has(.morning-theme) .report-cell.status-alert { background:linear-gradient(160deg,#fff,#fcf7f7) !important; border-color:#ecdede !important; border-top-color:#bc5a5c !important; }
body:has(.morning-theme) .report-cell.status-alert .report-value { color:#a64d4f; }
body:has(.morning-theme) .report-cell.status-special { background:linear-gradient(160deg,#fff,#f9f7fc) !important; border-color:#e5dfed !important; border-top-color:#896eb0 !important; }
body:has(.morning-theme) .report-cell.status-special .report-value { color:#765d99; }
body:has(.morning-theme) .report-cell.status-neutral { background:linear-gradient(160deg,#fff,#f8f9fa) !important; border-color:#e1e4e6 !important; border-top-color:#9ba2a7 !important; }
body:has(.morning-theme) .report-cell.status-neutral .report-value { color:#626a70; }
body:has(.morning-theme) .report-group,
body:has(.morning-theme) .report-cell .report-label,
body:has(.morning-theme) .report-cell .report-value { color:#171717 !important; }
body:has(.morning-theme) .focus-card.status-complete { background:linear-gradient(125deg,#101713 0%,#142d20 48%,#176743 100%) !important; border-color:#63f5a6 !important; box-shadow:0 0 0 1px rgba(99,245,166,.3),0 0 12px rgba(70,255,160,.24),0 12px 26px rgba(10,57,34,.3),0 24px 42px rgba(10,57,34,.18),inset 0 2px 0 rgba(218,255,235,.24),inset 0 -3px 6px rgba(0,0,0,.32) !important; }
body:has(.morning-theme) .focus-card.status-complete::after { background:radial-gradient(circle,rgba(116,255,181,.5),rgba(116,255,181,0) 72%); }
body:has(.morning-theme) .focus-card.status-complete .focus-label,
body:has(.morning-theme) .focus-card.status-complete .focus-unit { color:#d9ffea !important; }
body:has(.morning-theme) .focus-card.status-complete .focus-value { color:#7dffb5 !important; text-shadow:0 0 8px rgba(94,255,170,.5); }
body:has(.morning-theme) .focus-card.alert { background:linear-gradient(115deg,#21110d 0%,#791b12 34%,#d52f20 67%,#f47725 100%) !important; border-color:#f18a52 !important; box-shadow:0 2px 4px rgba(75,24,14,.16),0 12px 24px rgba(163,48,22,.25),0 22px 38px rgba(163,48,22,.12),inset 0 2px 0 rgba(255,230,199,.55),inset 0 -3px 6px rgba(44,13,9,.3) !important; }
body:has(.morning-theme) .focus-card.alert .focus-label,
body:has(.morning-theme) .focus-card.alert .focus-unit { color:#fff1e5 !important; }
body:has(.morning-theme) .focus-card.alert .focus-value { color:#ffe08a !important; text-shadow:0 1px 2px rgba(58,13,7,.5); }

/* compact sizing for the morning page */
.focus-grid { gap:10px; margin:3px 0; }
.focus-label { font-size:.76rem; }
.focus-value { font-size:1.7rem; margin:2px 0; }
.focus-unit { font-size:.68rem; }
.report-wrap.compact { padding:3px 5px 5px; }
.report-wrap.compact .report-head { padding:3px 7px 4px; font-size:.9rem; }
.report-wrap.compact .report-grid { gap:4px; }
.report-wrap.compact .report-group { font-size:.74rem; padding:1px 2px; }
.report-wrap.compact .report-cell { min-height:50px; padding:3px 4px; gap:2px; }
.report-wrap.compact .report-label { font-size:.72rem; line-height:1.14; }
.report-wrap.compact .report-value { font-size:1.02rem; }
[class*="st-key-card"] { padding:.4rem .5rem .25rem; }
.chart-title { font-size:.82rem; margin-bottom:2px; }
.chart-sub { font-size:.64rem; margin-bottom:3px; padding-bottom:3px; }
</style>
""",
        unsafe_allow_html=True,
    )
    st.markdown('<div class="morning-theme" aria-hidden="true"></div>', unsafe_allow_html=True)

# Keep both page-specific upload sections mounted so Streamlit preserves each
# page's selected files while the user switches between reports.
_page_id = "morning" if sheet_selection == NAV[1] else "history"


def _render_upload_section(page_id, title, expanded):
    with st.sidebar.expander(title, expanded=expanded):
        uploads = {
            "rp01_1610": st.file_uploader(
                "RP01 งานรับทั้งหมด (16.10)", type=["xlsx", "xls"], key=f"{page_id}_rp01_1610"
            ),
        }
        if page_id == "history":
            uploads["rp01_2200"] = st.file_uploader(
                "RP01 หลัง 22.00", type=["xlsx", "xls"], key=f"{page_id}_rp01_2200"
            )
        else:
            uploads["rp01_2200"] = None
        uploads.update({
            "rp03": st.file_uploader(
                "RP03 GRN", type=["xlsx", "xls"], key=f"{page_id}_rp03"
            ),
            "rp42": st.file_uploader(
                "RP42 ASN Status", type=["xlsx", "xls"], key=f"{page_id}_rp42"
            ),
        })
        return uploads


# Reserve the first sidebar slot for the active report date. The date is
# populated after reading the uploads, but remains visually above the importers.
report_date_slot = st.sidebar.empty()
st.sidebar.markdown('<div class="sb-label">Import ไฟล์แยกตามหน้า</div>', unsafe_allow_html=True)
_page_uploads = {
    "history": _render_upload_section("history", "🕘 รายงานย้อนหลัง", _page_id == "history"),
    "morning": _render_upload_section("morning", "🌅 กะเช้าปัจจุบัน", _page_id == "morning"),
}
_active_uploads = _page_uploads[_page_id]
uploaded_rp01_1610 = _active_uploads["rp01_1610"]
uploaded_rp01_2200 = _active_uploads["rp01_2200"]
uploaded_rp03 = _active_uploads["rp03"]
uploaded_rp42 = _active_uploads["rp42"]
_uploaded_files = {
    "rp01_1610": uploaded_rp01_1610,
    "rp01_2200": uploaded_rp01_2200,
    "rp03": uploaded_rp03,
    "rp42": uploaded_rp42,
}
if sheet_selection == NAV[1]:
    imported, import_error = _load_morning_metrics(_uploaded_files)
else:
    imported, import_error = _load_imported_metrics(_uploaded_files)
report_date_warnings = _report_date_warnings(_uploaded_files, imported.get("report_date"))
_uploads = (uploaded_rp01_1610, uploaded_rp01_2200, uploaded_rp03, uploaded_rp42)
_import_signature = hashlib.sha256(
    "|".join(
        f"{f.name}:{hashlib.sha256(f.getvalue()).hexdigest()}" if f is not None else "-"
        for f in _uploads
    ).encode("utf-8")
).hexdigest()
_import_signature_key = f"_last_import_signature_{_page_id}"
_new_import = _import_signature != st.session_state.get(_import_signature_key)
if sheet_selection == "สถานการณ์การรับ(กะเช้าปัจจุบัน)":
    rp01_date = imported.get("report_date")
    if rp01_date:
        st.session_state["m_date_input"] = rp01_date
    elif "m_date_input" not in st.session_state:
        latest = REPORT_STORE.get("morning_latest_date")
        st.session_state["m_date_input"] = date.fromisoformat(latest) if latest else date.today()
    with report_date_slot.container():
        st.markdown('<div class="sb-label">วันที่รายงาน</div>', unsafe_allow_html=True)
        morning_date = st.date_input(
            "ประจำวันที่ (กะเช้าปัจจุบัน)", value=st.session_state["m_date_input"], key="m_date_input"
        )
    selected_morning_date = morning_date.strftime("%Y-%m-%d")
    record = _get_report_record(selected_morning_date, page="morning")
    for field, default_value in MORNING_DEFAULTS.items():
        value = record.get(field, default_value) if record else default_value
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        st.session_state[field] = value
    st.session_state["_loaded_morning_date"] = selected_morning_date
    if _new_import and any(_uploads):
        for metric, value in imported.items():
            if metric in MORNING_DEFAULTS:
                st.session_state[metric] = value
        values = {field: st.session_state.get(field, default) for field, default in MORNING_DEFAULTS.items()}
        values["total_remaining"] = sum(values[field] for field in ("recv_start", "in_transit", "sup_night_missing", "sup_morning_missing"))
        _save_report_record(selected_morning_date, values, page="morning")
        REPORT_STORE = _load_report_store()
        st.session_state[_import_signature_key] = _import_signature
        st.sidebar.success(f"จำข้อมูลหน้ากะเช้าของวันที่ {selected_morning_date} แล้ว")
    else:
        st.sidebar.caption("วันที่อิงจาก RP01 · ตัวเลขและข้อมูลหน้านี้จำอัตโนมัติ")
    if st.session_state.get("_morning_autosave_error"):
        st.sidebar.error(f"บันทึกอัตโนมัติไม่สำเร็จ: {st.session_state['_morning_autosave_error']}")
else:
    rp01_date = imported.get("report_date")
    if rp01_date:
        st.session_state["h_date_input"] = rp01_date
    elif "h_date_input" not in st.session_state:
        latest = REPORT_STORE.get("latest_date")
        st.session_state["h_date_input"] = date.fromisoformat(latest) if latest else date.today()
    with report_date_slot.container():
        st.markdown('<div class="sb-label">วันที่รายงาน</div>', unsafe_allow_html=True)
        hist_date = st.date_input(
            "ประจำวันที่ (ย้อนหลัง)", value=st.session_state["h_date_input"], key="h_date_input"
        )
    selected_report_date = hist_date.strftime("%Y-%m-%d")
    record = _get_report_record(selected_report_date, page="history")
    for field, default_value in REPORT_DEFAULTS.items():
        value = record.get(field, default_value) if record else default_value
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        st.session_state[field] = value
    if st.session_state.get("_loaded_history_detail_date") != selected_report_date:
        st.session_state["hist_delayed_po_input"] = record.get("hist_delayed_po_input", "") if record else ""
        st.session_state["hist_reject_item_input"] = record.get("hist_reject_item_input", "") if record else ""
        st.session_state["hist_po_rp01_source"] = record.get("hist_po_rp01_source", "RP01 16.10") if record else "RP01 16.10"
        st.session_state["_loaded_history_detail_date"] = selected_report_date
    st.session_state["_loaded_report_date"] = selected_report_date
    if _new_import:
        for metric, value in imported.items():
            if metric.startswith("hist_"):
                st.session_state[metric] = value
        st.session_state[_import_signature_key] = _import_signature
    if _new_import and any(key.startswith("hist_") for key in imported):
        _autosave_current_report()
        REPORT_STORE = _load_report_store()
        st.sidebar.success(f"นำเข้าและจำข้อมูลของวันที่ {selected_report_date} แล้ว")
    elif record:
        st.sidebar.caption("โหลดข้อมูลล่าสุดของวันที่เลือกแล้ว · บันทึกอัตโนมัติ")
    else:
        st.sidebar.caption("ข้อมูลที่แก้ไขจะถูกจำไว้โดยอัตโนมัติ")
    if st.session_state.get("_autosave_error"):
        st.sidebar.error(f"บันทึกอัตโนมัติไม่สำเร็จ: {st.session_state['_autosave_error']}")
if import_error:
    st.sidebar.error(f"อ่านไฟล์ไม่สำเร็จ: {import_error}")
if imported.get("extra_join"):
    st.sidebar.info(imported["extra_join"])
_chips = "".join(
    f'<span class="sb-chip{" on" if f is not None else ""}">{n}</span>'
    for n, f in (("RP01 16.10", uploaded_rp01_1610), ("RP01 22.00", uploaded_rp01_2200), ("RP03", uploaded_rp03), ("RP42", uploaded_rp42))
)
st.sidebar.markdown(f'<div class="sb-foot">{_chips}</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="hero"><div class="hero-kicker">DCSB DISTRIBUTION CENTER · REPORT SYSTEM</div>'
    f'<div class="hero-title">{NAV_ICON[sheet_selection]} {sheet_selection}</div>'
    '<div class="hero-subtitle">ระบบสรุปยอดงานรับสินค้าและงานค้างรับคงเหลือ</div>'
    '<div class="hero-bar"></div></div>',
    unsafe_allow_html=True,
)
for warning in report_date_warnings:
    st.warning(warning)

# =========================================================================
# ชีตที่ 1: สถานการณ์การรับ (กะเช้าปัจจุบัน)
# =========================================================================
if sheet_selection == "สถานการณ์การรับ(กะเช้าปัจจุบัน)":
    inv_all = st.session_state.get("inv_all", 6058)
    night_received = st.session_state.get("night_received", 4240)
    morning_received = st.session_state.get("morning_received", 1818)
    complete_night = st.session_state.get("complete_night", 4475)
    complete_morning = st.session_state.get("complete_morning", 0)
    extra_work = st.session_state.get("extra_work", 0)
    complete_all = complete_night + complete_morning + extra_work
    pct_night_received = (complete_night / inv_all * 100) if inv_all > 0 else 0
    recv_start = st.session_state.get("recv_start", 473)
    in_transit = st.session_state.get("in_transit", 268)
    sup_night_missing = st.session_state.get("sup_night_missing", 239)
    sup_morning_missing = st.session_state.get("sup_morning_missing", 0)
    total_remaining = recv_start + in_transit + sup_night_missing + sup_morning_missing
    st.session_state["total_remaining"] = total_remaining
    pct_remaining = ((sup_night_missing + sup_morning_missing) / inv_all * 100) if inv_all > 0 else 0

    st.markdown(
        f"""
    <div class="focus-grid">
        <div class="focus-card status-complete">
        <div class="focus-label">📦 ยอดงาน Complete รวมทั้งหมด</div>
            <div class="focus-value">{complete_all:,} <span>Pallets</span></div>
            <div class="focus-unit">รวมกะดึก + กะเช้า + งานนอกรอบ</div>
        </div>
        <div class="focus-card alert">
            <div class="focus-label">🚨 Supplier ที่ยังไม่ Check-in</div>
            <div class="focus-value">{total_remaining:,} <span>Pallets</span></div>
            <div class="focus-unit">แยกช่วงเวลาไว้ในตารางสรุปด้านล่าง</div>
        </div>
    </div>
    """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">ตารางสรุปรายงาน</div>', unsafe_allow_html=True)
    report_values = [
        ("group", "ภาพรวมยอดประมาณการและงานรับ"),
        ("ยอดงานประมาณการรับทั้งหมด", f"{inv_all:,}", False),
        ("ยอดงานรับกะดึก (18.00-06.00)", f"{night_received:,}", False),
        ("ยอดงานรับกะเช้า (07.00-18.00)", f"{morning_received:,}", False),
        ("group", "งาน Complete"),
        ("ยอดงาน Complete รับกะดึก (18.00-06.00)", f"{complete_night:,}", False),
        ("ยอดงาน Complete รับกะเช้า (07.00-18.00)", f"{complete_morning:,}", False),
        ("งาน Complete นอกรอบ (RP03)", f"{extra_work:,}", False),
        ("(%) สรุปยอดงานรับกะดึก", f"{pct_night_received:.2f}%", False),
        ("group", "งานรับคงเหลือ"),
        ("สรุปงานรับคงเหลือทั้งหมด\n(Pallets)", f"{total_remaining:,}", False),
        ("งานรับคงเหลือทั้งหมด\nกำลังลงสินค้า (Receiving Start)", f"{recv_start:,}", False),
        ("งานรับคงเหลือทั้งหมด\nInTransit ที่ Check-in แล้ว", f"{in_transit:,}", False),
        ("งานรับคงเหลือทั้งหมด\nSupplier ที่ยังไม่มาส่งกะดึก (18.00-06.00)", f"{sup_night_missing:,}", False),
        ("งานรับคงเหลือทั้งหมด\nSupplier ที่ยังไม่มาส่งกะเช้า (07.00-18.00)", f"{sup_morning_missing:,}", False),
        ("(%) Supplier ค้างรับรวม", f"{pct_remaining:.2f}%", True),
    ]
    st.markdown(
        report_table(
            "รายงานการรับสินค้าประจำวันที่",
            morning_date.strftime("%d/%m/%Y"),
            "(ตาม Window Time 00.00-23.59 น.)",
            report_values,
            6,
            compact=True,
            accent_colors=["#83c89d", "#8eb7df", "#aa94d0", "#f1bd69", "#e789a9", "#efd479", "#e99c82"],
        ),
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">กราฟสรุปภาพรวม</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.35, 1], gap="small")
    with c1:
        with st.container(key="card_radial_1"):
            chart_card_title("ประมาณการรับเทียบยอดงาน Complete", "เปรียบเทียบยอดรวม ไม่แยกกะ · Complete รวมงานนอกรอบ")
            st.plotly_chart(
                metric_bar_3d_chart(
                    ["ประมาณการรับทั้งหมด", "งาน Complete รวม"],
                    [inv_all, complete_all],
                    palette=["#8b9399", "#45d894"],
                    reference=inv_all,
                    height=250,
                ),
                use_container_width=True,
            )
    with c2:
        with st.container(key="card_radial_2"):
            chart_card_title("สรุปงานรับคงเหลือทั้งหมด", "หน่วย Pallets แยกตามสถานะ")
            st.plotly_chart(
                metric_bar_3d_chart(
                    ["กำลังลงสินค้า", "InTransit (Check-in แล้ว)", "Supplier กะดึก", "Supplier กะเช้า"],
                    [recv_start, in_transit, sup_night_missing, sup_morning_missing],
                    palette=["#e4a62e", "#5d9bcb", "#e63e2d", "#fa7a27"],
                    height=240,
                ),
                use_container_width=True,
            )

    with st.expander("✏️ ปรับปรุงข้อมูลรายงานประจำวัน", expanded=False):
        col_a, col_b, col_c = st.columns(3, gap="small")
        with col_a:
            inv_all = st.number_input(
                "ยอดงานประมาณการรับทั้งหมด", value=inv_all, step=1, key="inv_all", on_change=_autosave_morning_report
            )
            night_received = st.number_input(
                "ยอดงานรับกะดึก (ทั้งหมด)", value=night_received, step=1, key="night_received", on_change=_autosave_morning_report
            )
            morning_received = st.number_input(
                "ยอดงานรับกะเช้า (ทั้งหมด)", value=morning_received, step=1, key="morning_received", on_change=_autosave_morning_report
            )
        with col_b:
            complete_night = st.number_input(
                "ยอดงาน Complete รับกะดึก", value=complete_night, step=1, key="complete_night", on_change=_autosave_morning_report
            )
            complete_morning = st.number_input(
                "ยอดงาน Complete รับกะเช้า", value=complete_morning, step=1, key="complete_morning", on_change=_autosave_morning_report
            )
            extra_work = st.number_input(
                "งาน Complete - นอกรอบ", value=extra_work, step=1, key="extra_work", on_change=_autosave_morning_report
            )
        with col_c:
            recv_start = st.number_input(
                "กำลังลงสินค้า (Receiving Start)", value=recv_start, step=1, key="recv_start", on_change=_autosave_morning_report
            )
            in_transit = st.number_input(
                "InTransit ที่ Check-in แล้ว", value=in_transit, step=1, key="in_transit", on_change=_autosave_morning_report
            )
            sup_night_missing = st.number_input(
                "Supplier ยังไม่มาส่งกะดึก", value=sup_night_missing, step=1, key="sup_night_missing", on_change=_autosave_morning_report
            )
            sup_morning_missing = st.number_input(
                "Supplier ยังไม่มาส่งกะเช้า", value=sup_morning_missing, step=1, key="sup_morning_missing", on_change=_autosave_morning_report
            )
        if st.session_state.get("_morning_autosaved_at"):
            st.caption(f"จำข้อมูลล่าสุดอัตโนมัติ · {st.session_state['_morning_autosaved_at']}")

# =========================================================================
# ชีตที่ 2: สรุปงานรับและงานค้างส่ง(ย้อนหลัง)
# =========================================================================
else:
    with st.expander("📋 วาง PO สำหรับคำนวณเลื่อนส่งและ Reject", expanded=False):
        st.caption("ใส่ PO สำหรับเลื่อนส่ง และใส่รหัส Item สำหรับ Reject ระบบจะเทียบกับไฟล์ที่เลือกและจำรายการแยกตามวันที่")
        source_name = st.selectbox(
            "เลือก RP01 สำหรับเทียบ PO และ Item",
            ["RP01 16.10", "RP01 22.00"],
            key="hist_po_rp01_source",
            on_change=_autosave_current_report,
        )
        po_col1, po_col2 = st.columns(2, gap="small")
        with po_col1:
            st.text_area(
                "PO เลื่อนส่ง (หนึ่ง PO ต่อบรรทัด)",
                key="hist_delayed_po_input",
                height=110,
                placeholder="ตัวอย่าง:\n5011110177\n5011110673",
                on_change=_autosave_current_report,
            )
        with po_col2:
            st.text_area(
                "Item สำหรับ Reject (หนึ่ง Item ต่อบรรทัด)",
                key="hist_reject_item_input",
                height=110,
                placeholder="วางรหัส Item ที่ต้องตรวจ Reject",
                on_change=_autosave_current_report,
            )

        rp01_source = uploaded_rp01_1610 if source_name == "RP01 16.10" else uploaded_rp01_2200
        delayed_pos = st.session_state.get("hist_delayed_po_input", "")
        reject_items = st.session_state.get("hist_reject_item_input", "")
        lookup_changed = False
        if rp01_source is None:
            st.info(f"กรุณา Import {source_name} ก่อน ระบบจึงจะเทียบ PO ได้")
        else:
            if _po_values(delayed_pos):
                try:
                    delayed_value, delayed_missing = _lookup_delayed_pallets(rp01_source, delayed_pos)
                    if st.session_state.get("hist_delayed") != delayed_value:
                        st.session_state["hist_delayed"] = delayed_value
                        lookup_changed = True
                    delayed_count = len(_po_values(delayed_pos))
                    st.caption(f"เลื่อนส่ง: {delayed_value:,} Pallets · พบ PO {delayed_count - len(delayed_missing)}/{delayed_count}")
                    if delayed_missing:
                        st.warning("PO เลื่อนส่งที่ไม่พบใน RP01: " + ", ".join(delayed_missing))
                except Exception as exc:
                    st.error(f"คำนวณเลื่อนส่งไม่สำเร็จ: {exc}")
            else:
                st.caption("ยังไม่ได้วาง PO เลื่อนส่ง · ใช้ยอดที่บันทึกไว้หรือกรอกยอดเองในตัวแก้ไขด้านล่าง")

            if _po_values(reject_items):
                if uploaded_rp03 is None:
                    st.info("กรุณา Import RP03 เพื่อคำนวณ Reject")
                else:
                    try:
                        reject_value, reject_missing, reject_invalid = _lookup_reject_pallets(
                            rp01_source, uploaded_rp03, reject_items
                        )
                        if st.session_state.get("hist_not_po") != reject_value:
                            st.session_state["hist_not_po"] = reject_value
                            lookup_changed = True
                        reject_count = len(_po_values(reject_items))
                        st.caption(f"Reject: {_format_quantity(reject_value)} Pallets · พบ Item {reject_count - len(reject_missing)}/{reject_count}")
                        if reject_missing:
                            st.warning("Item Reject ที่ไม่พบใน RP01: " + ", ".join(reject_missing))
                        if reject_invalid:
                            st.warning("มี Item ที่ไม่มีค่า TI/HI สำหรับคำนวณ หรือมี TI/HI หลายชุด: " + ", ".join(reject_invalid))
                    except Exception as exc:
                        st.error(f"คำนวณ Reject ไม่สำเร็จ: {exc}")
            else:
                st.caption("ยังไม่ได้วาง Item Reject · ใช้ยอดที่บันทึกไว้หรือกรอกยอดเองในตัวแก้ไขด้านล่าง")
        if lookup_changed:
            _autosave_current_report()

    hist_inv = st.session_state.get("hist_inv", 6126)
    hist_total_recv = st.session_state.get("hist_total_recv", 5904)
    hist_complete = st.session_state.get("hist_complete", 5722)
    hist_extra = st.session_state.get("hist_extra", 138)
    hist_extra_chart = 0 if hist_extra is None else hist_extra
    hist_extra_label = "ตรวจสอบคีย์" if hist_extra is None else f"{hist_extra:,}"
    hist_delayed = st.session_state.get("hist_delayed", 0)
    hist_not_po = st.session_state.get("hist_not_po", 73)
    hist_not_billed = st.session_state.get("hist_not_billed", 250)

    hist_not_recv = hist_not_po + hist_not_billed
    hist_pct_not_recv = (
        (hist_not_recv / hist_total_recv * 100) if hist_total_recv > 0 else 0
    )
    hist_complete_all = hist_complete + hist_extra_chart
    hist_complete_all_label = "ตรวจสอบคีย์" if hist_extra is None else f"{hist_complete_all:,}"

    st.markdown(
        """
        <style>
        body:has(.history-theme) [data-testid="stMainBlockContainer"] { padding-top:.35rem; padding-bottom:.4rem; }
        body:has(.history-theme) [data-testid="stVerticalBlock"] { gap:.42rem; }
        body:has(.history-theme) [data-testid="stAppViewContainer"] { background:linear-gradient(145deg,#f4f8ff 0%,#e9f1fc 55%,#f8faff 100%); }
        body:has(.history-theme) .hero { background:transparent; border:0; box-shadow:none; text-align:left; max-width:none; padding:1px 4px 0; margin:0; border-radius:0; }
        body:has(.history-theme) .hero-kicker { color:#48658b; font-size:.55rem; }
        body:has(.history-theme) .hero-title { color:#172c4b; font-size:1.45rem; margin:0; }
        body:has(.history-theme) .hero-subtitle { color:#48658b; font-size:.68rem; margin:0; }
        body:has(.history-theme) .hero-bar { display:none; }
        body:has(.history-theme) .section-title { color:#172c4b; font-size:.88rem; margin:5px 0 2px; padding-bottom:0; }
        body:has(.history-theme) .section-title::before { background:linear-gradient(160deg,#101b32,#24579b 65%,#80b5f2); box-shadow:0 2px 7px rgba(31,75,135,.3); }
        body:has(.history-theme) .focus-card { background:linear-gradient(125deg,#0b1426 0%,#14294a 48%,#285a9c 100%); border:1px solid #6e9ed8; box-shadow:0 0 0 1px rgba(83,142,216,.2),0 8px 18px rgba(18,48,88,.22),0 18px 32px rgba(18,48,88,.14),inset 0 2px 0 rgba(210,231,255,.28),inset 0 -3px 6px rgba(0,0,0,.28); }
        body:has(.history-theme) .focus-card::after { content:""; position:absolute; right:-34px; top:-34px; width:140px; height:140px; border-radius:50%; background:radial-gradient(circle,rgba(150,197,255,.38),rgba(150,197,255,0) 72%); pointer-events:none; }
        body:has(.history-theme) .focus-card .focus-label { color:#e1efff; }
        body:has(.history-theme) .focus-card .focus-value { color:#fff !important; text-shadow:0 0 10px rgba(111,174,255,.32); }
        body:has(.history-theme) .focus-card .focus-unit,
        body:has(.history-theme) .focus-card .focus-value span { color:#b7d5f8 !important; }
        body:has(.history-theme) .focus-grid { gap:8px; margin:1px 0 2px; }
        body:has(.history-theme) .focus-card { padding:8px 16px 7px; border-radius:18px; }
        body:has(.history-theme) .focus-label { font-size:.72rem; }
        body:has(.history-theme) .focus-value { font-size:1.8rem; line-height:1.15; margin:1px 0; }
        body:has(.history-theme) .focus-unit { font-size:.62rem; }
        body:has(.history-theme) .focus-card.status-complete { background:linear-gradient(125deg,#0b1426 0%,#14294a 48%,#285a9c 100%) !important; border-color:#6e9ed8 !important; box-shadow:0 0 0 1px rgba(83,142,216,.22),0 8px 18px rgba(18,48,88,.24),0 18px 32px rgba(18,48,88,.15),inset 0 2px 0 rgba(210,231,255,.28),inset 0 -3px 6px rgba(0,0,0,.28) !important; }
        body:has(.history-theme) .focus-card.alert { background:linear-gradient(135deg,#fff 0%,#eaf3ff 58%,#d8e9ff 100%) !important; border-color:#eaa341 !important; box-shadow:0 0 0 1px rgba(234,163,65,.18),0 7px 17px rgba(54,82,118,.16),0 0 16px rgba(245,178,63,.2),inset 0 2px 0 rgba(255,255,255,.98),inset 0 -3px 5px rgba(64,91,127,.08) !important; }
        body:has(.history-theme) .focus-card.alert::after { background:radial-gradient(circle,rgba(255,197,76,.62),rgba(255,146,45,0) 72%); }
        body:has(.history-theme) .focus-card.alert .focus-label { color:#142a49 !important; }
        body:has(.history-theme) .focus-card.alert .focus-unit,
        body:has(.history-theme) .focus-card.alert .focus-value span { color:#354b67 !important; text-shadow:none; }
        body:has(.history-theme) .focus-card.alert .focus-value .focus-number { color:transparent !important; -webkit-text-fill-color:transparent !important; font-size:inherit !important; font-weight:inherit !important; line-height:inherit !important; background:linear-gradient(100deg,#171717 0%,#a94316 48%,#e3a11b 78%,#f2c94c 100%); -webkit-background-clip:text; background-clip:text; filter:drop-shadow(0 1px 2px rgba(176,92,22,.2)); }
        body:has(.history-theme) .report-wrap { background:linear-gradient(150deg,#fff,#edf4ff); border-color:#b8cce5; box-shadow:0 2px 4px rgba(27,54,89,.07),0 12px 28px rgba(34,73,122,.15),0 24px 44px rgba(34,73,122,.07),inset 0 2px 0 #fff,inset 0 -2px 5px rgba(45,83,130,.05); }
        body:has(.history-theme) .report-head { border-color:#c6d7ec; color:#172c4b; }
        body:has(.history-theme) .report-date,
        body:has(.history-theme) .report-group { color:#294d7d; }
        body:has(.history-theme) .report-group { border-color:#c6d7ec; }
        body:has(.history-theme) .report-wrap.compact { padding:2px 4px 4px; border-radius:18px; }
        body:has(.history-theme) .report-wrap.compact .report-head { padding:2px 6px 3px; margin-bottom:0; font-size:.82rem; }
        body:has(.history-theme) .report-wrap.compact .report-grid { gap:3px; padding:1px; }
        body:has(.history-theme) .report-wrap.compact .report-group { font-size:.68rem; line-height:1.08; padding:0 2px 1px; }
        body:has(.history-theme) .report-wrap.compact .report-cell { min-height:46px; padding:2px 3px; border-radius:8px; gap:1px; }
        body:has(.history-theme) .report-wrap.compact .report-label { font-size:.66rem; line-height:1.08; }
        body:has(.history-theme) .report-wrap.compact .report-value { font-size:.95rem; }
        body:has(.history-theme) .report-cell { background:linear-gradient(145deg,#fff 0%,#eaf2fd 100%) !important; border-color:#c2d3e9 !important; border-top-color:#477bb9 !important; box-shadow:0 3px 7px rgba(27,54,89,.12),inset 0 2px 0 rgba(255,255,255,.96); }
        body:has(.history-theme) .report-label { color:#233d60 !important; }
        body:has(.history-theme) .report-value { color:#142a49 !important; }
        body:has(.history-theme) [class*="st-key-card"] { background:linear-gradient(145deg,#fff 0%,#e9f2ff 100%); border-color:#b7cce6; box-shadow:0 3px 8px rgba(27,54,89,.11),0 12px 24px rgba(34,73,122,.12),inset 0 2px 0 rgba(255,255,255,.96); }
        body:has(.history-theme) .chart-title { color:#172c4b; }
        body:has(.history-theme) .chart-sub { color:#48658b; border-color:#c6d7ec; }
        body:has(.history-theme) [class*="st-key-card_hist"] { padding:.35rem .45rem .2rem; border-radius:20px; }
        body:has(.history-theme) .chart-title { font-size:.8rem; margin-bottom:1px; }
        body:has(.history-theme) .chart-sub { font-size:.62rem; margin-bottom:2px; padding-bottom:2px; }
        </style>
        <div class="history-theme" aria-hidden="true"></div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="focus-grid">
            <div class="focus-card status-complete">
                <div class="focus-label">📦 งาน Complete รวมทั้งหมด</div>
                <div class="focus-value">{hist_complete_all_label} <span>Pallets</span></div>
                <div class="focus-unit">รวมงาน Complete (22.00 น.) และงานนอกรอบ</div>
            </div>
            <div class="focus-card alert">
                <div class="focus-label">🌙 งานไม่ได้รับทั้งหมด</div>
                <div class="focus-value"><span class="focus-number">{hist_not_recv:,}</span> <span>Pallets</span></div>
                <div class="focus-unit">รวม Reject / ส่งไม่ครบ PO และยังไม่ยื่นบิล</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-title">ตารางสรุปรายงาน</div>', unsafe_allow_html=True
    )
    hist_report_values = [
        ("group", "ภาพรวมยอดประมาณการและงานรับ"),
        ("ประมาณการรับ (Inventory)", f"{hist_inv:,}", False),
        ("งานรับทั้งหมด (16.10 น.)", f"{hist_total_recv:,}", False),
        ("group", "งาน Complete"),
        ("งาน Complete (22.00 น.)", f"{hist_complete:,}", False),
        ("งานนอกรอบ Complete", hist_extra_label, False),
        ("งาน Complete รวม", hist_complete_all_label, False),
        ("group", "งานค้างส่ง"),
        ("งานไม่ได้รับ (%)", f"{hist_pct_not_recv:.2f}%", True),
        ("เลื่อนส่ง", _format_quantity(hist_delayed), False),
        ("Reject / ส่งไม่ครบตาม PO", _format_quantity(hist_not_po), False),
        ("ยังไม่ยื่นบิล", f"{hist_not_billed:,}", False),
        ("งานไม่ได้รับทั้งหมด", f"{hist_not_recv:,}", True),
    ]
    st.markdown(
        report_table(
            "สรุปงานรับและงานค้างส่ง ประจำวันที่",
            hist_date.strftime("%d/%m/%Y"),
            "(00.00-22.00 น.)",
            hist_report_values,
            6,
            compact=True,
        ),
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">กราฟสรุปภาพรวม</div>', unsafe_allow_html=True)
    col_chart1, col_chart2 = st.columns([1.35, 1], gap="small")
    with col_chart1:
        with st.container(key="card_hist_tube"):
            chart_card_title(
                "📊 เปรียบเทียบยอดรับและงาน Complete",
                "แท่งแสดงจำนวน Pallets · % เทียบกับงานรับทั้งหมด (16.10 น.)",
            )
            st.plotly_chart(
                metric_bar_3d_chart(
                    ["งานรับทั้งหมด", "งาน Complete", "งานนอกรอบ"],
                    [hist_total_recv, hist_complete, hist_extra_chart],
                    reference=hist_total_recv,
                    height=250,
                    palette=["#29476d", "#3d78bd", "#83a9d4"],
                ),
                use_container_width=True,
            )
    with col_chart2:
        with st.container(key="card_hist_radial"):
            chart_card_title(
                "🚨 งานไม่ได้รับทั้งหมด · สัดส่วน",
                f"เลื่อนส่ง {_format_quantity(hist_delayed)} Pallets แสดงแยกจากยอดนี้",
            )
            st.plotly_chart(
                pending_donut_chart(
                    hist_not_billed,
                    hist_not_po,
                    hist_not_recv,
                    height=250,
                ),
                use_container_width=True,
            )

    with st.expander("✏️ ปรับปรุงข้อมูลสรุปงานรับและงานค้างส่ง(ย้อนหลัง)", expanded=False):
        hc1, hc2, hc3, hc4 = st.columns(4, gap="small")
        with hc1:
            hist_inv = st.number_input(
                "ประมาณการรับ (Inventory)", value=hist_inv, step=1, key="hist_inv", on_change=_autosave_current_report
            )
            hist_total_recv = st.number_input(
                "งานรับทั้งหมด (16.10 น.)", value=hist_total_recv, step=1, key="hist_total_recv", on_change=_autosave_current_report
            )
        with hc2:
            hist_complete = st.number_input(
                    "งาน Complete (22.00 น.)", value=hist_complete, step=1, key="hist_complete", on_change=_autosave_current_report
            )
            if hist_extra is not None:
                hist_extra = st.number_input(
                    "งานนอกรอบ Complete", value=hist_extra, step=1, key="hist_extra", on_change=_autosave_current_report
                )
            else:
                st.caption("งานนอกรอบยังไม่แสดงยอดจนกว่าจะจับคู่คีย์ RP01/RP03 ได้")
        with hc3:
            hist_delayed = st.number_input(
                "เลื่อนส่ง (Pallets)", value=float(hist_delayed), step=1.0, key="hist_delayed", on_change=_autosave_current_report
            )
            hist_not_po = st.number_input(
                "Reject / ส่งไม่ครบตาม PO (Pallets)", value=math.ceil(float(hist_not_po)), step=1, key="hist_not_po", on_change=_autosave_current_report
            )
        with hc4:
            hist_not_billed = st.number_input(
                "ยังไม่ยื่นบิล", value=hist_not_billed, step=1, key="hist_not_billed", on_change=_autosave_current_report
            )
        if st.session_state.get("_autosaved_at"):
            st.caption(f"จำข้อมูลล่าสุดอัตโนมัติ · {st.session_state['_autosaved_at']}")
