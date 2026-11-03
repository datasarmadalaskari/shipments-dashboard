"""China & F-550 Shipments dashboard.

Reads the MASTER(NEW) workbook from Google Drive, and refreshes the data
every Saturday and Wednesday (the first visit after each refresh time pulls
the latest file).
"""

import io
import os
import re
import datetime as dt
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

TZ = ZoneInfo("Asia/Amman")
REFRESH_DAYS = (2, 5)  # Wednesday=2, Saturday=5 (Monday=0)
REFRESH_HOUR = 8       # 08:00 Amman time

STATUS_ORDER = ["Loaded", "Delivered", "Other"]
STATUS_COLORS = {"Loaded": "#3b82f6", "Delivered": "#16a34a", "Other": "#dc2626"}

st.set_page_config(page_title="China & F-550 Shipments", page_icon=":material/local_shipping:", layout="wide")


# ---------------------------------------------------------------- settings
def setting(name, default=None):
    """Read a value from Streamlit secrets, then environment variables."""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


# ---------------------------------------------------------------- refresh schedule
def last_refresh_slot(now=None):
    """Most recent Wednesday/Saturday at REFRESH_HOUR, in Amman time."""
    now = now or dt.datetime.now(TZ)
    for back in range(0, 8):
        day = now - dt.timedelta(days=back)
        slot = day.replace(hour=REFRESH_HOUR, minute=0, second=0, microsecond=0)
        if day.weekday() in REFRESH_DAYS and slot <= now:
            return slot
    return now


# ---------------------------------------------------------------- loading the workbook
def download_workbook() -> bytes:
    local = setting("LOCAL_PATH")
    if local:
        with open(local, "rb") as f:
            return f.read()

    file_id = setting("DRIVE_FILE_ID")
    if not file_id:
        st.error("DRIVE_FILE_ID is not set. Add it to the app secrets.")
        st.stop()

    # Private file: read it with a Google service account.
    if setting("gcp_service_account"):
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload

        creds = service_account.Credentials.from_service_account_info(
            dict(setting("gcp_service_account")),
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        meta = drive.files().get(fileId=file_id, fields="mimeType", supportsAllDrives=True).execute()
        if meta["mimeType"] == "application/vnd.google-apps.spreadsheet":
            req = drive.files().export_media(
                fileId=file_id,
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            req = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        return buf.getvalue()

    # File shared as "Anyone with the link".
    urls = [
        f"https://drive.google.com/uc?export=download&id={file_id}",
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx",
    ]
    for url in urls:
        r = requests.get(url, timeout=60)
        if r.ok and r.content[:2] == b"PK":  # an .xlsx is a zip file
            return r.content
    st.error("Couldn't download the workbook. Check the file is shared and DRIVE_FILE_ID is right.")
    st.stop()


def pdate(v):
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return pd.Timestamp(v)
    if isinstance(v, str):
        m = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", v)
        if m:
            d, mo, y = map(int, m.groups())
            y += 2000 if y < 100 else 0
            try:
                return pd.Timestamp(y, mo, d)
            except ValueError:
                return pd.NaT
    return pd.NaT


def txt(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def clean_status(v):
    s = txt(v).lower()
    if "deliver" in s:
        return "Delivered"
    if "load" in s:
        return "Loaded"
    return "Other"


def clean_dest(v):
    s = txt(v)
    if not s:
        return "Not set"
    return s.split(",")[0].replace(" PORT", "").strip().title()


def read_sheet(xls, name_part, line_cols):
    name = next(s for s in xls.sheet_names if name_part.lower() in s.lower())
    df = pd.read_excel(xls, sheet_name=name, header=0)
    df.columns = [str(c).strip() for c in df.columns]
    rows = []
    for _, r in df.iterrows():
        booking = txt(r.get("Booking#"))
        cars = pd.to_numeric(r.get("# cars"), errors="coerce")
        if not booking or pd.isna(cars):  # notes and parts-only lines
            continue
        sail, eta, dlv = pdate(r.get("Sail Date")), pdate(r.get("ETA")), pdate(r.get("Delivery Date"))
        when = next((t for t in (sail, eta, dlv) if not pd.isna(t)), pd.NaT)
        line = next((txt(r.get(c)) for c in line_cols
                     if c in df.columns and txt(r.get(c)) not in ("", "-") and not txt(r.get(c)).isdigit()), "")
        rows.append({
            "Status": clean_status(r.get("Status")),
            "Cars": int(cars),
            "Booking #": booking,
            "Model": txt(r.get("UNIT - MODEL")),
            "Destination": clean_dest(r.get("Destin.")),
            "Shipping line": line,
            "Sail date": sail,
            "ETA": eta,
            "Supplier": txt(r.get("Supplier")).strip("- "),
            "Consignee": txt(r.get("Consignee - Tafweed")),
            "Year": "No date" if pd.isna(when) else str(when.year),
        })
    return pd.DataFrame(rows)


@st.cache_data(max_entries=2, show_spinner="Loading shipments…")
def load_data(slot_key: str):
    raw = download_workbook()
    xls = pd.ExcelFile(io.BytesIO(raw))
    china = read_sheet(xls, "China Shipments", ["Shipping Line.1", "Shipping Line"])
    f550 = read_sheet(xls, "F-550 Shipments", ["SL"])
    loaded_at = dt.datetime.now(TZ)
    return china, f550, loaded_at


# ---------------------------------------------------------------- styling
st.markdown(
    """
    <style>
      .block-container {padding-top: 2rem; max-width: 1300px;}
      h1 {font-weight: 500; margin-bottom: 0;}
      .sub {color: #8a8f98; font-size: .9rem; margin-top: -.4rem;}
      .kpi {border: 1px solid rgba(128,128,128,.25); border-radius: 10px; padding: .8rem 1rem;
            border-top-width: 3px; height: 100%;}
      .kpi b {font-size: 1.9rem; font-weight: 500; display: block; line-height: 1.15;}
      .kpi span {font-weight: 600; font-size: .9rem;}
      .kpi small {display: block; color: #8a8f98; font-size: .78rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def kpi(col, value, label, sub, color):
    col.markdown(
        f'<div class="kpi" style="border-top-color:{color}"><b>{value:,}</b><span>{label}</span><small>{sub}</small></div>',
        unsafe_allow_html=True,
    )


def nb(df):
    return df["Booking #"].nunique()


# ---------------------------------------------------------------- page
slot = last_refresh_slot()
china, f550, loaded_at = load_data(slot.isoformat())

st.title("China & F-550 Shipments")
st.markdown(f'<p class="sub">Last refresh: <b>{loaded_at:%d/%m/%Y, %H:%M}</b></p>', unsafe_allow_html=True)

tab_china, tab_f550 = st.tabs(["China Shipments", "F-550 Shipments"])


def render(df: pd.DataFrame, key: str):
    if df.empty:
        st.info("No shipments found in this sheet.")
        return

    years = sorted([y for y in df["Year"].unique() if y != "No date"], reverse=True)
    options = years + ["All years"]
    this_year = str(dt.datetime.now(TZ).year)
    default = options.index(this_year) if this_year in options else 0

    c1, c2 = st.columns([3, 1])
    with c2:
        year = st.selectbox("Year", options, index=default, key=f"{key}-year")
    with c1:
        status_pick = st.segmented_control(
            "Show", ["All"] + STATUS_ORDER, default="All", key=f"{key}-status",
        ) or "All"

    view = df if year == "All years" else df[df["Year"] == year]

    # Headline figures (distinct booking numbers)
    cols = st.columns(5)
    kpi(cols[0], nb(view), "All shipments", "booking numbers", "#8a8f98")
    kpi(cols[1], nb(view[view.Status == "Loaded"]), "Loaded", "booking numbers", STATUS_COLORS["Loaded"])
    kpi(cols[2], nb(view[view.Status == "Delivered"]), "Delivered", "booking numbers", STATUS_COLORS["Delivered"])
    kpi(cols[3], nb(view[view.Status == "Other"]), "No status / other", "booking numbers", STATUS_COLORS["Other"])
    kpi(cols[4], int(view["Cars"].sum()), "Cars", "units shipped", "#8a8f98")

    st.write("")
    st.subheader("By destination")
    dest = view.groupby("Destination")["Booking #"].nunique().sort_values()
    fig = go.Figure(go.Bar(
        x=dest.values, y=dest.index, orientation="h", marker_color="#14b8a6",
        text=dest.values, textposition="outside",
        hovertemplate="%{y}: %{x} booking numbers<extra></extra>",
    ))
    fig.update_layout(height=max(140, 34 * len(dest) + 30), margin=dict(l=0, r=30, t=0, b=0),
                      xaxis=dict(visible=False), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False}, key=f"{key}-dest")

    # Details table
    st.subheader("Shipment details")
    q = st.text_input("Search", placeholder="Search booking, model, consignee, line…",
                      key=f"{key}-q", label_visibility="collapsed")
    table = view if status_pick == "All" else view[view.Status == status_pick]
    if q:
        mask = table.astype(str).apply(lambda col: col.str.contains(q, case=False, regex=False)).any(axis=1)
        table = table[mask]
    table = table.sort_values("Sail date", ascending=False, na_position="last")
    st.caption(f"{nb(table):,} booking numbers · {len(table):,} lines")
    st.dataframe(
        table,
        hide_index=True,
        width="stretch",
        height=520,
        column_config={
            "Sail date": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "ETA": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Cars": st.column_config.NumberColumn(format="%d"),
        },
    )


with tab_china:
    render(china, "china")
with tab_f550:
    render(f550, "f550")
