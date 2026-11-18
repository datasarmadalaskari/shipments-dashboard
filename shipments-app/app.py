"""China & F-550 Shipments dashboard.

Reads the MASTER(NEW) workbook from Google Drive, and refreshes the data
every Saturday and Wednesday (the first visit after each refresh time pulls
the latest file).

Tabs:
  * China Shipments  (unchanged)
  * F-550 Shipments  (unchanged)
  * China VINs       (new: cars by China contract and booking number, VIN search)
"""

import io
import os
import re
import datetime as dt
from email.utils import parsedate_to_datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

TZ = ZoneInfo("Asia/Amman")
REFRESH_DAYS = (2, 5)  # Wednesday=2, Saturday=5 (Monday=0)
REFRESH_HOUR = 8       # 08:00 Amman time
VINS_DEFAULT_YEAR = "2026"   # the VINs tab opens on the new cars of this year

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
def download_workbook():
    """Returns (file bytes, when the file was last changed - or None if unknown)."""
    local = setting("LOCAL_PATH")
    if local:
        with open(local, "rb") as f:
            data = f.read()
        return data, dt.datetime.fromtimestamp(os.path.getmtime(local), TZ)

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
        meta = drive.files().get(fileId=file_id, fields="mimeType,modifiedTime",
                                 supportsAllDrives=True).execute()
        modified = dt.datetime.fromisoformat(meta["modifiedTime"].replace("Z", "+00:00")).astimezone(TZ)
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
        return buf.getvalue(), modified

    # File shared as "Anyone with the link".
    urls = [
        f"https://drive.google.com/uc?export=download&id={file_id}",
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx",
    ]
    for url in urls:
        r = requests.get(url, timeout=60)
        if r.ok and r.content[:2] == b"PK":  # an .xlsx is a zip file
            modified = None
            if r.headers.get("Last-Modified"):
                modified = parsedate_to_datetime(r.headers["Last-Modified"]).astimezone(TZ)
            return r.content, modified
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


def _sheet(xls, name_part):
    return next(s for s in xls.sheet_names if name_part.lower() in s.lower())


def read_sheet(xls, name_part, line_cols):
    name = _sheet(xls, name_part)
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


def read_vins(xls):
    """China VINs: one row per car, with its contract (Order) and booking number."""
    df = pd.read_excel(xls, sheet_name=_sheet(xls, "China VINs"))
    df.columns = [str(c).strip() for c in df.columns]   # 'Order date' first occurrence = the order date
    df = df[df["VIN"].map(txt) != ""]

    contracts = pd.read_excel(xls, sheet_name=_sheet(xls, "China Contracts"))
    known_col = next(c for c in contracts.columns if "china contract" in str(c).lower())
    known = {txt(v) for v in contracts[known_col] if txt(v)}

    order = df["Order"].map(txt)
    date = pd.to_datetime(df["Order date"], errors="coerce")
    out = pd.DataFrame({
        "VIN": df["VIN"].map(txt).str.upper(),
        "Model": df["Model, Exterior, Interior / Contract"].map(txt),
        "Color": df["Color"].map(txt),
        "Contract": order,
        "Contract matched": order.isin(known).map({True: "Yes", False: "No"}),
        "Booking #": df["Booking#"].map(txt),
        "Status": df["Status"].map(txt),
        "Order date": date,
    })
    # Year: from the contract code (e.g. JFEX-26-290F -> 2026), since most rows have no order date
    code = order.str.extract(r"-(\d{2})-")[0]
    from_code = code.map(lambda c: "" if pd.isna(c) else f"20{c}")
    from_date = date.dt.year.map(lambda y: "" if pd.isna(y) else str(int(y)))
    out["Year"] = from_code.where(from_code != "", from_date).replace("", "No date")

    # a car can appear twice (cancelled line + real line): keep the real one
    # and leave out cancelled / never-shipped lines
    out = out[~out["Status"].str.contains("Cancelled|Not Shipped", case=False, regex=True)]
    out = out.assign(_has_booking=out["Booking #"] != "")
    out = out.sort_values("_has_booking", ascending=False).drop_duplicates("VIN")
    full = df.loc[out.index]                      # every column of the sheet, same rows
    out = out.drop(columns="_has_booking")
    return out.reset_index(drop=True), full.reset_index(drop=True)


@st.cache_data(max_entries=2, show_spinner="Loading shipments…")
def load_data(slot_key: str):
    raw, modified = download_workbook()
    xls = pd.ExcelFile(io.BytesIO(raw))
    china = read_sheet(xls, "China Shipments", ["Shipping Line.1", "Shipping Line"])
    f550 = read_sheet(xls, "F-550 Shipments", ["SL"])
    vins, vins_full = read_vins(xls)
    china_raw = pd.read_excel(xls, sheet_name=_sheet(xls, "China Shipments"))
    f550_raw = pd.read_excel(xls, sheet_name=_sheet(xls, "F-550 Shipments"))
    loaded_at = dt.datetime.now(TZ)
    return china, f550, vins, vins_full, china_raw, f550_raw, modified, loaded_at


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
china, f550, vins, vins_full, china_raw, f550_raw, modified, loaded_at = load_data(slot.isoformat())

t_col, b_col = st.columns([6, 1])
with t_col:
    st.title("China & F-550 Shipments")
with b_col:
    st.write("")
    if st.button("↻ Refresh", width="stretch", help="Read the latest file from Google Drive now"):
        st.cache_data.clear()
        st.rerun()
file_txt = f"{modified:%d/%m/%Y, %H:%M}" if modified else "not available"
st.markdown(f'<p class="sub">File last updated: <b>{file_txt}</b> · '
            f'Loaded here: <b>{loaded_at:%d/%m/%Y, %H:%M}</b></p>', unsafe_allow_html=True)

tab_china, tab_f550, tab_vins = st.tabs(["China Shipments", "F-550 Shipments", "China VINs"])


def drill_down(key: str, table: pd.DataFrame, raw: pd.DataFrame, vins_full=None):
    """Pick one booking: its shipment lines, every sheet column, and (China) its VINs."""
    st.subheader("Open a shipment")
    opts = sorted(b for b in table["Booking #"].unique() if b)
    if not opts:
        return
    pick = st.selectbox("Booking #", opts, index=None, key=f"{key}-open",
                        placeholder="Pick a booking number to see its details and VINs")
    if not pick:
        return

    lines = table[table["Booking #"] == pick]
    st.markdown(f"**Booking {pick}** · {int(lines['Cars'].sum())} cars · "
                f"{', '.join(sorted(set(lines['Status'])))}")
    st.dataframe(lines.drop(columns=["Year"]), hide_index=True, width="stretch")

    sheet_rows = raw[raw["Booking#"].map(txt) == pick].dropna(axis=1, how="all")
    with st.expander("Everything in the sheet for this booking", expanded=False):
        st.dataframe(sheet_rows, hide_index=True, width="stretch")

    if vins_full is None:
        st.info("F-550 VINs are not linked to a booking number yet, so they are not shown here.")
        return
    cars = (vins_full[vins_full["Booking#"].map(txt) == pick]
            .drop(columns=["Inventory", "SOLD YEAR"], errors="ignore")
            .dropna(axis=1, how="all"))
    st.markdown(f"**VINs in this booking: {len(cars)}**")
    if cars.empty:
        st.caption("No VINs in the China VINs sheet for this booking.")
    else:
        st.dataframe(cars, hide_index=True, width="stretch", height=420)


def render(df: pd.DataFrame, key: str, raw: pd.DataFrame, vins_full=None):
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
            "Show", ["All"] + STATUS_ORDER, default="Loaded", key=f"{key}-status",
        ) or "All"

    view = df if year == "All years" else df[df["Year"] == year]
    model_q = st.text_input("Car model", placeholder="e.g. Elantra, Tucson", key=f"{key}-model")
    if model_q.strip():
        view = view[view["Model"].str.contains(model_q.strip(), case=False, regex=False)]
    if view.empty:
        st.warning("No shipments match these filters.")
        return

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

    drill_down(key, table, raw, vins_full)


def render_vins(v: pd.DataFrame):
    if v.empty:
        st.info("No cars found in the China VINs sheet.")
        return

    # filters
    years = sorted([y for y in v["Year"].unique() if y != "No date"], reverse=True)
    options = ["All years"] + years
    default = options.index(VINS_DEFAULT_YEAR) if VINS_DEFAULT_YEAR in options else 0
    f1, f2, f3 = st.columns([1, 2, 2])
    year = f1.selectbox("Year (order date)", options, index=default, key="vins-year")
    sub = v if year == "All years" else v[v["Year"] == year]
    model_q = st.text_input("Car model", placeholder="e.g. Elantra, Tucson", key="vins-model")
    if model_q.strip():
        sub = sub[sub["Model"].str.contains(model_q.strip(), case=False, regex=False)]

    contracts = sorted(c for c in sub["Contract"].unique() if c)
    pick_c = f2.multiselect("Contract", contracts, key="vins-contract", placeholder="All contracts")
    bookings = sorted(b for b in sub["Booking #"].unique() if b)
    pick_b = f3.multiselect("Booking #", bookings, key="vins-booking", placeholder="All bookings")
    q = st.text_input("Search VIN", placeholder="Paste a VIN or part of it", key="vins-q")
    only_unmatched = st.checkbox("Only cars without a matching China contract", key="vins-unmatched")

    if pick_c:
        sub = sub[sub["Contract"].isin(pick_c)]
    if pick_b:
        sub = sub[sub["Booking #"].isin(pick_b)]
    if only_unmatched:
        sub = sub[sub["Contract matched"] == "No"]
    if q.strip():
        sub = sub[sub["VIN"].str.contains(q.strip().upper(), regex=False)]
    if sub.empty:
        st.warning("No cars match these filters.")
        return

    k = st.columns(4)
    kpi(k[0], len(sub), "Cars", "VINs", "#8a8f98")
    kpi(k[1], sub["Contract"].nunique(), "Contracts", "China contracts", "#14b8a6")
    kpi(k[2], sub["Booking #"].nunique(), "Bookings", "booking numbers", STATUS_COLORS["Loaded"])
    kpi(k[3], int((sub["Contract matched"] == "No").sum()), "No matching contract", "cars",
        STATUS_COLORS["Other"])

    # by booking
    st.subheader("By booking")
    by_b = (sub.groupby(["Booking #", "Contract"], dropna=False)
              .agg(Cars=("VIN", "count"), VINs=("VIN", lambda s: ", ".join(s)))
              .reset_index()
              .sort_values("Cars", ascending=False))
    st.dataframe(
        by_b,
        hide_index=True,
        width="stretch",
        height=280,
        column_config={
            "Cars": st.column_config.NumberColumn(format="%d"),
            "VINs": st.column_config.TextColumn(width="large"),
        },
    )

    # every car
    st.subheader("Cars")
    cols = ["VIN", "Model", "Color", "Contract", "Contract matched", "Booking #", "Status", "Order date"]
    st.dataframe(
        sub[cols].sort_values("Order date", ascending=False, na_position="last"),
        hide_index=True,
        width="stretch",
        height=420,
        column_config={
            "Order date": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Contract matched": st.column_config.TextColumn("Matches a China contract"),
        },
    )
    st.caption(f"{len(sub):,} car(s) with the filters above")


with tab_china:
    render(china, "china", china_raw, vins_full)
with tab_f550:
    render(f550, "f550", f550_raw, None)
with tab_vins:
    render_vins(vins)
