"""
Container Cargo Load Planner — Streamlit Edition
=====================================================
Supports:
  - Upload cargo Excel/CSV + auto-load transit days from bundled Transit_Days.xlsx
  - Priority & Required_Arrival_Date columns
  - Predicted departure + estimated arrival dates
  - Stuffing rate shown as percentage
  - Dynamic loading suggestions (weight + fill context)
  - Interactive 3-D container visualisation via Plotly
  - NLP chatbot powered by Claude API (claude-sonnet-4-20250514)
  - One-click deploy to Streamlit Cloud / GitHub
"""
import streamlit as st

# =========================
# 🔐 SIMPLE LOGIN SYSTEM
# =========================

def login():
   st.title("🔐 Login")
   username = st.text_input("Username")
   password = st.text_input("Password", type="password")
   if st.button("Login"):
       if username == "Rohith" and password == "Rohith@1234":
           st.session_state["logged_in"] = True
       else:
           st.error("Invalid credentials")
# Session check
if "logged_in" not in st.session_state:
   st.session_state["logged_in"] = False
if not st.session_state["logged_in"]:
   login()
   st.stop()

import json
import os
from datetime import datetime, timedelta

import pandas as pd
import numpy as np
import plotly.graph_objects as go
#import streamlit as st

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

# ──────────────────────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────────────────────

CONTAINERS = {
    "20FT": {"max_cbm": 33,  "max_weight": 28000, "l": 5.90,  "w": 2.35, "h": 2.39},
    "40FT": {"max_cbm": 67,  "max_weight": 26000, "l": 12.03, "w": 2.35, "h": 2.39},
    "40HQ": {"max_cbm": 76,  "max_weight": 26000, "l": 12.03, "w": 2.35, "h": 2.72},
}

DEFAULT_TRANSIT_DAYS = 14
CONSOLIDATION_WINDOW = 5

PRIORITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}

BOX_COLORS = [
    "#E74C3C", "#3498DB", "#2ECC71", "#F39C12", "#9B59B6",
    "#1ABC9C", "#E67E22", "#34495E", "#E91E63", "#00BCD4",
]

ANTHROPIC_MODEL = "claude-sonnet-4-20250514"

# Path to the bundled transit days file



TRANSIT_FILE = os.path.join(os.path.dirname(__file__), "data", "Transit_Days.xlsx")

# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────

def excel_serial_to_date(val) -> str:
    """Convert Excel date serial number or string to YYYY-MM-DD."""
    try:
        v = float(val)
        if v > 40000:  # looks like an Excel serial
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=int(v))).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        pass
    try:
        return pd.to_datetime(val).strftime("%Y-%m-%d")
    except Exception:
        return str(val)


def add_days(date_str: str, n: int) -> str:
    d = datetime.strptime(date_str[:10], "%Y-%m-%d") + timedelta(days=n)
    return d.strftime("%Y-%m-%d")


def date_gap(a: str, b: str) -> int:
    da = datetime.strptime(a[:10], "%Y-%m-%d")
    db = datetime.strptime(b[:10], "%Y-%m-%d")
    return abs((da - db).days)


def select_container(total_cbm: float, total_weight: float) -> str:
    if total_cbm <= 33 and total_weight <= 28000:
        return "20FT"
    elif total_cbm <= 67 and total_weight <= 26000:
        return "40FT"
    else:
        return "40HQ"


def loading_suggestion(avg_weight_kg: float, stuffing_pct: float) -> str:
    parts = []
    if avg_weight_kg > 700:
        parts.append("Heavy cargo — load heaviest cartons at bottom-centre and distribute weight evenly over axles.")
    elif avg_weight_kg > 400:
        parts.append("Medium-weight cargo — distribute evenly side to side to prevent tipping.")
    else:
        parts.append("Light cargo — maximise height utilisation; use void fillers to prevent shifting.")

    if stuffing_pct < 50:
        parts.append("Utilisation critically low (< 50 %) — strongly consider LCL consolidation.")
    elif stuffing_pct < 70:
        parts.append("Utilisation below 70 % — review consolidation opportunities.")
    elif stuffing_pct >= 85:
        parts.append("High fill (≥ 85 %) — ensure securing straps are in place; verify door-sealing clearance.")
    return " ".join(parts)


def load_transit_map(uploaded_file=None) -> dict:
    """Load lane → transit_days from uploaded file or bundled Transit_Days.xlsx."""
    lane_map = {}
    try:
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
        elif os.path.exists(TRANSIT_FILE):
            df = pd.read_excel(TRANSIT_FILE)
        else:
            return lane_map
        df.columns = [c.strip() for c in df.columns]
        for _, row in df.iterrows():
            key = f"{str(row['From_Port']).strip()} -> {str(row['To_Port']).strip()}"
            lane_map[key] = int(row["Transit_Days"])
    except Exception as e:
        st.warning(f"Could not load transit days: {e}")
    return lane_map


# ──────────────────────────────────────────────────────────────
# OPTIMISATION ENGINE
# ──────────────────────────────────────────────────────────────

def run_optimization(df: pd.DataFrame, lane_map: dict) -> pd.DataFrame:
    df.columns = [c.strip() for c in df.columns]

    required = ["BoxID", "Customer", "From_Port", "To_Port",
                "Shipment_Type", "Length_cm", "Width_cm", "Height_cm",
                "Weight_kg", "Requested_Handover_Date"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {missing}")

    df = df.copy()

    # Handle Excel serial dates
    df["Requested_Handover_Date"] = df["Requested_Handover_Date"].apply(excel_serial_to_date)

    if "Required_Arrival_Date" in df.columns:
        df["Required_Arrival_Date"] = df["Required_Arrival_Date"].apply(excel_serial_to_date)
    else:
        df["Required_Arrival_Date"] = None

    if "Priority" not in df.columns:
        df["Priority"] = "Medium"

    df["CBM"] = (
        df["Length_cm"].astype(float)
        * df["Width_cm"].astype(float)
        * df["Height_cm"].astype(float)
    ) / 1_000_000

    df["Lane"] = df["From_Port"].str.strip() + " -> " + df["To_Port"].str.strip()
    df["Shipment_Type"] = df["Shipment_Type"].str.upper().str.strip()
    df["Priority_Rank"] = df["Priority"].map(PRIORITY_ORDER).fillna(2)

    # Sort by priority within lane/type/date
    df = df.sort_values(
        ["Lane", "Shipment_Type", "Priority_Rank", "Requested_Handover_Date"]
    ).reset_index(drop=True)

    results = []
    counter = 1

    for (lane, stype), grp in df.groupby(["Lane", "Shipment_Type"]):
        grp = grp.sort_values(["Priority_Rank", "Requested_Handover_Date"]).reset_index(drop=True)
        transit = lane_map.get(lane, DEFAULT_TRANSIT_DAYS)

        # ── FCL ──
        if stype == "FCL":
            for customer, cdf in grp.groupby("Customer"):
                total_cbm    = cdf["CBM"].sum()
                total_weight = cdf["Weight_kg"].astype(float).sum()
                box_ids      = cdf["BoxID"].tolist()
                earliest     = cdf["Requested_Handover_Date"].min()
                priority     = cdf["Priority"].iloc[0] if "Priority" in cdf.columns else "Medium"

                ctype    = select_container(total_cbm, total_weight)
                max_cbm  = CONTAINERS[ctype]["max_cbm"]
                rate_pct = round(min(total_cbm / max_cbm * 100, 100), 1)

                if rate_pct >= 85:
                    cutoff = 1
                elif rate_pct >= 60:
                    cutoff = 3
                else:
                    cutoff = 5

                dep_date = add_days(earliest, cutoff)
                arr_date = add_days(dep_date, transit)
                avg_wt   = total_weight / max(len(box_ids), 1)

                # Check arrival vs required
                req_arr  = cdf["Required_Arrival_Date"].dropna().min() if "Required_Arrival_Date" in cdf.columns else None
                arrival_flag = ""
                if req_arr and req_arr != "NaT":
                    try:
                        if arr_date > req_arr:
                            arrival_flag = f"⚠️ Late by {date_gap(arr_date, req_arr)} days"
                        else:
                            arrival_flag = f"✅ On time ({date_gap(arr_date, req_arr)} days buffer)"
                    except Exception:
                        pass

                results.append({
                    "Container_ID":         f"CONT_{counter:04d}",
                    "Shipment_Type":        stype,
                    "Customer":             customer,
                    "Lane":                 lane,
                    "Container_Type":       ctype,
                    "Priority":             priority,
                    "Box_Count":            len(box_ids),
                    "BoxIDs":               ", ".join(map(str, box_ids)),
                    "Total_CBM":            round(total_cbm, 2),
                    "Total_Weight_KG":      round(total_weight, 0),
                    "Stuffing_Rate_%":      f"{rate_pct} %",
                    "Stuffing_Rate_num":    rate_pct,
                    "Suggested_Departure":  dep_date,
                    "Estimated_Arrival":    arr_date,
                    "Required_Arrival":     req_arr or "—",
                    "Arrival_Status":       arrival_flag,
                    "Transit_Days":         transit,
                    "Loading_Suggestion":   loading_suggestion(avg_wt, rate_pct),
                    "Optimization_Insight": (
                        "Low utilisation — consider LCL consolidation."
                        if rate_pct < 50 else
                        "FCL container utilisation acceptable."
                    ),
                })
                counter += 1

        # ── LCL ──
        else:
            cur_boxes  = []
            cur_cbm    = 0.0
            cur_weight = 0.0
            earliest   = None
            cur_priority = "Low"
            cur_req_arr  = None

            def flush_lcl():
                nonlocal cur_boxes, cur_cbm, cur_weight, earliest, counter, cur_priority, cur_req_arr
                if not cur_boxes:
                    return
                ctype    = select_container(cur_cbm, cur_weight)
                max_cbm  = CONTAINERS[ctype]["max_cbm"]
                rate_pct = round(min(cur_cbm / max_cbm * 100, 100), 1)
                cutoff   = 1 if rate_pct >= 85 else 3
                dep_date = add_days(earliest, cutoff)
                arr_date = add_days(dep_date, transit)
                avg_wt   = cur_weight / max(len(cur_boxes), 1)

                arrival_flag = ""
                if cur_req_arr:
                    try:
                        if arr_date > cur_req_arr:
                            arrival_flag = f"⚠️ Late by {date_gap(arr_date, cur_req_arr)} days"
                        else:
                            arrival_flag = f"✅ On time ({date_gap(arr_date, cur_req_arr)} days buffer)"
                    except Exception:
                        pass

                results.append({
                    "Container_ID":         f"CONT_{counter:04d}",
                    "Shipment_Type":        stype,
                    "Customer":             "MULTIPLE",
                    "Lane":                 lane,
                    "Container_Type":       ctype,
                    "Priority":             cur_priority,
                    "Box_Count":            len(cur_boxes),
                    "BoxIDs":               ", ".join(map(str, cur_boxes)),
                    "Total_CBM":            round(cur_cbm, 2),
                    "Total_Weight_KG":      round(cur_weight, 0),
                    "Stuffing_Rate_%":      f"{rate_pct} %",
                    "Stuffing_Rate_num":    rate_pct,
                    "Suggested_Departure":  dep_date,
                    "Estimated_Arrival":    arr_date,
                    "Required_Arrival":     cur_req_arr or "—",
                    "Arrival_Status":       arrival_flag,
                    "Transit_Days":         transit,
                    "Loading_Suggestion":   loading_suggestion(avg_wt, rate_pct),
                    "Optimization_Insight": (
                        "Low fill LCL batch — widen consolidation window?"
                        if rate_pct < 50 else
                        "LCL consolidated shipment."
                    ),
                })
                counter    += 1
                cur_boxes   = []
                cur_cbm     = 0.0
                cur_weight  = 0.0
                earliest    = None
                cur_priority = "Low"
                cur_req_arr  = None

            for _, row in grp.iterrows():
                proj_cbm    = cur_cbm + row["CBM"]
                proj_weight = cur_weight + float(row["Weight_kg"])
                tmp_type    = select_container(proj_cbm, proj_weight)
                max_cbm     = CONTAINERS[tmp_type]["max_cbm"]
                gap         = date_gap(row["Requested_Handover_Date"], earliest) if earliest else 0

                if (proj_cbm > max_cbm or gap > CONSOLIDATION_WINDOW) and cur_boxes:
                    flush_lcl()

                if earliest is None:
                    earliest = row["Requested_Handover_Date"]

                cur_boxes.append(row["BoxID"])
                cur_cbm    += row["CBM"]
                cur_weight += float(row["Weight_kg"])

                row_priority = row.get("Priority", "Medium")
                if PRIORITY_ORDER.get(row_priority, 2) < PRIORITY_ORDER.get(cur_priority, 3):
                    cur_priority = row_priority

                row_req = row.get("Required_Arrival_Date")
                if row_req and str(row_req) not in ("None", "nan", "NaT", "—"):
                    if cur_req_arr is None or row_req < cur_req_arr:
                        cur_req_arr = row_req

            flush_lcl()

    out = pd.DataFrame(results)
    if not out.empty:
        out = out.sort_values(["Lane", "Shipment_Type", "Suggested_Departure"]).reset_index(drop=True)
    return out


# ──────────────────────────────────────────────────────────────
# 3-D VISUALISATION (Plotly)
# ──────────────────────────────────────────────────────────────

def _box_mesh(ox, oy, oz, l, w, h, color, name=""):
    x0, x1 = ox, ox + l
    y0, y1 = oy, oy + h
    z0, z1 = oz, oz + w
    vx = [x0, x1, x1, x0, x0, x1, x1, x0]
    vy = [y0, y0, y1, y1, y0, y0, y1, y1]
    vz = [z0, z0, z0, z0, z1, z1, z1, z1]
    i  = [0, 0, 0, 1, 1, 2, 4, 4, 4, 5, 5, 6]
    j  = [1, 2, 4, 2, 5, 3, 5, 6, 0, 6, 1, 7]
    k  = [2, 3, 5, 6, 6, 7, 6, 7, 3, 7, 2, 3]
    return go.Mesh3d(
        x=vx, y=vy, z=vz,
        i=i, j=j, k=k,
        color=color, opacity=0.85,
        name=name,
        hoverinfo="name",
        flatshading=True,
    )


def build_3d_figure(result_row: dict) -> go.Figure:
    ctype = result_row["Container_Type"]
    spec  = CONTAINERS[ctype]
    L, W, H = spec["l"], spec["w"], spec["h"]

    box_count  = int(result_row["Box_Count"])
    fill_frac  = min(result_row["Stuffing_Rate_num"] / 100.0, 0.98)
    total_vol  = spec["max_cbm"] * fill_frac
    per_box    = max(total_vol / max(box_count, 1), 0.001)
    side       = per_box ** (1 / 3) * 0.88

    traces = []

    # Container outline (wireframe via scatter)
    edges = [
        [(0,0,0),(L,0,0)],[(L,0,0),(L,W,0)],[(L,W,0),(0,W,0)],[(0,W,0),(0,0,0)],
        [(0,0,H),(L,0,H)],[(L,0,H),(L,W,H)],[(L,W,H),(0,W,H)],[(0,W,H),(0,0,H)],
        [(0,0,0),(0,0,H)],[(L,0,0),(L,0,H)],[(L,W,0),(L,W,H)],[(0,W,0),(0,W,H)],
    ]
    for (x0,z0,y0),(x1,z1,y1) in edges:
        traces.append(go.Scatter3d(
            x=[x0,x1,None], y=[y0,y1,None], z=[z0,z1,None],
            mode="lines",
            line=dict(color="#4488cc", width=2),
            showlegend=False, hoverinfo="skip",
        ))

    col_w  = side * 1.12
    rows_z = max(1, int(W / col_w))
    rows_x = max(1, int(L / (side * 1.12)))

    placed = 0
    y = side * 0.05
    while placed < min(box_count, 80) and y + side * 0.8 <= H + 0.01:
        x = side * 0.05
        for _ in range(rows_x):
            if placed >= min(box_count, 80):
                break
            z = side * 0.05
            for _ in range(rows_z):
                if placed >= min(box_count, 80):
                    break
                if x + side <= L + 0.01 and z + side * 0.9 <= W + 0.01:
                    color = BOX_COLORS[placed % len(BOX_COLORS)]
                    traces.append(_box_mesh(x, y, z, side, side * 0.9, side * 0.8,
                                            color=color, name=f"Box {placed+1}"))
                    placed += 1
                z += side * 0.9 + 0.02
            x += side + 0.02
        y += side * 0.8 + 0.02

    fig = go.Figure(data=traces)
    fig.update_layout(
        title=dict(
            text=(
                f"<b>{result_row['Container_ID']}</b>  ·  {ctype}  ·  {result_row['Lane']}<br>"
                f"Stuffing: {result_row['Stuffing_Rate_%']}  |  "
                f"{result_row['Total_CBM']} CBM  |  "
                f"{int(result_row['Total_Weight_KG']):,} kg  |  "
                f"{box_count} boxes"
            ),
            font=dict(size=13),
        ),
        scene=dict(
            xaxis=dict(title="Length (m)", backgroundcolor="#0e1117", gridcolor="#333"),
            yaxis=dict(title="Height (m)", backgroundcolor="#0e1117", gridcolor="#333"),
            zaxis=dict(title="Width (m)",  backgroundcolor="#0e1117", gridcolor="#333"),
            bgcolor="#0e1117",
        ),
        paper_bgcolor="#0e1117",
        font_color="#cdd6f4",
        height=560,
        margin=dict(l=0, r=0, t=80, b=0),
    )
    return fig


# ──────────────────────────────────────────────────────────────
# CHATBOT
# ──────────────────────────────────────────────────────────────

def ask_claude(question: str, data_context: str, history: list, api_key: str) -> str:
    if not ANTHROPIC_AVAILABLE:
        return "Install anthropic SDK: `pip install anthropic`"
    try:
        client = anthropic.Anthropic(api_key=api_key)
        history.append({"role": "user", "content": question})
        response = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=1000,
            system=data_context,
            messages=history,
        )
        answer = response.content[0].text
        history.append({"role": "assistant", "content": answer})
        return answer
    except Exception as exc:
        return f"API error: {exc}"


# ──────────────────────────────────────────────────────────────
# STREAMLIT APP
# ──────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Container Cargo Load Planner Pro",
        page_icon="🚢",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── Custom CSS ──
    st.markdown("""
    <style>
    .metric-card {
        background: #1e2a3a;
        border-radius: 10px;
        padding: 14px 18px;
        text-align: center;
        border: 1px solid #2d3f55;
    }
    .metric-value { font-size: 2rem; font-weight: 700; color: #89b4fa; }
    .metric-label { font-size: 0.78rem; color: #6c7086; margin-top: 2px; }
    .stTabs [data-baseweb="tab"] { font-size: 15px; font-weight: 600; }
    .priority-critical { color: #f38ba8; font-weight: 600; }
    .priority-high     { color: #fab387; font-weight: 600; }
    .stAlert { border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

    # ── Session State ──
    for key, default in [
        ("raw_df",      None),
        ("results",     None),
        ("lane_map",    {}),
        ("chat_history", []),
        ("data_context", ""),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    st.title("🚢 Container Cargo Load Planner Pro")
    st.caption("Streamlit Edition — upload cargo data, optimise container loads, visualise in 3D, and chat with an AI logistics assistant.")

    # ── SIDEBAR ──
    with st.sidebar:
        st.header("⚙️ Configuration")

        # API Key
        api_key = st.text_input(
            "Anthropic API Key",
            type="password",
            value=os.environ.get("ANTHROPIC_API_KEY", ""),
            help="Required only for the AI Chatbot tab. Get yours at console.anthropic.com",
        )
        st.divider()

        # Upload cargo file
        st.subheader("📂 Upload Cargo Data")
        cargo_file = st.file_uploader(
            "Cargo file (.xlsx / .xls / .csv)",
            type=["xlsx", "xls", "csv"],
            help="Must contain: BoxID, Customer, From_Port, To_Port, Shipment_Type, "
                 "Length_cm, Width_cm, Height_cm, Weight_kg, Requested_Handover_Date. "
                 "Optional: Priority, Required_Arrival_Date",
        )

        st.subheader("🗺️ Transit Days")
        transit_file = st.file_uploader(
            "Transit days file (.xlsx) — optional",
            type=["xlsx"],
            help="Columns: From_Port, To_Port, Transit_Days. "
                 "If not uploaded, the bundled Transit_Days.xlsx is used.",
        )
        if st.button("🔄 Reload Transit Days"):
            st.session_state.lane_map = load_transit_map(transit_file)
            st.success(f"Loaded {len(st.session_state.lane_map)} lanes.")

        st.divider()

        # Manual lane override
        st.subheader("✏️ Override a Lane")
        c1, c2 = st.columns(2)
        from_port = c1.text_input("From", key="from_port").upper().strip()
        to_port   = c2.text_input("To",   key="to_port").upper().strip()
        transit_d = st.number_input("Transit days", 1, 120, 14, key="transit_d")
        if st.button("Add / Update Lane"):
            if from_port and to_port:
                key_ = f"{from_port} -> {to_port}"
                st.session_state.lane_map[key_] = int(transit_d)
                st.success(f"Set {key_} = {transit_d} days")
            else:
                st.warning("Enter both ports.")

        if st.session_state.lane_map:
            st.caption(f"**{len(st.session_state.lane_map)} lanes configured**")
            with st.expander("View all lanes"):
                for k, v in sorted(st.session_state.lane_map.items()):
                    st.text(f"{k}  →  {v} days")

        st.divider()
        st.caption("v2.0 · Streamlit Edition · Powered by Anthropic Claude")

    # ── Load defaults on first run ──
    if not st.session_state.lane_map:
        st.session_state.lane_map = load_transit_map(None)

    # ── Auto-load cargo file ──
    if cargo_file:
        try:
            ext = cargo_file.name.split(".")[-1].lower()
            if ext == "csv":
                st.session_state.raw_df = pd.read_csv(cargo_file)
            else:
                st.session_state.raw_df = pd.read_excel(cargo_file)
            st.sidebar.success(f"✓ {cargo_file.name} — {len(st.session_state.raw_df)} rows")
        except Exception as e:
            st.sidebar.error(f"Load error: {e}")

    # ─────────────────────────────────────────────────────────
    # TABS
    # ─────────────────────────────────────────────────────────
    tab_import, tab_results, tab_visual, tab_chat = st.tabs([
        "① Import & Optimize", "② Results Table", "③ 3D Visualizer", "④ AI Chatbot"
    ])

    # ═══════════════════════════════
    # TAB 1 — IMPORT & OPTIMIZE
    # ═══════════════════════════════
    with tab_import:
        col_left, col_right = st.columns([2, 1])

        with col_left:
            st.subheader("Loaded Data Preview")
            if st.session_state.raw_df is not None:
                df = st.session_state.raw_df
                st.info(f"**{len(df)} boxes** across **{df['From_Port'].nunique() if 'From_Port' in df.columns else '?'} origin ports**")
                # Show dates as-is then convert for display
                display_df = df.copy()
                for col in ["Requested_Handover_Date", "Required_Arrival_Date"]:
                    if col in display_df.columns:
                        display_df[col] = display_df[col].apply(excel_serial_to_date)
                st.dataframe(display_df, use_container_width=True, height=320)
            else:
                st.info("Upload a cargo file in the sidebar, or use the demo data below.")
                if st.button("✨ Load Demo Data"):
                    import random
                    random.seed(42)
                    ports  = [("INMAA","DEHAM"),("INBLR","USLAX"),("THBKK","CNHKG"),("SGSIN","NLRTM"),("MYTPP","DEHAM")]
                    custs  = ["AlphaShip","BetaCargo","GammaTrade","DeltaEx"]
                    stypes = ["FCL","FCL","LCL","LCL","FCL"]
                    prios  = ["Critical","High","Medium","Low"]
                    base   = datetime(2026, 6, 1)
                    rows   = []
                    for i in range(1, 41):
                        fr, to = ports[i % len(ports)]
                        d  = base + timedelta(days=random.randint(0, 20))
                        ra = d + timedelta(days=random.randint(30, 50))
                        rows.append({
                            "BoxID":  f"BOX{i:04d}",
                            "Customer":             custs[i % len(custs)],
                            "From_Port":            fr,
                            "To_Port":              to,
                            "Shipment_Type":        stypes[i % len(stypes)],
                            "Length_cm":            random.randint(80, 250),
                            "Width_cm":             random.randint(60, 150),
                            "Height_cm":            random.randint(60, 180),
                            "Weight_kg":            random.randint(100, 1200),
                            "Requested_Handover_Date": d.strftime("%Y-%m-%d"),
                            "Priority":             prios[i % len(prios)],
                            "Required_Arrival_Date": ra.strftime("%Y-%m-%d"),
                        })
                    st.session_state.raw_df = pd.DataFrame(rows)
                    st.rerun()

        with col_right:
            st.subheader("Run Optimization")
            st.markdown(f"""
            **Transit Days Source**  
            {'✅ ' + str(len(st.session_state.lane_map)) + ' lanes loaded' if st.session_state.lane_map else '⚠️ No lane data — using default ' + str(DEFAULT_TRANSIT_DAYS) + ' days'}
            
            **Consolidation Window**  
            {CONSOLIDATION_WINDOW} days (LCL batching)

            **Priority handling**  
            Critical → High → Medium → Low sort within each lane
            """)

            if st.button("▶ Run Optimization", type="primary", use_container_width=True,
                         disabled=st.session_state.raw_df is None):
                with st.spinner("Optimising container loads…"):
                    try:
                        results = run_optimization(
                            st.session_state.raw_df.copy(),
                            st.session_state.lane_map,
                        )
                        st.session_state.results = results
                        sample = results.head(30).to_dict(orient="records")
                        st.session_state.data_context = (
                            "You are an expert logistics and container shipping assistant. "
                            "Here is the current container optimization output (up to 30 rows):\n"
                            + json.dumps(sample, default=str, indent=2)
                            + "\n\nAnswer concisely based on this data. "
                            "Format numbers cleanly. Always reference Container_IDs when relevant."
                        )
                        st.session_state.chat_history = []
                        st.success(f"✅ Optimisation complete — **{len(results)} containers** planned.")
                    except Exception as e:
                        st.error(f"Optimisation error: {e}")

    # ═══════════════════════════════
    # TAB 2 — RESULTS TABLE
    # ═══════════════════════════════
    with tab_results:
        if st.session_state.results is None:
            st.info("Run the optimisation first (Tab ①).")
        else:
            r = st.session_state.results

            # Metric cards
            cols = st.columns(7)
            metrics = [
                ("Containers",    str(len(r))),
                ("Total CBM",     f"{r['Total_CBM'].sum():.1f}"),
                ("Total Weight",  f"{r['Total_Weight_KG'].sum()/1000:.1f} t"),
                ("Avg Stuffing",  f"{r['Stuffing_Rate_num'].mean():.1f} %"),
                ("FCL",           str((r['Shipment_Type']=='FCL').sum())),
                ("LCL",           str((r['Shipment_Type']=='LCL').sum())),
                ("Lanes",         str(r['Lane'].nunique())),
            ]
            for col, (lbl, val) in zip(cols, metrics):
                col.markdown(
                    f'<div class="metric-card">'
                    f'<div class="metric-value">{val}</div>'
                    f'<div class="metric-label">{lbl}</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

            st.divider()

            # Filters
            fc1, fc2, fc3 = st.columns(3)
            ftype = fc1.multiselect("Shipment Type", options=r["Shipment_Type"].unique().tolist(),
                                     default=r["Shipment_Type"].unique().tolist())
            flane = fc2.multiselect("Lane", options=r["Lane"].unique().tolist(),
                                     default=r["Lane"].unique().tolist())
            fprio = fc3.multiselect("Priority", options=r["Priority"].unique().tolist(),
                                     default=r["Priority"].unique().tolist())

            filtered = r[
                r["Shipment_Type"].isin(ftype) &
                r["Lane"].isin(flane) &
                r["Priority"].isin(fprio)
            ]

            # Highlight low-stuffing rows
            def highlight_row(row):
                # Use the original filtered dataframe because Stuffing_Rate_num
                # is not included in show_df displayed to the user.
                try:
                    stuffing_rate = filtered.loc[row.name, "Stuffing_Rate_num"]
                except Exception:
                    stuffing_rate = None

                if stuffing_rate is not None and pd.notna(stuffing_rate) and stuffing_rate < 50:
                    return ["background-color: #3a1a1a"] * len(row)

                if row.get("Priority") == "Critical":
                    return ["background-color: #1a2a1a"] * len(row)

                return [""] * len(row)

            display_cols = [
                "Container_ID", "Shipment_Type", "Lane", "Customer", "Priority",
                "Container_Type", "Box_Count", "Total_CBM", "Total_Weight_KG",
                "Stuffing_Rate_%", "Suggested_Departure", "Estimated_Arrival",
                "Required_Arrival", "Arrival_Status", "Optimization_Insight",
            ]
            show_df = filtered[[c for c in display_cols if c in filtered.columns]]
            st.dataframe(
                show_df.style.apply(highlight_row, axis=1),
                use_container_width=True,
                height=420,
            )

            # Export
            st.divider()
            export_df = r.drop(columns=["Stuffing_Rate_num"], errors="ignore")
            import io
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                export_df.to_excel(writer, index=False, sheet_name="Optimized_Output")
            st.download_button(
                "💾 Download Results as Excel",
                data=buf.getvalue(),
                file_name="container_optimization_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ═══════════════════════════════
    # TAB 3 — 3D VISUALIZER
    # ═══════════════════════════════
    with tab_visual:
        if st.session_state.results is None:
            st.info("Run the optimisation first (Tab ①).")
        else:
            r = st.session_state.results
            container_ids = r["Container_ID"].tolist()
            selected_id = st.selectbox(
                "Select Container to Visualise",
                options=container_ids,
                format_func=lambda cid: (
                    f"{cid}  —  {r[r['Container_ID']==cid]['Lane'].values[0]}  "
                    f"({r[r['Container_ID']==cid]['Shipment_Type'].values[0]})  "
                    f"Stuffing: {r[r['Container_ID']==cid]['Stuffing_Rate_%'].values[0]}"
                ),
            )
            row_data = r[r["Container_ID"] == selected_id].iloc[0].to_dict()

            col_a, col_b = st.columns([3, 1])
            with col_b:
                st.markdown("**Container Details**")
                for field in ["Container_Type", "Box_Count", "Total_CBM",
                              "Total_Weight_KG", "Stuffing_Rate_%",
                              "Suggested_Departure", "Estimated_Arrival",
                              "Required_Arrival", "Arrival_Status", "Priority"]:
                    if field in row_data:
                        st.metric(field.replace("_", " "), row_data[field])

                st.markdown("**Loading Suggestion**")
                st.info(row_data.get("Loading_Suggestion", ""))

            with col_a:
                fig = build_3d_figure(row_data)
                st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════
    # TAB 4 — AI CHATBOT
    # ═══════════════════════════════
    with tab_chat:
        st.subheader("🤖 AI Logistics Assistant")

        if not api_key:
            st.warning("Enter your Anthropic API key in the sidebar to use the chatbot.")
        elif st.session_state.results is None:
            st.warning("Run the optimisation first — the chatbot needs data to answer questions.")
        else:
            st.caption("Ask anything about your container plan. The assistant has access to the optimisation results.")

            # Quick chips
            chips = [
                "Which containers have low stuffing rates?",
                "Summarise FCL vs LCL",
                "Which is the busiest lane?",
                "Any containers at risk of late arrival?",
                "What are the Critical priority containers?",
                "Average stuffing rate by lane?",
            ]
            st.markdown("**Quick questions:**")
            chip_cols = st.columns(3)
            for i, chip in enumerate(chips):
                if chip_cols[i % 3].button(chip, key=f"chip_{i}"):
                    with st.spinner("Thinking…"):
                        ans = ask_claude(chip, st.session_state.data_context,
                                         st.session_state.chat_history, api_key)
                    st.success("Response added to chat history below.")
                    st.rerun()

            st.divider()

            # Chat history
            for msg in st.session_state.chat_history:
                role = "🧑 You" if msg["role"] == "user" else "🤖 Assistant"
                with st.chat_message(msg["role"]):
                    st.markdown(f"**{role}:** {msg['content']}")

            # Input
            user_input = st.chat_input("Ask a logistics question…")
            if user_input:
                with st.chat_message("user"):
                    st.markdown(f"**🧑 You:** {user_input}")
                with st.chat_message("assistant"):
                    with st.spinner("Thinking…"):
                        answer = ask_claude(
                            user_input,
                            st.session_state.data_context,
                            st.session_state.chat_history,
                            api_key,
                        )
                    st.markdown(f"**🤖 Assistant:** {answer}")


if __name__ == "__main__":
    main()