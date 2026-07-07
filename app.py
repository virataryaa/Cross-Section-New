"""
Cross Section Monitor
Momentum (Vol-Adj) vs Roll Yield scatter — cross-sectional ranking across 31 commodities.
"""

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path

try:
    import win32com.client as _win32
    _OUTLOOK_OK = True
except ImportError:
    _OUTLOOK_OK = False

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Cross Section Monitor", layout="wide")
st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f5f5f7;}
h1,h2,h3{color:#1d1d1f;}
</style>
""", unsafe_allow_html=True)

st.title("Cross Section Monitor")

# ── Constants ─────────────────────────────────────────────────────────────────
HERE = Path(__file__).parent

# (name, spot_ric, fut_ric, is_soft, gsci_ric)
# gsci_ric = None  →  no GSCI index; spot_ric (c2) is used for momentum instead
COMMODITIES = [
    ("Brent Oil",    "LCOc1",  "LCOc13",  False, ".SPGSBRP"),
    ("Cocoa",        "CCc2",   "CCc7",    True,  ".SPGSCCP"),
    ("Coffee",       "KCc2",   "KCc7",    True,  ".SPGSKCP"),
    ("Corn",         "Cc1",    "Cc6",     False, ".SPGSCNP"),
    ("Cotton",       "CTc2",   "CTc7",    True,  ".SPGSCTP"),
    ("Gas Oil",      "LGOc1",  "LGOc13",  False, ".SPGSGOP"),
    ("Gold",         "GCc1",   "GCc8",    False, ".SPGSGCP"),
    ("Heating Oil",  "HOc1",   "HOc13",   False, ".SPGSHOP"),
    ("HG Copper",    "HGc1",   "HGc13",   False, ".SPGSICP"),
    ("LDN Cocoa",    "LCCc2",  "LCCc7",   True,  None),
    ("Lean Hog",     "LHc1",   "LHc9",    False, ".SPGSLHP"),
    ("Live Cattle",  "LCc1",   "LCc7",    False, ".SPGSLCP"),
    ("Natural Gas",  "NGc1",   "NGc13",   False, ".SPGSNGP"),
    ("Silver",       "SIc1",   "SIc8",    False, ".SPGSSIP"),
    ("Soy Meal",     "SMc1",   "SMc9",    False, ".SPGSSMP"),
    ("Soy Bean",     "Sc1",    "Sc8",     False, ".SPGSSOP"),
    ("Soy Oil",      "BOc1",   "BOc9",    False, ".SPGSBOP"),
    ("Sugar",        "SBc1",   "SBc5",    True,  ".SPGSSBP"),
    ("Gasoline",     "RBc1",   "RBc13",   False, ".SPGSHUP"),
    ("Wheat",        "Wc1",    "Wc6",     False, ".SPGSWHP"),
    ("Wheat (KCB)",  "KWc1",   "KWc6",    False, ".SPGSKWP"),
    ("WTI Crude",    "CLc1",   "CLc13",   False, ".SPGSCLP"),
    ("White Sugar",  "LSUc1",  "LSUc6",   True,  None),
    ("Robusta",      "LRCc2",  "LRCc8",   True,  None),
    ("Canola",       "RSc2",   "RSc7",    False, None),
    ("Zinc",         "MZNc2",  "MZNc13",  False, ".SPGSIZP"),
    ("Aluminium",    "MALc2",  "MALc13",  False, ".SPGSIAP"),
    ("Lead",         "MPBc2",  "MPBc13",  False, ".SPGSILP"),
    ("Copper",       "MCUc2",  "MCUc13",  False, ".SPGSICP"),
    ("Nickel",       "MNIc2",  "MNIc13",  False, ".SPGSIKP"),
    ("Tin",          "MSNc2",  "MSNc13",  False, ".SPGSIS"),
    ("Orange Juice", "OJc2",   "OJc7",    True,  None),
]

NAME2SPOT   = {r[0]: r[1] for r in COMMODITIES}
NAME2FUT    = {r[0]: r[2] for r in COMMODITIES}
NAME2GSCI   = {r[0]: r[4] for r in COMMODITIES}
ALL_NAMES   = [r[0] for r in COMMODITIES]
SOFT_NAMES  = [r[0] for r in COMMODITIES if r[3]]

COMMODITY_CATEGORY: dict[str, str] = {
    # Softs
    "Coffee": "Softs", "Cocoa": "Softs", "LDN Cocoa": "Softs", "Cotton": "Softs",
    "Sugar": "Softs", "White Sugar": "Softs", "Robusta": "Softs", "Orange Juice": "Softs",
    # Energy
    "Brent Oil": "Energy", "WTI Crude": "Energy", "Gas Oil": "Energy",
    "Heating Oil": "Energy", "Gasoline": "Energy", "Natural Gas": "Energy",
    # Precious Metals
    "Gold": "Precious Metals", "Silver": "Precious Metals",
    # Base Metals (COMEX + LME merged)
    "HG Copper": "Base Metals", "Zinc": "Base Metals", "Aluminium": "Base Metals",
    "Lead": "Base Metals", "Copper": "Base Metals", "Nickel": "Base Metals", "Tin": "Base Metals",
    # Grains & Oilseeds
    "Corn": "Grains & Oilseeds", "Wheat": "Grains & Oilseeds", "Wheat (KCB)": "Grains & Oilseeds",
    "Soy Bean": "Grains & Oilseeds", "Soy Meal": "Grains & Oilseeds",
    "Soy Oil": "Grains & Oilseeds", "Canola": "Grains & Oilseeds",
    # Livestock
    "Lean Hog": "Livestock", "Live Cattle": "Livestock",
}

CATEGORY_COLOR: dict[str, str] = {
    "Softs":             "#1565C0",   # bold vivid blue
    "Energy":            "#E8976A",   # pastel orange
    "Precious Metals":   "#C9A028",   # muted gold
    "Base Metals":       "#7A9BB5",   # pastel steel blue
    "Grains & Oilseeds": "#7AAF82",   # pastel sage green
    "Livestock":         "#A88BC4",   # pastel lavender
}

# Per-commodity color for time series — softs each get a unique vivid color
COMMODITY_TS_COLOR: dict[str, str] = {
    "Coffee":       "#1565C0",  # vivid blue
    "Cocoa":        "#D62728",  # vivid red
    "Cotton":       "#9467BD",  # vivid purple
    "LDN Cocoa":    "#17BECF",  # vivid cyan
    "Sugar":        "#E67E22",  # vivid orange
    "White Sugar":  "#F5C518",  # vivid amber
    "Robusta":      "#2CA02C",  # vivid green
    "Orange Juice": "#E377C2",  # vivid pink
}

IS_SOFT = {name: (COMMODITY_CATEGORY[name] == "Softs") for name in ALL_NAMES}

AG_CATEGORIES = {"Softs", "Grains & Oilseeds", "Livestock"}
AG_NAMES      = [n for n in ALL_NAMES if COMMODITY_CATEGORY[n] in AG_CATEGORIES]

MAX_ANIM_DAYS = 130   # ~6 months; blocks Previous Year from blowing up the browser

MOM_DAYS    = {"3 Months": 63, "6 Months": 126, "12 Months": 252}
VOL_DAYS    = {"20d": 20, "60d": 60, "120d": 120}
PREV_DAYS   = {"Previous Day": 1, "Previous Week": 5, "Previous Month": 20,
               "Previous Quarter": 60, "Previous Year": 250}

N           = len(ALL_NAMES)

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def load_data() -> pd.DataFrame:
    p = HERE / "prices.parquet"
    if not p.exists():
        return pd.DataFrame()
    df = pd.read_parquet(p)
    df["Date"] = pd.to_datetime(df["Date"])
    return df


@st.cache_data(ttl=3600)
def build_metrics(df_raw: pd.DataFrame, mom_days: int, vol_days: int) -> pd.DataFrame:
    """
    Returns daily DataFrame with columns:
      commodity, Date, spot, future, return_N, vol, mom_vol_adj, roll_yield
    """
    if df_raw.empty:
        return pd.DataFrame()

    frames = []
    for name, spot_ric, fut_ric, _, gsci_ric in COMMODITIES:
        # Roll yield always uses c2 (spot_ric) and c7 (fut_ric)
        spot_df = (df_raw[df_raw["RIC"] == spot_ric][["Date","Price"]]
                   .rename(columns={"Price":"spot"})
                   .drop_duplicates("Date"))
        fut_df  = (df_raw[df_raw["RIC"] == fut_ric][["Date","Price"]]
                   .rename(columns={"Price":"future"})
                   .drop_duplicates("Date"))
        roll    = spot_df.merge(fut_df, on="Date", how="inner").set_index("Date").sort_index()

        # Momentum uses GSCI index if available, else falls back to c2
        mom_ric = gsci_ric if gsci_ric else spot_ric
        mom_df  = (df_raw[df_raw["RIC"] == mom_ric][["Date","Price"]]
                   .rename(columns={"Price":"mom_price"})
                   .drop_duplicates("Date")
                   .set_index("Date").sort_index())

        # Align GSCI onto the futures settlement calendar (ffill for holiday mismatches)
        m = roll.copy()
        m["mom_price"] = mom_df["mom_price"].reindex(roll.index, method="ffill")
        m = m.dropna(subset=["spot", "future", "mom_price"])
        if len(m) < mom_days + 5:
            continue

        # Roll yield (c2 / c7)
        m["roll_yield"] = m["spot"] / m["future"] - 1

        # Momentum calcs on GSCI (or c2 fallback)
        m["log_ret"]     = np.log(m["mom_price"] / m["mom_price"].shift(1))
        m["return_N"]    = m["mom_price"].pct_change(mom_days)
        m["vol"]         = m["log_ret"].rolling(vol_days).std() * np.sqrt(252)
        m["mom_vol_adj"] = m["return_N"] / m["vol"]

        m["commodity"] = name
        frames.append(m.reset_index()[["Date","commodity","spot","future",
                                       "return_N","vol","mom_vol_adj","roll_yield"]])

    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)

    # Each commodity trades on its own exchange calendar, so different commodities
    # have different missing dates (US holidays vs UK holidays vs LME, etc.).
    # Reindex every commodity onto a shared business-day spine and forward-fill
    # gaps up to 5 days so all commodities appear on every date in the snapshot.
    spine = pd.date_range(df["Date"].min(), df["Date"].max(), freq="B")
    metric_cols = ["spot", "future", "return_N", "vol", "mom_vol_adj", "roll_yield"]
    filled = []
    for name, grp in df.groupby("commodity"):
        g = grp.set_index("Date")[metric_cols].reindex(spine).ffill(limit=5)
        g["commodity"] = name
        filled.append(g.reset_index().rename(columns={"index": "Date"}))

    return pd.concat(filled, ignore_index=True)


def cross_rank(series: pd.Series) -> pd.Series:
    """Scale cross-sectional ranks to -20 … +20."""
    n = series.notna().sum()
    if n < 2:
        return pd.Series(np.nan, index=series.index)
    ranked = series.rank(method="average", na_option="keep")
    return ((ranked - 1) / (n - 1) * 40 - 20).round(1)


def snapshot(metrics: pd.DataFrame, date: pd.Timestamp) -> pd.DataFrame:
    """Cross-sectional snapshot at a given date (use nearest available date)."""
    target_str  = pd.Timestamp(date).strftime("%Y-%m-%d")
    avail_strs  = sorted([pd.Timestamp(d).strftime("%Y-%m-%d") for d in metrics["Date"].unique()])
    if not avail_strs:
        return pd.DataFrame()
    idx         = max(0, sum(1 for s in avail_strs if s <= target_str) - 1)
    actual_date = pd.Timestamp(avail_strs[idx])
    s = metrics[metrics["Date"] == actual_date].copy()

    s["rank_mom"]   = cross_rank(s["mom_vol_adj"])
    s["rank_yield"] = cross_rank(s["roll_yield"])
    s["rank_vol"]   = cross_rank(s["vol"])
    return s


# ── Helper: build scatter figure ──────────────────────────────────────────────
def scatter_fig(snap: pd.DataFrame, x_col: str, y_col: str,
                title: str, xlab: str, ylab: str,
                x_pct: bool = False, y_pct: bool = False,
                ranked: bool = False, names: list = None) -> go.Figure:
    fig = go.Figure()
    names = names if names is not None else ALL_NAMES

    # Non-softs rendered first so softs always appear on top
    render_order = [n for n in names if not IS_SOFT[n]] + \
                   [n for n in names if IS_SOFT[n]]
    for name in render_order:
        row = snap[snap["commodity"] == name]
        if row.empty:
            continue
        xv = row[x_col].values[0]
        yv = row[y_col].values[0]
        if pd.isna(xv) or pd.isna(yv):
            continue
        is_soft  = IS_SOFT[name]
        color    = CATEGORY_COLOR[COMMODITY_CATEGORY[name]]
        size     = 12 if is_soft else 7

        fig.add_trace(go.Scatter(
            x=[xv], y=[yv], mode="markers+text",
            marker=dict(color=color, size=size, line=dict(width=0)),
            text=[name], textposition="middle right",
            textfont=dict(size=11 if is_soft else 8,
                          color=color,
                          family="Arial Black" if is_soft else "Arial"),
            name=name, showlegend=False, cliponaxis=False,
            hovertemplate=f"<b>{name}</b><br>{xlab}: %{{x:.2f}}<br>{ylab}: %{{y:.2f}}<extra></extra>",
        ))

    categories_present = {COMMODITY_CATEGORY[n] for n in names}
    for cat, col in CATEGORY_COLOR.items():
        if cat not in categories_present:
            continue
        fig.add_trace(go.Scatter(
            x=[None], y=[None], mode="markers",
            marker=dict(color=col, size=9),
            name=cat, showlegend=True,
        ))

    xfmt = ".0%" if x_pct else ".1f"
    yfmt = ".0%" if y_pct else ".1f"

    # Ranked charts: fix axis range ±22 with crosshair shapes
    # Actual charts: auto-scale, zeroline handles the zero axis
    shapes = []
    yaxis_extra = {}
    xaxis_extra = {}
    if ranked:
        shapes = [
            dict(type="line", x0=0, x1=0, y0=-22, y1=22,
                 line=dict(color="#bbb", width=1, dash="dot")),
            dict(type="line", x0=-22, x1=22, y0=0, y1=0,
                 line=dict(color="#bbb", width=1, dash="dot")),
        ]
        yaxis_extra = {"range": [-22, 22]}
        xaxis_extra = {"range": [-22, 22]}

    fig.update_layout(
        title=dict(text=title, font=dict(size=13, color="#1d1d1f")),
        xaxis=dict(title=xlab, zeroline=True, zerolinecolor="#bbb", zerolinewidth=1,
                   tickformat=xfmt, showgrid=True, gridcolor="#f0f0f0", **xaxis_extra),
        yaxis=dict(title=ylab, zeroline=True, zerolinecolor="#bbb", zerolinewidth=1,
                   tickformat=yfmt, showgrid=True, gridcolor="#f0f0f0", **yaxis_extra),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        margin=dict(l=50, r=20, t=50, b=70),
        height=620,
        shapes=shapes,
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.08,
                    xanchor="center", x=0.5, font=dict(size=10)),
    )
    return fig


# ── Animated scatter ─────────────────────────────────────────────────────────
# Transition slightly longer than frame duration → dots are always mid-glide,
# creating a continuous plane-glide feel rather than step-by-step jumps.
_FRAME_MS      = 350
_TRANSITION_MS = 420
_EASING        = "sin-in-out"   # sinusoidal: gentlest natural acceleration curve


def build_anim_fig(metrics: pd.DataFrame, anim_dates: list,
                   mom_label: str, on_progress=None) -> go.Figure:
    render_order = [n for n in ALL_NAMES if not IS_SOFT[n]] + \
                   [n for n in ALL_NAMES if IS_SOFT[n]]

    def _xy(snap, name):
        row = snap[snap["commodity"] == name]
        if row.empty:
            return None, None
        xv, yv = row["rank_mom"].values[0], row["rank_yield"].values[0]
        return (None, None) if (pd.isna(xv) or pd.isna(yv)) else (float(xv), float(yv))

    snap0 = snapshot(metrics, anim_dates[0])

    # One persistent trace per commodity (non-softs first so softs render on top)
    traces = []
    for name in render_order:
        xv, yv  = _xy(snap0, name)
        is_soft = IS_SOFT[name]
        color   = CATEGORY_COLOR[COMMODITY_CATEGORY[name]]
        traces.append(go.Scatter(
            x=[xv], y=[yv], mode="markers+text",
            marker=dict(color=color, size=12 if is_soft else 7, line=dict(width=0)),
            text=[name], textposition="middle right", cliponaxis=False,
            textfont=dict(size=11 if is_soft else 8, color=color,
                          family="Arial Black" if is_soft else "Arial"),
            name=name, showlegend=False,
            hovertemplate=f"<b>{name}</b><br>Mom Rank: %{{x:.1f}}<br>Yield Rank: %{{y:.1f}}<extra></extra>",
        ))

    # Static legend traces — after animated traces, so frames don't touch them
    for cat, col in CATEGORY_COLOR.items():
        traces.append(go.Scatter(x=[None], y=[None], mode="markers",
                                 marker=dict(color=col, size=9),
                                 name=cat, showlegend=True))

    # One frame per date — progress callback fired after each
    total  = len(anim_dates)
    frames = []
    for i, dt in enumerate(anim_dates):
        snap = snapshot(metrics, dt)
        frames.append(go.Frame(
            data=[go.Scatter(x=[_xy(snap, n)[0]], y=[_xy(snap, n)[1]])
                  for n in render_order],
            name=dt.strftime("%Y-%m-%d"),
        ))
        if on_progress:
            on_progress(i + 1, total)

    _t = dict(duration=_TRANSITION_MS, easing=_EASING)
    _f = dict(duration=_FRAME_MS, redraw=False)

    slider_steps = [
        dict(method="animate",
             args=[[f.name], dict(mode="immediate", frame=_f, transition=_t)],
             label=f.name)
        for f in frames
    ]

    fig = go.Figure(data=traces, frames=frames)
    fig.update_layout(
        title=dict(text=f"Cross-Section Animation — Momentum (Vol Adj) {mom_label} vs Roll Yield (Ranked)",
                   font=dict(size=13, color="#1d1d1f")),
        xaxis=dict(title=f"Momentum (Vol Adj) {mom_label}", range=[-22, 22],
                   zeroline=True, zerolinecolor="#bbb", zerolinewidth=1,
                   showgrid=True, gridcolor="#f0f0f0", tickformat=".1f"),
        yaxis=dict(title="1 Year Spread (Roll Yield)", range=[-22, 22],
                   zeroline=True, zerolinecolor="#bbb", zerolinewidth=1,
                   showgrid=True, gridcolor="#f0f0f0", tickformat=".1f"),
        plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
        height=820,
        margin=dict(l=50, r=20, t=60, b=130),
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=-0.12,
                    xanchor="center", x=0.5, font=dict(size=10)),
        shapes=[
            dict(type="line", x0=0, x1=0, y0=-22, y1=22,
                 line=dict(color="#bbb", width=1, dash="dot")),
            dict(type="line", x0=-22, x1=22, y0=0, y1=0,
                 line=dict(color="#bbb", width=1, dash="dot")),
        ],
        updatemenus=[dict(
            type="buttons", showactive=False,
            y=1.07, x=0.5, xanchor="center",
            pad=dict(r=10, t=10),
            buttons=[
                dict(label="▶  Play", method="animate",
                     args=[None, dict(frame=_f, fromcurrent=True,
                                     mode="immediate", transition=_t)]),
                dict(label="⏸  Pause", method="animate",
                     args=[[None], dict(frame=dict(duration=0, redraw=False),
                                       mode="immediate")]),
            ],
        )],
        sliders=[dict(
            active=0, steps=slider_steps,
            x=0, len=1.0, y=0, yanchor="top",
            pad=dict(b=10, t=55),
            currentvalue=dict(visible=False),
            transition=_t,
        )],
    )
    return fig


# ── Outlook email helpers ─────────────────────────────────────────────────────
def _build_snapshot_html(snap: pd.DataFrame, date_str: str,
                         mom_label: str, vol_label: str) -> str:
    tbl = snap[["commodity", "rank_mom", "rank_yield", "rank_vol",
                "mom_vol_adj", "roll_yield"]].copy()
    tbl = tbl.dropna(subset=["rank_mom"]).sort_values("rank_mom", ascending=False).reset_index(drop=True)

    rows = ""
    for i, row in tbl.iterrows():
        cat   = COMMODITY_CATEGORY.get(row["commodity"], "")
        color = CATEGORY_COLOR.get(cat, "#333")
        bold  = "font-weight:bold;" if cat == "Softs" else ""
        bg    = "#fafafa" if i % 2 == 0 else "#ffffff"
        ry    = f"{row['roll_yield']*100:.1f}%" if pd.notna(row["roll_yield"]) else "—"
        mom   = f"{row['mom_vol_adj']:.2f}"      if pd.notna(row["mom_vol_adj"]) else "—"
        rows += (
            f'<tr style="background:{bg}">'
            f'<td style="color:{color};{bold}padding:5px 8px">{row["commodity"]}</td>'
            f'<td style="padding:5px 8px;color:#555">{cat}</td>'
            f'<td style="text-align:center;padding:5px 8px">{row["rank_mom"]:.1f}</td>'
            f'<td style="text-align:center;padding:5px 8px">{row["rank_yield"]:.1f}</td>'
            f'<td style="text-align:center;padding:5px 8px">{row["rank_vol"]:.1f}</td>'
            f'<td style="text-align:center;padding:5px 8px">{mom}</td>'
            f'<td style="text-align:center;padding:5px 8px">{ry}</td>'
            f'</tr>'
        )

    th = "background:#1565C0;color:#fff;padding:6px 8px;text-align:left;"
    thc = "background:#1565C0;color:#fff;padding:6px 8px;text-align:center;"
    return f"""
<html><body style="font-family:Arial,sans-serif;font-size:13px;color:#1d1d1f">
<h2 style="color:#1565C0;margin-bottom:4px">Cross Section Monitor</h2>
<p style="margin:0;color:#555">
  <b>Date:</b> {date_str} &nbsp;|&nbsp;
  <b>Momentum:</b> {mom_label} &nbsp;|&nbsp;
  <b>Vol Window:</b> {vol_label}
</p><br>
<table border="0" cellpadding="0" cellspacing="0"
       style="border-collapse:collapse;width:680px;font-size:12px">
  <tr>
    <th style="{th}">Commodity</th>
    <th style="{th}">Category</th>
    <th style="{thc}">Mom Rank</th>
    <th style="{thc}">Yield Rank</th>
    <th style="{thc}">Vol Rank</th>
    <th style="{thc}">Mom (Vol Adj)</th>
    <th style="{thc}">Roll Yield</th>
  </tr>
  {rows}
</table>
<br><p style="color:#aaa;font-size:10px">Cross Section Monitor — auto-generated</p>
</body></html>"""


def _send_outlook_snapshot(snap: pd.DataFrame, date_str: str,
                           mom_label: str, vol_label: str) -> None:
    html    = _build_snapshot_html(snap, date_str, mom_label, vol_label)
    outlook = _win32.Dispatch("Outlook.Application")
    mail    = outlook.CreateItem(0)
    mail.To      = "viratarya30@gmail.com"
    mail.Subject = f"Cross Section Monitor — {date_str} | {mom_label}"
    mail.HTMLBody = html
    mail.Display()


# ── Controls ──────────────────────────────────────────────────────────────────
df_raw = load_data()

if df_raw.empty:
    st.error("No data found. Run `ingest.py` first.")
    st.stop()

# Store dates as sorted ISO strings — pure Python, zero numpy involvement
all_dates_str = sorted([pd.Timestamp(d).strftime("%Y-%m-%d") for d in df_raw["Date"].unique()])
latest_date   = pd.Timestamp(all_dates_str[-1])

# ── Sidebar controls ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Controls")

    st.markdown("**Date**")
    use_latest = st.button("Latest Date", use_container_width=True)
    chosen_date = st.date_input("Select Date", value=latest_date.date(),
                                min_value=pd.Timestamp(all_dates_str[0]).date(),
                                max_value=latest_date.date())
    chosen_str = str(chosen_date)
    if use_latest:
        chosen_str = all_dates_str[-1]

    st.markdown("---")
    mom_label = st.radio("Momentum Period", list(MOM_DAYS.keys()), index=1)
    mom_days  = MOM_DAYS[mom_label]

    st.markdown("---")
    vol_label = st.radio("Volatility Window", list(VOL_DAYS.keys()), index=1)
    vol_days  = VOL_DAYS[vol_label]

    st.markdown("---")
    st.markdown("**Previous Period**")
    prev_label    = st.selectbox("Preset", list(PREV_DAYS.keys()), index=2)
    prev_override = st.number_input("Custom Days", min_value=1, max_value=500,
                                    value=PREV_DAYS[prev_label])
    prev_days = prev_override

    st.markdown("---")
    st.markdown("**Email Snapshot**")
    _send_clicked = st.button("Send Snapshot via Outlook",
                              use_container_width=True,
                              disabled=not _OUTLOOK_OK)
    if not _OUTLOOK_OK:
        st.caption("pywin32 not available")

# Pure string comparison (ISO dates sort lexicographically) — no numpy possible
curr_idx = max(0, sum(1 for s in all_dates_str if s <= chosen_str) - 1)
prev_idx = max(0, curr_idx - int(prev_days))
curr_ts  = pd.Timestamp(all_dates_str[curr_idx])
prev_ts  = pd.Timestamp(all_dates_str[prev_idx])

st.caption(f"Current: **{all_dates_str[curr_idx]}**  |  Previous: **{all_dates_str[prev_idx]}**")

# ── Compute metrics ───────────────────────────────────────────────────────────
metrics = build_metrics(df_raw, mom_days, vol_days)

if metrics.empty:
    st.error("Could not compute metrics. Check data quality.")
    st.stop()

snap_curr = snapshot(metrics, curr_ts)
snap_prev = snapshot(metrics, prev_ts)

if _send_clicked:
    try:
        _send_outlook_snapshot(snap_curr, all_dates_str[curr_idx], mom_label, vol_label)
        st.sidebar.success("Opened in Outlook — review and hit Send.")
    except Exception as e:
        st.sidebar.error(f"Outlook error: {e}")

mom_xlab  = f"Momentum (Vol Adj) {mom_label}"

# Compute cross-sectional rank per date for a given commodity subset
@st.cache_data(ttl=3600)
def build_rank_ts(metrics: pd.DataFrame, signal: str) -> pd.DataFrame:
    """Vectorised cross-sectional rank per date."""
    df = metrics[["Date", "commodity", signal]].copy()
    def _rank(s):
        n = s.notna().sum()
        if n < 2:
            return pd.Series(np.nan, index=s.index)
        return ((s.rank(method="average", na_option="keep") - 1) / (n - 1) * 40 - 20).round(1)
    df["rank"] = df.groupby("Date")[signal].transform(_rank)
    return df[["Date", "commodity", "rank"]].dropna(subset=["rank"])


def render_rank_ts_chart(rank_ts: pd.DataFrame, select: list) -> go.Figure:
    fig = go.Figure()
    for name in select:
        sub = rank_ts[rank_ts["commodity"] == name].sort_values("Date")
        color = COMMODITY_TS_COLOR.get(name, CATEGORY_COLOR[COMMODITY_CATEGORY[name]])
        fig.add_trace(go.Scatter(
            x=sub["Date"], y=sub["rank"],
            mode="lines", name=name,
            line=dict(color=color, width=2.5 if IS_SOFT[name] else 1.5),
        ))
    fig.update_layout(
        xaxis=dict(title="Date", showgrid=True, gridcolor="#f0f0f0"),
        yaxis=dict(title="Cross-Sectional Rank (−20 to +20)",
                   range=[-22, 22], zeroline=True, zerolinecolor="#bbb",
                   showgrid=True, gridcolor="#f0f0f0"),
        plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
        margin=dict(l=50, r=20, t=30, b=80),
        height=420,
    )
    return fig


tab_all, tab_ags = st.tabs(["All Commodities", "Ags Only"])

with tab_all:
    # ── Section 1: Scatter charts ─────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Momentum vs Roll Yield Scatter")

    col_l, col_r = st.columns(2)

    with col_l:
        fig_rank_curr = scatter_fig(
            snap_curr, "rank_mom", "rank_yield",
            "Momentum vs Spread Scatter (Relative Rank Basis)",
            mom_xlab, "1 Year Spread",
            ranked=True,
        )
        st.plotly_chart(fig_rank_curr, use_container_width=True)

        fig_act_curr = scatter_fig(
            snap_curr, "mom_vol_adj", "roll_yield",
            "Momentum vs Spread Scatter (Actual)",
            mom_xlab, "1 Year Spread",
            x_pct=False, y_pct=True,
            ranked=False,
        )
        st.plotly_chart(fig_act_curr, use_container_width=True)

    with col_r:
        fig_rank_prev = scatter_fig(
            snap_prev, "rank_mom", "rank_yield",
            f"Momentum vs Spread Scatter (Relative Rank Basis) : {prev_label}",
            mom_xlab, "1 Year Spread",
            ranked=True,
        )
        st.plotly_chart(fig_rank_prev, use_container_width=True)

        fig_act_prev = scatter_fig(
            snap_prev, "mom_vol_adj", "roll_yield",
            f"Momentum vs Spread Scatter (Actual) : {prev_label}",
            mom_xlab, "1 Year Spread",
            x_pct=False, y_pct=True,
            ranked=False,
        )
        st.plotly_chart(fig_act_prev, use_container_width=True)

    # ── Section 2: Animated cross-section ─────────────────────────────────────
    st.markdown("---")
    with st.expander("Animated Cross-Section", expanded=False):
        _anim_span = curr_idx - prev_idx
        if _anim_span > MAX_ANIM_DAYS:
            st.warning(
                f"Period too long for animation ({_anim_span} trading days). "
                f"Capped at {MAX_ANIM_DAYS} days (~6 months). "
                f"Choose Previous Quarter or shorter."
            )
            _anim_start = curr_idx - MAX_ANIM_DAYS
        else:
            _anim_start = prev_idx

        _anim_dates = [pd.Timestamp(all_dates_str[i])
                       for i in range(_anim_start, curr_idx + 1)]
        _anim_key   = (all_dates_str[_anim_start], all_dates_str[curr_idx], mom_label, vol_label)

        st.caption(f"{len(_anim_dates)} frames · {all_dates_str[_anim_start]} → {all_dates_str[curr_idx]}")

        if "anim_cache" not in st.session_state:
            st.session_state.anim_cache = (None, None)

        if st.button("Generate Animation", use_container_width=True):
            _prog = st.progress(0, text=f"Building frame 1 / {len(_anim_dates)}…")
            def _on_progress(i, total):
                _prog.progress(i / total, text=f"Building frame {i} / {total}…")
            _fig_anim = build_anim_fig(metrics, _anim_dates, mom_label,
                                       on_progress=_on_progress)
            _prog.empty()
            st.session_state.anim_cache = (_anim_key, _fig_anim)

        _cached_key, _cached_fig = st.session_state.anim_cache
        if _cached_fig is not None and _cached_key == _anim_key:
            st.plotly_chart(_cached_fig, use_container_width=True)
        elif _cached_fig is not None:
            st.info("Date or parameters changed — click Generate to refresh.")

    # ── Section 3: Momentum ranking time series ────────────────────────────────
    st.markdown("---")
    st.subheader(f"Momentum Ranking Time Series ({mom_label})")

    ts_cols_default = SOFT_NAMES
    ts_select = st.multiselect(
        "Select commodities for time series",
        options=ALL_NAMES,
        default=ts_cols_default,
        key="all_ts_multiselect",
    )

    rank_ts = build_rank_ts(metrics, "mom_vol_adj")
    st.plotly_chart(render_rank_ts_chart(rank_ts, ts_select), use_container_width=True)

    # ── Section 3b: Spread (Roll Yield) ranking time series ────────────────────
    st.markdown("---")
    st.subheader("Spread Ranking Time Series (Roll Yield)")

    spread_ts_select = st.multiselect(
        "Select commodities for time series",
        options=ALL_NAMES,
        default=ts_cols_default,
        key="spread_ts_multiselect",
    )

    rank_ts_spread = build_rank_ts(metrics, "roll_yield")
    st.plotly_chart(render_rank_ts_chart(rank_ts_spread, spread_ts_select), use_container_width=True)

    # ── Section 4: Volatility ranking (collapsible) ────────────────────────────
    st.markdown("---")
    with st.expander("Volatility Ranking", expanded=False):
        st.markdown(f"##### Cross-Sectional Volatility Rank ({vol_label} window) — {all_dates_str[curr_idx]}")

        vol_select = st.multiselect(
            "Select commodities",
            options=ALL_NAMES,
            default=SOFT_NAMES,
            key="vol_multiselect",
        )

        snap_vol = snap_curr[snap_curr["commodity"].isin(vol_select)].copy()
        snap_vol = snap_vol.dropna(subset=["rank_vol","vol"]).sort_values("rank_vol", ascending=False)

        if snap_vol.empty:
            st.info("No volatility data available.")
        else:
            tbl = snap_vol[["commodity","vol","rank_vol","mom_vol_adj","rank_mom","roll_yield","rank_yield"]].copy()
            tbl.columns = ["Commodity","Vol (ann.)","Vol Rank","Mom (Vol Adj)","Mom Rank","Roll Yield","Yield Rank"]
            tbl["Vol (ann.)"]    = (tbl["Vol (ann.)"]   * 100).round(1).astype(str) + "%"
            tbl["Roll Yield"]    = (tbl["Roll Yield"]   * 100).round(1).astype(str) + "%"
            tbl["Mom (Vol Adj)"] = tbl["Mom (Vol Adj)"].round(2)
            st.dataframe(tbl.set_index("Commodity"), use_container_width=True)

with tab_ags:
    # Ranks recomputed within the Ags-only universe (Softs + Grains & Oilseeds +
    # Livestock), not sliced from the 31-name global rank — so a name's rank here
    # reflects its standing among ags only, and will differ from the main tab.
    metrics_ag  = metrics[metrics["commodity"].isin(AG_NAMES)]
    snap_curr_ag = snapshot(metrics_ag, curr_ts)
    snap_prev_ag = snapshot(metrics_ag, prev_ts)

    st.markdown("---")
    st.subheader("Momentum vs Roll Yield Scatter — Ags Only")

    col_l, col_r = st.columns(2)

    with col_l:
        fig_rank_curr_ag = scatter_fig(
            snap_curr_ag, "rank_mom", "rank_yield",
            "Momentum vs Spread Scatter (Relative Rank Basis) — Ags",
            mom_xlab, "1 Year Spread",
            ranked=True, names=AG_NAMES,
        )
        st.plotly_chart(fig_rank_curr_ag, use_container_width=True, key="ag_rank_curr")

        fig_act_curr_ag = scatter_fig(
            snap_curr_ag, "mom_vol_adj", "roll_yield",
            "Momentum vs Spread Scatter (Actual) — Ags",
            mom_xlab, "1 Year Spread",
            x_pct=False, y_pct=True,
            ranked=False, names=AG_NAMES,
        )
        st.plotly_chart(fig_act_curr_ag, use_container_width=True, key="ag_act_curr")

    with col_r:
        fig_rank_prev_ag = scatter_fig(
            snap_prev_ag, "rank_mom", "rank_yield",
            f"Momentum vs Spread Scatter (Relative Rank Basis) — Ags : {prev_label}",
            mom_xlab, "1 Year Spread",
            ranked=True, names=AG_NAMES,
        )
        st.plotly_chart(fig_rank_prev_ag, use_container_width=True, key="ag_rank_prev")

        fig_act_prev_ag = scatter_fig(
            snap_prev_ag, "mom_vol_adj", "roll_yield",
            f"Momentum vs Spread Scatter (Actual) — Ags : {prev_label}",
            mom_xlab, "1 Year Spread",
            x_pct=False, y_pct=True,
            ranked=False, names=AG_NAMES,
        )
        st.plotly_chart(fig_act_prev_ag, use_container_width=True, key="ag_act_prev")

    # ── Momentum ranking time series (Ags) ─────────────────────────────────────
    st.markdown("---")
    st.subheader(f"Momentum Ranking Time Series ({mom_label}) — Ags Only")

    ag_ts_select = st.multiselect(
        "Select commodities for time series",
        options=AG_NAMES,
        default=SOFT_NAMES,
        key="ag_ts_multiselect",
    )

    rank_ts_ag = build_rank_ts(metrics_ag, "mom_vol_adj")
    st.plotly_chart(render_rank_ts_chart(rank_ts_ag, ag_ts_select), use_container_width=True, key="ag_mom_ts_chart")

    # ── Spread (Roll Yield) ranking time series (Ags) ──────────────────────────
    st.markdown("---")
    st.subheader("Spread Ranking Time Series (Roll Yield) — Ags Only")

    ag_spread_ts_select = st.multiselect(
        "Select commodities for time series",
        options=AG_NAMES,
        default=SOFT_NAMES,
        key="ag_spread_ts_multiselect",
    )

    rank_ts_spread_ag = build_rank_ts(metrics_ag, "roll_yield")
    st.plotly_chart(render_rank_ts_chart(rank_ts_spread_ag, ag_spread_ts_select), use_container_width=True, key="ag_spread_ts_chart")

# ── Footer ────────────────────────────────────────────────────────────────────
st.caption(f"Data updated: {latest_date.date()} | {N} commodities | Vol window: {vol_label} | Mom: {mom_label}")
