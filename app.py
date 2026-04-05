"""
Cross Section Monitor
Momentum (Vol-Adj) vs Roll Yield scatter — cross-sectional ranking across 31 commodities.
"""

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Cross Section Monitor", layout="wide")
st.markdown("""
<style>
[data-testid="stAppViewContainer"]{background:#f5f5f7;}
[data-testid="stSidebar"]{background:#1d1d1f;}
h1,h2,h3{color:#1d1d1f;}
</style>
""", unsafe_allow_html=True)

st.title("Cross Section Monitor")

# ── Constants ─────────────────────────────────────────────────────────────────
HERE = Path(__file__).parent

COMMODITIES = [
    ("Brent Oil",    "LCOc1",  "LCOc13",  False),
    ("Cocoa",        "CCc2",   "CCc7",    True),
    ("Coffee",       "KCc2",   "KCc7",    True),
    ("Corn",         "Cc1",    "Cc6",     False),
    ("Cotton",       "CTc2",   "CTc7",    True),
    ("Gas Oil",      "LGOc1",  "LGOc13",  False),
    ("Gold",         "GCc1",   "GCc8",    False),
    ("Heating Oil",  "HOc1",   "HOc13",   False),
    ("HG Copper",    "HGc1",   "HGc13",   False),
    ("LDN Cocoa",    "LCCc2",  "LCCc7",   True),
    ("Lean Hog",     "LHc1",   "LHc9",    False),
    ("Live Cattle",  "LCc1",   "LCc7",    False),
    ("Natural Gas",  "NGc1",   "NGc13",   False),
    ("Silver",       "SIc1",   "SIc8",    False),
    ("Soy Meal",     "SMc1",   "SMc9",    False),
    ("Soy Bean",     "Sc1",    "Sc8",     False),
    ("Soy Oil",      "BOc1",   "BOc9",    False),
    ("Sugar",        "SBc1",   "SBc5",    True),
    ("Gasoline",     "RBc1",   "RBc13",   False),
    ("Wheat",        "Wc1",    "Wc6",     False),
    ("Wheat (KCB)",  "KWc1",   "KWc6",    False),
    ("WTI Crude",    "CLc1",   "CLc13",   False),
    ("White Sugar",  "LSUc1",  "LSUc6",   True),
    ("Robusta",      "LRCc2",  "LRCc8",   True),
    ("Canola",       "RSc2",   "RSc7",    False),
    ("Zinc",         "MZNc2",  "MZNc13",  False),
    ("Aluminium",    "MALc2",  "MALc13",  False),
    ("Lead",         "MPBc2",  "MPBc13",  False),
    ("Copper",       "MCUc2",  "MCUc13",  False),
    ("Nickel",       "MNIc2",  "MNIc13",  False),
    ("Tin",          "MSNc2",  "MSNc13",  False),
    ("Orange Juice", "OJc2",   "OJc7",    True),
]

NAME2SPOT   = {r[0]: r[1] for r in COMMODITIES}
NAME2FUT    = {r[0]: r[2] for r in COMMODITIES}
IS_SOFT     = {r[0]: r[3] for r in COMMODITIES}
ALL_NAMES   = [r[0] for r in COMMODITIES]
SOFT_NAMES  = [r[0] for r in COMMODITIES if r[3]]

MOM_DAYS    = {"3 Months": 63, "6 Months": 126, "12 Months": 252}
VOL_DAYS    = {"20d": 20, "60d": 60, "120d": 120}
PREV_DAYS   = {"Previous Day": 1, "Previous Week": 5, "Previous Month": 20,
               "Previous Quarter": 60, "Previous Year": 250}

DARK_RED    = "#8b1a1a"
SOFT_COLOR  = "#1a3a5c"
OTHER_COLOR = "#8b0000"
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
    for name, spot_ric, fut_ric, _ in COMMODITIES:
        spot_df = df_raw[df_raw["RIC"] == spot_ric][["Date","Price"]].rename(columns={"Price":"spot"})
        fut_df  = df_raw[df_raw["RIC"] == fut_ric][["Date","Price"]].rename(columns={"Price":"future"})

        m = spot_df.merge(fut_df, on="Date", how="inner").set_index("Date").sort_index()
        if len(m) < mom_days + 5:
            continue

        # Roll yield
        m["roll_yield"] = m["spot"] / m["future"] - 1

        # Log returns on spot
        m["log_ret"] = np.log(m["spot"] / m["spot"].shift(1))

        # N-month return (simple)
        m["return_N"] = m["spot"].pct_change(mom_days)

        # Vol (annualised)
        m["vol"] = m["log_ret"].rolling(vol_days).std() * np.sqrt(252)

        # Vol-adj momentum
        m["mom_vol_adj"] = m["return_N"] / m["vol"]

        m["commodity"] = name
        frames.append(m.reset_index()[["Date","commodity","spot","future",
                                       "return_N","vol","mom_vol_adj","roll_yield"]])

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


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
                ranked: bool = False) -> go.Figure:
    fig = go.Figure()

    for name in ALL_NAMES:
        row = snap[snap["commodity"] == name]
        if row.empty:
            continue
        xv = row[x_col].values[0]
        yv = row[y_col].values[0]
        if pd.isna(xv) or pd.isna(yv):
            continue
        is_soft = IS_SOFT[name]
        color   = SOFT_COLOR if is_soft else OTHER_COLOR
        size    = 10 if is_soft else 7

        fig.add_trace(go.Scatter(
            x=[xv], y=[yv], mode="markers+text",
            marker=dict(color=color, size=size),
            text=[name], textposition="middle right",
            textfont=dict(size=10 if is_soft else 8,
                          color=color,
                          family="Arial Black" if is_soft else "Arial"),
            name=name, showlegend=False,
            hovertemplate=f"<b>{name}</b><br>{xlab}: %{{x:.2f}}<br>{ylab}: %{{y:.2f}}<extra></extra>",
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
        margin=dict(l=50, r=20, t=50, b=50),
        height=620,
        shapes=shapes,
    )
    return fig


# ── Controls ──────────────────────────────────────────────────────────────────
df_raw = load_data()

if df_raw.empty:
    st.error("No data found. Run `ingest.py` first.")
    st.stop()

# Store dates as sorted ISO strings — pure Python, zero numpy involvement
all_dates_str = sorted([pd.Timestamp(d).strftime("%Y-%m-%d") for d in df_raw["Date"].unique()])
latest_date   = pd.Timestamp(all_dates_str[-1])

col_c1, col_c2, col_c3, col_c4, col_c5 = st.columns([1.2, 1, 1.2, 1, 1.5])

with col_c1:
    st.markdown("**Choose Date**")
    use_latest = st.button("Latest Date", use_container_width=True)

with col_c2:
    chosen_date = st.date_input("", value=latest_date.date(),
                                min_value=pd.Timestamp(all_dates_str[0]).date(),
                                max_value=latest_date.date(),
                                label_visibility="collapsed")
    chosen_str = str(chosen_date)          # "YYYY-MM-DD" — plain Python string
    if use_latest:
        chosen_str = all_dates_str[-1]

with col_c3:
    mom_label = st.radio("Momentum Period", list(MOM_DAYS.keys()),
                         index=1, horizontal=True)
    mom_days = MOM_DAYS[mom_label]

with col_c4:
    vol_label = st.radio("Volatility Window", list(VOL_DAYS.keys()),
                         index=1, horizontal=True)
    vol_days = VOL_DAYS[vol_label]

with col_c5:
    st.markdown("**Previous Period**")
    pc1, pc2 = st.columns([2, 1])
    with pc1:
        prev_label = st.selectbox("", list(PREV_DAYS.keys()),
                                  index=2, label_visibility="collapsed")
    with pc2:
        prev_override = st.number_input("Days", min_value=1, max_value=500,
                                        value=PREV_DAYS[prev_label],
                                        label_visibility="collapsed")
    prev_days = prev_override

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

mom_xlab  = f"Momentum (Vol Adj) {mom_label}"

# ── Section 1: Scatter charts ─────────────────────────────────────────────────
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

# ── Section 2: Momentum ranking time series ───────────────────────────────────
st.markdown("---")
st.subheader(f"Momentum Ranking Time Series ({mom_label})")

# Compute cross-sectional rank per date for all commodities
@st.cache_data(ttl=3600)
def build_rank_ts(metrics: pd.DataFrame, signal: str) -> pd.DataFrame:
    """Pivot: rows=Date, cols=commodity, values=cross-sectional rank."""
    dates = metrics["Date"].unique()
    rows  = []
    for d in dates:
        day = metrics[metrics["Date"] == d][["commodity", signal]].copy()
        day["rank"] = cross_rank(day[signal])
        for _, r in day.iterrows():
            rows.append({"Date": d, "commodity": r["commodity"], "rank": r["rank"]})
    return pd.DataFrame(rows)

ts_cols_default = SOFT_NAMES
ts_select = st.multiselect(
    "Select commodities for time series",
    options=ALL_NAMES,
    default=ts_cols_default,
)

rank_ts = build_rank_ts(metrics, "mom_vol_adj")

fig_ts = go.Figure()
# Unique color per commodity — consistent regardless of selection
PALETTE = [
    "#1f77b4","#d62728","#2ca02c","#ff7f0e","#9467bd",
    "#8c564b","#e377c2","#17becf","#bcbd22","#7f7f7f",
    "#aec7e8","#ff9896","#98df8a","#ffbb78","#c5b0d5",
    "#c49c94","#f7b6d2","#dbdb8d","#9edae5","#393b79",
    "#6b6ecf","#b5cf6b","#e7969c","#9c9ede","#cedb9c",
    "#e7cb94","#e7ba52","#843c39","#ad494a","#d6616b",
    "#3182bd",
]
for name in ts_select:
    sub = rank_ts[rank_ts["commodity"] == name].sort_values("Date")
    cidx = ALL_NAMES.index(name)
    color = PALETTE[cidx % len(PALETTE)]
    fig_ts.add_trace(go.Scatter(
        x=sub["Date"], y=sub["rank"],
        mode="lines", name=name,
        line=dict(color=color, width=2.5 if IS_SOFT[name] else 1.5),
    ))

fig_ts.update_layout(
    xaxis=dict(title="Date", showgrid=True, gridcolor="#f0f0f0"),
    yaxis=dict(title="Cross-Sectional Rank (−20 to +20)",
               range=[-22, 22], zeroline=True, zerolinecolor="#bbb",
               showgrid=True, gridcolor="#f0f0f0"),
    plot_bgcolor="#ffffff", paper_bgcolor="#ffffff",
    legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
    margin=dict(l=50, r=20, t=30, b=80),
    height=420,
)
st.plotly_chart(fig_ts, use_container_width=True)

# ── Section 3: Volatility ranking (collapsible) ───────────────────────────────
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

# ── Footer ────────────────────────────────────────────────────────────────────
st.caption(f"Data updated: {latest_date.date()} | {N} commodities | Vol window: {vol_label} | Mom: {mom_label}")
