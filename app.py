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
IS_SOFT     = {r[0]: r[3] for r in COMMODITIES}
NAME2GSCI   = {r[0]: r[4] for r in COMMODITIES}
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
    """Vectorised cross-sectional rank per date."""
    df = metrics[["Date", "commodity", signal]].copy()
    def _rank(s):
        n = s.notna().sum()
        if n < 2:
            return pd.Series(np.nan, index=s.index)
        return ((s.rank(method="average", na_option="keep") - 1) / (n - 1) * 40 - 20).round(1)
    df["rank"] = df.groupby("Date")[signal].transform(_rank)
    return df[["Date", "commodity", "rank"]].dropna(subset=["rank"])

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
