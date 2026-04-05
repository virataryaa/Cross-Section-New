"""
Cross Section Ingest
Fetches 3 years of daily settlement prices for 31 spot + 31 future RICs.
Saves to prices.parquet with columns: Date, RIC, Price
"""

import datetime
import pandas as pd
import refinitiv.data as rd
from pathlib import Path

HERE = Path(__file__).parent
OUT  = HERE / "prices.parquet"

# ── RIC config ────────────────────────────────────────────────────────────────
# (commodity, spot_ric, future_ric, is_soft)
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

ALL_SPOT    = [r[1] for r in COMMODITIES]
ALL_FUTURE  = [r[2] for r in COMMODITIES]
ALL_RICS    = ALL_SPOT + ALL_FUTURE   # 64 total


def _get_start_date() -> str:
    """3 years back, or incremental (last 10 days) if parquet exists."""
    if OUT.exists():
        df = pd.read_parquet(OUT)
        last = pd.to_datetime(df["Date"]).max()
        start = last - pd.Timedelta(days=14)
        print(f"Incremental: fetching from {start.date()}")
        return start.strftime("%Y-%m-%d")
    three_yrs = datetime.date.today() - datetime.timedelta(days=3*365)
    print(f"Full fetch: from {three_yrs}")
    return three_yrs.strftime("%Y-%m-%d")


def fetch() -> pd.DataFrame:
    start = _get_start_date()
    end   = datetime.date.today().isoformat()

    print(f"Fetching {len(ALL_RICS)} RICs from {start} to {end} ...")
    raw = rd.get_history(
        universe=ALL_RICS,
        fields=["TR.SETTLEMENTPRICE"],
        start=start,
        end=end,
        interval="daily",
    )
    # raw: DatetimeIndex × MultiIndex columns (RIC, field)
    raw.index = pd.to_datetime(raw.index)

    # Flatten to long format
    frames = []
    for ric in ALL_RICS:
        try:
            col = raw[ric]           # already a Series indexed by Date
            tmp = col.dropna().reset_index()
            tmp.columns = ["Date", "Price"]
            tmp["RIC"] = ric
            frames.append(tmp)
        except Exception as e:
            print(f"  Warning: could not extract {ric}: {e}")

    df = pd.concat(frames, ignore_index=True)
    df["Date"] = pd.to_datetime(df["Date"])
    return df


def save(df_new: pd.DataFrame):
    if OUT.exists():
        df_old = pd.read_parquet(OUT)
        df_old["Date"] = pd.to_datetime(df_old["Date"])
        # Remove overlapping dates then append
        cutoff = df_new["Date"].min()
        df_old = df_old[df_old["Date"] < cutoff]
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new

    df = df.sort_values(["RIC", "Date"]).reset_index(drop=True)
    df.to_parquet(OUT, index=False)
    print(f"Saved {len(df):,} rows to {OUT}")
    print(f"Date range: {df['Date'].min().date()} to {df['Date'].max().date()}")
    print(f"RICs: {df['RIC'].nunique()}")


if __name__ == "__main__":
    rd.open_session()
    try:
        df = fetch()
        save(df)
    finally:
        rd.close_session()
