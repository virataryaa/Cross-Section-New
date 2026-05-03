"""
Cross Section Ingest
Fetches 3 years of daily prices for 31 commodities.
  - GSCI Excess Return RICs  → TR.PriceClose       (used for momentum)
  - c2 / c7 continuous RICs  → TR.SETTLEMENTPRICE  (used for roll yield)
Saves to prices.parquet with columns: Date, RIC, Price
"""

import datetime
import pandas as pd
import lseg.data as rd
from pathlib import Path

HERE = Path(__file__).parent
OUT  = HERE / "prices.parquet"

# (name, spot_ric, fut_ric, is_soft, gsci_ric)
# gsci_ric = None  →  no GSCI index; c2 (spot_ric) is used for momentum instead
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
    ("LDN Cocoa",    "LCCc2",  "LCCc7",   True,  None),        # no GSCI
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
    ("White Sugar",  "LSUc1",  "LSUc6",   True,  None),        # no GSCI
    ("Robusta",      "LRCc2",  "LRCc8",   True,  None),        # no GSCI
    ("Canola",       "RSc2",   "RSc7",    False, None),        # no GSCI
    ("Zinc",         "MZNc2",  "MZNc13",  False, ".SPGSIZP"),
    ("Aluminium",    "MALc2",  "MALc13",  False, ".SPGSIAP"),
    ("Lead",         "MPBc2",  "MPBc13",  False, ".SPGSILP"),
    ("Copper",       "MCUc2",  "MCUc13",  False, ".SPGSICP"),
    ("Nickel",       "MNIc2",  "MNIc13",  False, ".SPGSIKP"),
    ("Tin",          "MSNc2",  "MSNc13",  False, ".SPGSIS"),
    ("Orange Juice", "OJc2",   "OJc7",    True,  None),        # no GSCI
]

# RIC lists
GSCI_RICS  = [r[4] for r in COMMODITIES if r[4] is not None]
ROLL_RICS  = [r[1] for r in COMMODITIES] + [r[2] for r in COMMODITIES]  # c2 + c7 for all


def _incremental_start() -> str:
    """Last 14 days if parquet exists, else 3 years back."""
    if OUT.exists():
        df = pd.read_parquet(OUT)
        last = pd.to_datetime(df["Date"]).max()
        start = last - pd.Timedelta(days=14)
        print(f"Incremental (roll): fetching from {start.date()}")
        return start.strftime("%Y-%m-%d")
    return (datetime.date.today() - datetime.timedelta(days=3*365)).strftime("%Y-%m-%d")


def _gsci_start() -> str:
    """Full 3-year fetch if any GSCI RIC has less than 400 rows, else incremental."""
    three_yrs = (datetime.date.today() - datetime.timedelta(days=3*365)).strftime("%Y-%m-%d")
    if not OUT.exists():
        return three_yrs
    df_existing = pd.read_parquet(OUT)
    gsci_counts = df_existing[df_existing["RIC"].isin(GSCI_RICS)].groupby("RIC").size()
    if gsci_counts.empty or gsci_counts.min() < 400:
        print("GSCI RICs have insufficient history — full 3-year fetch.")
        return three_yrs
    return _incremental_start()


def _to_long(raw, rics: list[str]) -> pd.DataFrame:
    """Flatten a get_history result to long format (Date, RIC, Price)."""
    frames = []
    for ric in rics:
        try:
            col = raw[ric]
            tmp = col.dropna().reset_index()
            tmp.columns = ["Date", "Price"]
            tmp["RIC"] = ric
            frames.append(tmp)
        except Exception as e:
            print(f"  Warning: could not extract {ric}: {e}")
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def fetch() -> pd.DataFrame:
    gsci_start = _gsci_start()
    roll_start = _incremental_start()
    end        = datetime.date.today().isoformat()

    # ── 1. GSCI indices → TR.PriceClose ───────────────────────────────────────
    print(f"Fetching {len(GSCI_RICS)} GSCI RICs (TR.PriceClose) from {gsci_start} to {end} ...")
    try:
        raw_gsci = rd.get_history(
            universe=GSCI_RICS,
            fields=["TR.PriceClose"],
            start=gsci_start,
            end=end,
            interval="daily",
        )
        raw_gsci.index = pd.to_datetime(raw_gsci.index)
        df_gsci = _to_long(raw_gsci, GSCI_RICS)
        print(f"  GSCI: {len(df_gsci):,} rows fetched.")
    except Exception as e:
        print(f"  WARNING: GSCI fetch failed ({e}). Momentum will fall back to spot prices.")
        df_gsci = pd.DataFrame()

    # ── 2. Continuous futures → TR.SETTLEMENTPRICE (roll yield + fallback mom) ─
    print(f"Fetching {len(ROLL_RICS)} continuous RICs (TR.SETTLEMENTPRICE) from {roll_start} to {end} ...")
    raw_roll = rd.get_history(
        universe=ROLL_RICS,
        fields=["TR.SETTLEMENTPRICE"],
        start=roll_start,
        end=end,
        interval="daily",
    )
    raw_roll.index = pd.to_datetime(raw_roll.index)
    df_roll = _to_long(raw_roll, ROLL_RICS)

    df = pd.concat([df_gsci, df_roll], ignore_index=True)
    df["Date"] = pd.to_datetime(df["Date"])
    return df


def save(df_new: pd.DataFrame):
    if OUT.exists():
        df_old = pd.read_parquet(OUT)
        df_old["Date"] = pd.to_datetime(df_old["Date"])
        # Per-RIC cutoff: only remove old rows from the date that RIC's new data starts
        filtered = df_old.copy()
        for ric, grp in df_new.groupby("RIC"):
            ric_cutoff = grp["Date"].min()
            filtered = filtered[~((filtered["RIC"] == ric) & (filtered["Date"] >= ric_cutoff))]
        df = pd.concat([filtered, df_new], ignore_index=True)
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
