"""
Script 4: Filter Trading Signals
==================================
Reads:  db/eq_data.parquet, db/ath.parquet, config/simulation_params.json
Applies BOTH filters for each configured filter-param combination:

  Filter 1 (5-day dip):
    pct_from_low = (close - min_5day_close) / min_5day_close
    PASS: pct_min <= pct_from_low <= pct_max   (both negative, e.g. -0.10 to -0.05)

  Filter 2 (ATH distance):
    pct_from_ath = (close - ath) / ath
    PASS: ath_min <= pct_from_ath <= ath_max   (both negative, e.g. -0.60 to -0.30)

Output: db/signals.parquet  (one row per passing signal day per symbol per filter-combo)
"""

import os
import json
import itertools
import numpy as np
import pandas as pd
from datetime import datetime

CONFIG_FILE = "config/simulation_params.json"
EQ_FILE     = "db/eq_data.parquet"
ATH_FILE    = "db/ath.parquet"
OUTPUT_FILE = "db/signals.parquet"

def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)

def build_filter_combos(cfg):
    """Generate all valid filter-param combinations."""
    f = cfg["filter"]
    combos = []
    for db, pmin, pmax, amin, amax in itertools.product(
        f["days_back"], f["pct_min"], f["pct_max"], f["ath_min"], f["ath_max"]
    ):
        # Validate: pct_min < pct_max (both negative, min more negative)
        if pmin >= pmax:
            continue
        # Validate: ath_min < ath_max (both negative, min more negative)
        if amin >= amax:
            continue
        combos.append((db, pmin, pmax, amin, amax))
    return combos

def filter_symbol(sym_df, ath_price, days_back, pct_min, pct_max,
                  ath_min, ath_max, start_date, end_date):
    """
    Apply both filters to a single symbol's price history.
    Returns list of dicts for each passing signal day.
    """
    arr = sym_df[["DATE1", "CLOSE_PRICE"]].values  # shape (N, 2)
    dates  = arr[:, 0]
    closes = arr[:, 1].astype(float)
    n = len(dates)
    results = []

    for i in range(days_back, n):
        sig_date = dates[i]
        if sig_date < start_date or sig_date > end_date:
            continue

        today_close = closes[i]
        if today_close <= 0:
            continue

        # --- Filter 1: 5-day dip (daysBack previous trading sessions) ---
        lookback = closes[i - days_back: i]
        if len(lookback) < days_back:
            continue
        min_close = np.min(lookback)
        if min_close <= 0:
            continue
        pct_from_low = (today_close - min_close) / min_close

        if not (pct_min <= pct_from_low <= pct_max):
            continue

        # --- Filter 2: ATH distance ---
        pct_from_ath = (today_close - ath_price) / ath_price
        if not (ath_min <= pct_from_ath <= ath_max):
            continue

        # Find min_close date
        min_idx = i - days_back + int(np.argmin(lookback))
        min_date = dates[min_idx]

        # Previous close for 1-day change
        prev_close = closes[i - 1] if i > 0 else 0
        pct_from_prev = ((today_close - prev_close) / prev_close) if prev_close > 0 else 0

        results.append({
            "SYMBOL":       sym_df["SYMBOL"].iloc[0],
            "SIGNAL_DATE":  sig_date,
            "SIGNAL_CLOSE": today_close,
            "MIN_5D_CLOSE": min_close,
            "MIN_5D_DATE":  min_date,
            "PCT_FROM_LOW": round(pct_from_low, 6),
            "PCT_FROM_ATH": round(pct_from_ath, 6),
            "PCT_1D_CHANGE":round(pct_from_prev, 6),
            "ATH_PRICE":    ath_price,
            "DAYSBACK":     days_back,
            "PCTMIN":       pct_min,
            "PCTMAX":       pct_max,
            "ATHMIN":       ath_min,
            "ATHMAX":       amax,
        })

    return results

def main():
    for f in [CONFIG_FILE, EQ_FILE, ATH_FILE]:
        if not os.path.exists(f):
            print(f"❌ Missing: {f}")
            raise SystemExit(1)

    cfg = load_config()
    start_date = pd.Timestamp(cfg["signal_start_date"])
    end_date   = pd.Timestamp(cfg["signal_end_date"])

    filter_combos = build_filter_combos(cfg)
    print(f"Filter combinations to test: {len(filter_combos)}")

    print("Loading EQ data...")
    eq = pd.read_parquet(EQ_FILE, columns=["SYMBOL", "DATE1", "CLOSE_PRICE"])
    eq["DATE1"] = pd.to_datetime(eq["DATE1"])
    eq.sort_values(["SYMBOL", "DATE1"], inplace=True)

    print("Loading ATH data...")
    ath_df = pd.read_parquet(ATH_FILE, columns=["SYMBOL", "ATH_PRICE"])
    ath_map = dict(zip(ath_df["SYMBOL"], ath_df["ATH_PRICE"]))

    symbols = eq["SYMBOL"].unique()
    print(f"Processing {len(symbols):,} symbols × {len(filter_combos)} filter combos...")

    all_signals = []
    for sym_idx, sym in enumerate(symbols):
        if (sym_idx + 1) % 500 == 0:
            print(f"  [{sym_idx+1}/{len(symbols)}] signals so far: {len(all_signals):,}")

        ath_price = ath_map.get(sym, 0)
        if ath_price <= 0:
            continue

        sym_df = eq[eq["SYMBOL"] == sym].reset_index(drop=True)
        if len(sym_df) < 10:
            continue

        for (db, pmin, pmax, amin, amax) in filter_combos:
            signals = filter_symbol(sym_df, ath_price, db, pmin, pmax,
                                    amin, amax, start_date, end_date)
            all_signals.extend(signals)

    print(f"\nTotal signals found: {len(all_signals):,}")

    if not all_signals:
        print("⚠️  No signals found. Check your filter parameters and data date range.")
        # Write empty file so downstream scripts don't fail
        pd.DataFrame(columns=[
            "SYMBOL","SIGNAL_DATE","SIGNAL_CLOSE","MIN_5D_CLOSE","MIN_5D_DATE",
            "PCT_FROM_LOW","PCT_FROM_ATH","PCT_1D_CHANGE","ATH_PRICE",
            "DAYSBACK","PCTMIN","PCTMAX","ATHMIN","ATHMAX"
        ]).to_parquet(OUTPUT_FILE, index=False)
        raise SystemExit(0)

    sig_df = pd.DataFrame(all_signals)
    sig_df.to_parquet(OUTPUT_FILE, index=False)

    print(f"✅ Signals saved: {OUTPUT_FILE}")
    print(f"✅ Unique symbols with signals: {sig_df['SYMBOL'].nunique():,}")
    print(f"✅ Date range of signals: {sig_df['SIGNAL_DATE'].min()} → {sig_df['SIGNAL_DATE'].max()}")

if __name__ == "__main__":
    main()
