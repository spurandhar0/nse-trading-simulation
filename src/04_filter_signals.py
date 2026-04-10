"""
Script 4: Filter Trading Signals
==================================
Reads:  db/eq_data.parquet, db/ath.parquet, config/simulation_params.json

Modes:
  --mode full  (default)  Sweeps all filter combos from param_sweep.filter
  --mode quick            Uses single filter set from quick_run section

Applies BOTH filters for each configured filter-param combination:

  Filter 1 (5-day dip):
    pct_from_low = (close - min_5day_close) / min_5day_close
    PASS: pct_min <= pct_from_low <= pct_max   (both negative, e.g. -0.10 to -0.05)

  Filter 2 (ATH distance):
    pct_from_ath = (close - ath) / ath
    PASS: ath_min <= pct_from_ath <= ath_max   (both negative, e.g. -0.60 to -0.30)

Output: db/signals.parquet  (one row per passing signal day per symbol per filter-combo)
        db/signals_quick.parquet  (quick mode output, does not overwrite full)
"""

import os
import sys
import json
import argparse
import itertools
import numpy as np
import pandas as pd
from datetime import datetime

CONFIG_FILE      = "config/simulation_params.json"
EQ_FILE          = "db/eq_data.parquet"
ATH_FILE         = "db/ath.parquet"
OUTPUT_FILE_FULL = "db/signals.parquet"
OUTPUT_FILE_QUICK= "db/signals_quick.parquet"

def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)

def build_filter_combos(cfg, mode):
    """Generate filter-param combinations based on mode."""
    if mode == "quick":
        q = cfg["quick_run"]
        # Single combo — wrap in list
        return [(q["days_back"], q["pct_min"], q["pct_max"],
                 q["ath_min"],   q["ath_max"])]

    # Full mode: sweep all param_sweep.filter ranges
    f = cfg["param_sweep"]["filter"]
    combos = []
    for db, pmin, pmax, amin, amax in itertools.product(
        f["days_back"], f["pct_min"], f["pct_max"], f["ath_min"], f["ath_max"]
    ):
        if pmin >= pmax:   # invalid: min must be more negative than max
            continue
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
    arr    = sym_df[["DATE1", "CLOSE_PRICE"]].values  # shape (N, 2)
    dates  = arr[:, 0]
    closes = arr[:, 1].astype(float)
    n      = len(dates)
    results = []

    for i in range(days_back, n):
        sig_date = dates[i]
        if sig_date < start_date or sig_date > end_date:
            continue

        today_close = closes[i]
        if today_close <= 0:
            continue

        # --- Filter 1: N-day dip (days_back previous trading sessions) ---
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

        # Min-close date
        min_idx  = i - days_back + int(np.argmin(lookback))
        min_date = dates[min_idx]

        # Previous close for 1-day change
        prev_close    = closes[i - 1] if i > 0 else 0
        pct_from_prev = ((today_close - prev_close) / prev_close) if prev_close > 0 else 0

        results.append({
            "SYMBOL":        sym_df["SYMBOL"].iloc[0],
            "SIGNAL_DATE":   sig_date,
            "SIGNAL_CLOSE":  today_close,
            "MIN_5D_CLOSE":  min_close,
            "MIN_5D_DATE":   min_date,
            "PCT_FROM_LOW":  round(pct_from_low,  6),
            "PCT_FROM_ATH":  round(pct_from_ath,  6),
            "PCT_1D_CHANGE": round(pct_from_prev, 6),
            "ATH_PRICE":     ath_price,
            "DAYSBACK":      days_back,
            "PCTMIN":        pct_min,
            "PCTMAX":        pct_max,
            "ATHMIN":        ath_min,
            "ATHMAX":        ath_max,   # ← fixed (was `amax` variable scope bug)
        })

    return results

def resolve_symbols(cli_symbols_str, cfg, mode):
    """
    Priority order:
      1. --symbols CLI arg (comma-separated string)  ← highest priority
      2. watch_symbols in config (quick_run or param_sweep section)
      3. Empty list → run ALL symbols
    Returns a set of uppercase symbol strings, or empty set (= all).
    """
    # 1. CLI arg
    if cli_symbols_str and cli_symbols_str.strip():
        syms = [s.strip().upper() for s in cli_symbols_str.split(",") if s.strip()]
        if syms:
            print(f"Symbol filter   : CLI override → {syms}")
            return set(syms)

    # 2. Config watch_symbols
    section = cfg.get("quick_run" if mode == "quick" else "param_sweep", {})
    config_syms = section.get("watch_symbols", [])
    if config_syms:
        syms = [s.strip().upper() for s in config_syms if s.strip()]
        if syms:
            print(f"Symbol filter   : config watch_symbols → {syms}")
            return set(syms)

    # 3. All symbols
    print("Symbol filter   : ALL symbols")
    return set()


def main():
    parser = argparse.ArgumentParser(description="Filter NSE trading signals")
    parser.add_argument("--mode", choices=["quick", "full"], default="full",
                        help="quick = single param set; full = all param sweep combos")
    parser.add_argument("--symbols", default="",
                        help="Comma-separated symbols to test, e.g. TCS,WIPRO,INFY (empty = all)")
    args = parser.parse_args()
    mode = args.mode

    for f in [CONFIG_FILE, EQ_FILE, ATH_FILE]:
        if not os.path.exists(f):
            print(f"❌ Missing: {f}")
            raise SystemExit(1)

    cfg        = load_config()
    start_date = pd.Timestamp(cfg["signal_start_date"])
    end_date   = pd.Timestamp(cfg["signal_end_date"])
    output_file = OUTPUT_FILE_QUICK if mode == "quick" else OUTPUT_FILE_FULL

    filter_combos = build_filter_combos(cfg, mode)
    print(f"Mode            : {mode.upper()}")
    print(f"Filter combos   : {len(filter_combos)}")

    # Resolve symbol filter (CLI > config > all)
    symbol_filter = resolve_symbols(args.symbols, cfg, mode)

    print("Loading EQ data...")
    eq = pd.read_parquet(EQ_FILE, columns=["SYMBOL", "DATE1", "CLOSE_PRICE"])
    eq["DATE1"] = pd.to_datetime(eq["DATE1"])
    eq.sort_values(["SYMBOL", "DATE1"], inplace=True)

    # Apply symbol filter if specified
    if symbol_filter:
        eq = eq[eq["SYMBOL"].isin(symbol_filter)]
        missing = symbol_filter - set(eq["SYMBOL"].unique())
        if missing:
            print(f"⚠️  Symbols not found in data: {sorted(missing)}")
        if eq.empty:
            print("❌ No data found for specified symbols. Check symbol names (must match NSE exact name).")
            raise SystemExit(1)

    print("Loading ATH data...")
    ath_df  = pd.read_parquet(ATH_FILE, columns=["SYMBOL", "ATH_PRICE"])
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

    EMPTY_COLS = [
        "SYMBOL","SIGNAL_DATE","SIGNAL_CLOSE","MIN_5D_CLOSE","MIN_5D_DATE",
        "PCT_FROM_LOW","PCT_FROM_ATH","PCT_1D_CHANGE","ATH_PRICE",
        "DAYSBACK","PCTMIN","PCTMAX","ATHMIN","ATHMAX"
    ]

    if not all_signals:
        print("⚠️  No signals found. Check your filter parameters and data date range.")
        pd.DataFrame(columns=EMPTY_COLS).to_parquet(output_file, index=False)
        raise SystemExit(0)

    sig_df = pd.DataFrame(all_signals)
    sig_df.to_parquet(output_file, index=False)

    print(f"✅ Signals saved       : {output_file}")
    print(f"✅ Unique symbols      : {sig_df['SYMBOL'].nunique():,}")
    print(f"✅ Date range          : {sig_df['SIGNAL_DATE'].min()} → {sig_df['SIGNAL_DATE'].max()}")

if __name__ == "__main__":
    main()
