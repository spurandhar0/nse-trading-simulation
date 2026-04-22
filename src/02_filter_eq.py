"""
Script 2: Filter EQ Series Only
=================================
Reads: db/consolidated.parquet
Keeps: only rows where SERIES == 'EQ' AND symbol's LATEST series is still 'EQ'

Why the latest-series check matters:
  A stock like FAZE3Q may have been EQ for years but recently moved to BE/BZ/SM.
  Its old rows still say SERIES='EQ', so a simple row filter keeps them.
  We must check the symbol's most recent bhav record — if that record shows a
  non-EQ series, the symbol is excluded entirely (all historical rows dropped).

Output: db/eq_data.parquet
        db/series_changed.parquet  (symbols that were EQ but are now non-EQ)
"""

import os
import pandas as pd

INPUT_FILE          = "db/consolidated.parquet"
OUTPUT_FILE         = "db/eq_data.parquet"
SERIES_CHANGED_FILE = "db/series_changed.parquet"

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Input not found: {INPUT_FILE}")
        print("   Run 01_consolidate_csv.py first.")
        raise SystemExit(1)

    print(f"Loading {INPUT_FILE} ...")
    df = pd.read_parquet(INPUT_FILE)
    print(f"  Total rows (all series): {len(df):,}")
    print(f"  Series found: {sorted(df['SERIES'].unique())}")

    # ── Step 1: Find each symbol's LATEST series ──────────────────────────────
    # Sort by DATE1 so idxmax gives the most recent row per symbol
    df["DATE1"] = pd.to_datetime(df["DATE1"])
    latest_idx    = df.groupby("SYMBOL")["DATE1"].idxmax()
    latest_series = df.loc[latest_idx, ["SYMBOL", "SERIES", "DATE1"]].copy()
    latest_series.set_index("SYMBOL", inplace=True)

    # ── Step 2: Symbols whose latest series is NOT EQ ────────────────────────
    was_eq_now_not = latest_series[latest_series["SERIES"] != "EQ"].copy()
    was_eq_now_not = was_eq_now_not[
        was_eq_now_not.index.isin(df[df["SERIES"] == "EQ"]["SYMBOL"].unique())
    ]

    if len(was_eq_now_not):
        print(f"\n⚠️  Symbols that HAD EQ history but latest series is NOT EQ ({len(was_eq_now_not)}):")
        for sym, row in was_eq_now_not.iterrows():
            print(f"   {sym:20s} → latest series: {row['SERIES']}  (as of {row['DATE1'].date()})")
        # Save for use by simulation script to generate correct Invalid remarks
        was_eq_now_not.reset_index().rename(
            columns={"SERIES": "LATEST_SERIES", "DATE1": "LATEST_DATE"}
        ).to_parquet(SERIES_CHANGED_FILE, index=False)
        print(f"✅ Saved series-changed list: {SERIES_CHANGED_FILE}")
    else:
        print("✅ No symbols found with series change from EQ.")
        # Write empty file so simulation script can safely read it
        pd.DataFrame(columns=["SYMBOL", "LATEST_SERIES", "LATEST_DATE"]).to_parquet(
            SERIES_CHANGED_FILE, index=False
        )

    # ── Step 3: Symbols whose LATEST series IS EQ (safe to include) ──────────
    eq_symbols = set(latest_series[latest_series["SERIES"] == "EQ"].index)

    # ── Step 4: Keep only EQ-series rows AND only for symbols still on EQ ────
    eq = df[(df["SERIES"] == "EQ") & (df["SYMBOL"].isin(eq_symbols))].copy()
    print(f"\n  EQ rows kept (latest-series validated): {len(eq):,}")
    print(f"  EQ symbols                             : {eq['SYMBOL'].nunique():,}")
    excluded = df[df["SERIES"] == "EQ"]["SYMBOL"].nunique() - eq["SYMBOL"].nunique()
    if excluded:
        print(f"  Symbols excluded (series changed)      : {excluded}")

    if len(eq) == 0:
        print("❌ No EQ rows found! Check SERIES column in your CSV files.")
        raise SystemExit(1)

    eq.sort_values(["SYMBOL", "DATE1"], inplace=True)
    eq.reset_index(drop=True, inplace=True)
    eq.to_parquet(OUTPUT_FILE, index=False)

    print(f"\n✅ EQ data saved: {OUTPUT_FILE}")
    print(f"✅ Date range  : {eq['DATE1'].min().date()} → {eq['DATE1'].max().date()}")

if __name__ == "__main__":
    main()
