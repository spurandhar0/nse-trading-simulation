"""
Script 2: Filter EQ Series Only
=================================
Reads: db/consolidated.parquet
Keeps: only rows where SERIES == 'EQ'
Output: db/eq_data.parquet
"""

import os
import pandas as pd

INPUT_FILE  = "db/consolidated.parquet"
OUTPUT_FILE = "db/eq_data.parquet"

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Input not found: {INPUT_FILE}")
        print("   Run 01_consolidate_csv.py first.")
        raise SystemExit(1)

    print(f"Loading {INPUT_FILE} ...")
    df = pd.read_parquet(INPUT_FILE)
    print(f"  Total rows (all series): {len(df):,}")
    print(f"  Series found: {sorted(df['SERIES'].unique())}")

    eq = df[df["SERIES"] == "EQ"].copy()
    print(f"  EQ rows kept          : {len(eq):,}")
    print(f"  EQ symbols            : {eq['SYMBOL'].nunique():,}")

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
