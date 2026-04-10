"""
Script 3: Find All Time High (ATH) Per Symbol
===============================================
Reads: db/eq_data.parquet
Finds: max HIGH_PRICE and date for each symbol across all history
Output: db/ath.parquet  (columns: SYMBOL, ATH_PRICE, ATH_DATE)

This is equivalent to modTemplateUpdate: Update_ATH_In_Template
"""

import os
import pandas as pd

INPUT_FILE  = "db/eq_data.parquet"
OUTPUT_FILE = "db/ath.parquet"

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ Input not found: {INPUT_FILE}")
        print("   Run 02_filter_eq.py first.")
        raise SystemExit(1)

    print(f"Loading {INPUT_FILE} ...")
    df = pd.read_parquet(INPUT_FILE, columns=["SYMBOL", "DATE1", "HIGH_PRICE"])
    print(f"  Rows: {len(df):,}  |  Symbols: {df['SYMBOL'].nunique():,}")

    # For each symbol find the row with max HIGH_PRICE
    print("Computing ATH per symbol...")
    idx = df.groupby("SYMBOL")["HIGH_PRICE"].idxmax()
    ath = df.loc[idx, ["SYMBOL", "DATE1", "HIGH_PRICE"]].copy()
    ath.rename(columns={"DATE1": "ATH_DATE", "HIGH_PRICE": "ATH_PRICE"}, inplace=True)
    ath.reset_index(drop=True, inplace=True)

    # Remove any with ATH_PRICE <= 0
    ath = ath[ath["ATH_PRICE"] > 0]

    print(f"\nATH computed for {len(ath):,} symbols")
    print(ath.describe())

    ath.to_parquet(OUTPUT_FILE, index=False)
    print(f"\n✅ ATH saved: {OUTPUT_FILE}")

    # Sample output
    print("\nSample ATH records:")
    print(ath.head(10).to_string(index=False))

if __name__ == "__main__":
    main()
