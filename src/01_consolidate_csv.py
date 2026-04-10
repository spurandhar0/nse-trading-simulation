"""
Script 1: Consolidate NSE Bhavcopy CSVs
========================================
Reads all CSV files from bhav_data/**/*.csv
Merges into a single Parquet file: db/consolidated.parquet

Expected CSV columns (NSE sec_bhavdata_full format):
  SYMBOL, SERIES, DATE1, PREV_CLOSE, OPEN_PRICE, HIGH_PRICE, LOW_PRICE,
  LAST_PRICE, CLOSE_PRICE, AVG_PRICE, TTL_TRD_QNTY, TURNOVER_LACS,
  NO_OF_TRADES, DELIV_QTY, DELIV_PER

DATE1 formats handled: DD-Mon-YYYY (01-JAN-2025) or YYYY-MM-DD
"""

import os
import glob
import pandas as pd
from datetime import datetime

BHAV_DIR   = "bhav_data"
OUTPUT_DIR = "db"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated.parquet")

REQUIRED_COLS = [
    "SYMBOL", "SERIES", "DATE1", "PREV_CLOSE", "OPEN_PRICE",
    "HIGH_PRICE", "LOW_PRICE", "LAST_PRICE", "CLOSE_PRICE",
    "AVG_PRICE", "TTL_TRD_QNTY", "TURNOVER_LACS",
    "NO_OF_TRADES", "DELIV_QTY", "DELIV_PER"
]

def parse_date(val):
    """Parse NSE date formats: DD-Mon-YYYY or YYYY-MM-DD."""
    if pd.isna(val):
        return pd.NaT
    s = str(val).strip()
    for fmt in ("%d-%b-%Y", "%d-%B-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return pd.NaT

def load_csv(filepath):
    """Load a single bhav copy CSV, normalize column names."""
    try:
        df = pd.read_csv(filepath, dtype=str, low_memory=False)
        # Normalize: strip whitespace, uppercase column names
        df.columns = [c.strip().upper().replace(" ", "_") for c in df.columns]

        # Rename common alternate column names
        rename_map = {
            "DATE": "DATE1",
            "TOTTRDQTY": "TTL_TRD_QNTY",
            "TOTTRDVAL": "TURNOVER_LACS",
            "TOTALTRADES": "NO_OF_TRADES",
        }
        df.rename(columns=rename_map, inplace=True)

        # Keep only known columns (fill missing with empty)
        for col in REQUIRED_COLS:
            if col not in df.columns:
                df[col] = ""

        df = df[REQUIRED_COLS].copy()

        # Clean symbol and series
        df["SYMBOL"] = df["SYMBOL"].str.strip().str.upper()
        df["SERIES"] = df["SERIES"].str.strip().str.upper()

        # Parse date
        df["DATE1"] = df["DATE1"].apply(parse_date)
        df = df.dropna(subset=["DATE1", "SYMBOL"])
        df = df[df["SYMBOL"] != ""]

        # Numeric columns
        num_cols = ["PREV_CLOSE", "OPEN_PRICE", "HIGH_PRICE", "LOW_PRICE",
                    "LAST_PRICE", "CLOSE_PRICE", "AVG_PRICE",
                    "TTL_TRD_QNTY", "TURNOVER_LACS", "NO_OF_TRADES",
                    "DELIV_QTY", "DELIV_PER"]
        for col in num_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

        return df

    except Exception as e:
        print(f"  ⚠️  Skipped {filepath}: {e}")
        return None

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    csv_files = sorted(glob.glob(os.path.join(BHAV_DIR, "**", "*.csv"), recursive=True))
    print(f"Found {len(csv_files)} CSV files in {BHAV_DIR}/")

    if not csv_files:
        print("❌ No CSV files found. Upload bhav copy CSVs to bhav_data/ folder.")
        raise SystemExit(1)

    frames = []
    for i, f in enumerate(csv_files, 1):
        print(f"  [{i}/{len(csv_files)}] Loading: {os.path.basename(f)}")
        df = load_csv(f)
        if df is not None and len(df) > 0:
            frames.append(df)

    if not frames:
        print("❌ No valid data loaded.")
        raise SystemExit(1)

    print("Merging all files...")
    combined = pd.concat(frames, ignore_index=True)

    # Drop exact duplicates (same symbol + date)
    before = len(combined)
    combined.drop_duplicates(subset=["SYMBOL", "DATE1"], keep="last", inplace=True)
    after = len(combined)
    print(f"Deduplication: {before:,} → {after:,} rows ({before - after:,} duplicates removed)")

    # Sort by symbol + date
    combined.sort_values(["SYMBOL", "DATE1"], inplace=True)
    combined.reset_index(drop=True, inplace=True)

    combined.to_parquet(OUTPUT_FILE, index=False)

    print(f"\n✅ Consolidated: {after:,} rows")
    print(f"✅ Date range : {combined['DATE1'].min().date()} → {combined['DATE1'].max().date()}")
    print(f"✅ Symbols    : {combined['SYMBOL'].nunique():,}")
    print(f"✅ Saved to   : {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
