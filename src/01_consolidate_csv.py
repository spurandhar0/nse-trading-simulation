"""
Script 1: Consolidate NSE Bhavcopy CSVs
========================================
Reads all CSV files from bhav_data/**/*.csv
Merges into a single Parquet file: db/consolidated.parquet

File naming:  NSE bhav files end with 8 digits = DDMMYYYY
              e.g.  sec_bhavdata_full_01032025.csv  ->  01-Mar-2025
              Files are sorted date-ascending (oldest first) before loading.
              Duplicate files sharing the same date (last 8 chars) are skipped.

Row-level deduplication: keeps last row for each (SYMBOL, DATE1) pair.

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

BHAV_DIR    = "bhav_data"
OUTPUT_DIR  = "db"
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


def extract_file_date(filepath):
    """
    Extract date from last 8 characters of filename (excluding extension).
    NSE bhav files end with DDMMYYYY, e.g. sec_bhavdata_full_01032025.csv
    Returns a datetime object or datetime.min if not parseable.
    """
    name = os.path.splitext(os.path.basename(filepath))[0]
    if len(name) >= 8:
        date_str = name[-8:]   # last 8 chars = DDMMYYYY
        try:
            return datetime.strptime(date_str, "%d%m%Y")
        except ValueError:
            pass
    return datetime.min  # fallback — cannot parse, will sort first


def load_csv(filepath):
    """Load a single bhav copy CSV, normalize column names."""
    try:
        df = pd.read_csv(filepath, dtype=str, low_memory=False)
        # Normalize: strip whitespace, uppercase column names
        df.columns = [c.strip().upper().replace(" ", "_") for c in df.columns]

        # Rename common alternate column names
        rename_map = {
            "DATE":        "DATE1",
            "TOTTRDQTY":   "TTL_TRD_QNTY",
            "TOTTRDVAL":   "TURNOVER_LACS",
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

    # ── Discover all CSV files ──────────────────────────────────────────────
    csv_files = glob.glob(os.path.join(BHAV_DIR, "**", "*.csv"), recursive=True)
    print(f"Found {len(csv_files)} CSV files in {BHAV_DIR}/")

    if not csv_files:
        print("❌ No CSV files found. Upload bhav copy CSVs to bhav_data/ folder.")
        raise SystemExit(1)

    # ── Sort files date-ascending by last 8 digits of filename (DDMMYYYY) ──
    csv_files = sorted(csv_files, key=extract_file_date)

    # ── Remove duplicate files with the same date (keep first occurrence) ──
    seen_dates = {}
    unique_files = []
    skipped_dup = []
    for f in csv_files:
        d = extract_file_date(f)
        # Use the date string as key; files with unparseable names get unique key
        if d == datetime.min:
            key = f  # can't determine date — include unconditionally
        else:
            key = d.strftime("%Y%m%d")

        if key not in seen_dates:
            seen_dates[key] = f
            unique_files.append(f)
        else:
            skipped_dup.append(os.path.basename(f))

    if skipped_dup:
        print(f"Duplicate date files skipped ({len(skipped_dup)}):")
        for name in skipped_dup:
            print(f"  ⚠️  {name}")

    csv_files = unique_files
    print(f"Unique date files to load: {len(csv_files)}")

    # Print date range detected from filenames
    parseable = [extract_file_date(f) for f in csv_files if extract_file_date(f) != datetime.min]
    if parseable:
        print(f"File date range  : {min(parseable).strftime('%d-%b-%Y')} → {max(parseable).strftime('%d-%b-%Y')}")

    # ── Load each file ──────────────────────────────────────────────────────
    frames = []
    for i, f in enumerate(csv_files, 1):
        d = extract_file_date(f)
        date_label = d.strftime("%d-%b-%Y") if d != datetime.min else "?"
        print(f"  [{i:3d}/{len(csv_files)}] {date_label}  {os.path.basename(f)}")
        df = load_csv(f)
        if df is not None and len(df) > 0:
            frames.append(df)

    if not frames:
        print("❌ No valid data loaded.")
        raise SystemExit(1)

    # ── Merge ───────────────────────────────────────────────────────────────
    print("Merging all files...")
    combined = pd.concat(frames, ignore_index=True)

    # Row-level deduplication: same symbol + same date → keep last
    before = len(combined)
    combined.drop_duplicates(subset=["SYMBOL", "DATE1"], keep="last", inplace=True)
    after = len(combined)
    print(f"Row deduplication: {before:,} → {after:,} rows ({before - after:,} removed)")

    # Sort by symbol + date (ascending)
    combined.sort_values(["SYMBOL", "DATE1"], inplace=True)
    combined.reset_index(drop=True, inplace=True)

    combined.to_parquet(OUTPUT_FILE, index=False)

    print(f"\n✅ Consolidated: {after:,} rows")
    print(f"✅ Date range : {combined['DATE1'].min().date()} → {combined['DATE1'].max().date()}")
    print(f"✅ Symbols    : {combined['SYMBOL'].nunique():,}")
    print(f"✅ Saved to   : {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
