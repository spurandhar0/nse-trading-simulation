"""
compact_checkpoint.py — Run at the start of each sweep to keep partial
chunk files clean of duplicate rows.

Each append-mode resume accumulates duplicate rows within a chunk when the
job is killed mid-batch. This script deduplicates every chunk file in
output/partial/, keeping only the last row per Test number.

Also handles the legacy single-file case (full_sweep_partial.csv) for
backwards compatibility.
"""
import os
import sys
import glob

PARTIAL_DIR = "output/partial"
LEGACY_CSV  = os.path.join(PARTIAL_DIR, "full_sweep_partial.csv")

if not os.path.isdir(PARTIAL_DIR):
    print("No partial/ directory found — starting fresh.")
    sys.exit(0)

import pandas as pd

# ── Compact legacy single CSV (if it exists from an old run) ──────────────────
if os.path.exists(LEGACY_CSV):
    before_size = os.path.getsize(LEGACY_CSV)
    df = pd.read_csv(LEGACY_CSV)
    before_rows = len(df)
    df = (
        df.drop_duplicates(subset=["Test"], keep="last")
          .sort_values("Test")
          .reset_index(drop=True)
    )
    df.to_csv(LEGACY_CSV, index=False)
    after_size = os.path.getsize(LEGACY_CSV)
    print(
        f"[legacy] Compacted full_sweep_partial.csv: "
        f"{before_rows} -> {len(df)} rows  "
        f"({before_size/1024/1024:.1f} MB -> {after_size/1024/1024:.1f} MB)"
    )

# ── Compact each chunk_*.csv file ─────────────────────────────────────────────
chunk_files = sorted(glob.glob(os.path.join(PARTIAL_DIR, "chunk_*.csv")))

if not chunk_files:
    print("No chunk files found — nothing to compact.")
    sys.exit(0)

total_before = 0
total_after  = 0
total_before_mb = 0.0
total_after_mb  = 0.0

for cf in chunk_files:
    try:
        before_size = os.path.getsize(cf)
        df = pd.read_csv(cf)
        before_rows = len(df)
        df = (
            df.drop_duplicates(subset=["Test"], keep="last")
              .sort_values("Test")
              .reset_index(drop=True)
        )
        df.to_csv(cf, index=False)
        after_size = os.path.getsize(cf)
        total_before     += before_rows
        total_after      += len(df)
        total_before_mb  += before_size / 1024 / 1024
        total_after_mb   += after_size  / 1024 / 1024
        if before_rows != len(df):
            print(
                f"  {os.path.basename(cf)}: {before_rows} -> {len(df)} rows  "
                f"({before_size/1024/1024:.1f} MB -> {after_size/1024/1024:.1f} MB)"
            )
    except Exception as e:
        print(f"  WARN: could not compact {cf}: {e}")

print(
    f"\nCompact complete: {len(chunk_files)} chunk(s)  |  "
    f"{total_before:,} -> {total_after:,} rows  |  "
    f"{total_before_mb:.1f} MB -> {total_after_mb:.1f} MB total"
)
