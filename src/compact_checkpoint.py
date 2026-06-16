"""
compact_checkpoint.py — Run at the start of each sweep to keep partial
chunk files clean of duplicate rows.

Each append-mode resume accumulates duplicate rows within a chunk when the
job is killed mid-batch. This script deduplicates every chunk file in
output/partial/, keeping only the last row per Test number.

Also handles the legacy single-file case (full_sweep_partial.csv) for
backwards compatibility.

FIX: Now also strips duplicate header rows (where "Test" is the string
"Test") that were injected by a cross-day resume bug (write_hdr logic).
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


def _compact_csv(path: str) -> tuple[int, int, float, float]:
    """
    Deduplicate a single CSV in-place. Returns (before_rows, after_rows,
    before_mb, after_mb). Raises on unrecoverable errors.
    """
    before_size = os.path.getsize(path)
    df = pd.read_csv(path)
    before_rows = len(df)

    # FIX: coerce "Test" column to numeric and drop non-numeric rows.
    # This removes any duplicate header lines that were accidentally appended
    # when a cross-day resume wrote write_hdr=True to an existing file.
    df["Test"] = pd.to_numeric(df["Test"], errors="coerce")
    df = df.dropna(subset=["Test"])
    df["Test"] = df["Test"].astype(int)

    df = (
        df.drop_duplicates(subset=["Test"], keep="last")
          .sort_values("Test")
          .reset_index(drop=True)
    )
    df.to_csv(path, index=False)
    after_size = os.path.getsize(path)
    return before_rows, len(df), before_size / 1024 / 1024, after_size / 1024 / 1024


# ── Compact legacy single CSV (if it exists from an old run) ──────────────────
if os.path.exists(LEGACY_CSV):
    try:
        br, ar, bmb, amb = _compact_csv(LEGACY_CSV)
        print(
            f"[legacy] Compacted full_sweep_partial.csv: "
            f"{br} -> {ar} rows ({bmb:.1f} MB -> {amb:.1f} MB)"
        )
    except Exception as e:
        print(f"WARN: could not compact {LEGACY_CSV}: {e}")

# ── Compact each chunk_*.csv file ─────────────────────────────────────────────
chunk_files = sorted(glob.glob(os.path.join(PARTIAL_DIR, "chunk_*.csv")))

if not chunk_files:
    print("No chunk files found — nothing to compact.")
    sys.exit(0)

total_before    = 0
total_after     = 0
total_before_mb = 0.0
total_after_mb  = 0.0

for cf in chunk_files:
    try:
        br, ar, bmb, amb = _compact_csv(cf)
        total_before    += br
        total_after     += ar
        total_before_mb += bmb
        total_after_mb  += amb
        if br != ar:
            print(
                f"  {os.path.basename(cf)}: {br} -> {ar} rows "
                f"({bmb:.1f} MB -> {amb:.1f} MB)"
            )
    except Exception as e:
        print(f"  WARN: could not compact {cf}: {e}")

print(
    f"\nCompact complete: {len(chunk_files)} chunk(s) | "
    f"{total_before:,} -> {total_after:,} rows | "
    f"{total_before_mb:.1f} MB -> {total_after_mb:.1f} MB total"
)
