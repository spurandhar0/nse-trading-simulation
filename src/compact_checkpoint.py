"""
compact_checkpoint.py — Run at the start of each sweep to keep the partial
checkpoint CSV from growing too large.

Each append-mode resume accumulates duplicate rows. After 5 days of 350-min
runs the file reached ~93 MB (GitHub hard limit = 100 MB). This script keeps
only the last row per Test number, bringing the file back to ~1-3 MB.
"""
import os
import sys

CSV = "output/partial/full_sweep_partial.csv"

if not os.path.exists(CSV):
    print("No checkpoint CSV found — starting fresh.")
    sys.exit(0)

import pandas as pd

before_size = os.path.getsize(CSV)
df = pd.read_csv(CSV)
before_rows = len(df)

df = (
    df.drop_duplicates(subset=["Test"], keep="last")
    .sort_values("Test")
    .reset_index(drop=True)
)
df.to_csv(CSV, index=False)

after_size = os.path.getsize(CSV)
after_rows = len(df)

print(
    f"Compacted: {before_rows} -> {after_rows} rows  "
    f"({before_size/1024/1024:.1f} MB -> {after_size/1024/1024:.1f} MB)"
)
