"""
NSE Bhavcopy Fetcher
=====================
Downloads today's NSE EQ bhavcopy and saves to bhav_data/
Outputs GitHub Actions step output: new_file=true/false
"""

import argparse
import csv
import glob
import io
import os
from datetime import date, datetime, timedelta, timezone

import requests

IST     = timezone(timedelta(hours=5, minutes=30))
NOW_IST = datetime.now(tz=IST)
BHAV_ROOT = "bhav_data"

# NSE capital market circular NSE/CMTR/71775.
NSE_HOLIDAYS_2026 = {
    date(2026, 1, 26), date(2026, 3, 3), date(2026, 3, 26),
    date(2026, 3, 31), date(2026, 4, 3), date(2026, 4, 14),
    date(2026, 5, 1), date(2026, 5, 28), date(2026, 6, 26),
    date(2026, 9, 14), date(2026, 10, 2), date(2026, 10, 20),
    date(2026, 11, 10), date(2026, 11, 24), date(2026, 12, 25),
}
NSE_MUHURAT_2026 = date(2026, 11, 8)


def set_gha_output(key, value):
    gha_output = os.environ.get("GITHUB_OUTPUT", "")
    if gha_output:
        with open(gha_output, "a") as f:
            f.write(f"{key}={value}\n")


def fetch_bhav(session_date, *, allow_missing=False):
    dd, mm, yyyy, mon3 = session_date.strftime("%d %m %Y %b").split()

    fname    = f"sec_bhavdata_full_{dd}{mm}{yyyy}.csv"
    savepath = os.path.join(BHAV_ROOT, fname)

    date_pattern = f"*_{dd}{mm}{yyyy}.csv"
    existing = (
        glob.glob(os.path.join(BHAV_ROOT, "**", date_pattern), recursive=True)
        + glob.glob(os.path.join(BHAV_ROOT, date_pattern))
    )
    existing = list(set(existing))

    if existing:
        print(f"\u26a0\ufe0f  File for {dd}-{mon3}-{yyyy} already exists: {existing[0]}")
        print("\u23ed\ufe0f  Skipping download.")
        set_gha_output("new_file", "false")
        return True

    url = (
        f"https://nsearchives.nseindia.com/products/content/"
        f"sec_bhavdata_full_{dd}{mm}{yyyy}.csv"
    )
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/124.0.0.0 Safari/537.36",
        "Referer": "https://www.nseindia.com",
        "Accept": "text/html,application/xhtml+xml,*/*",
        "Accept-Language": "en-US,en;q=0.9",
    }
    print(f"Downloading bhav: {url}")
    resp = requests.get(url, headers=headers, timeout=60)
    if allow_missing and resp.status_code == 404:
        print(f"No archive file for {session_date}; skipping.")
        return False
    resp.raise_for_status()

    reader = csv.DictReader(io.StringIO(resp.content.decode("utf-8-sig")))
    reader.fieldnames = [name.strip().upper() for name in (reader.fieldnames or [])]
    expected_date = session_date.strftime("%d-%b-%Y").upper()
    if not {"SYMBOL", "SERIES", "DATE1", "OPEN_PRICE", "HIGH_PRICE",
            "LOW_PRICE", "CLOSE_PRICE"}.issubset(reader.fieldnames):
        raise ValueError(f"Unexpected bhavcopy columns for {session_date}")
    if not any(
        (row["SERIES"] or "").strip().upper() == "EQ"
        and (row["DATE1"] or "").strip().upper() == expected_date
        for row in reader
    ):
        raise ValueError(f"Bhavcopy contains no EQ rows for {session_date}")

    os.makedirs(BHAV_ROOT, exist_ok=True)
    with open(savepath, "wb") as f:
        f.write(resp.content)

    size_kb = os.path.getsize(savepath) / 1024
    with open(savepath, "r") as f:
        row_count = sum(1 for _ in f) - 1

    print(f"\u2705 Saved   : {savepath} ({size_kb:.1f} KB)")
    print(f"\u2705 Date    : {dd}-{mon3}-{yyyy}")
    print(f"\u2705 Rows    : {row_count:,}")
    set_gha_output("new_file", "true")
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download NSE equity bhavcopies")
    parser.add_argument("--session-date", type=date.fromisoformat, default=NOW_IST.date())
    parser.add_argument("--from-date", type=date.fromisoformat)
    args = parser.parse_args()
    day = args.from_date or args.session_date
    if day > args.session_date:
        parser.error("--from-date must be on or before --session-date")
    while day < args.session_date:
        if day == NSE_MUHURAT_2026 or (day.weekday() < 5 and day not in NSE_HOLIDAYS_2026):
            fetch_bhav(day, allow_missing=True)
        day += timedelta(days=1)
    fetch_bhav(args.session_date)
