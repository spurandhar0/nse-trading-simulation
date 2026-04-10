# NSE Trading Simulation — GitHub Automation

> Zero PC dependency. Upload NSE data → click Run → download Excel results.

---

## Folder Structure

```
├── src/
│   ├── nse_bhavcopy_fetch.py       # Auto-fetches daily bhav from NSE
│   ├── 01_consolidate_csv.py       # Merge all bhav CSVs → db/
│   ├── 02_filter_eq.py             # Keep EQ series only
│   ├── 03_find_ath.py              # All Time High per symbol
│   ├── 04_filter_signals.py        # 5-day dip + ATH distance filter
│   ├── 05_run_simulation.py        # 2000-combo parameter test → Excel
│   └── 06_consolidate_results.py   # Merge multiple result Excels
├── config/
│   └── simulation_params.json      # ← Edit filter/trade params here
├── .github/workflows/
│   ├── nse_bhavcopy.yml            # Daily 6:15 PM IST auto-fetch
│   └── run_simulation.yml          # Manual pipeline trigger
├── bhav_data/                      # Upload historical CSVs here
│   └── Apr-2026/                   # Auto-created monthly folders
├── db/                             # Auto-generated intermediate files
└── output/                         # Excel results saved here
```

---

## One-Time Setup

### 1. Upload Historical Bhav Copy Files

Upload your past NSE bhav copy CSV files into `bhav_data/` using this structure:
```
bhav_data/
  Jan-2020/sec_bhavdata_full_01012020.csv
  Feb-2020/sec_bhavdata_full_03022020.csv
  ...
```

Each CSV must contain these NSE columns:
```
SYMBOL, SERIES, DATE1, PREV_CLOSE, OPEN_PRICE, HIGH_PRICE, LOW_PRICE,
LAST_PRICE, CLOSE_PRICE, AVG_PRICE, TTL_TRD_QNTY, TURNOVER_LACS,
NO_OF_TRADES, DELIV_QTY, DELIV_PER
```

### 2. Configure Parameters (optional)

Edit `config/simulation_params.json` to change:
- Signal date range (`signal_start_date`, `signal_end_date`)
- Filter thresholds (`days_back`, `pct_min`, `pct_max`, `ath_min`, `ath_max`)
- Trade parameters (`max_buys`, `buy_drop`, `target`, `stoploss`, `max_duration`)

---

## Running the Simulation

### Go to → Actions tab → "Run Full Simulation Pipeline" → Run workflow

You can select which steps to run:

| Step | Script | When to skip |
|------|--------|--------------|
| 1 – Consolidate CSVs | `01_consolidate_csv.py` | If you only added new bhav files and db/ is old |
| 2 – Filter EQ | `02_filter_eq.py` | Same as above |
| 3 – Compute ATH | `03_find_ath.py` | If ATH data is current |
| 4 – Filter Signals | `04_filter_signals.py` | If filter params unchanged |
| 5 – Run Simulation | `05_run_simulation.py` | Never skip this |
| 6 – Consolidate Results | `06_consolidate_results.py` | Only if merging multiple runs |

### First-ever run: Enable ALL steps (default).

Results appear in `output/Results_YYYYMMDD_HHMMSS.xlsx` — committed automatically.

---

## Output: 43 Columns (exact order)

| # | Column | Description |
|---|--------|-------------|
| 1 | Test | Row number |
| 2 | DAYSBACK | Lookback days for 5-day dip |
| 3 | PCTMIN | Min dip threshold (e.g. -0.10) |
| 4 | PCTMAX | Max dip threshold (e.g. -0.05) |
| 5 | ATHMIN | Min ATH distance (e.g. -0.60) |
| 6 | ATHMAX | Max ATH distance (e.g. -0.30) |
| 7 | MAXBUYS | Max additional buys |
| 8 | BUYDROP | Buy-down % per additional buy |
| 9 | TARGET | Target profit % |
| 10 | STOPLOSS | Stoploss % |
| 11 | MAXDURA | Max market days to hold |
| 12 | WinRate | Win % of closed trades |
| 13 | TotalTrade | Total executed trades |
| 14 | Executed | Trades that got first buy |
| 15 | Open | Currently open trades |
| 16 | Closed | Closed trades |
| 17 | ProfitTGT | Target achieved count |
| 18 | LossSL | Stoploss triggered count |
| 19 | LossFEMD | Force Exit Market Days — loss |
| 20 | LossFECD | Force Exit Calendar Days — loss |
| 21 | ProfitFEMD | Force Exit Market Days — profit |
| 22 | ProfitFECD | Force Exit Calendar Days — profit |
| 23 | Pending | Pending (not triggered yet) |
| 24 | Expired | Expired (never triggered) |
| 25 | Invalid | Invalid (no data) |
| 26 | TotalRows | Total input rows |
| 27 | Wins | Total winning trades |
| 28 | Losses | Total losing trades |
| 29 | TotalStock | Total symbols tested |
| 30 | SumProfit | Sum of profit/loss in ₹ |
| 31 | SumGainFin | Sum of gain% across closed trades |
| 32–39 | Dur5–Dur40 | Duration bucket counts (≤5d, ≤10d, …, ≤40d) |
| 40 | ExitTGT | Exit: target count |
| 41 | ExitSL | Exit: stoploss count |
| 42 | ExitFEMD | Exit: force exit market days count |
| 43 | ExitFECD | Exit: force exit calendar days count |

---

## Daily Automation

Every weekday at 6:15 PM IST, GitHub automatically:
1. Downloads today's NSE bhav copy from NSE archives
2. Saves to `bhav_data/Mon-YYYY/sec_bhavdata_full_DDMMYYYY.csv`
3. Commits to repo

No action needed from you — data grows automatically.

---

## Simulation Logic (mirrors VBA exactly)

### Filter 1 — 5-Day Dip
```
pct_from_low = (close - min_5day_close) / min_5day_close
PASS: pct_min ≤ pct_from_low ≤ pct_max   (e.g. -10% to -5%)
```

### Filter 2 — ATH Distance  
```
pct_from_ath = (close - ATH) / ATH
PASS: ath_min ≤ pct_from_ath ≤ ath_max   (e.g. -60% to -30%)
```

### Exit Rules
| Exit | Trigger | Type |
|------|---------|------|
| Target Achieved | High ≥ target price | ProfitTGT |
| Stoploss Triggered | Low ≤ stop price | LossSL |
| Force Exit — Market Days | Held ≥ MaxDuration trading days | Profit/LossFEMD |
| Force Exit — Calendar Days | Held ≥ 90 calendar days | Profit/LossFECD |

### Trade Entry
- Buy triggered when: `Low ≤ signal_close` (first buy)
- Additional buys: `Low ≤ avg_buy × (1 - buy_drop%)`, up to MaxBuys
- Investment per buy: ₹10,000 (configurable)
- Target recalculates after each buy (based on avg buy price)
- Stoploss stays fixed at signal price × (1 - stoploss%)
