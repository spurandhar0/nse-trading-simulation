"""
Script 5: Run Parameter Test Simulation
=========================================
Reads:  db/signals.parquet        (full mode)
        db/signals_quick.parquet  (quick mode)
        db/eq_data.parquet
        config/simulation_params.json

Modes:
  --mode full  (default)  All 2000 parameter combinations → monthly sweep
  --mode quick            Single param set from quick_run config → 1 output row

Output:
  output/Results_YYYYMMDD_HHMMSS.xlsx       (full mode)
  output/QuickRun_YYYYMMDD_HHMMSS.xlsx      (quick mode)

43 Output Columns (exact order):
  Test, DAYSBACK, PCTMIN, PCTMAX, ATHMIN, ATHMAX, MAXBUYS, BUYDROP,
  TARGET, STOPLOSS, MAXDURA, WinRate, TotalTrade, Executed, Open/Closed,
  ProfitTGT, LossSL, LossFEMD, LossFECD, ProfitFEMD, ProfitFECD,
  Pending, Expired, Invalid, TotalRows, Wins, Losses, TotalStock,
  SumProfit, SumGainFin, Dur5, Dur10, Dur15, Dur20, Dur25, Dur30,
  Dur35, Dur40, ExitTGT, ExitSL, ExitFEMD, ExitFECD
"""

import os
import sys
import json
import argparse
import itertools
import numpy as np
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side)
from openpyxl.utils import get_column_letter

CONFIG_FILE        = "config/simulation_params.json"
SIGNALS_FILE_FULL  = "db/signals.parquet"
SIGNALS_FILE_QUICK = "db/signals_quick.parquet"
EQ_FILE            = "db/eq_data.parquet"
OUTPUT_DIR         = "output"

COLUMNS_43 = [
    "Test", "DAYSBACK", "PCTMIN", "PCTMAX", "ATHMIN", "ATHMAX",
    "MAXBUYS", "BUYDROP", "TARGET", "STOPLOSS", "MAXDURA",
    "WinRate", "TotalTrade", "Executed", "Open/Closed",
    "ProfitTGT", "LossSL", "LossFEMD", "LossFECD", "ProfitFEMD", "ProfitFECD",
    "Pending", "Expired", "Invalid", "TotalRows", "Wins", "Losses", "TotalStock",
    "SumProfit", "SumGainFin",
    "Dur5", "Dur10", "Dur15", "Dur20", "Dur25", "Dur30", "Dur35", "Dur40",
    "ExitTGT", "ExitSL", "ExitFEMD", "ExitFECD"
]

# ─── CONFIG ───────────────────────────────────────────────────────────────────

def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)

def build_trade_combos(cfg, mode):
    """Generate trade parameter combinations based on mode."""
    if mode == "quick":
        q = cfg["quick_run"]
        return [(q["max_buys"], q["buy_drop"], q["target"],
                 q["stoploss"],  q["max_duration"])]

    # Full mode: sweep all param_sweep.trade ranges
    t = cfg["param_sweep"]["trade"]
    return list(itertools.product(
        t["max_buys"], t["buy_drop"], t["target"],
        t["stoploss"], t["max_duration"]
    ))

# ─── PRICE DATA ───────────────────────────────────────────────────────────────

def build_price_dict(eq_df):
    """
    Build lookup dict: symbol → dict with dates/closes/highs/lows/day_map
    date → index mapping for fast lookup.
    """
    print("Building price dictionary...")
    price_dict = {}

    for sym, grp in eq_df.groupby("SYMBOL"):
        grp    = grp.sort_values("DATE1").reset_index(drop=True)
        dates  = grp["DATE1"].values
        closes = grp["CLOSE_PRICE"].values.astype(float)
        highs  = grp["HIGH_PRICE"].values.astype(float)
        lows   = grp["LOW_PRICE"].values.astype(float)
        opens  = grp["OPEN_PRICE"].values.astype(float)
        prevs  = grp["PREV_CLOSE"].values.astype(float)

        day_map = {pd.Timestamp(d): i for i, d in enumerate(dates)}

        price_dict[sym] = {
            "dates":   dates,
            "closes":  closes,
            "highs":   highs,
            "lows":    lows,
            "opens":   opens,
            "prevs":   prevs,
            "day_map": day_map,
        }

    print(f"  Price dict built for {len(price_dict):,} symbols")
    return price_dict

# ─── SINGLE TRADE SIMULATION ──────────────────────────────────────────────────

def simulate_trade(sym, signal_date, signal_close, price_dict,
                   max_buys, buy_drop, target_pct, stoploss_pct, max_duration,
                   investment_per_buy, force_exit_calendar_days, pending_window_days):
    """
    Simulate one trade. Returns dict with order/status/profit/exit_type/etc.
    Mirrors VBA: SimulateTrades + WriteTradeResults + WriteClosedTradeResult
    """
    if sym not in price_dict:
        return {"order": "Invalid"}

    pd_data = price_dict[sym]
    dates   = pd_data["dates"]
    closes  = pd_data["closes"]
    highs   = pd_data["highs"]
    lows    = pd_data["lows"]
    day_map = pd_data["day_map"]

    sig_ts = pd.Timestamp(signal_date)
    if sig_ts not in day_map:
        return {"order": "Invalid"}

    start_idx = day_map[sig_ts]

    # Fixed stoploss based on signal close
    fixed_stop   = round(signal_close * (1 - stoploss_pct), 2)
    target_price = round(signal_close * (1 + target_pct), 2)   # updated on avg buy
    stop_price   = fixed_stop

    buy_count        = 0
    avg_buy_price    = 0.0
    total_qty        = 0.0
    total_investment = 0.0
    first_buy_date   = None
    buy_day_indices  = set()
    market_days      = 0
    exit_found       = False
    exit_type        = None
    exit_price       = 0.0

    last_idx = len(dates) - 1

    for i in range(start_idx + 1, last_idx + 1):
        curr_ts  = pd.Timestamp(dates[i])
        low_px   = lows[i]
        high_px  = highs[i]
        close_px = closes[i]

        if buy_count == 0:
            # First buy: low touches signal_close
            if low_px <= signal_close:
                buy_count      = 1
                first_buy_date = curr_ts
                buy_day_indices.add(i)
                qty = int(investment_per_buy / signal_close)
                if qty < 1:
                    qty = 1
                total_qty        += qty
                total_investment += signal_close * qty
                avg_buy_price     = round(total_investment / total_qty, 2)
                target_price      = round(avg_buy_price * (1 + target_pct), 2)
        else:
            is_buy_day = i in buy_day_indices
            if not is_buy_day:
                market_days += 1

            cal_days = (curr_ts - first_buy_date).days if first_buy_date else 0

            if not is_buy_day:
                # 1. Stoploss
                if low_px <= stop_price:
                    exit_found = True; exit_price = stop_price
                    exit_type  = "Stoploss Triggered"; exit_date = curr_ts; break
                # 2. Target
                if high_px >= target_price:
                    exit_found = True; exit_price = target_price
                    exit_type  = "Target Achieved"; exit_date = curr_ts; break
                # 3. Force exit — market days
                if market_days >= max_duration:
                    exit_found = True; exit_price = round(close_px, 2)
                    exit_type  = "Force Exit - Market Days"; exit_date = curr_ts; break
                # 4. Force exit — calendar days
                if cal_days >= force_exit_calendar_days:
                    exit_found = True; exit_price = round(close_px, 2)
                    exit_type  = "Force Exit - Calendar Days"; exit_date = curr_ts; break

            # Additional buys
            if not exit_found and not is_buy_day and buy_count < max_buys:
                buy_level = round(avg_buy_price * (1 - buy_drop), 2)
                if buy_level >= stop_price and low_px <= buy_level:
                    buy_count        += 1
                    buy_day_indices.add(i)
                    qty = int(investment_per_buy / buy_level)
                    if qty < 1: qty = 1
                    total_qty        += qty
                    total_investment += buy_level * qty
                    avg_buy_price     = round(total_investment / total_qty, 2)
                    target_price      = round(avg_buy_price * (1 + target_pct), 2)

    # ── Determine result ──────────────────────────────────────────────────────
    if buy_count == 0:
        last_ts   = pd.Timestamp(dates[last_idx])
        days_from = (last_ts - sig_ts).days
        order     = "Pending" if days_from <= pending_window_days else "Expired"
        return {"order": order, "status": order, "profit": 0, "gain_pct": 0,
                "exit_type": None, "result": None, "market_days": 0}

    if not exit_found:
        return {"order": "Executed", "status": "Open", "profit": 0, "gain_pct": 0,
                "exit_type": None, "result": None, "market_days": market_days}

    profit   = round((exit_price - avg_buy_price) * total_qty, 2)
    gain_pct = round(((exit_price - avg_buy_price) / avg_buy_price) * 100, 2) \
               if avg_buy_price > 0 else 0

    if   exit_type == "Target Achieved":          result = "Profit-TGT"
    elif exit_type == "Stoploss Triggered":       result = "Loss-SL"
    elif exit_type == "Force Exit - Market Days": result = ("Profit" if profit >= 0 else "Loss") + "-FE-MD"
    elif exit_type == "Force Exit - Calendar Days": result = ("Profit" if profit >= 0 else "Loss") + "-FE-CD"
    else:                                         result = "Profit" if profit >= 0 else "Loss"

    return {
        "order":       "Executed",
        "status":      "Closed",
        "profit":      profit,
        "gain_pct":    gain_pct,
        "exit_type":   exit_type,
        "result":      result,
        "market_days": market_days,
    }

# ─── AGGREGATE STATS ──────────────────────────────────────────────────────────

def aggregate_stats(results):
    """Aggregate trade results into the 43-column stats dict."""
    wins=0; losses=0
    c_profit_tgt=0; c_loss_sl=0
    c_loss_femd=0;  c_profit_femd=0
    c_loss_fecd=0;  c_profit_fecd=0
    dur5=0; dur10=0; dur15=0; dur20=0
    dur25=0; dur30=0; dur35=0; dur40=0
    pending=0; expired=0; invalid=0
    exec_count=0; open_count=0; closed_count=0
    sum_profit=0.0; sum_gain=0.0
    exit_tgt=0; exit_sl=0; exit_femd=0; exit_fecd=0

    for r in results:
        order = r.get("order", "Invalid")
        if   order == "Pending":  pending += 1
        elif order == "Expired":  expired += 1
        elif order == "Invalid":  invalid += 1
        elif order == "Executed":
            exec_count += 1
            status = r.get("status", "")
            if   status == "Open":   open_count += 1
            elif status == "Closed":
                closed_count += 1
                result    = r.get("result", "")    or ""
                exit_type = r.get("exit_type", "") or ""
                profit    = r.get("profit",   0)   or 0
                gain_pct  = r.get("gain_pct", 0)   or 0
                mdays     = r.get("market_days", 0) or 0

                sum_profit += profit
                sum_gain   += gain_pct

                ru = result.upper()
                if   "PROFIT-TGT"  in ru: c_profit_tgt  += 1; wins   += 1
                elif "LOSS-SL"     in ru: c_loss_sl      += 1; losses += 1
                elif ru == "PROFIT-FE-MD": c_profit_femd += 1; wins   += 1
                elif ru == "LOSS-FE-MD":   c_loss_femd   += 1; losses += 1
                elif ru == "PROFIT-FE-CD": c_profit_fecd += 1; wins   += 1
                elif ru == "LOSS-FE-CD":   c_loss_fecd   += 1; losses += 1
                elif "PROFIT" in ru:       wins   += 1
                else:                      losses += 1

                eu = exit_type.upper()
                if   "TARGET ACHIEVED"        in eu: exit_tgt  += 1
                elif "STOPLOSS TRIGGERED"     in eu: exit_sl   += 1
                elif "FORCE EXIT - MARKET"    in eu: exit_femd += 1
                elif "FORCE EXIT - CALENDAR"  in eu: exit_fecd += 1

                if   mdays <= 5:  dur5  += 1
                elif mdays <= 10: dur10 += 1
                elif mdays <= 15: dur15 += 1
                elif mdays <= 20: dur20 += 1
                elif mdays <= 25: dur25 += 1
                elif mdays <= 30: dur30 += 1
                elif mdays <= 35: dur35 += 1
                elif mdays <= 40: dur40 += 1

    total_trades = exec_count
    total_stocks = len(results)
    win_rate     = round((wins / closed_count * 100) if closed_count > 0 else 0, 2)

    return {
        "WinRate":    win_rate,    "TotalTrade": total_trades,
        "Executed":   exec_count,  "Open/Closed": f"{open_count}/{closed_count}",
        "ProfitTGT":  c_profit_tgt,"LossSL":     c_loss_sl,
        "LossFEMD":   c_loss_femd, "LossFECD":   c_loss_fecd,
        "ProfitFEMD": c_profit_femd,"ProfitFECD": c_profit_fecd,
        "Pending":    pending,     "Expired":    expired,
        "Invalid":    invalid,     "TotalRows":  total_stocks,
        "Wins":       wins,        "Losses":     losses,
        "TotalStock": total_stocks,"SumProfit":  round(sum_profit, 2),
        "SumGainFin": round(sum_gain, 2),
        "Dur5":  dur5,  "Dur10": dur10, "Dur15": dur15, "Dur20": dur20,
        "Dur25": dur25, "Dur30": dur30, "Dur35": dur35, "Dur40": dur40,
        "ExitTGT":  exit_tgt,  "ExitSL":   exit_sl,
        "ExitFEMD": exit_femd, "ExitFECD": exit_fecd,
    }

# ─── EXCEL OUTPUT ─────────────────────────────────────────────────────────────

NAVY  = "00203864"
WHITE = "00FFFFFF"
LIGHT = "00F2F2F2"

def style_sheet(ws, mode_label):
    thin = Side(style="thin", color="BFBFBF")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title row (row 1)
    ws.insert_rows(1)
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1,   end_column=len(COLUMNS_43))
    tc          = ws.cell(row=1, column=1)
    tc.value    = (f"NSE Simulation Results [{mode_label}] — "
                   f"Generated {datetime.now().strftime('%d-%b-%Y %H:%M')}")
    tc.font     = Font(bold=True, size=14, color=WHITE)
    tc.fill     = PatternFill("solid", fgColor=NAVY)
    tc.alignment= Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Header row (row 2)
    for col_idx, col_name in enumerate(COLUMNS_43, 1):
        cell           = ws.cell(row=2, column=col_idx)
        cell.value     = col_name
        cell.font      = Font(bold=True, size=10, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = bdr
    ws.row_dimensions[2].height = 18

    # Zebra stripes
    last_row = ws.max_row
    for row in range(3, last_row + 1):
        fill_color = LIGHT if row % 2 == 0 else "00FFFFFF"
        for col in range(1, len(COLUMNS_43) + 1):
            cell           = ws.cell(row=row, column=col)
            cell.fill      = PatternFill("solid", fgColor=fill_color)
            cell.border    = bdr
            cell.alignment = Alignment(horizontal="center")

    ws.freeze_panes = ws.cell(row=3, column=1)
    ws.auto_filter.ref = (ws.cell(row=2, column=1).coordinate + ":" +
                          ws.cell(row=2, column=len(COLUMNS_43)).coordinate)

    for col_idx in range(1, len(COLUMNS_43) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 13

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0

# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="NSE Trading Simulation")
    parser.add_argument("--mode", choices=["quick", "full"], default="full",
                        help="quick = 1 param combo (daily); full = all combos (monthly)")
    parser.add_argument("--symbols", default="",
                        help="Comma-separated symbols to simulate, e.g. TCS,WIPRO,INFY (empty = all)")
    args = parser.parse_args()
    mode = args.mode

    signals_file = SIGNALS_FILE_QUICK if mode == "quick" else SIGNALS_FILE_FULL

    for f in [CONFIG_FILE, signals_file, EQ_FILE]:
        if not os.path.exists(f):
            print(f"❌ Missing: {f}")
            raise SystemExit(1)

    cfg = load_config()
    investment_per_buy        = cfg.get("investment_per_buy",        10000)
    force_exit_calendar_days  = cfg.get("force_exit_calendar_days",  90)
    pending_window_days       = cfg.get("pending_window_days",       30)
    max_rows_per_sheet        = cfg.get("max_rows_per_sheet",        25000)

    trade_combos = build_trade_combos(cfg, mode)
    mode_label   = "QUICK RUN" if mode == "quick" else "FULL PARAMETER SWEEP"
    print(f"Mode            : {mode_label}")
    print(f"Trade combos    : {len(trade_combos):,}")

    # ── Resolve symbol filter (CLI > config > all) ─────────────────────────────
    symbol_filter = set()
    cli_syms = args.symbols.strip()
    if cli_syms:
        symbol_filter = {s.strip().upper() for s in cli_syms.split(",") if s.strip()}
        print(f"Symbol filter   : CLI override → {sorted(symbol_filter)}")
    else:
        section = cfg.get("quick_run" if mode == "quick" else "param_sweep", {})
        config_syms = section.get("watch_symbols", [])
        if config_syms:
            symbol_filter = {s.strip().upper() for s in config_syms if s.strip()}
            print(f"Symbol filter   : config watch_symbols → {sorted(symbol_filter)}")
        else:
            print("Symbol filter   : ALL symbols")

    print("Loading signals...")
    sig_df = pd.read_parquet(signals_file)
    sig_df["SIGNAL_DATE"] = pd.to_datetime(sig_df["SIGNAL_DATE"])

    # Apply symbol filter to signals
    if symbol_filter:
        sig_df = sig_df[sig_df["SYMBOL"].isin(symbol_filter)]
        missing = symbol_filter - set(sig_df["SYMBOL"].unique())
        if missing:
            print(f"⚠️  Symbols not in signals (no trades triggered): {sorted(missing)}")

    print(f"  Total signals : {len(sig_df):,}")

    if len(sig_df) == 0:
        print("⚠️  No signals found. Check symbol names or run 04_filter_signals.py first.")
        raise SystemExit(0)

    print("Loading EQ price data...")
    eq_df          = pd.read_parquet(EQ_FILE)
    eq_df["DATE1"] = pd.to_datetime(eq_df["DATE1"])
    # Filter EQ price dict to only needed symbols (saves memory)
    if symbol_filter:
        eq_df = eq_df[eq_df["SYMBOL"].isin(symbol_filter)]
    price_dict     = build_price_dict(eq_df)
    del eq_df

    filter_cols   = ["DAYSBACK", "PCTMIN", "PCTMAX", "ATHMIN", "ATHMAX"]
    filter_groups = sig_df.groupby(filter_cols, dropna=False)
    total_expected = len(filter_groups) * len(trade_combos)
    print(f"Filter combos in signals : {len(filter_groups)}")
    print(f"Total output rows        : {total_expected:,}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts_str   = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix   = "QuickRun" if mode == "quick" else "Results"
    out_path = os.path.join(OUTPUT_DIR, f"{prefix}_{ts_str}.xlsx")

    wb           = Workbook()
    ws           = wb.active
    ws.title     = "Data_1"
    ws.append(COLUMNS_43)

    test_num      = 0
    sheet_num     = 1
    rows_on_sheet = 0

    for filter_key, filter_group in filter_groups:
        db, pmin, pmax, amin, amax = filter_key

        signal_list = list(zip(
            filter_group["SYMBOL"],
            filter_group["SIGNAL_DATE"],
            filter_group["SIGNAL_CLOSE"],
        ))
        total_stocks = len(signal_list)

        for (mb, bd, tgt, sl, mdur) in trade_combos:
            test_num += 1

            results = [
                simulate_trade(
                    sym, sig_date, sig_close, price_dict,
                    mb, bd, tgt, sl, mdur,
                    investment_per_buy, force_exit_calendar_days, pending_window_days
                )
                for sym, sig_date, sig_close in signal_list
            ]

            stats = aggregate_stats(results)
            stats["TotalStock"] = total_stocks

            row  = [test_num, db, pmin, pmax, amin, amax, mb, bd, tgt, sl, mdur]
            row += [stats[c] for c in COLUMNS_43[11:]]

            if rows_on_sheet >= max_rows_per_sheet:
                sheet_num    += 1
                rows_on_sheet = 0
                ws = wb.create_sheet(title=f"Data_{sheet_num}")
                ws.append(COLUMNS_43)

            ws.append(row)
            rows_on_sheet += 1

        if test_num % 200 == 0:
            print(f"  Completed {test_num:,} parameter combos...")

    print(f"\nApplying formatting to {sheet_num} sheet(s)...")
    for sn in range(1, sheet_num + 1):
        sname = f"Data_{sn}"
        if sname in wb.sheetnames:
            style_sheet(wb[sname], mode_label)

    wb.save(out_path)
    print(f"\n✅ Results saved  : {out_path}")
    print(f"✅ Total rows     : {test_num:,}")
    print(f"✅ Sheets used    : {sheet_num}")

if __name__ == "__main__":
    start = datetime.now()
    main()
    elapsed = (datetime.now() - start).total_seconds()
    print(f"\n⏱  Total time: {elapsed:.1f}s ({elapsed/60:.1f} min)")
