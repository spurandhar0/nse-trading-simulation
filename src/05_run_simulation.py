"""
Script 5: Run Parameter Test Simulation
=========================================
Reads:  db/signals.parquet        (full mode)
        db/signals_quick.parquet  (quick mode)
        db/eq_data.parquet
        config/simulation_params.json

Modes:
  --mode quick  (default for daily run)
        Single fixed param set from quick_run config.
        Output: per-trade PICKS SHEET — one row per trade.
        Sheet name: Pickse
        File  : output/YYYY-MM/QuickRun_Picks_YYYYMMDD_HHMMSS.xlsx

  --mode full
        All parameter combinations from param_sweep -> aggregate stats.
        Output: 43-column summary — one row per param combo.
        File  : output/YYYY-MM/Results_YYYYMMDD_HHMMSS.xlsx

SIMULATION RULES:
  - Entry  : D+1 low <= signal_close  -> buy at signal_close price (B0)
  - Additional buys: when buy_count <= max_buys AND next-buy-level >= stop_price
  - Same-day buy and sell NOT allowed (exit checks skip the buy day itself)
  - Invalid: signal date is the last available date for that symbol (no D+1 data)
  - Pending: buy not triggered yet, within pending_window_days of signal
  - Expired: buy not triggered, beyond pending_window_days
  - FE-MD  : market_days >= max_duration (trading days counted after first buy)
  - FE-CD  : calendar_days >= force_exit_calendar_days (90 by default)
  - Stop   : based on signal_close (USE_AVGBUY_FOR_STOPLOSS = False)
  - Target : updated on each additional buy (USE_AVGBUY_FOR_TARGET = True)

PICKS SHEET COLUMNS (quick mode — 50 columns for max_buys=2):
  1DChange%, StockName, 5DLow%, 5DLowPrice, RecentLTP,
  BuyDate, BuyClPrice, 5DLowDate, TodayDate,
  BuyCount, AvgBuyPrice, TotalQty, TargetPrice, StoplossPrice, TotalInvestment,
  Order, Status, Duration, DurationGroup, Profits, GainLoss%, Result, ExitType,
  Action, BuyChance, SoldDate, SoldPrice, SoldPrevClose, SoldOpen,
  SoldHigh, SoldLow, SoldClose,
  B0_BoughtDate/PrevClose/Open/High/Low/Close,
  B1_BoughtDate/PrevClose/Open/High/Low/Close,
  B2_BoughtDate/PrevClose/Open/High/Low/Close

43 AGGREGATE COLUMNS (full mode — exact order):
  Test, DAYSBACK, PCTMIN, PCTMAX, ATHMIN, ATHMAX, MAXBUYS, BUYDROP,
  TARGET, STOPLOSS, MAXDURA, WinRate, TotalTrade, Executed, Open, Closed,
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
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CONFIG_FILE        = "config/simulation_params.json"
SIGNALS_FILE_FULL  = "db/signals.parquet"
SIGNALS_FILE_QUICK = "db/signals_quick.parquet"
EQ_FILE            = "db/eq_data.parquet"
OUTPUT_DIR         = "output"

COLUMNS_43 = [
    "Test", "DAYSBACK", "PCTMIN", "PCTMAX", "ATHMIN", "ATHMAX",
    "MAXBUYS", "BUYDROP", "TARGET", "STOPLOSS", "MAXDURA",
    "WinRate", "TotalTrade", "Executed", "Open", "Closed",
    "ProfitTGT", "LossSL", "LossFEMD", "LossFECD", "ProfitFEMD", "ProfitFECD",
    "Pending", "Expired", "Invalid", "TotalRows", "Wins", "Losses", "TotalStock",
    "SumProfit", "SumGainFin",
    "Dur5", "Dur10", "Dur15", "Dur20", "Dur25", "Dur30", "Dur35", "Dur40",
    "ExitTGT", "ExitSL", "ExitFEMD", "ExitFECD"
]

NAVY  = "00203864"
WHITE = "00FFFFFF"
LIGHT = "00F2F2F2"

# Excel number format strings
FMT_DATE  = "DD-MM-YYYY"
FMT_PRICE = "0.00"
FMT_PCT   = "0.00%"
FMT_INT   = "General"


# ─── CONFIG ───────────────────────────────────────────────────────────────────

def load_config():
    with open(CONFIG_FILE) as f:
        return json.load(f)


def build_trade_combos(cfg, mode):
    if mode == "quick":
        q = cfg["quick_run"]
        return [(q["max_buys"], q["buy_drop"], q["target"],
                 q["stoploss"],  q["max_duration"])]
    t = cfg["param_sweep"]["trade"]
    return list(itertools.product(
        t["max_buys"], t["buy_drop"], t["target"],
        t["stoploss"], t["max_duration"]
    ))


# ─── PRICE DATA ───────────────────────────────────────────────────────────────

def build_price_dict(eq_df):
    print("Building price dictionary...")
    price_dict = {}
    for sym, grp in eq_df.groupby("SYMBOL"):
        grp   = grp.sort_values("DATE1").reset_index(drop=True)
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


# ─── DURATION GROUP ───────────────────────────────────────────────────────────

def duration_group(market_days):
    """Return DurationGroup: integer bucket of 5 up to 50, '50+' for >50."""
    if market_days is None:
        return None
    if market_days > 50:
        return "50+"
    return ((market_days - 1) // 5 + 1) * 5


# ─── SIMULATION (DETAILED — Picks Sheet / Quick Run) ─────────────────────────

def simulate_trade_detailed(sym, signal_date, signal_close, price_dict,
                             max_buys, buy_drop, target_pct, stoploss_pct,
                             max_duration, investment_per_buy,
                             force_exit_calendar_days, pending_window_days):
    """
    Simulate one trade and return all per-trade details for the picks sheet.

    Key rules:
      - If start_idx >= last_idx  (signal date is the last available date for
        this symbol) -> INVALID (no D+1 data to attempt entry).
      - Loop: range(start_idx+1, last_idx)  -- excludes the very last data day
        ("today") from both buy attempts and exit checks.
      - Same-day buy and sell NOT allowed: exit checks skip the buy day itself.
      - max_buys = number of ADDITIONAL dip buys;
        total buy slots = max_buys+1 (B0..B(max_buys))
    """
    invalid = {
        "order": "Invalid", "status": None, "action": "Skip",
        "buy_count": None, "avg_buy_price": None, "total_qty": None,
        "total_investment": None, "target_price": None, "stop_price": None,
        "first_buy_date": None, "exit_found": False,
        "exit_date": None, "exit_price": None, "exit_type": None,
        "profit": None, "gain_pct": None, "market_days": None,
        "result_str": "Invalid: No market data found",
        "duration_group": None,
        "buy_chance": None,
        "sold_prev_close": None, "sold_open": None,
        "sold_high": None, "sold_low": None, "sold_close": None,
        "buys": [],
        "had_buy_chance": False,
    }

    if sym not in price_dict:
        return invalid

    pd_data  = price_dict[sym]
    dates    = pd_data["dates"]
    closes   = pd_data["closes"]
    highs    = pd_data["highs"]
    lows     = pd_data["lows"]
    opens    = pd_data["opens"]
    prevs    = pd_data["prevs"]
    day_map  = pd_data["day_map"]

    sig_ts = pd.Timestamp(signal_date)
    if sig_ts not in day_map:
        return invalid

    start_idx = day_map[sig_ts]
    last_idx  = len(dates) - 1

    # ── INVALID: signal on last available date for this symbol — no D+1 ──────
    if start_idx >= last_idx:
        return invalid

    # Prices for target and stop (based on signal_close per VBA logic)
    stop_price   = round(signal_close * (1 - stoploss_pct), 2)
    target_price = round(signal_close * (1 + target_pct),  2)

    buy_count        = 0
    avg_buy_price    = 0.0
    total_qty        = 0
    total_investment = 0.0
    first_buy_date   = None
    buy_day_indices  = set()
    market_days      = 0
    exit_found       = False
    exit_type        = None
    exit_price       = 0.0
    exit_date        = None
    exit_idx         = -1
    buys             = []
    had_buy_chance   = False

    # ── Main loop — D+1 through D+(N-2); last data point excluded ────────────
    for i in range(start_idx + 1, last_idx):
        curr_ts  = pd.Timestamp(dates[i])
        low_px   = float(lows[i])
        high_px  = float(highs[i])
        close_px = float(closes[i])
        open_px  = float(opens[i])
        prev_px  = float(prevs[i])

        if buy_count == 0:
            # ── First buy (B0): day low touches signal_close ─────────────────
            if low_px <= signal_close:
                buy_count      = 1
                first_buy_date = curr_ts
                buy_day_indices.add(i)
                qty = int(investment_per_buy / signal_close)
                if qty < 1:
                    qty = 1
                total_qty        = qty
                total_investment = signal_close * qty
                avg_buy_price    = round(total_investment / total_qty, 2)
                # Target updates on each buy (USE_AVGBUY_FOR_TARGET = True)
                target_price     = round(avg_buy_price * (1 + target_pct), 2)
                buys.append({
                    "date":       curr_ts,
                    "prev_close": prev_px,
                    "open":       open_px,
                    "high":       high_px,
                    "low":        low_px,
                    "close":      close_px,
                })
        else:
            is_buy_day = i in buy_day_indices
            if not is_buy_day:
                market_days += 1

            cal_days = (curr_ts - first_buy_date).days if first_buy_date else 0

            # ── Check exits (only on non-buy days — same-day sell not allowed) ─
            if not is_buy_day:
                if low_px <= stop_price:
                    exit_found = True
                    exit_price = stop_price
                    exit_type  = "Stoploss Triggered"
                    exit_date  = curr_ts
                    exit_idx   = i
                    break
                if high_px >= target_price:
                    exit_found = True
                    exit_price = target_price
                    exit_type  = "Target Achieved"
                    exit_date  = curr_ts
                    exit_idx   = i
                    break
                if market_days >= max_duration:
                    exit_found = True
                    exit_price = round(close_px, 2)
                    exit_type  = "Force Exit - Market Days"
                    exit_date  = curr_ts
                    exit_idx   = i
                    break
                if cal_days >= force_exit_calendar_days:
                    exit_found = True
                    exit_price = round(close_px, 2)
                    exit_type  = "Force Exit - Calendar Days"
                    exit_date  = curr_ts
                    exit_idx   = i
                    break

            # ── Additional buys (buy_count <= max_buys means we can buy more) ─
            if not exit_found and not is_buy_day and buy_count <= max_buys:
                buy_level = round(avg_buy_price * (1 - buy_drop), 2)
                if low_px <= buy_level:
                    had_buy_chance = True
                    if buy_level >= stop_price:
                        buy_count        += 1
                        buy_day_indices.add(i)
                        qty = int(investment_per_buy / buy_level)
                        if qty < 1:
                            qty = 1
                        total_qty        += qty
                        total_investment += buy_level * qty
                        avg_buy_price     = round(total_investment / total_qty, 2)
                        target_price      = round(avg_buy_price * (1 + target_pct), 2)
                        buys.append({
                            "date":       curr_ts,
                            "prev_close": prev_px,
                            "open":       open_px,
                            "high":       high_px,
                            "low":        low_px,
                            "close":      close_px,
                        })

    # ── Determine order / status / action ─────────────────────────────────────
    if buy_count == 0:
        last_ts   = pd.Timestamp(dates[last_idx])
        days_from = (last_ts - sig_ts).days
        if days_from <= pending_window_days:
            order = "Pending"
            action = "Buy"
        else:
            order = "Expired"
            action = "Skip"
        return {
            "order": order,
            "status": "Not Triggered" if order == "Pending" else "Expired",
            "action": action,
            "buy_count": 0, "avg_buy_price": None, "total_qty": None,
            "total_investment": None,
            "target_price": target_price, "stop_price": stop_price,
            "first_buy_date": None, "exit_found": False,
            "exit_date": None, "exit_price": None, "exit_type": None,
            "profit": None, "gain_pct": None, "market_days": None,
            "result_str": None, "duration_group": None,
            "buy_chance": None,
            "sold_prev_close": None, "sold_open": None,
            "sold_high": None, "sold_low": None, "sold_close": None,
            "buys": [], "had_buy_chance": False,
        }

    # ── Executed — Open ───────────────────────────────────────────────────────
    if not exit_found:
        return {
            "order": "Executed", "status": "Open", "action": "Hold",
            "buy_count": buy_count, "avg_buy_price": avg_buy_price,
            "total_qty": total_qty, "total_investment": round(total_investment, 2),
            "target_price": target_price, "stop_price": stop_price,
            "first_buy_date": first_buy_date, "exit_found": False,
            "exit_date": None, "exit_price": None, "exit_type": None,
            "profit": None,   # computed in build_picks_row using recent_ltp
            "gain_pct": None, # computed in build_picks_row using recent_ltp
            "market_days": market_days,
            "result_str": None,
            "duration_group": duration_group(market_days),
            "buy_chance": "Buy Chance" if had_buy_chance else None,
            "sold_prev_close": None, "sold_open": None,
            "sold_high": None, "sold_low": None, "sold_close": None,
            "buys": buys, "had_buy_chance": had_buy_chance,
        }

    # ── Executed — Closed ─────────────────────────────────────────────────────
    profit   = round((exit_price - avg_buy_price) * total_qty, 2)
    gain_pct = round((exit_price - avg_buy_price) / avg_buy_price, 4) \
               if avg_buy_price > 0 else 0.0

    if   exit_type == "Target Achieved":           result_str = "Profit-TGT"
    elif exit_type == "Stoploss Triggered":         result_str = "Loss-SL"
    elif exit_type == "Force Exit - Market Days":   result_str = ("Profit" if profit >= 0 else "Loss") + "-FE-MD"
    elif exit_type == "Force Exit - Calendar Days": result_str = ("Profit" if profit >= 0 else "Loss") + "-FE-CD"
    else:                                           result_str = "Profit" if profit >= 0 else "Loss"

    sold_prev_close = sold_open = sold_high = sold_low = sold_close = None
    if exit_idx >= 0:
        sold_prev_close = round(float(prevs[exit_idx]),  2)
        sold_open       = round(float(opens[exit_idx]),  2)
        sold_high       = round(float(highs[exit_idx]),  2)
        sold_low        = round(float(lows[exit_idx]),   2)
        sold_close      = round(float(closes[exit_idx]), 2)

    return {
        "order": "Executed", "status": "Closed", "action": "Exit",
        "buy_count": buy_count, "avg_buy_price": avg_buy_price,
        "total_qty": total_qty, "total_investment": round(total_investment, 2),
        "target_price": target_price, "stop_price": stop_price,
        "first_buy_date": first_buy_date, "exit_found": True,
        "exit_date": exit_date, "exit_price": exit_price, "exit_type": exit_type,
        "profit": profit, "gain_pct": gain_pct,
        "market_days": market_days,
        "result_str": result_str,
        "duration_group": duration_group(market_days),
        "buy_chance": "Buy Chance" if had_buy_chance else None,
        "sold_prev_close": sold_prev_close, "sold_open": sold_open,
        "sold_high": sold_high, "sold_low": sold_low, "sold_close": sold_close,
        "buys": buys, "had_buy_chance": had_buy_chance,
    }


# ─── SIMULATION (SIMPLE — Full mode aggregate stats) ─────────────────────────

def simulate_trade(sym, signal_date, signal_close, price_dict,
                   max_buys, buy_drop, target_pct, stoploss_pct,
                   max_duration, investment_per_buy,
                   force_exit_calendar_days, pending_window_days):
    """Lightweight simulation for full-mode parameter sweep aggregate stats."""
    if sym not in price_dict:
        return {"order": "Invalid"}

    pd_data  = price_dict[sym]
    dates    = pd_data["dates"]
    closes   = pd_data["closes"]
    highs    = pd_data["highs"]
    lows     = pd_data["lows"]
    day_map  = pd_data["day_map"]

    sig_ts = pd.Timestamp(signal_date)
    if sig_ts not in day_map:
        return {"order": "Invalid"}

    start_idx = day_map[sig_ts]
    last_idx  = len(dates) - 1

    # ── INVALID: signal on last available date for this symbol ────────────────
    if start_idx >= last_idx:
        return {"order": "Invalid"}

    stop_price   = round(signal_close * (1 - stoploss_pct), 2)
    target_price = round(signal_close * (1 + target_pct),  2)

    buy_count        = 0
    avg_buy_price    = 0.0
    total_qty        = 0
    total_investment = 0.0
    first_buy_date   = None
    buy_day_indices  = set()
    market_days      = 0
    exit_found       = False
    exit_type        = None
    exit_price       = 0.0
    exit_date        = None

    for i in range(start_idx + 1, last_idx):
        curr_ts  = pd.Timestamp(dates[i])
        low_px   = float(lows[i])
        high_px  = float(highs[i])
        close_px = float(closes[i])

        if buy_count == 0:
            if low_px <= signal_close:
                buy_count      = 1
                first_buy_date = curr_ts
                buy_day_indices.add(i)
                qty = int(investment_per_buy / signal_close)
                if qty < 1: qty = 1
                total_qty        = qty
                total_investment = signal_close * qty
                avg_buy_price    = round(total_investment / total_qty, 2)
                target_price     = round(avg_buy_price * (1 + target_pct), 2)
        else:
            is_buy_day = i in buy_day_indices
            if not is_buy_day:
                market_days += 1
            cal_days = (curr_ts - first_buy_date).days if first_buy_date else 0
            if not is_buy_day:
                if low_px <= stop_price:
                    exit_found = True; exit_price = stop_price
                    exit_type  = "Stoploss Triggered"; exit_date = curr_ts; break
                if high_px >= target_price:
                    exit_found = True; exit_price = target_price
                    exit_type  = "Target Achieved"; exit_date = curr_ts; break
                if market_days >= max_duration:
                    exit_found = True; exit_price = round(close_px, 2)
                    exit_type  = "Force Exit - Market Days"; exit_date = curr_ts; break
                if cal_days >= force_exit_calendar_days:
                    exit_found = True; exit_price = round(close_px, 2)
                    exit_type  = "Force Exit - Calendar Days"; exit_date = curr_ts; break
            if not exit_found and not is_buy_day and buy_count <= max_buys:
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

    if buy_count == 0:
        last_ts   = pd.Timestamp(dates[last_idx])
        days_from = (last_ts - sig_ts).days
        order = "Pending" if days_from <= pending_window_days else "Expired"
        return {"order": order, "status": order, "profit": 0, "gain_pct": 0,
                "exit_type": None, "result": None, "market_days": 0}

    if not exit_found:
        return {"order": "Executed", "status": "Open", "profit": 0, "gain_pct": 0,
                "exit_type": None, "result": None, "market_days": market_days}

    profit   = round((exit_price - avg_buy_price) * total_qty, 2)
    gain_pct = round((exit_price - avg_buy_price) / avg_buy_price, 4) \
               if avg_buy_price > 0 else 0

    if   exit_type == "Target Achieved":           result = "Profit-TGT"
    elif exit_type == "Stoploss Triggered":         result = "Loss-SL"
    elif exit_type == "Force Exit - Market Days":   result = ("Profit" if profit >= 0 else "Loss") + "-FE-MD"
    elif exit_type == "Force Exit - Calendar Days": result = ("Profit" if profit >= 0 else "Loss") + "-FE-CD"
    else:                                           result = "Profit" if profit >= 0 else "Loss"

    return {
        "order": "Executed", "status": "Closed",
        "profit": profit, "gain_pct": gain_pct,
        "exit_type": exit_type, "result": result,
        "market_days": market_days,
    }


# ─── AGGREGATE STATS (full mode) ─────────────────────────────────────────────

def aggregate_stats(results):
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
                result    = r.get("result",     "") or ""
                exit_type = r.get("exit_type",  "") or ""
                profit    = r.get("profit",     0)  or 0
                gain_pct  = r.get("gain_pct",   0)  or 0
                mdays     = r.get("market_days",0)  or 0

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
                if   "TARGET ACHIEVED"     in eu: exit_tgt  += 1
                elif "STOPLOSS TRIGGERED"  in eu: exit_sl   += 1
                elif "FORCE EXIT - MARKET" in eu: exit_femd += 1
                elif "FORCE EXIT - CALEND" in eu: exit_fecd += 1

                dg = duration_group(mdays)
                dg_num = dg if isinstance(dg, int) else 40
                if   dg_num <= 5:  dur5  += 1
                elif dg_num <= 10: dur10 += 1
                elif dg_num <= 15: dur15 += 1
                elif dg_num <= 20: dur20 += 1
                elif dg_num <= 25: dur25 += 1
                elif dg_num <= 30: dur30 += 1
                elif dg_num <= 35: dur35 += 1
                else:              dur40 += 1

    total_stocks = len(results)
    win_rate     = round((wins / closed_count * 100) if closed_count > 0 else 0, 2)

    return {
        "WinRate":    win_rate,
        "TotalTrade": exec_count + pending + expired,
        "Executed":   exec_count,
        "Open":       open_count,
        "Closed":     closed_count,
        "ProfitTGT":  c_profit_tgt, "LossSL":     c_loss_sl,
        "LossFEMD":   c_loss_femd,  "LossFECD":   c_loss_fecd,
        "ProfitFEMD": c_profit_femd,"ProfitFECD": c_profit_fecd,
        "Pending":    pending,       "Expired":    expired,
        "Invalid":    invalid,       "TotalRows":  total_stocks,
        "Wins":       wins,          "Losses":     losses,
        "TotalStock": total_stocks,  "SumProfit":  round(sum_profit, 2),
        "SumGainFin": round(sum_gain, 4),
        "Dur5":  dur5,  "Dur10": dur10, "Dur15": dur15, "Dur20": dur20,
        "Dur25": dur25, "Dur30": dur30, "Dur35": dur35, "Dur40": dur40,
        "ExitTGT":  exit_tgt,  "ExitSL":   exit_sl,
        "ExitFEMD": exit_femd, "ExitFECD": exit_fecd,
    }


# ─── PICKS SHEET HELPERS ─────────────────────────────────────────────────────

def get_picks_columns(max_buys):
    """
    Column headers for Picks Sheet.
    max_buys additional dip buys -> total buy blocks = max_buys+1 (B0..B(max_buys))
    """
    cols = [
        "1DChange%", "StockName", "5DLow%", "5DLowPrice", "RecentLTP",
        "BuyDate", "BuyClPrice", "5DLowDate", "TodayDate",
        "BuyCount", "AvgBuyPrice", "TotalQty", "TargetPrice", "StoplossPrice",
        "TotalInvestment", "Order", "Status", "Duration", "DurationGroup",
        "Profits", "GainLoss%", "Result", "ExitType", "Action", "BuyChance",
        "SoldDate", "SoldPrice", "SoldPrevClose", "SoldOpen",
        "SoldHigh", "SoldLow", "SoldClose",
    ]
    for b in range(max_buys + 1):   # B0 through B(max_buys)
        cols += [
            f"B{b}_BoughtDate", f"B{b}_PrevClose", f"B{b}_Open",
            f"B{b}_High",       f"B{b}_Low",        f"B{b}_Close",
        ]
    return cols


def _col_fmt(col_name):
    """Return openpyxl number format string for a column name."""
    if "Date" in col_name or col_name.endswith("Date"):
        return FMT_DATE
    if col_name in ("1DChange%", "5DLow%", "GainLoss%"):
        return FMT_PCT
    if col_name in ("BuyCount", "TotalQty", "Duration", "DurationGroup", "BuyChance",
                    "Order", "Status", "Result", "ExitType", "Action",
                    "StockName", "WinRate"):
        return FMT_INT
    # Price / money columns
    if col_name in ("5DLowPrice", "RecentLTP", "BuyClPrice", "AvgBuyPrice",
                    "TargetPrice", "StoplossPrice", "TotalInvestment",
                    "Profits", "SoldPrice", "SoldPrevClose", "SoldOpen",
                    "SoldHigh", "SoldLow", "SoldClose"):
        return FMT_PRICE
    # Buy-block price columns  (B0_PrevClose, B0_Open, etc.)
    if "_PrevClose" in col_name or "_Open" in col_name or "_High" in col_name \
       or "_Low" in col_name or "_Close" in col_name:
        return FMT_PRICE
    return FMT_INT


def _to_dt(ts):
    """Convert to Python datetime or None."""
    if ts is None:
        return None
    try:
        t = pd.Timestamp(ts)
        if pd.isna(t):
            return None
        return t.to_pydatetime().replace(tzinfo=None)
    except Exception:
        return None


def build_picks_row(sig, sim, price_dict, max_buys, last_data_date):
    """
    Assemble one picks-sheet row from a signal dict and simulation result.
    All date fields are Python datetime objects (formatted DD-MM-YYYY by writer).
    Price fields are rounded to 2 decimal places.
    GainLoss%, 1DChange%, 5DLow% are stored as decimal fractions (e.g. -0.08),
    formatted as percentage in Excel via 0.00% format.
    """
    sym   = str(sig["SYMBOL"])
    order = sim["order"]

    # RecentLTP = latest available close for this symbol
    recent_ltp = None
    if sym in price_dict:
        recent_ltp = round(float(price_dict[sym]["closes"][-1]), 2)
    if recent_ltp is None:
        recent_ltp = round(float(sig["SIGNAL_CLOSE"]), 2)

    # Date fields as Python datetime objects
    buy_date   = _to_dt(sig["SIGNAL_DATE"])
    min5d_date = _to_dt(sig["MIN_5D_DATE"])
    today_date = last_data_date if order != "Invalid" else None
    sold_date  = _to_dt(sim["exit_date"]) if sim["exit_found"] else None

    # BuyClPrice — the signal-day close price
    # For Invalid rows this is still the signal close (same as expected output)
    buy_cl_price = round(float(sig["SIGNAL_CLOSE"]), 2)

    # Unrealized P&L for Open trades (computed using recent_ltp)
    profit   = sim["profit"]
    gain_pct = sim["gain_pct"]
    if order == "Executed" and sim["status"] == "Open":
        avg_buy = sim["avg_buy_price"]
        qty     = sim["total_qty"]
        if avg_buy and avg_buy > 0 and recent_ltp and qty:
            profit   = round((recent_ltp - avg_buy) * qty, 2)
            gain_pct = round((recent_ltp - avg_buy) / avg_buy, 4)

    row = [
        float(sig["PCT_1D_CHANGE"]),           # 1DChange%  (decimal fraction)
        sym,                                    # StockName
        float(sig["PCT_FROM_LOW"]),             # 5DLow%     (decimal fraction)
        round(float(sig["MIN_5D_CLOSE"]), 2),  # 5DLowPrice
        recent_ltp,                             # RecentLTP
        buy_date,                               # BuyDate (datetime)
        buy_cl_price,                           # BuyClPrice
        min5d_date,                             # 5DLowDate (datetime)
        today_date,                             # TodayDate (datetime or None)
        sim["buy_count"],                       # BuyCount
        sim["avg_buy_price"],                   # AvgBuyPrice
        sim["total_qty"],                       # TotalQty
        sim["target_price"],                    # TargetPrice
        sim["stop_price"],                      # StoplossPrice
        sim["total_investment"],                # TotalInvestment
        sim["order"],                           # Order
        sim["status"],                          # Status
        sim["market_days"],                     # Duration
        sim["duration_group"],                  # DurationGroup
        profit,                                 # Profits
        gain_pct,                               # GainLoss% (decimal fraction)
        sim["result_str"],                      # Result
        sim["exit_type"],                       # ExitType
        sim["action"],                          # Action
        sim["buy_chance"],                      # BuyChance
        sold_date,                              # SoldDate (datetime)
        sim["exit_price"],                      # SoldPrice
        sim["sold_prev_close"],                 # SoldPrevClose
        sim["sold_open"],                       # SoldOpen
        sim["sold_high"],                       # SoldHigh
        sim["sold_low"],                        # SoldLow
        sim["sold_close"],                      # SoldClose
    ]

    # Per-buy blocks B0..B(max_buys)
    buys = sim.get("buys", [])
    for b in range(max_buys + 1):
        if b < len(buys):
            bdata = buys[b]
            row += [
                _to_dt(bdata["date"]),
                round(float(bdata["prev_close"]), 2),
                round(float(bdata["open"]),       2),
                round(float(bdata["high"]),       2),
                round(float(bdata["low"]),        2),
                round(float(bdata["close"]),      2),
            ]
        else:
            row += [None, None, None, None, None, None]

    return row


def write_picks_excel(rows, columns, out_path):
    """
    Write picks sheet to Excel.
      Sheet name : Pickse
      Row 1      : Bold header — navy background, white text
      Row 2+     : Data with alternating row fill
      Formats    : prices = 0.00 | percentages = 0.00% | dates = DD-MM-YYYY
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Pickse"

    thin    = Side(style="thin", color="BFBFBF")
    bdr     = Border(left=thin, right=thin, top=thin, bottom=thin)
    ncols   = len(columns)

    # Pre-compute per-column format strings
    col_fmts = [_col_fmt(c) for c in columns]

    # ── Header row (row 1) ────────────────────────────────────────────────────
    ws.append(columns)
    for col_idx, col_name in enumerate(columns, 1):
        cell           = ws.cell(row=1, column=col_idx)
        cell.font      = Font(bold=True, size=10, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = bdr
    ws.row_dimensions[1].height = 18

    # ── Data rows (row 2+) ────────────────────────────────────────────────────
    for r_idx, row in enumerate(rows):
        ws.append(row)
        fill_color  = LIGHT if (r_idx % 2 == 0) else "00FFFFFF"
        actual_row  = r_idx + 2
        for col_idx in range(1, ncols + 1):
            cell               = ws.cell(row=actual_row, column=col_idx)
            cell.fill          = PatternFill("solid", fgColor=fill_color)
            cell.border        = bdr
            cell.alignment     = Alignment(horizontal="center")
            cell.number_format = col_fmts[col_idx - 1]

    # ── Freeze panes & auto-filter ────────────────────────────────────────────
    ws.freeze_panes = ws.cell(row=2, column=1)
    ws.auto_filter.ref = (ws.cell(row=1, column=1).coordinate + ":" +
                          ws.cell(row=1, column=ncols).coordinate)

    # ── Column widths ─────────────────────────────────────────────────────────
    for col_idx, col_name in enumerate(columns, 1):
        w = 13
        if col_name == "StockName":                    w = 16
        elif "Date" in col_name:                       w = 13
        elif col_name in ("Result", "ExitType"):       w = 22
        elif col_name in ("Order", "Status", "Action"): w = 14
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    # ── Page setup ────────────────────────────────────────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0

    wb.save(out_path)


# ─── STYLE (full mode aggregate sheet) ───────────────────────────────────────

def style_sheet(ws, mode_label):
    thin = Side(style="thin", color="BFBFBF")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)

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

    for col_idx, col_name in enumerate(COLUMNS_43, 1):
        cell           = ws.cell(row=2, column=col_idx)
        cell.value     = col_name
        cell.font      = Font(bold=True, size=10, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = bdr
    ws.row_dimensions[2].height = 18

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


# ─── NO-SIGNALS DIAGNOSTIC FILE ──────────────────────────────────────────────

def save_no_signals_file(cfg, mode, mode_label, symbol_filter, signals_file, month_dir, ts_str):
    prefix   = "QuickRun_Picks" if mode == "quick" else "Results"
    out_path = os.path.join(month_dir, f"{prefix}_NoSignals_{ts_str}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "No_Signals"
    ws.append(["NSE Simulation — No Signals Found"])
    ws.cell(1, 1).font = Font(bold=True, size=13, color=WHITE)
    ws.cell(1, 1).fill = PatternFill("solid", fgColor=NAVY)
    ws.append([])
    ws.append(["Field", "Value"])
    for c in [ws.cell(3, 1), ws.cell(3, 2)]:
        c.font = Font(bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY)
    q_or_p = cfg.get("quick_run" if mode == "quick" else "param_sweep", {})
    rows_d = [
        ("Run mode",          mode_label),
        ("Signal start date", cfg.get("signal_start_date", "?")),
        ("Signal end date",   cfg.get("signal_end_date",   "?")),
        ("Signals file",      signals_file),
        ("ATH min filter",    q_or_p.get("ath_min", "?")),
        ("ATH max filter",    q_or_p.get("ath_max", "?")),
        ("Pct min (5d dip)",  q_or_p.get("pct_min", "?")),
        ("Pct max (5d dip)",  q_or_p.get("pct_max", "?")),
        ("Symbol filter",     ", ".join(sorted(symbol_filter)) if symbol_filter else "ALL"),
        ("", ""),
        ("Root Cause", "ATH computed from uploaded data only. Short data -> ATH near current price."),
        ("",           "Most stocks appear near ATH and fail the -30% to -60% ATH filter."),
        ("Solution",   "Upload 1-2 years of historical NSE bhav CSVs to bhav_data/ folder."),
        ("",           "Re-run with rebuild_db=true after uploading."),
    ]
    for r in rows_d:
        ws.append(list(r))
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 80
    wb.save(out_path)
    print(f"✅ Diagnostic file saved : {out_path}")
    return out_path


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="NSE Trading Simulation")
    parser.add_argument("--mode", choices=["quick", "full"], default="full",
                        help="quick = per-trade picks sheet (daily); full = aggregate stats (monthly)")
    parser.add_argument("--symbols", default="",
                        help="Comma-separated symbols e.g. TCS,WIPRO,INFY (empty = all)")
    args = parser.parse_args()
    mode = args.mode

    signals_file = SIGNALS_FILE_QUICK if mode == "quick" else SIGNALS_FILE_FULL

    for f in [CONFIG_FILE, signals_file, EQ_FILE]:
        if not os.path.exists(f):
            print(f"❌ Missing: {f}")
            raise SystemExit(1)

    cfg = load_config()
    investment_per_buy       = cfg.get("investment_per_buy",       10000)
    force_exit_calendar_days = cfg.get("force_exit_calendar_days", 90)
    pending_window_days      = cfg.get("pending_window_days",       30)
    max_rows_per_sheet       = cfg.get("max_rows_per_sheet",        25000)

    trade_combos = build_trade_combos(cfg, mode)
    mode_label   = "QUICK RUN" if mode == "quick" else "FULL PARAMETER SWEEP"
    print(f"Mode            : {mode_label}")
    print(f"Trade combos    : {len(trade_combos):,}")

    # ── Symbol filter (CLI > config > all) ────────────────────────────────────
    symbol_filter = set()
    cli_syms = args.symbols.strip()
    if cli_syms:
        symbol_filter = {s.strip().upper() for s in cli_syms.split(",") if s.strip()}
        print(f"Symbol filter   : CLI override -> {sorted(symbol_filter)}")
    else:
        section     = cfg.get("quick_run" if mode == "quick" else "param_sweep", {})
        config_syms = section.get("watch_symbols", [])
        if config_syms:
            symbol_filter = {s.strip().upper() for s in config_syms if s.strip()}
            print(f"Symbol filter   : config watch_symbols -> {len(symbol_filter)} symbols")
        else:
            print("Symbol filter   : ALL EQ symbols")

    # ── Load signals ──────────────────────────────────────────────────────────
    print("Loading signals...")
    sig_df = pd.read_parquet(signals_file)
    sig_df["SIGNAL_DATE"] = pd.to_datetime(sig_df["SIGNAL_DATE"])

    if symbol_filter:
        sig_df = sig_df[sig_df["SYMBOL"].isin(symbol_filter)]
        missing = symbol_filter - set(sig_df["SYMBOL"].unique())
        if missing:
            print(f"⚠️  Symbols not in signals: {sorted(missing)}")

    print(f"  Total signals : {len(sig_df):,}")

    # ── Output path (month subfolder) ─────────────────────────────────────────
    ts_now    = datetime.now()
    ts_str    = ts_now.strftime("%Y%m%d_%H%M%S")
    month_dir = os.path.join(OUTPUT_DIR, ts_now.strftime("%Y-%m"))
    os.makedirs(month_dir, exist_ok=True)

    if len(sig_df) == 0:
        print("⚠️  No signals found. Saving diagnostic Excel.")
        save_no_signals_file(cfg, mode, mode_label, symbol_filter, signals_file, month_dir, ts_str)
        raise SystemExit(0)

    # ── Load EQ price data ────────────────────────────────────────────────────
    print("Loading EQ price data...")
    eq_df          = pd.read_parquet(EQ_FILE)
    eq_df["DATE1"] = pd.to_datetime(eq_df["DATE1"])
    if symbol_filter:
        eq_df = eq_df[eq_df["SYMBOL"].isin(symbol_filter)]
    price_dict = build_price_dict(eq_df)

    # Last data date (= "TodayDate" for non-Invalid rows)
    last_data_ts   = eq_df["DATE1"].max()
    last_data_date = last_data_ts.to_pydatetime().replace(tzinfo=None)
    print(f"Last data date  : {last_data_date.strftime('%d-%m-%Y')}")
    del eq_df

    # ══════════════════════════════════════════════════════════════════════════
    # QUICK MODE -> Per-trade PICKS SHEET
    # ══════════════════════════════════════════════════════════════════════════
    if mode == "quick":
        q    = cfg["quick_run"]
        mb   = q["max_buys"]
        bd   = q["buy_drop"]
        tgt  = q["target"]
        sl   = q["stoploss"]
        mdur = q["max_duration"]

        print(f"Parameters      : DaysBack={q['days_back']} "
              f"PctMin={q['pct_min']:.0%} PctMax={q['pct_max']:.0%} "
              f"ATHMin={q['ath_min']:.0%} ATHMax={q['ath_max']:.0%} "
              f"MaxBuys={mb} BuyDrop={bd:.0%} Target={tgt:.0%} SL={sl:.0%} MaxDur={mdur}")
        print(f"Running picks simulation for {len(sig_df):,} signals...")

        # Deduplicate: same symbol+date -> keep first
        sig_df_picks = sig_df.drop_duplicates(
            subset=["SYMBOL", "SIGNAL_DATE"]).reset_index(drop=True)
        print(f"Unique signals  : {len(sig_df_picks):,}")

        columns  = get_picks_columns(mb)
        all_rows = []
        counts   = {"Executed": 0, "Pending": 0, "Expired": 0, "Invalid": 0,
                    "Open": 0, "Closed": 0, "Profit": 0, "Loss": 0}

        for _, sig in sig_df_picks.iterrows():
            sim = simulate_trade_detailed(
                str(sig["SYMBOL"]),
                sig["SIGNAL_DATE"],
                float(sig["SIGNAL_CLOSE"]),
                price_dict,
                mb, bd, tgt, sl, mdur,
                investment_per_buy, force_exit_calendar_days, pending_window_days
            )
            row = build_picks_row(sig, sim, price_dict, mb, last_data_date)
            all_rows.append(row)

            o = sim["order"]
            counts[o] = counts.get(o, 0) + 1
            if o == "Executed":
                s = sim["status"] or ""
                counts[s] = counts.get(s, 0) + 1
                if sim.get("result_str"):
                    rs = sim["result_str"].upper()
                    if "PROFIT" in rs: counts["Profit"] += 1
                    elif "LOSS"  in rs: counts["Loss"]  += 1

        prefix   = "QuickRun_Picks"
        out_path = os.path.join(month_dir, f"{prefix}_{ts_str}.xlsx")

        print(f"Writing {len(all_rows):,} rows to Excel (sheet: Pickse)...")
        write_picks_excel(all_rows, columns, out_path)

        print(f"\n✅ Picks sheet saved : {out_path}")
        print(f"✅ Total signals     : {len(all_rows):,}")
        print(f"   Executed          : {counts.get('Executed',0)}  "
              f"(Open={counts.get('Open',0)}  Closed={counts.get('Closed',0)})")
        print(f"   Pending           : {counts.get('Pending',0)}")
        print(f"   Expired           : {counts.get('Expired',0)}")
        print(f"   Invalid           : {counts.get('Invalid',0)}")
        print(f"   Profit trades     : {counts.get('Profit',0)}")
        print(f"   Loss trades       : {counts.get('Loss',0)}")
        return

    # ══════════════════════════════════════════════════════════════════════════
    # FULL MODE -> Aggregate 43-column stats (parameter sweep)
    # ══════════════════════════════════════════════════════════════════════════
    filter_cols   = ["DAYSBACK", "PCTMIN", "PCTMAX", "ATHMIN", "ATHMAX"]
    filter_groups = sig_df.groupby(filter_cols, dropna=False)
    total_expected = len(filter_groups) * len(trade_combos)
    print(f"Filter combos in signals : {len(filter_groups)}")
    print(f"Total output rows        : {total_expected:,}")

    out_path  = os.path.join(month_dir, f"Results_{ts_str}.xlsx")

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

        for (mb, bd_p, tgt, sl, mdur) in trade_combos:
            test_num += 1

            results = [
                simulate_trade(
                    sym, sig_date, sig_close, price_dict,
                    mb, bd_p, tgt, sl, mdur,
                    investment_per_buy, force_exit_calendar_days, pending_window_days
                )
                for sym, sig_date, sig_close in signal_list
            ]

            stats = aggregate_stats(results)
            stats["TotalStock"] = total_stocks

            row  = [test_num, db, pmin, pmax, amin, amax, mb, bd_p, tgt, sl, mdur]
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
