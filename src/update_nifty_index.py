"""
Fetches NIFTY 50 and BANKNIFTY closing prices via yfinance
and writes nifty_index.json to the path given as argv[1].
"""
import sys, json, pytz
from datetime import datetime
import yfinance as yf

out_path = sys.argv[1] if len(sys.argv) > 1 else "nifty_index.json"

IST = pytz.timezone('Asia/Kolkata')
now = datetime.now(IST)

def get_idx(ticker, label):
    try:
        hist = yf.Ticker(ticker).history(period='5d')
        if len(hist) < 2:
            return None
        row, prev = hist.iloc[-1], hist.iloc[-2]
        close = round(float(row['Close']), 2)
        pc    = round(float(prev['Close']), 2)
        chg   = round(close - pc, 2)
        return {
            "label":      label,
            "ticker":     ticker,
            "close":      close,
            "prev_close": pc,
            "open":       round(float(row['Open']),  2),
            "high":       round(float(row['High']),  2),
            "low":        round(float(row['Low']),   2),
            "change":     chg,
            "change_pct": round((chg / pc) * 100, 2) if pc else 0,
            "date":       hist.index[-1].strftime('%d-%b-%Y')
        }
    except Exception as e:
        print(f"  Warning — {label}: {e}")
        return None

data = {
    "updated_at": now.strftime('%d-%b-%Y %I:%M %p IST'),
    "indices": []
}

for t, l in [('^NSEI', 'NIFTY 50'), ('^NSEBANK', 'BANKNIFTY')]:
    idx = get_idx(t, l)
    if idx:
        data['indices'].append(idx)
        print(f"  {l}: {idx['close']} ({idx['change_pct']:+.2f}%)")

with open(out_path, 'w') as f:
    json.dump(data, f, indent=2)

print(f"✅ nifty_index.json written to {out_path}")
