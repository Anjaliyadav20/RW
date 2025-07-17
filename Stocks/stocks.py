#!/usr/bin/env python3
"""
Kite Stock CLI ― real‑time + last‑2‑days OHLC
Author : <you>
Date   : 2025‑07‑16
"""

import sys
import json
import requests
from datetime import datetime, timedelta
from kiteconnect import KiteConnect

API_KEY     = "ixick4apw3b5oh22"
API_SECRET  = "whlu97153a2mbhcer97g7mqtkabj3rl3"

kite = KiteConnect(api_key=API_KEY)

print("\n🔗  Open this URL in your browser →  (log in & look at the address bar)")
print(kite.login_url())

request_token = input("\nPaste request_token from the URL: ").strip()
if not request_token:
    sys.exit("  No request_token provided.")

try:
    session = kite.generate_session(request_token, api_secret=API_SECRET)
    kite.set_access_token(session["access_token"])
except Exception as e:
    sys.exit(f"❌  Failed to generate access token → {e}")

print("✅  Login successful!")

# ──────────────────────────────────────────
# 3. ASK USER FOR SYMBOL
# ──────────────────────────────────────────
symbol = input("\nEnter symbol (e.g. NSE:INFY, BSE:500209): ").strip().upper()
if not symbol:
    sys.exit("Empty symbol.")

# ──────────────────────────────────────────
# 4. REAL‑TIME QUOTE + TODAY’S OHLC
# ──────────────────────────────────────────
try:
    quote = kite.ltp(symbol)
    ohlc_today = kite.ohlc(symbol)
except Exception as e:
    sys.exit(f"API error while fetching quote: {e}")

last_price = quote[symbol]["last_price"]
tod_ohlc   = ohlc_today[symbol]["ohlc"]

# ──────────────────────────────────────────
# 5. HISTORICAL ― LAST 2 COMPLETE DAYS
#    Requires historical‑data add‑on enabled
# ──────────────────────────────────────────
def get_instrument_token(sym: str) -> int:
    """Look up instrument token from full instrument dump (once per run)."""
    for inst in kite.instruments():
        if inst["tradingsymbol"] == sym.split(":")[1] and inst["exchange"] == sym.split(":")[0]:
            return inst["instrument_token"]
    raise ValueError("Instrument token not found.")

hist_rows = []
try:
    token = get_instrument_token(symbol)
    today  = datetime.today().date()
    start  = today - timedelta(days=7)   # cover weekends
    candles = kite.historical_data(token, start, today, "day")
    # Keep only fully‑completed days (exclude today)
    candles = [c for c in candles if c["date"].date() < today][-2:]
    hist_rows = candles[-2:]
except Exception as e:
    print("⚠️  Couldn’t fetch historical (need add‑on?). Skipping…", e)

# ──────────────────────────────────────────
# 6. DISPLAY
# ──────────────────────────────────────────
def rupees(x):  # nice rounding
    return f"{x:,.2f}" if x is not None else "—"

print(f"\n📊  {symbol}  |  {datetime.now().strftime('%d‑%b‑%Y %H:%M:%S')}")
print("────────────────────────────────────────────")
print(f"Current Price : ₹ {rupees(last_price)}")
print("────────────────────────────────────────────")
print("Today")
print(f"  Open        : ₹ {rupees(tod_ohlc['open'])}")
print(f"  High        : ₹ {rupees(tod_ohlc['high'])}")
print(f"  Low         : ₹ {rupees(tod_ohlc['low'])}")
print(f"  Prev Close  : ₹ {rupees(tod_ohlc['close'])}")

if hist_rows:
    print("────────────────────────────────────────────")
    print("Last 2 completed sessions")
    for row in hist_rows:
        d = row["date"].strftime("%Y‑%m‑%d")
        print(f"{d} → O:{rupees(row['open'])}  H:{rupees(row['high'])} "
              f"L:{rupees(row['low'])}  C:{rupees(row['close'])}")

print("────────────────────────────────────────────")
print("Done ✔")








from openpyxl import load_workbook

def rupees(x):  # nice rounding
    return f"{x:,.2f}" if x is not None else "—"

print(f"\n📊  {symbol}  |  {datetime.now().strftime('%d‑%b‑%Y %H:%M:%S')}")
print("────────────────────────────────────────────")
print(f"Current Price : ₹ {rupees(last_price)}")
print("────────────────────────────────────────────")
print("Today")
print(f"  Open        : ₹ {rupees(tod_ohlc['open'])}")
print(f"  High        : ₹ {rupees(tod_ohlc['high'])}")
print(f"  Low         : ₹ {rupees(tod_ohlc['low'])}")
print(f"  Prev Close  : ₹ {rupees(tod_ohlc['close'])}")

if hist_rows:
    print("────────────────────────────────────────────")
    print("Last 2 completed sessions")
    for row in hist_rows:
        d = row["date"].strftime("%Y‑%m‑%d")
        print(f"{d} → O:{rupees(row['open'])}  H:{rupees(row['high'])} "
              f"L:{rupees(row['low'])}  C:{rupees(row['close'])}")

print("────────────────────────────────────────────")
print("✔ Real‑time stock fetch complete.")

# ─────────────────────────────────────────────
# Write fixed symbols to Excel
# ─────────────────────────────────────────────
print("\n📄 Updating Excel file…")

excel_path = r"C:\Users\HP5CD\OneDrive\Desktop\Stocks\Stocks\RW leverage Calac.xlsx"

symbols_to_update = {
    "NSE:SGBFEB32IV-GB": "G11",
    "MCX:GOLDM25OCTFUT": "G12"
}

wb = load_workbook(excel_path)
sheet = wb["V1- Code"]

for sym, cell in symbols_to_update.items():
    try:
        quote = kite.ltp(sym)
        price = quote[sym]["last_price"]
        print(f"✅ Fetched {sym}: ₹{price}")
        sheet[cell] = price
    except Exception as e:
        print(f"❌ Failed to fetch/write {sym}: {e}")

wb.save(excel_path)
print(f"📄 Excel file updated & saved → {excel_path}")
print("────────────────────────────────────────────")
print("🎯 Done ✔")
