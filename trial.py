#!/usr/bin/env python3

import sys
import os
import re
from datetime import datetime, timedelta
from kiteconnect import KiteConnect
from openpyxl import load_workbook

API_KEY     = "ixick4apw3b5oh22"
API_SECRET  = "whlu97153a2mbhcer97g7mqtkabj3rl3"
TOKEN_FILE  = "access_token.txt"
EXCEL_PATH  = r"C:\Users\HP5CD\OneDrive\Desktop\Stocks\Stocks\RW leverage Calac.xlsx"
SHEET_NAME  = "V1- Code"

kite = KiteConnect(api_key=API_KEY)

def login_and_save_token():
    print("\n🔗  Open this URL in your browser →")
    print(kite.login_url())
    request_token = input("\nPaste request_token from the URL: ").strip()
    if not request_token:
        sys.exit("❌ No request_token provided.")
    try:
        session = kite.generate_session(request_token, api_secret=API_SECRET)
        access_token = session["access_token"]
        with open(TOKEN_FILE, "w") as f:
            f.write(access_token)
        print("✅ Login successful and token saved!")
        return access_token
    except Exception as e:
        sys.exit(f"❌ Failed to generate access token → {e}")

def load_saved_token():
    if not os.path.exists(TOKEN_FILE):
        return None
    with open(TOKEN_FILE, "r") as f:
        return f.read().strip()

def rupees(x):
    return f"{x:,.2f}" if x is not None else "—"

def get_instrument_token(sym: str) -> int:
    for inst in kite.instruments():
        if inst["tradingsymbol"] == sym.split(":")[1] and inst["exchange"] == sym.split(":")[0]:
            return inst["instrument_token"]
    raise ValueError("Instrument token not found.")

def normalize(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9 ]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def format_for_excel(symbol):
    parts = symbol.replace(":", " ").split()
    if len(parts) != 2:
        return symbol  # fallback
    exch, rest = parts
    rest = rest.lower()
    rest = re.sub(r'(\d{2})([A-Z]{3})', r'\1 \2', rest)
    words = re.findall(r'[a-z]+|\d+', rest)
    words.append(exch)
    return ' '.join([w.capitalize() for w in words])

# ─────────────────────────────────────────────
# login
# ─────────────────────────────────────────────
access_token = load_saved_token()
if access_token:
    try:
        kite.set_access_token(access_token)
        kite.profile()
        print("✅ Logged in using saved token.")
    except:
        print("⚠️ Saved token invalid or expired. Please log in again.")
        access_token = login_and_save_token()
        kite.set_access_token(access_token)
else:
    access_token = login_and_save_token()
    kite.set_access_token(access_token)

# ─────────────────────────────────────────────
# main loop
# ─────────────────────────────────────────────
while True:
    print("\n────────────────────────────────────────────")
    symbol = input("Enter symbol (e.g. NSE:INFY, MCX:GOLDM25OCTFUT) or 'q' to quit: ").strip().upper()

    if symbol in {"Q", "QUIT", "EXIT"}:
        print("👋 Bye!")
        break

    if not symbol:
        print("❌ Empty symbol. Try again.")
        continue

    try:
        quote = kite.ltp(symbol)
        ohlc_today = kite.ohlc(symbol)
        last_price = quote[symbol]["last_price"]
        tod_ohlc   = ohlc_today[symbol]["ohlc"]
    except Exception as e:
        print(f"❌ API error while fetching quote: {e}")
        continue

    print(f"\n📊  {symbol}  |  {datetime.now().strftime('%d-%b-%Y %H:%M:%S')}")
    print("────────────────────────────────────────────")
    print(f"Current Price : ₹ {rupees(last_price)}")
    print("Today")
    print(f"  Open        : ₹ {rupees(tod_ohlc['open'])}")
    print(f"  High        : ₹ {rupees(tod_ohlc['high'])}")
    print(f"  Low         : ₹ {rupees(tod_ohlc['low'])}")
    print(f"  Prev Close  : ₹ {rupees(tod_ohlc['close'])}")

    hist_rows = []
    try:
        token = get_instrument_token(symbol)
        today = datetime.today().date()
        start = today - timedelta(days=7)
        candles = kite.historical_data(token, start, today, "day")
        candles = [c for c in candles if c["date"].date() < today][-2:]
        hist_rows = candles[-2:]
    except Exception as e:
        print("⚠️ Couldn’t fetch historical (need add‑on?). Skipping…", e)

    if hist_rows:
        print("────────────────────────────────────────────")
        print("Last 2 completed sessions")
        for row in hist_rows:
            d = row["date"].strftime("%Y-%m-%d")
            print(f"{d} → O:{rupees(row['open'])}  H:{rupees(row['high'])} "
                  f"L:{rupees(row['low'])}  C:{rupees(row['close'])}")

    # ─────────────────────────────────────────────
    # Excel Update (find last stock row properly)
    # ─────────────────────────────────────────────
    print("\n📄 Updating Excel file…")
    try:
        wb = load_workbook(EXCEL_PATH)
        sheet = wb[SHEET_NAME]

        found = False
        last_stock_row = None

        normalized_symbol = normalize(symbol.replace(":", " "))
        symbol_words = normalized_symbol.split()

        row = 12
        while True:
            cell_value = sheet[f"B{row}"].value

            if cell_value:
                cell_str = str(cell_value).strip()
                if 'total' in cell_str.lower():
                    break  # stop at total
                if cell_str != "":
                    last_stock_row = row  # update last stock row

                cell_normalized = normalize(cell_str)
                if all(word in cell_normalized for word in symbol_words):
                    sheet[f"G{row}"] = last_price
                    found = True
                    print(f"✅ Updated existing {symbol} (matched to '{cell_value}') in row {row} with ₹{last_price}")
                    break
            row += 1

        if not found:
            formatted_name = format_for_excel(symbol)
            insert_row = (last_stock_row or 9) + 1
            print(f"ℹ️ No match found. Adding new entry: {formatted_name} → ₹{last_price}")
            sheet[f"B{insert_row}"] = formatted_name
            sheet[f"G{insert_row}"] = last_price
            print(f"✅ Added new {formatted_name} in row {insert_row} with ₹{last_price}")

        wb.save(EXCEL_PATH)

    except Exception as e:
        print(f"❌ Failed to write to Excel: {e}")

    print("────────────────────────────────────────────")
    print("✅ Ready for next input.")
