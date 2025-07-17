from kiteconnect import KiteConnect
from openpyxl import load_workbook
from datetime import datetime
import os, sys

# ========= Config ===============
API_KEY     = "ixick4apw3b5oh22"
API_SECRET  = "whlu97153a2mbhcer97g7mqtkabj3rl3"
TOKEN_FILE  = "access_token.txt"
EXCEL_PATH  = r"C:\Users\HP5CD\OneDrive\Desktop\Stocks\Stocks\RW leverage Calac.xlsx"
SHEET_NAME  = "V1- Code"
# =================================

# Symbol to Excel row mapping
SYMBOL_ROW_MAP = {
    "MCX:GOLDM25OCTFUT": 12,
    "NSE:SGBFEB32IV-GB": 11,
}

# Symbol-specific details for margin calculation
SYMBOL_DETAILS = {
    "MCX:GOLDM25OCTFUT": {"exchange": "MCX", "tradingsymbol": "GOLDM25OCTFUT", "multiplier": 10},
    "NSE:SGBFEB32IV-GB": {"exchange": "NSE", "tradingsymbol": "SGBFEB32IV-GB", "multiplier": 1},
}

kite = KiteConnect(api_key=API_KEY)

def login():
    print("🔗 Login URL →", kite.login_url())
    req_token = input("Paste request_token from URL: ").strip()
    session = kite.generate_session(req_token, api_secret=API_SECRET)
    with open(TOKEN_FILE, "w") as f:
        f.write(session["access_token"])
    kite.set_access_token(session["access_token"])
    print("✅ Logged in.")
    return

if os.path.exists(TOKEN_FILE):
    with open(TOKEN_FILE) as f:
        token = f.read().strip()
    kite.set_access_token(token)
    try:
        kite.profile()
        print("✅ Using saved token.")
    except:
        print("⚠️ Token expired. Login again.")
        login()
else:
    login()

# Load Excel
try:
    wb = load_workbook(EXCEL_PATH)
    sheet = wb[SHEET_NAME]
except Exception as e:
    sys.exit(f"❌ Excel error: {e}")

while True:
    symbol = input("\n📥 Enter symbol (e.g. NSE:INFY) or 'q' to quit: ").strip().upper()
    if symbol in {"Q", "QUIT", "EXIT"}:
        break

    if symbol not in SYMBOL_ROW_MAP or symbol not in SYMBOL_DETAILS:
        print("⚠️ Symbol not mapped. Add it to SYMBOL_ROW_MAP and SYMBOL_DETAILS.")
        continue

    row_number = SYMBOL_ROW_MAP[symbol]
    details = SYMBOL_DETAILS[symbol]

    try:
        price = kite.ltp(symbol)[symbol]['last_price']
        print(f"✅ Fetched price for {symbol}: ₹{price}")
    except Exception as e:
        print(f"❌ Failed to fetch price: {e}")
        continue

    # Ask for lot size
    lot_input = input("📦 Enter lot size (e.g., 1): ").strip()
    try:
        lot = float(lot_input)
    except ValueError:
        print("⚠️ Invalid lot size. Using 1 as default.")
        lot = 1

    # Get margin required
    try:
        quantity = int(lot * details["multiplier"])
        margin_data = kite.order_margins([{
            "exchange": details["exchange"],
            "tradingsymbol": details["tradingsymbol"],
            "transaction_type": "BUY",       # or "SELL" depending on your use-case
            "variety": "regular",
            "product": "NRML",               # "NRML" for carry-forward, "MIS" for intraday
            "order_type": "MARKET",
            "quantity": quantity
        }])
        required_margin = margin_data[0]['total']
        print(f"💰 Required Margin for {symbol} (Lot: {lot}, Qty: {quantity}): ₹{required_margin}")
    except Exception as e:
        print(f"❌ Failed to fetch margin: {e}")

    # Write to Excel
    try:
        sheet[f"G{row_number}"] = price
        sheet[f"E{row_number}"] = lot
        sheet[f"D{row_number}"] = price * lot

        wb.save(EXCEL_PATH)
        print(f"📄 Excel updated → Row {row_number} | G{row_number} = ₹{price}, E{row_number} = {lot}, D{row_number} = ₹{price * lot}")
    except Exception as e:
        print(f"❌ Failed to write to Excel: {e}")
