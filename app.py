import os
import pandas as pd
from datetime import datetime
import MetaTrader5 as mt5
from openpyxl import load_workbook

# === CONFIG ===
FILE_PATH = "finance_data.xlsx"
SHEET_NAME = "Gold Data"
TICKER = "XAUUSDm"
TIMEFRAME = mt5.TIMEFRAME_D1
NUM_CANDLES = 100

def fetch_ticker_data(ticker: str, timeframe, num_bars: int) -> pd.DataFrame:
    """Fetches ticker data from MT5 and returns a cleaned pandas DataFrame."""
    if not mt5.initialize():
        print("⚠️ Failed to initialize MT5.")
        return None

    if mt5.account_info() is None:
        print("⚠️ No MT5 account info found.")
        mt5.shutdown()
        return None

    rates = mt5.copy_rates_from(ticker, timeframe, datetime.now(), num_bars)
    mt5.shutdown()

    if rates is None or len(rates) == 0:
        print("⚠️ No market data received.")
        return None

    df = pd.DataFrame(rates)
    df["time"] = pd.to_datetime(df["time"], unit="s")
    return df[["time", "open", "high", "low", "close", "tick_volume", "spread", "real_volume"]]

def append_to_excel(df: pd.DataFrame, file_path: str, sheet_name: str):
    """Appends data to an Excel sheet, avoiding duplicates based on time column."""
    if os.path.exists(file_path):
        book = load_workbook(file_path)
        if sheet_name in book.sheetnames:
            existing_df = pd.read_excel(file_path, sheet_name=sheet_name, parse_dates=["time"])
            df = pd.concat([existing_df, df]).drop_duplicates(subset=["time"])
    df.sort_values("time", inplace=True)

    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a" if os.path.exists(file_path) else "w", if_sheet_exists="replace") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    print(f"✅ Gold data written to '{sheet_name}' in {file_path} without duplicates.")

def main():
    df = fetch_ticker_data(TICKER, TIMEFRAME, NUM_CANDLES)
    if df is not None and not df.empty:
        append_to_excel(df, FILE_PATH, SHEET_NAME)
    else:
        print("❌ No data to write.")

if __name__ == "__main__":
    main()
    # print(fetch_ticker_data(TICKER, TIMEFRAME, NUM_CANDLES))
