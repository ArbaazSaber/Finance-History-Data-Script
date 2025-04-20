import os
import pandas as pd
import MetaTrader5 as mt5

from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.styles import Font, numbers

# === CONFIG ===
FILE_PATH = "finance_data.xlsx"
TIMEFRAME = mt5.TIMEFRAME_D1
NUM_CANDLES = 100
TICKERS_SHEET_NAME = "Tickers"
DEFAULT_TICKER = "XAUUSDm"

def initialize_mt5():
    if not mt5.initialize():
        raise RuntimeError("⚠️ Failed to initialize MetaTrader 5.")
    if mt5.account_info() is None:
        mt5.shutdown()
        raise RuntimeError("⚠️ No MT5 account info found.")

def fetch_market_data(ticker: str) -> pd.DataFrame:
    """Fetch market data for the given ticker and return a DataFrame."""
    initialize_mt5()
    rates = mt5.copy_rates_from(ticker, TIMEFRAME, datetime.now(), NUM_CANDLES)
    mt5.shutdown()

    if rates is None or len(rates) == 0:
        print(f"⚠️ No data returned for {ticker}.")
        return pd.DataFrame()

    df = pd.DataFrame(rates)
    df["time"] = pd.to_datetime(df["time"], unit="s")
    return df[["time", "open", "high", "low", "close", "tick_volume", "spread", "real_volume"]]

def append_data_to_sheet(df: pd.DataFrame, ticker: str, file_path: str):
    """Append the given DataFrame to the Excel sheet named after the ticker."""
    sheet_name = f"{ticker} Data"
    if os.path.exists(file_path):
        book = load_workbook(file_path)
        if sheet_name in book.sheetnames:
            existing_df = pd.read_excel(file_path, sheet_name=sheet_name, parse_dates=["time"])
            df = pd.concat([existing_df, df]).drop_duplicates(subset=["time"])
    df.sort_values("time", inplace=True)

    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a" if os.path.exists(file_path) else "w", if_sheet_exists="replace") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    print(f"✅ {ticker} data written to '{sheet_name}' in {file_path}.")

def ensure_workbook_exists(file_path: str, default_ticker: str):
    """Ensure the Excel file and Tickers sheet exist, with at least one default ticker."""
    if not os.path.exists(file_path):
        wb = Workbook()
        ws = wb.active
        ws.title = TICKERS_SHEET_NAME
        ws.append(["Ticker"])
        ws.append([default_ticker])
        wb.save(file_path)
        print(f"📘 Created new workbook with default ticker: {default_ticker}")

def read_tickers(file_path: str) -> list:
    """Return a list of tickers from the Tickers sheet."""
    try:
        tickers_df = pd.read_excel(file_path, sheet_name=TICKERS_SHEET_NAME)
        return tickers_df["Ticker"].dropna().tolist()
    except Exception as e:
        print(f"❌ Error reading tickers: {e}")
        return []
    

def update_overview_sheet(file_path: str, tickers: list, days: int = 15):
    """Creates an 'Overview' sheet showing the last `days` of data and charts for each ticker."""
    book = load_workbook(file_path)

    if "Overview" in book.sheetnames:
        del book["Overview"]

    overview_ws = book.create_sheet("Overview")
    current_row = 1

    for ticker in tickers:
        sheet_name = f"{ticker} Data"
        if sheet_name not in book.sheetnames:
            print(f"⚠️ Sheet '{sheet_name}' not found. Skipping.")
            continue

        df = pd.read_excel(file_path, sheet_name=sheet_name, parse_dates=["time"])
        if df.empty:
            continue

        df = df.sort_values("time").tail(days)

        # --- Write header for each ticker
        overview_ws.cell(row=current_row, column=1, value=f"Ticker: {ticker}")
        overview_ws.cell(row=current_row, column=1).font = Font(bold=True)
        current_row += 1

        for r in dataframe_to_rows(df[["time", "open", "high", "low", "close"]], index=False, header=True):
            for col_idx, value in enumerate(r, start=1):
                overview_ws.cell(row=current_row, column=col_idx, value=value)
            current_row += 1

        # Format the Time column to 'DD-MMM'
        for row in overview_ws.iter_rows(min_row=current_row - len(df), max_row=current_row - 1, min_col=1, max_col=1):
            for cell in row:
                cell.number_format = 'DD-MMM'


        chart = LineChart()
        chart.title = f"{ticker} - Close Price (Last {days} Days)"
        chart.style = 13
        chart.width = 15
        chart.height = 7

        chart.legend = None
        chart.x_axis.title = "Date"
        chart.y_axis.title = "Close Price"

        chart.x_axis.number_format = 'DD-MMM'
        chart.x_axis.majorTimeUnit = "days"

        # Ensure axis lines and labels are shown
        chart.x_axis.majorTickMark = "out"
        chart.y_axis.majorTickMark = "out"
        chart.x_axis.tickLblPos = "nextTo"
        chart.y_axis.tickLblPos = "nextTo"
        chart.x_axis.visible = True
        chart.y_axis.visible = True

        chart.x_axis.majorGridlines = None
        chart.y_axis.majorGridlines = None

        # Reference for data
        header_row = current_row - len(df) - 1
        data_start = header_row + 1
        data_end = data_start + len(df) - 1

        values = Reference(overview_ws, min_col=5, min_row=header_row, max_row=data_end)  # Close column
        categories = Reference(overview_ws, min_col=1, min_row=data_start, max_row=data_end)  # Time column

        chart.add_data(values, titles_from_data=True)
        chart.set_categories(categories)
        overview_ws.add_chart(chart, f"G{data_start}")

        current_row += 2  # spacing between tickers

    book.save(file_path)
    print("📊 Overview sheet updated with charts.")

def main():
    ensure_workbook_exists(FILE_PATH, DEFAULT_TICKER)
    tickers = read_tickers(FILE_PATH)
    
    if not tickers:
        print("❌ No tickers found in the workbook.")
        return

    for ticker in tickers:
        df = fetch_market_data(ticker)
        if not df.empty:
            append_data_to_sheet(df, ticker, FILE_PATH)

    update_overview_sheet(FILE_PATH, tickers, days=15)


if __name__ == "__main__":
    main()