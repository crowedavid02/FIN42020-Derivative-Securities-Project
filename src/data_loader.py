from pathlib import Path
import pandas as pd


# ----- paths and key dates -----

# this file lives in src/, so go up one level to the project root
BASE_DIR = Path(__file__).resolve().parents[1]

# your main Excel file
DATA_FILE = BASE_DIR / "data" / "raw" / "CAT_Options_Data.xlsx"

# project dates from the brief
TRADE_DATE = pd.Timestamp("2025-09-19")
EXPIRY_DATE = pd.Timestamp("2025-10-17")


# ----- simple loaders for each sheet -----

def load_calls():
    """Call option chain on the trade date."""
    return pd.read_excel(DATA_FILE, sheet_name="Calls")


def load_puts():
    """Put option chain on the trade date."""
    return pd.read_excel(DATA_FILE, sheet_name="Puts")


def load_price(sort_ascending: bool = True):
    """
    Underlying CAT price series.
    """
    df = pd.read_excel(DATA_FILE, sheet_name="Price")
    if sort_ascending:
        df = df.sort_values("Date").reset_index(drop=True)
    return df


def load_usgg1m(sort_ascending: bool = True):
    """
    1-month rate series (USGG1M).
    """
    df = pd.read_excel(DATA_FILE, sheet_name="USGG1M")
    if sort_ascending:
        df = df.sort_values("Date").reset_index(drop=True)
    return df

def load_atm_series():
    """
    Clean ATM option daily series for CAT 10/17/25 C470.

    Uses the 'ATM' sheet, which must have:
        'Date', 'Last Price', 'Bid Price'

    Returns a DataFrame with:
        'Date', 'Price'
    restricted to TRADE_DATE..EXPIRY_DATE, sorted by date.
    """
    df = pd.read_excel(DATA_FILE, sheet_name="ATM")

    # basic column checks
    for col in ("Date", "Last Price"):
        if col not in df.columns:
            raise ValueError(f"ATM sheet missing required column: {col}")

    # drop rows with missing Date or Last Price (e.g. empty 01/09 row if any)
    df = df.dropna(subset=["Date", "Last Price"])

    # normalise types
    df["Date"] = pd.to_datetime(df["Date"])
    df["Last Price"] = pd.to_numeric(df["Last Price"])

    # window: trade date to expiry
    mask = (df["Date"] >= TRADE_DATE) & (df["Date"] <= EXPIRY_DATE)
    df = df.loc[mask].copy()

    # sort and rename
    df = df.sort_values("Date").reset_index(drop=True)
    df = df.rename(columns={"Last Price": "Price"})

    return df[["Date", "Price"]]


