from pathlib import Path

import pandas as pd
import yfinance as yf


YFINANCE_CACHE_DIR = Path("data/.yf-cache")


def _configure_yfinance_cache() -> None:
    YFINANCE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    if hasattr(yf, "set_tz_cache_location"):
        yf.set_tz_cache_location(str(YFINANCE_CACHE_DIR.resolve()))


def get_price_data(ticker: str, period: str = "1y") -> pd.DataFrame:
    """
    Fetch historical stock price data for the given ticker.
    """
    _configure_yfinance_cache()

    df = yf.download(
        tickers=ticker,
        period=period,
        interval="1d",
        auto_adjust=False,
        progress=False,
        threads=False,
    )

    if df is None or df.empty:
        raise ValueError(f"No price data returned for {ticker}")

    df = df.reset_index()

    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [col[0] if isinstance(col, tuple) else col for col in df.columns]

    expected_cols = ["Date", "Open", "High", "Low", "Close", "Volume"]
    missing = [col for col in expected_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Missing expected columns for {ticker}: {missing}")

    df = df[expected_cols].copy()
    df = df[df["Close"].notna()].copy()

    if df.empty:
        raise ValueError(f"Price dataframe is empty after cleaning for {ticker}")

    return df
