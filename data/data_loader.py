"""
data/data_loader.py
-------------------
Market data ingestion via yfinance.
Falls back to realistic simulated OHLCV data when network is unavailable.
"""

import numpy as np
import pandas as pd
from datetime import datetime, timedelta


def fetch_data(tickers: list[str], start: str, end: str) -> dict[str, pd.DataFrame]:
    """
    Fetch OHLCV data for a list of tickers.
    Uses yfinance; falls back to simulation if unavailable.

    Returns:
        dict mapping ticker -> DataFrame with columns [Open, High, Low, Close, Volume]
    """
    try:
        import yfinance as yf
        data = {}
        for ticker in tickers:
            df = yf.download(ticker, start=start, end=end, progress=False)
            if df.empty:
                raise ValueError(f"No data for {ticker}")
            data[ticker] = df[["Open", "High", "Low", "Close", "Volume"]]
        return data
    except Exception:
        print("[DataLoader] yfinance unavailable — using simulated market data.")
        return simulate_market_data(tickers, start, end)


def simulate_market_data(
    tickers: list[str], start: str, end: str, seed: int = 42
) -> dict[str, pd.DataFrame]:
    """
    Generate realistic OHLCV data via geometric Brownian motion.
    Annualised drift ~8%, vol ~20% — representative of large-cap equities.
    """
    rng = np.random.default_rng(seed)
    dates = pd.bdate_range(start=start, end=end)
    n = len(dates)

    # Distinct starting prices and vol profiles per ticker
    profiles = {
        "AAPL": (180.0, 0.22),
        "MSFT": (370.0, 0.20),
        "GOOGL": (140.0, 0.24),
        "AMZN": (180.0, 0.26),
        "TSLA": (250.0, 0.45),
        "SPY":  (450.0, 0.14),
        "QQQ":  (380.0, 0.18),
        "NVDA": (500.0, 0.40),
    }

    data = {}
    for ticker in tickers:
        s0, annual_vol = profiles.get(ticker, (100.0, 0.20))
        dt = 1 / 252
        drift = 0.08 * dt
        vol = annual_vol * np.sqrt(dt)

        # Correlated returns via shared market factor
        market = rng.normal(0, 1, n)
        idio   = rng.normal(0, 1, n)
        beta   = 0.7
        shocks = drift + vol * (beta * market + np.sqrt(1 - beta**2) * idio)
        log_prices = np.log(s0) + np.cumsum(shocks)
        closes = np.exp(log_prices)

        daily_range = closes * annual_vol / np.sqrt(252) * rng.uniform(0.5, 1.5, n)
        opens  = closes * np.exp(rng.normal(0, vol * 0.3, n))
        highs  = np.maximum(opens, closes) + daily_range * rng.uniform(0.2, 0.6, n)
        lows   = np.minimum(opens, closes) - daily_range * rng.uniform(0.2, 0.6, n)
        volume = (s0 * 1e6 * rng.lognormal(0, 0.5, n)).astype(int)

        data[ticker] = pd.DataFrame(
            {"Open": opens, "High": highs, "Low": lows, "Close": closes, "Volume": volume},
            index=dates,
        )
    return data


def compute_returns(price_series: pd.Series) -> pd.Series:
    """Daily log returns."""
    return np.log(price_series / price_series.shift(1)).dropna()
