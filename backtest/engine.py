"""
backtest/engine.py
-------------------
Vectorised backtesting engine.

- Equal-weight allocation across active positions
- Transaction cost of 0.1% per trade (round-trip = 0.2%)
- All signals are shifted by 1 day (no lookahead bias)
- Returns daily portfolio returns and a trade log
"""

import numpy as np
import pandas as pd


TRANSACTION_COST = 0.001   # 10 bps per leg


def run_backtest(
    signals: pd.DataFrame,
    price_data: dict[str, pd.DataFrame],
    initial_capital: float = 1_000_000.0,
) -> dict:
    """
    Run a vectorised backtest.

    Args:
        signals:         DataFrame (dates × tickers), values ∈ {0, 1}
        price_data:      dict {ticker: OHLCV DataFrame}
        initial_capital: Starting portfolio value in USD

    Returns:
        dict with keys:
            equity_curve    : pd.Series  (portfolio value over time)
            daily_returns   : pd.Series  (daily % returns)
            trade_log       : pd.DataFrame
            final_value     : float
    """
    closes = pd.DataFrame({t: df["Close"] for t, df in price_data.items()})
    closes = closes.reindex(columns=signals.columns)

    # Daily returns of each asset
    asset_returns = closes.pct_change().fillna(0)

    # Align signals with price dates
    signals = signals.reindex(closes.index, fill_value=0)

    # Number of active positions each day
    n_active = signals.sum(axis=1).replace(0, np.nan)

    # Equal-weight portfolio return = mean of active-asset returns
    weighted_returns = (signals * asset_returns).sum(axis=1) / n_active
    weighted_returns = weighted_returns.fillna(0)

    # Transaction cost: apply cost whenever signal changes
    signal_changes = signals.diff().abs().sum(axis=1) / 2   # number of position flips
    costs = signal_changes * TRANSACTION_COST
    net_returns = weighted_returns - costs

    # Equity curve
    equity = initial_capital * (1 + net_returns).cumprod()

    # ── Trade log ──────────────────────────────────────────────────────────────
    trades = _build_trade_log(signals, closes, initial_capital)

    return {
        "equity_curve":  equity,
        "daily_returns": net_returns,
        "trade_log":     trades,
        "final_value":   float(equity.iloc[-1]),
        "total_return":  float((equity.iloc[-1] / initial_capital) - 1),
    }


def _build_trade_log(
    signals: pd.DataFrame, closes: pd.DataFrame, capital: float
) -> pd.DataFrame:
    records = []
    position_open = {}   # ticker → {"entry_date": date, "entry_price": float}

    for date in signals.index:
        for ticker in signals.columns:
            sig = signals.loc[date, ticker]
            in_pos = ticker in position_open

            if sig == 1 and not in_pos:
                # Open long
                position_open[ticker] = {
                    "entry_date":  date,
                    "entry_price": closes.loc[date, ticker],
                }
            elif sig == 0 and in_pos:
                # Close position
                entry = position_open.pop(ticker)
                exit_price = closes.loc[date, ticker]
                pct_return = (exit_price - entry["entry_price"]) / entry["entry_price"]
                pct_return -= 2 * TRANSACTION_COST   # round-trip cost
                records.append({
                    "Ticker":       ticker,
                    "Entry Date":   entry["entry_date"].date(),
                    "Exit Date":    date.date(),
                    "Entry Price":  round(entry["entry_price"], 2),
                    "Exit Price":   round(exit_price, 2),
                    "Return (%)":   round(pct_return * 100, 2),
                    "P&L ($)":      round(pct_return * (capital / len(signals.columns)), 2),
                    "Duration (days)": (date - entry["entry_date"]).days,
                    "Result":       "Win" if pct_return > 0 else "Loss",
                })

    # Close any still-open positions at last date
    last_date = signals.index[-1]
    for ticker, entry in position_open.items():
        exit_price = closes.loc[last_date, ticker]
        pct_return = (exit_price - entry["entry_price"]) / entry["entry_price"]
        pct_return -= 2 * TRANSACTION_COST
        records.append({
            "Ticker":           ticker,
            "Entry Date":       entry["entry_date"].date(),
            "Exit Date":        last_date.date(),
            "Entry Price":      round(entry["entry_price"], 2),
            "Exit Price":       round(exit_price, 2),
            "Return (%)":       round(pct_return * 100, 2),
            "P&L ($)":          round(pct_return * (capital / len(signals.columns)), 2),
            "Duration (days)":  (last_date - entry["entry_date"]).days,
            "Result":           "Win" if pct_return > 0 else "Loss",
        })

    return pd.DataFrame(records)
