"""
strategies/momentum.py
-----------------------
Cross-sectional Momentum strategy.

Each period, rank all tickers by their past `lookback`-day return.
Go long the top `top_n` performers; hold equal weight.
"""

import pandas as pd
import numpy as np
from .base_strategy import BaseStrategy


class MomentumStrategy(BaseStrategy):
    name = "Momentum"

    def __init__(self, lookback: int = 20, top_n: int = 3):
        self.lookback = lookback
        self.top_n = top_n

    def generate_signals(self, data: dict[str, pd.DataFrame]) -> pd.DataFrame:
        # Build combined close price DataFrame
        closes = pd.DataFrame({t: df["Close"] for t, df in data.items()})

        # Past returns over lookback window
        past_returns = closes.pct_change(self.lookback)

        signals = pd.DataFrame(0, index=closes.index, columns=closes.columns)

        for date in closes.index[self.lookback:]:
            row = past_returns.loc[date].dropna()
            if len(row) < self.top_n:
                continue
            top_tickers = row.nlargest(self.top_n).index
            signals.loc[date, top_tickers] = 1

        # Shift to avoid lookahead
        return signals.shift(1).fillna(0)
