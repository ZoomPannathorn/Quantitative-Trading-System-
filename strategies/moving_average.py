"""
strategies/moving_average.py
-----------------------------
Dual Moving Average Crossover strategy.

Entry:  Buy when fast MA crosses above slow MA.
Exit:   Sell when fast MA crosses below slow MA.
"""

import pandas as pd
from .base_strategy import BaseStrategy


class MovingAverageCrossover(BaseStrategy):
    name = "MA Crossover"

    def __init__(self, fast: int = 20, slow: int = 50):
        assert fast < slow, "fast window must be shorter than slow"
        self.fast = fast
        self.slow = slow

    def generate_signals(self, data: dict[str, pd.DataFrame]) -> pd.DataFrame:
        signals = {}
        for ticker, df in data.items():
            close = df["Close"]
            fast_ma = close.rolling(self.fast).mean()
            slow_ma = close.rolling(self.slow).mean()

            sig = pd.Series(0, index=close.index)
            sig[fast_ma > slow_ma] = 1   # bullish regime
            sig[fast_ma < slow_ma] = 0   # bearish — go flat

            # Shift by 1: trade on next day's open to avoid lookahead
            signals[ticker] = sig.shift(1).fillna(0)

        return pd.DataFrame(signals)
