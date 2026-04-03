"""
strategies/mean_reversion.py
-----------------------------
Bollinger Band Mean-Reversion strategy.

Entry Long:   Close drops below lower band  → expect bounce back to mean
Exit Long:    Close crosses back above middle band (SMA)
Entry Short:  Close rises above upper band  → expect reversion (if allow_short)
"""

import pandas as pd
import numpy as np
from .base_strategy import BaseStrategy


class MeanReversionStrategy(BaseStrategy):
    name = "Mean Reversion"
    allow_short = False   # long-only variant

    def __init__(self, window: int = 20, num_std: float = 2.0):
        self.window = window
        self.num_std = num_std

    def generate_signals(self, data: dict[str, pd.DataFrame]) -> pd.DataFrame:
        signals = {}
        for ticker, df in data.items():
            close = df["Close"]
            sma   = close.rolling(self.window).mean()
            std   = close.rolling(self.window).std()
            lower = sma - self.num_std * std

            sig = pd.Series(0, index=close.index)
            in_trade = False

            for i in range(self.window, len(close)):
                if not in_trade:
                    if close.iloc[i] < lower.iloc[i]:   # price below lower band
                        in_trade = True
                        sig.iloc[i] = 1
                else:
                    sig.iloc[i] = 1
                    if close.iloc[i] >= sma.iloc[i]:    # price reverted to mean
                        in_trade = False

            signals[ticker] = sig.shift(1).fillna(0)

        return pd.DataFrame(signals)
