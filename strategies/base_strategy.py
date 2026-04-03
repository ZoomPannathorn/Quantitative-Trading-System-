"""
strategies/base_strategy.py
----------------------------
Abstract base class all strategies must implement.
"""

from abc import ABC, abstractmethod
import pandas as pd


class BaseStrategy(ABC):
    """
    All strategies subclass this and implement `generate_signals`.

    Signals convention:
        +1  → long
         0  → flat / no position
        -1  → short  (set self.allow_short = False to disable)
    """

    name: str = "BaseStrategy"
    allow_short: bool = False

    @abstractmethod
    def generate_signals(self, data: dict[str, pd.DataFrame]) -> pd.DataFrame:
        """
        Given OHLCV data for multiple tickers, return a DataFrame of signals.

        Args:
            data: dict {ticker: OHLCV DataFrame}

        Returns:
            DataFrame with DatetimeIndex, one column per ticker, values in {-1, 0, 1}
        """

    def __repr__(self) -> str:
        return f"<Strategy: {self.name}>"
