"""
utils/metrics.py
-----------------
Shared helper functions reused across modules.
"""

import numpy as np
import pandas as pd


def rolling_sharpe(returns: pd.Series, window: int = 60, rf: float = 0.0) -> pd.Series:
    """Rolling Sharpe ratio over a given window."""
    excess = returns - rf / 252
    return (excess.rolling(window).mean() / excess.rolling(window).std()) * np.sqrt(252)


def rolling_volatility(returns: pd.Series, window: int = 21) -> pd.Series:
    """Rolling annualised volatility."""
    return returns.rolling(window).std() * np.sqrt(252)


def cumulative_returns(returns: pd.Series) -> pd.Series:
    """Compound cumulative returns starting from 1.0."""
    return (1 + returns).cumprod()


def drawdown_series(returns: pd.Series) -> pd.Series:
    """Drawdown at each point in time."""
    cum = cumulative_returns(returns)
    return (cum - cum.cummax()) / cum.cummax()


def annualised_return(returns: pd.Series) -> float:
    n = len(returns)
    return float((1 + returns).prod() ** (252 / n) - 1)
