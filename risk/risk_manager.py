"""
risk/risk_manager.py
---------------------
Portfolio risk metrics and position sizing.

Metrics:
  - Sharpe Ratio (annualised)
  - Sortino Ratio
  - Max Drawdown
  - VaR  (Historical, 95% confidence)
  - CVaR (Expected Shortfall, 95%)
  - Calmar Ratio
  - Win Rate, Profit Factor
"""

import numpy as np
import pandas as pd
from scipy import stats


TRADING_DAYS = 252
CONFIDENCE   = 0.95


def compute_all_metrics(daily_returns: pd.Series, trade_log: pd.DataFrame) -> dict:
    r = daily_returns.dropna()

    sharpe   = _sharpe(r)
    sortino  = _sortino(r)
    mdd, mdd_start, mdd_end = _max_drawdown(r)
    var      = _var(r)
    cvar     = _cvar(r)
    ann_ret  = _annualised_return(r)
    calmar   = ann_ret / abs(mdd) if mdd != 0 else np.nan
    vol      = r.std() * np.sqrt(TRADING_DAYS)

    trade_metrics = _trade_statistics(trade_log) if len(trade_log) > 0 else {}

    return {
        "Annualised Return (%)":  round(ann_ret * 100, 2),
        "Annualised Volatility (%)": round(vol * 100, 2),
        "Sharpe Ratio":           round(sharpe, 3),
        "Sortino Ratio":          round(sortino, 3),
        "Max Drawdown (%)":       round(mdd * 100, 2),
        "Max DD Start":           str(mdd_start),
        "Max DD End":             str(mdd_end),
        "VaR 95% (daily %)":      round(var * 100, 2),
        "CVaR 95% (daily %)":     round(cvar * 100, 2),
        "Calmar Ratio":           round(calmar, 3) if not np.isnan(calmar) else "N/A",
        **trade_metrics,
    }


# ── Private helpers ────────────────────────────────────────────────────────────

def _annualised_return(r: pd.Series) -> float:
    n = len(r)
    return float((1 + r).prod() ** (TRADING_DAYS / n) - 1)


def _sharpe(r: pd.Series, rf: float = 0.0) -> float:
    excess = r - rf / TRADING_DAYS
    return float(excess.mean() / excess.std() * np.sqrt(TRADING_DAYS)) if excess.std() > 0 else 0.0


def _sortino(r: pd.Series, rf: float = 0.0) -> float:
    excess = r - rf / TRADING_DAYS
    downside = excess[excess < 0].std()
    return float(excess.mean() / downside * np.sqrt(TRADING_DAYS)) if downside > 0 else 0.0


def _max_drawdown(r: pd.Series):
    cumulative = (1 + r).cumprod()
    peak = cumulative.cummax()
    drawdown = (cumulative - peak) / peak
    mdd = float(drawdown.min())
    end_idx   = drawdown.idxmin()
    start_idx = cumulative[:end_idx].idxmax()
    return mdd, start_idx.date(), end_idx.date()


def _var(r: pd.Series) -> float:
    return float(-np.percentile(r, (1 - CONFIDENCE) * 100))


def _cvar(r: pd.Series) -> float:
    var = _var(r)
    return float(-r[r <= -var].mean())


def _trade_statistics(trade_log: pd.DataFrame) -> dict:
    if trade_log.empty:
        return {}
    wins = trade_log[trade_log["Result"] == "Win"]
    losses = trade_log[trade_log["Result"] == "Loss"]
    win_rate = len(wins) / len(trade_log) if len(trade_log) > 0 else 0
    avg_win  = wins["Return (%)"].mean() if len(wins) > 0 else 0
    avg_loss = losses["Return (%)"].mean() if len(losses) > 0 else 0
    profit_factor = abs(wins["P&L ($)"].sum() / losses["P&L ($)"].sum()) if losses["P&L ($)"].sum() != 0 else np.nan

    return {
        "Total Trades":       len(trade_log),
        "Win Rate (%)":       round(win_rate * 100, 1),
        "Avg Win (%)":        round(avg_win, 2),
        "Avg Loss (%)":       round(avg_loss, 2),
        "Profit Factor":      round(profit_factor, 2) if not np.isnan(profit_factor) else "N/A",
    }


def position_size(capital: float, var_pct: float, risk_per_trade: float = 0.01) -> float:
    """
    Kelly-inspired position sizing.
    Limits single-position risk to `risk_per_trade` fraction of capital
    relative to the daily VaR estimate.
    """
    if var_pct <= 0:
        return 0
    max_loss_dollars = capital * risk_per_trade
    return max_loss_dollars / var_pct
