"""
paper_trading/simulator.py
---------------------------
Live Paper Trading Simulator.

Replays the most recent N days of market data tick-by-tick (one bar = one day),
simulating a live trading session with:

  - Real-time order book (market & limit orders)
  - Position & P&L tracking per ticker
  - Slippage model  (0.05% market impact)
  - Commission model (flat $1 per trade + 0.05% of notional)
  - Risk guardrails  (max position size, daily loss limit)
  - Execution log with timestamps
  - End-of-session performance snapshot
"""

from __future__ import annotations
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from typing import Literal, Optional
import numpy as np
import pandas as pd


# ── Constants ──────────────────────────────────────────────────────────────────
SLIPPAGE_PCT   = 0.0005   # 5 bps market impact
COMMISSION_PCT = 0.0005   # 5 bps of notional
FLAT_COMM      = 1.00     # $1 per order


# ── Data classes ───────────────────────────────────────────────────────────────

@dataclass
class Order:
    order_id:   str
    ticker:     str
    side:       Literal["BUY", "SELL"]
    order_type: Literal["MARKET", "LIMIT"]
    qty:        int
    limit_px:   Optional[float] = None
    status:     str = "OPEN"        # OPEN | FILLED | CANCELLED | REJECTED
    filled_px:  Optional[float] = None
    filled_at:  Optional[datetime] = None
    commission: float = 0.0


@dataclass
class Position:
    ticker:      str
    qty:         int   = 0
    avg_cost:    float = 0.0
    realised_pnl: float = 0.0

    @property
    def market_value(self) -> float:
        return self.qty * self._last_px if hasattr(self, "_last_px") else 0.0

    def update_last_price(self, px: float):
        self._last_px = px

    @property
    def unrealised_pnl(self) -> float:
        if not hasattr(self, "_last_px") or self.qty == 0:
            return 0.0
        return self.qty * (self._last_px - self.avg_cost)


# ── Main Simulator ─────────────────────────────────────────────────────────────

class PaperTradingSimulator:
    """
    Tick-by-tick paper trading simulator driven by historical OHLCV bars.

    Usage:
        sim = PaperTradingSimulator(data, strategy, capital=100_000, lookback_days=120)
        sim.run()
        results = sim.get_results()
    """

    def __init__(
        self,
        price_data: dict[str, pd.DataFrame],
        strategy,
        initial_capital: float = 100_000.0,
        lookback_days:   int   = 120,
        max_position_pct: float = 0.20,   # max 20% of capital per ticker
        daily_loss_limit: float = 0.03,   # halt trading if daily loss > 3%
        verbose: bool = False,
    ):
        self.data            = price_data
        self.strategy        = strategy
        self.initial_capital = initial_capital
        self.capital         = initial_capital
        self.max_pos_pct     = max_position_pct
        self.daily_loss_limit = daily_loss_limit
        self.verbose         = verbose

        # Build trading timeline from the last `lookback_days` bars
        sample_df = next(iter(price_data.values()))
        self.timeline: pd.DatetimeIndex = sample_df.index[-lookback_days:]

        self.positions:  dict[str, Position] = {}
        self.orders:     list[Order]          = []
        self.order_book: list[Order]          = []   # pending limit orders

        # Snapshot series
        self.equity_snapshots: list[dict] = []
        self.daily_pnl_log:    list[dict] = []
        self.execution_log:    list[dict] = []

        self._session_start_equity = initial_capital

    # ── Main loop ──────────────────────────────────────────────────────────────

    def run(self) -> None:
        """Replay market data bar-by-bar, executing strategy decisions."""
        prev_date = None

        for bar_idx, date in enumerate(self.timeline):
            # Current market snapshot
            bar = self._get_bar(date)
            if not bar:
                continue

            # Update last prices for all positions
            for ticker, pos in self.positions.items():
                if ticker in bar:
                    pos.update_last_price(bar[ticker]["close"])

            # ── Day open: try to fill pending limit orders at open price ──────
            self._fill_limit_orders(bar, at_open=True)

            # ── Strategy signal → generate orders ────────────────────────────
            data_slice = {t: df.loc[:date] for t, df in self.data.items()
                          if date in df.index}
            signals = self.strategy.generate_signals(data_slice)
            if date in signals.index:
                self._process_signals(signals.loc[date], bar, date)

            # ── Day close: fill remaining limits + mark-to-market ─────────────
            self._fill_limit_orders(bar, at_open=False)
            self._mark_to_market(date, bar)

            # ── Daily risk check ──────────────────────────────────────────────
            if self._daily_loss_breached():
                if self.verbose:
                    print(f"  [RISK] Daily loss limit hit on {date.date()} — halting day")
                self._flatten_all(bar, date, reason="daily_loss_limit")

            prev_date = date

        if self.verbose:
            print(f"\n[Simulator] Session complete — {len(self.timeline)} bars replayed.")

    # ── Order management ───────────────────────────────────────────────────────

    def submit_market_order(
        self, ticker: str, side: str, qty: int, bar: dict, date: datetime
    ) -> Optional[Order]:
        """Submit and immediately fill a market order."""
        if qty <= 0:
            return None

        px = bar[ticker]["close"]
        # Apply slippage
        slip = px * SLIPPAGE_PCT
        fill_px = px + slip if side == "BUY" else px - slip
        commission = FLAT_COMM + fill_px * qty * COMMISSION_PCT

        # Capital check
        if side == "BUY":
            cost = fill_px * qty + commission
            if cost > self.capital:
                qty = max(0, int((self.capital - commission) / fill_px))
                if qty == 0:
                    return None
                cost = fill_px * qty + commission

        order = Order(
            order_id   = str(uuid.uuid4())[:8],
            ticker     = ticker,
            side       = side,
            order_type = "MARKET",
            qty        = qty,
            status     = "FILLED",
            filled_px  = round(fill_px, 4),
            filled_at  = date,
            commission = round(commission, 2),
        )
        self._apply_fill(order)
        self.orders.append(order)

        self.execution_log.append({
            "Date":       date.date(),
            "Ticker":     ticker,
            "Side":       side,
            "Type":       "MARKET",
            "Qty":        qty,
            "Fill Price": round(fill_px, 2),
            "Commission": round(commission, 2),
            "Order ID":   order.order_id,
        })
        return order

    def _apply_fill(self, order: Order):
        """Update position and cash after a fill."""
        ticker = order.ticker
        px = order.filled_px
        qty = order.qty

        if ticker not in self.positions:
            self.positions[ticker] = Position(ticker=ticker)

        pos = self.positions[ticker]

        if order.side == "BUY":
            new_qty = pos.qty + qty
            pos.avg_cost = (pos.qty * pos.avg_cost + qty * px) / new_qty if new_qty > 0 else px
            pos.qty = new_qty
            self.capital -= (px * qty + order.commission)

        elif order.side == "SELL":
            sold_qty = min(qty, pos.qty)
            if sold_qty > 0:
                realised = sold_qty * (px - pos.avg_cost)
                pos.realised_pnl += realised
                pos.qty -= sold_qty
                self.capital += (px * sold_qty - order.commission)

    def _fill_limit_orders(self, bar: dict, at_open: bool):
        for order in list(self.order_book):
            if order.status != "OPEN":
                self.order_book.remove(order)
                continue
            ticker = order.ticker
            if ticker not in bar:
                continue
            px_key = "open" if at_open else "close"
            cur_px = bar[ticker][px_key]
            if order.side == "BUY" and cur_px <= order.limit_px:
                order.filled_px = order.limit_px
                order.status    = "FILLED"
                order.commission = FLAT_COMM + order.limit_px * order.qty * COMMISSION_PCT
                self._apply_fill(order)
                self.order_book.remove(order)
            elif order.side == "SELL" and cur_px >= order.limit_px:
                order.filled_px = order.limit_px
                order.status    = "FILLED"
                order.commission = FLAT_COMM + order.limit_px * order.qty * COMMISSION_PCT
                self._apply_fill(order)
                self.order_book.remove(order)

    # ── Signal processing ──────────────────────────────────────────────────────

    def _process_signals(self, signal_row: pd.Series, bar: dict, date: datetime):
        for ticker, sig in signal_row.items():
            if ticker not in bar:
                continue
            pos = self.positions.get(ticker, Position(ticker=ticker))
            cur_px = bar[ticker]["close"]
            max_shares = int((self.capital * self.max_pos_pct) / cur_px)

            if sig == 1 and pos.qty == 0 and max_shares > 0:
                self.submit_market_order(ticker, "BUY", max_shares, bar, date)
            elif sig == 0 and pos.qty > 0:
                self.submit_market_order(ticker, "SELL", pos.qty, bar, date)

    def _flatten_all(self, bar: dict, date: datetime, reason: str = ""):
        for ticker, pos in self.positions.items():
            if pos.qty > 0 and ticker in bar:
                self.submit_market_order(ticker, "SELL", pos.qty, bar, date)

    # ── Mark-to-market ─────────────────────────────────────────────────────────

    def _mark_to_market(self, date: datetime, bar: dict):
        unrealised = sum(
            pos.unrealised_pnl for pos in self.positions.values()
        )
        portfolio_value = self.capital + unrealised
        daily_pnl = portfolio_value - self._session_start_equity \
                    if not self.equity_snapshots \
                    else portfolio_value - self.equity_snapshots[-1]["Portfolio Value"]

        self.equity_snapshots.append({
            "Date":            date.date(),
            "Cash":            round(self.capital, 2),
            "Unrealised P&L":  round(unrealised, 2),
            "Portfolio Value": round(portfolio_value, 2),
            "Daily P&L":       round(daily_pnl, 2),
            "Open Positions":  sum(1 for p in self.positions.values() if p.qty > 0),
        })

    def _daily_loss_breached(self) -> bool:
        if not self.equity_snapshots:
            return False
        cur_val  = self.equity_snapshots[-1]["Portfolio Value"]
        sod_val  = self.equity_snapshots[-1]["Portfolio Value"] \
                   if len(self.equity_snapshots) < 2 \
                   else self.equity_snapshots[-2]["Portfolio Value"]
        return (sod_val - cur_val) / sod_val > self.daily_loss_limit

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _get_bar(self, date) -> Optional[dict]:
        bar = {}
        for ticker, df in self.data.items():
            if date in df.index:
                row = df.loc[date]
                bar[ticker] = {
                    "open":  float(row["Open"]),
                    "high":  float(row["High"]),
                    "low":   float(row["Low"]),
                    "close": float(row["Close"]),
                    "volume": int(row["Volume"]),
                }
        return bar if bar else None

    # ── Results ────────────────────────────────────────────────────────────────

    def get_results(self) -> dict:
        equity_df  = pd.DataFrame(self.equity_snapshots)
        exec_df    = pd.DataFrame(self.execution_log) if self.execution_log \
                     else pd.DataFrame(columns=["Date","Ticker","Side","Type",
                                                "Qty","Fill Price","Commission","Order ID"])

        final_val  = equity_df["Portfolio Value"].iloc[-1] if len(equity_df) > 0 \
                     else self.initial_capital
        total_ret  = (final_val - self.initial_capital) / self.initial_capital

        daily_rets = equity_df["Portfolio Value"].pct_change().dropna()
        sharpe     = (daily_rets.mean() / daily_rets.std() * np.sqrt(252)) \
                     if daily_rets.std() > 0 else 0.0

        peak       = equity_df["Portfolio Value"].cummax()
        drawdown   = ((equity_df["Portfolio Value"] - peak) / peak)
        max_dd     = float(drawdown.min())

        total_commissions = exec_df["Commission"].sum() if not exec_df.empty else 0
        n_trades    = len(exec_df)
        win_trades  = 0
        if not exec_df.empty:
            sells = exec_df[exec_df["Side"] == "SELL"]
            win_trades = len(sells[sells["Fill Price"] > 0])  # simplified

        pos_summary = []
        for ticker, pos in self.positions.items():
            pos_summary.append({
                "Ticker":          ticker,
                "Open Qty":        pos.qty,
                "Avg Cost":        round(pos.avg_cost, 2),
                "Realised P&L ($)": round(pos.realised_pnl, 2),
                "Unrealised P&L ($)": round(pos.unrealised_pnl, 2) if pos.qty > 0 else 0.0,
            })

        return {
            "equity_curve":      equity_df,
            "execution_log":     exec_df,
            "positions":         pd.DataFrame(pos_summary),
            "final_value":       round(final_val, 2),
            "total_return":      round(total_ret * 100, 2),
            "sharpe_ratio":      round(sharpe, 3),
            "max_drawdown":      round(max_dd * 100, 2),
            "total_commissions": round(total_commissions, 2),
            "n_trades":          n_trades,
            "summary": {
                "Initial Capital ($)": f"${self.initial_capital:,.0f}",
                "Final Portfolio ($)":  f"${final_val:,.0f}",
                "Total Return (%)":     f"{total_ret*100:.2f}%",
                "Sharpe Ratio":         f"{sharpe:.3f}",
                "Max Drawdown (%)":     f"{max_dd*100:.2f}%",
                "Total Orders":         n_trades,
                "Total Commissions ($)": f"${total_commissions:,.2f}",
                "Strategy":             self.strategy.name,
                "Lookback Days":        len(self.timeline),
            }
        }
