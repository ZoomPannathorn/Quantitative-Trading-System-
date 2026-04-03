"""
main.py
--------
Entry point for the Quantitative Trading System.

Pipeline:
  1. Load / simulate market data
  2. Generate signals + run vectorised backtests (full history)
  3. Compute risk metrics
  4. Run paper trading simulator (recent 120 days, tick-by-tick)
  5. Export combined multi-sheet Excel report
"""

import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from data.data_loader import fetch_data
from strategies.moving_average import MovingAverageCrossover
from strategies.momentum import MomentumStrategy
from strategies.mean_reversion import MeanReversionStrategy
from backtest.engine import run_backtest
from risk.risk_manager import compute_all_metrics
from paper_trading.simulator import PaperTradingSimulator
from reporting.excel_reporter import generate_report


# ── Configuration ─────────────────────────────────────────────────────────────

TICKERS          = ["AAPL", "MSFT", "GOOGL", "AMZN", "TSLA", "SPY", "QQQ", "NVDA"]
START            = "2020-01-01"
END              = "2024-12-31"
CAPITAL          = 1_000_000
PAPER_CAPITAL    = 100_000
PAPER_LOOKBACK   = 120
OUTPUT           = "quant_report.xlsx"

STRATEGIES = [
    MovingAverageCrossover(fast=20, slow=50),
    MomentumStrategy(lookback=20, top_n=3),
    MeanReversionStrategy(window=20, num_std=2.0),
]


def main():
    print("=" * 60)
    print("  Quantitative Trading System")
    print("=" * 60)

    print(f"\n[1/5] Loading market data  ({START} -> {END})...")
    data = fetch_data(TICKERS, START, END)
    n_days = len(next(iter(data.values())))
    print(f"      Loaded {len(data)} tickers, {n_days} trading days each.")

    print("\n[2/5] Running full-history backtests...")
    results = {}
    for strategy in STRATEGIES:
        print(f"      > {strategy.name:<25}", end=" ", flush=True)
        signals = strategy.generate_signals(data)
        result  = run_backtest(signals, data, initial_capital=CAPITAL)
        results[strategy.name] = result
        tr = result["total_return"] * 100
        fv = result["final_value"]
        print(f"Return: {tr:+.1f}%   NAV: ${fv:,.0f}")

    print("\n[3/5] Computing risk metrics...")
    metrics = {}
    for strat_name, res in results.items():
        metrics[strat_name] = compute_all_metrics(
            res["daily_returns"], res["trade_log"]
        )

    print(f"\n[4/5] Running paper trading sessions  (last {PAPER_LOOKBACK} days)...")
    paper_results = {}
    for strategy in STRATEGIES:
        print(f"      > {strategy.name:<25}", end=" ", flush=True)
        sim = PaperTradingSimulator(
            price_data       = data,
            strategy         = strategy,
            initial_capital  = PAPER_CAPITAL,
            lookback_days    = PAPER_LOOKBACK,
            max_position_pct = 0.20,
            daily_loss_limit = 0.03,
            verbose          = False,
        )
        sim.run()
        res = sim.get_results()
        paper_results[strategy.name] = res
        print(f"Return: {res['total_return']:+.2f}%   "
              f"Sharpe: {res['sharpe_ratio']:.3f}   "
              f"Orders: {res['n_trades']}   "
              f"Comm: ${res['total_commissions']:,.2f}")

    print("\n[5/5] Generating Excel report...")
    generate_report(results, metrics, paper_results, output_path=OUTPUT)

    print("\n" + "=" * 60)
    print(f"{'Strategy':<25} {'BT Ann.Ret':>10} {'Sharpe':>8} {'PT Return':>10}")
    print("-" * 60)
    for strat in STRATEGIES:
        n = strat.name
        bt_ret = metrics[n]["Annualised Return (%)"]
        sharpe = metrics[n]["Sharpe Ratio"]
        pt_ret = paper_results[n]["total_return"]
        print(f"{n:<25} {bt_ret:>9.1f}%  {sharpe:>7.3f}  {pt_ret:>+9.2f}%")
    print("=" * 60)
    print(f"\nReport saved: {OUTPUT}")
    print("Sheets: Summary | Equity Curves | Trade Log | Risk Metrics |")
    print("        Monthly Returns | Drawdown | PT Overview |")
    print("        PT Execution Log | PT Live Positions")


if __name__ == "__main__":
    main()
