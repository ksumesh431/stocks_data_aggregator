"""Small deterministic tests for the accounting layer."""

import unittest

import pandas as pd

from portfolio_report.calculations import (
    build_holding_performance,
    calculate_metrics,
    run_cost_ledger,
)
from portfolio_report.currency import calculate_inr_metrics, weighted_funding_rate
from portfolio_report.returns import build_xirr_cash_flows, calculate_xirr


class CalculationTests(unittest.TestCase):
    def setUp(self):
        self.trades = pd.DataFrame(
            [
                {
                    "execution_time": pd.Timestamp("2025-06-01"),
                    "ticker": "ABC",
                    "side": "BUY",
                    "quantity": 10.0,
                    "gross_amount_usd": 100.0,
                    "brokerage_usd": 1.0,
                },
                {
                    "execution_time": pd.Timestamp("2026-06-01"),
                    "ticker": "ABC",
                    "side": "SELL",
                    "quantity": 4.0,
                    "gross_amount_usd": 60.0,
                    "brokerage_usd": 1.0,
                },
            ]
        )
        self.holdings = pd.DataFrame(
            [{"ticker": "ABC", "quantity": 6.0, "cost_basis_usd": 60.0}]
        )
        self.prices = pd.DataFrame(
            [
                {
                    "ticker": "ABC",
                    "current_price_usd": 20.0,
                    "quote_time": "2026-08-20 16:00 EDT",
                }
            ]
        )
        self.wallet = pd.DataFrame(
            [
                {
                    "date": pd.Timestamp("2026-07-01"),
                    "activity_type": "DIVIDEND",
                    "amount_usd": 2.0,
                    "updated_trade_balance_usd": 10.0,
                },
                {
                    "date": pd.Timestamp("2026-07-01"),
                    "activity_type": "DIVIDEND TAX",
                    "amount_usd": 0.5,
                    "updated_trade_balance_usd": 9.5,
                },
            ]
        )

    def test_all_time_metrics_reconcile(self):
        positions, sales, _unused_opening = run_cost_ledger(self.trades, self.holdings)
        holdings = build_holding_performance(positions, self.holdings, self.prices)
        metrics = calculate_metrics(self.trades, self.wallet, holdings, sales)

        self.assertAlmostEqual(metrics["invested_usd"], 60.0)
        self.assertAlmostEqual(metrics["current_usd"], 120.0)
        self.assertAlmostEqual(metrics["realized_pnl_usd"], 20.0)
        self.assertAlmostEqual(metrics["unrealized_pnl_usd"], 60.0)
        self.assertAlmostEqual(metrics["net_pnl_usd"], 79.5)

    def test_current_year_marks_opening_position_to_market(self):
        period_start = pd.Timestamp("2026-01-01")
        positions, sales, opening = run_cost_ledger(
            self.trades,
            self.holdings,
            period_start=period_start,
            start_prices={"ABC": 12.0},
        )
        holdings = build_holding_performance(positions, self.holdings, self.prices)

        self.assertAlmostEqual(opening["ABC"], 10.0)
        self.assertAlmostEqual(sales.iloc[0]["realized_pnl_usd"], 12.0)
        self.assertAlmostEqual(holdings.iloc[0]["cost_basis_usd"], 72.0)
        self.assertAlmostEqual(holdings.iloc[0]["unrealized_pnl_usd"], 48.0)

    def test_xirr_for_one_year_ten_percent_return(self):
        cash_flows = pd.DataFrame(
            [
                {"date": pd.Timestamp("2025-01-01"), "cash_flow_usd": -100.0},
                {"date": pd.Timestamp("2026-01-01"), "cash_flow_usd": 110.0},
            ]
        )
        self.assertAlmostEqual(calculate_xirr(cash_flows), 0.10, places=7)

    def test_xirr_cash_flows_use_external_movements_and_terminal_value(self):
        movements = pd.DataFrame(
            [
                {
                    "date": pd.Timestamp("2026-02-01"),
                    "movement": "Deposit",
                    "usd_amount": 100.0,
                },
                {
                    "date": pd.Timestamp("2026-03-01"),
                    "movement": "Withdrawal",
                    "usd_amount": -25.0,
                },
            ]
        )
        cash_flows = build_xirr_cash_flows(
            movements,
            terminal_date=pd.Timestamp("2026-12-31"),
            terminal_value=120.0,
        )
        self.assertEqual(cash_flows["cash_flow_usd"].tolist(), [-100.0, 25.0, 120.0])

    def test_inr_unrealized_return_includes_currency_movement(self):
        metrics = {
            "invested_usd": 100.0,
            "current_usd": 110.0,
            "realized_pnl_usd": 0.0,
            "gross_income_usd": 0.0,
            "withholding_tax_usd": 0.0,
            "net_income_usd": 0.0,
            "brokerage_usd": 0.0,
            "cash_balance_usd": 0.0,
        }
        movements = pd.DataFrame(
            [
                {
                    "date": pd.Timestamp("2026-01-01"),
                    "movement": "Deposit",
                    "usd_amount": 100.0,
                    "exchange_rate": 80.0,
                }
            ]
        )
        inr = calculate_inr_metrics(metrics, movements, current_fx_rate=90.0)

        self.assertAlmostEqual(inr["invested_inr"], 8000.0)
        self.assertAlmostEqual(inr["current_inr"], 9900.0)
        self.assertAlmostEqual(inr["unrealized_pnl_inr"], 1900.0)
        self.assertAlmostEqual(inr["unrealized_pnl_pct"], 0.2375)
        self.assertAlmostEqual(inr["fx_unrealized_contribution_inr"], 1000.0)

    def test_weighted_funding_rate_uses_usd_weights(self):
        movements = pd.DataFrame(
            [
                {
                    "date": pd.Timestamp("2026-01-01"),
                    "movement": "Deposit",
                    "usd_amount": 100.0,
                    "exchange_rate": 80.0,
                },
                {
                    "date": pd.Timestamp("2026-02-01"),
                    "movement": "Deposit",
                    "usd_amount": 300.0,
                    "exchange_rate": 88.0,
                },
            ]
        )
        self.assertAlmostEqual(weighted_funding_rate(movements), 86.0)


if __name__ == "__main__":
    unittest.main()
