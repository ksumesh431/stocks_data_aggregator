"""Deterministic tests for portfolio and QQQ benchmark comparisons."""

import unittest

import pandas as pd

from portfolio_report.benchmark import (
    build_cash_flow_matched_benchmark,
    calculate_portfolio_twr,
    calculate_same_period_benchmark_twr,
)


class BenchmarkTests(unittest.TestCase):
    def setUp(self):
        self.trades = pd.DataFrame(
            [
                {
                    "execution_time": pd.Timestamp("2025-01-01 20:00"),
                    "ticker": "ABC",
                    "side": "BUY",
                    "quantity": 10.0,
                },
                {
                    "execution_time": pd.Timestamp("2025-01-02 20:00"),
                    "ticker": "ABC",
                    "side": "BUY",
                    "quantity": 10.0,
                },
            ]
        )
        # Wallet rows are newest-first, including within each calendar date.
        self.wallet = pd.DataFrame(
            [
                {
                    "date": pd.Timestamp("2025-01-02"),
                    "activity_type": "BUY",
                    "amount_usd": 100.0,
                    "updated_trade_balance_usd": 0.0,
                },
                {
                    "date": pd.Timestamp("2025-01-02"),
                    "activity_type": "DEPOSIT",
                    "amount_usd": 100.0,
                    "updated_trade_balance_usd": 100.0,
                },
                {
                    "date": pd.Timestamp("2025-01-01"),
                    "activity_type": "BUY",
                    "amount_usd": 100.0,
                    "updated_trade_balance_usd": 0.0,
                },
                {
                    "date": pd.Timestamp("2025-01-01"),
                    "activity_type": "DEPOSIT",
                    "amount_usd": 100.0,
                    "updated_trade_balance_usd": 100.0,
                },
            ]
        )
        self.portfolio_prices = pd.DataFrame(
            {"ABC": [10.0, 10.0, 11.0]},
            index=pd.to_datetime(["2025-01-01", "2025-01-02", "2025-01-03"]),
        )
        self.qqq_prices = pd.Series(
            [10.0, 10.0, 11.0],
            index=pd.to_datetime(["2025-01-01", "2025-01-02", "2025-01-03"]),
            name="QQQ",
        )

    def test_portfolio_twr_removes_deposit_timing(self):
        result = calculate_portfolio_twr(
            self.trades,
            self.wallet,
            self.portfolio_prices,
            pd.Timestamp("2025-01-01"),
            pd.Timestamp("2025-01-03"),
        )

        self.assertAlmostEqual(result["return"], 0.10, places=7)
        self.assertAlmostEqual(
            result["daily"].loc[pd.Timestamp("2025-01-02"), "daily_return"],
            0.0,
            places=7,
        )

    def test_qqq_twr_uses_the_same_funded_days(self):
        portfolio = calculate_portfolio_twr(
            self.trades,
            self.wallet,
            self.portfolio_prices,
            pd.Timestamp("2025-01-01"),
            pd.Timestamp("2025-01-03"),
        )

        result = calculate_same_period_benchmark_twr(
            self.qqq_prices,
            portfolio["daily"],
        )

        self.assertAlmostEqual(result, 0.10, places=7)

    def test_cash_flow_matched_qqq_uses_exact_dates_and_amounts(self):
        cash_flows = pd.DataFrame(
            [
                {
                    "date": pd.Timestamp("2025-01-01"),
                    "description": "Deposit",
                    "cash_flow_usd": -100.0,
                },
                {
                    "date": pd.Timestamp("2026-01-01"),
                    "description": "Ending securities plus cash",
                    "cash_flow_usd": 110.0,
                },
            ]
        )
        benchmark_prices = pd.Series(
            [10.0, 11.0],
            index=pd.to_datetime(["2025-01-01", "2026-01-01"]),
            name="QQQ",
        )

        result = build_cash_flow_matched_benchmark(cash_flows, benchmark_prices)

        self.assertAlmostEqual(result["terminal_value_usd"], 110.0, places=7)
        self.assertAlmostEqual(result["xirr"], 0.10, places=7)


if __name__ == "__main__":
    unittest.main()
