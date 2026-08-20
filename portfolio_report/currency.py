"""Home-currency performance using actual historical funding exchange rates."""

import pandas as pd


def weighted_funding_rate(fund_movements, period_start=None):
    """Return the USD-weighted bank conversion rate for portfolio deposits."""
    deposits = fund_movements[
        fund_movements["movement"].str.strip().str.lower().eq("deposit")
    ].copy()
    deposits = deposits[deposits["usd_amount"].gt(0) & deposits["exchange_rate"].gt(0)]
    if period_start is not None:
        deposits = deposits[deposits["date"] >= pd.Timestamp(period_start)]
    if deposits.empty:
        return None
    converted_total = (deposits["usd_amount"] * deposits["exchange_rate"]).sum()
    return converted_total / deposits["usd_amount"].sum()


def calculate_inr_metrics(
    usd_metrics,
    fund_movements,
    current_fx_rate,
    period_start=None,
):
    """Calculate INR values with historical cost FX and current valuation FX.

    The structured reports do not identify which remittance funded each stock.
    A USD-weighted average of actual deposit exchange rates is therefore used as
    the portfolio cost rate. This mirrors the portfolio-level INR view without
    pretending that every historical dollar was acquired at today's FX rate.
    """
    historical_rate = weighted_funding_rate(fund_movements, period_start)
    used_current_rate_fallback = historical_rate is None
    if used_current_rate_fallback:
        historical_rate = current_fx_rate

    invested = usd_metrics["invested_usd"] * historical_rate
    current = usd_metrics["current_usd"] * current_fx_rate
    realized = usd_metrics["realized_pnl_usd"] * current_fx_rate
    unrealized = current - invested
    gross_income = usd_metrics["gross_income_usd"] * current_fx_rate
    withholding_tax = usd_metrics["withholding_tax_usd"] * current_fx_rate
    net_income = usd_metrics["net_income_usd"] * current_fx_rate
    brokerage = usd_metrics["brokerage_usd"] * current_fx_rate
    cash_balance = usd_metrics["cash_balance_usd"] * current_fx_rate
    net_pnl = realized + unrealized + net_income - brokerage
    fx_contribution = usd_metrics["invested_usd"] * (current_fx_rate - historical_rate)

    return {
        "historical_funding_fx": historical_rate,
        "current_fx": current_fx_rate,
        "used_current_rate_fallback": used_current_rate_fallback,
        "invested_inr": invested,
        "current_inr": current,
        "realized_pnl_inr": realized,
        "unrealized_pnl_inr": unrealized,
        "net_pnl_inr": net_pnl,
        "gross_income_inr": gross_income,
        "withholding_tax_inr": withholding_tax,
        "net_income_inr": net_income,
        "brokerage_inr": brokerage,
        "cash_balance_inr": cash_balance,
        "fx_unrealized_contribution_inr": fx_contribution,
        "realized_pnl_pct": realized / invested if invested else 0.0,
        "unrealized_pnl_pct": unrealized / invested if invested else 0.0,
        "net_pnl_pct": net_pnl / invested if invested else 0.0,
    }
