"""Orchestrate discovery, parsing, market data, calculations and output."""

import logging

import pandas as pd

from portfolio_report.benchmark import build_benchmark_comparison
from portfolio_report.calculations import (
    build_holding_performance,
    build_validations,
    calculate_metrics,
    opening_quantities,
    run_cost_ledger,
)
from portfolio_report.currency import calculate_inr_metrics
from portfolio_report.excel_writer import ensure_output_available, write_workbook
from portfolio_report.market_data import (
    fetch_benchmark_history,
    fetch_latest_market_data,
    fetch_period_start_prices,
)
from portfolio_report.parsers import (
    parse_brokerage_report,
    parse_fund_movement_reports,
    parse_holding_report,
    parse_trade_report,
    parse_wallet_report,
)
from portfolio_report.report_discovery import build_source_inventory, discover_reports
from portfolio_report.returns import (
    build_xirr_cash_flows,
    calculate_xirr,
    opening_cash_balance,
)

LOGGER = logging.getLogger(__name__)


def _build_period_result(
    label,
    table_name,
    trades,
    holdings,
    wallet,
    prices,
    period_start=None,
    start_prices=None,
):
    """Run the complete calculation for one reporting period."""
    positions, sales, opening = run_cost_ledger(
        trades,
        holdings,
        period_start=period_start,
        start_prices=start_prices,
    )
    holding_performance = build_holding_performance(positions, holdings, prices)
    metrics = calculate_metrics(
        trades,
        wallet,
        holding_performance,
        sales,
        period_start=period_start,
    )
    quote_summary = f"latest available for {len(prices)} securities"
    return {
        "label": label,
        "table_name": table_name,
        "metrics": metrics,
        "holdings": holding_performance,
        "sales": sales,
        "opening_quantities": opening,
        "quote_summary": quote_summary,
    }


def _add_xirr(
    result,
    fund_movements,
    wallet,
    valuation_date,
    period_start=None,
    start_prices=None,
):
    """Add auditable investor cash flows and money-weighted XIRR to a result."""
    start_prices = start_prices or {}
    opening_securities = sum(
        quantity * start_prices[ticker]
        for ticker, quantity in result["opening_quantities"].items()
    )
    opening_cash = (
        opening_cash_balance(wallet, period_start) if period_start is not None else 0.0
    )
    terminal_value = (
        result["metrics"]["current_usd"] + result["metrics"]["cash_balance_usd"]
    )
    cash_flows = build_xirr_cash_flows(
        fund_movements,
        terminal_date=valuation_date,
        terminal_value=terminal_value,
        period_start=period_start,
        opening_value=opening_securities + opening_cash,
    )
    result["cash_flows"] = cash_flows
    result["metrics"]["xirr"] = calculate_xirr(cash_flows)


def _add_inr_metrics(result, fund_movements, fx, period_start=None):
    """Add home-currency values using actual historical funding rates."""
    # If positions were carried into the period, their funding predates the
    # period; use all available deposit rates rather than dropping their basis.
    funding_start = period_start if not result["opening_quantities"] else None
    result["inr_metrics"] = calculate_inr_metrics(
        result["metrics"],
        fund_movements,
        fx["rate"],
        period_start=funding_start,
    )


def generate_portfolio_workbook(reports_directory, output_path, current_year=None):
    """Run the application and return the generated workbook path."""
    ensure_output_available(output_path)
    LOGGER.info("Discovering reports in %s", reports_directory)
    reports = discover_reports(reports_directory)

    LOGGER.info("Parsing structured reports")
    trades = parse_trade_report(reports["trades"])
    holding_report = parse_holding_report(reports["holdings"])
    holdings = holding_report["data"]
    wallet = parse_wallet_report(reports["wallet"])
    brokerage = parse_brokerage_report(reports["brokerage"])
    fund_movements = parse_fund_movement_reports(reports["fund_movements"])

    as_of = holding_report["as_of"]
    if pd.isna(as_of):
        as_of = trades["execution_time"].max().normalize()
    if current_year is None:
        current_year = int(as_of.year)
    period_start = pd.Timestamp(year=current_year, month=1, day=1)

    LOGGER.info("Fetching current Yahoo prices and USD/INR")
    prices, fx = fetch_latest_market_data(holdings["ticker"].tolist())

    opening = opening_quantities(trades, period_start)
    LOGGER.info("Fetching period-start prices for %s carried position(s)", len(opening))
    start_prices = fetch_period_start_prices(opening.keys(), period_start)

    LOGGER.info("Calculating all-time and %s performance", current_year)
    all_time_result = _build_period_result(
        "All Time",
        "AllTimeHoldings",
        trades,
        holdings,
        wallet,
        prices,
    )
    current_year_result = _build_period_result(
        str(current_year),
        "CurrentYearHoldings",
        trades,
        holdings,
        wallet,
        prices,
        period_start=period_start,
        start_prices=start_prices,
    )
    _add_xirr(
        all_time_result,
        fund_movements,
        wallet,
        valuation_date=as_of,
    )
    _add_xirr(
        current_year_result,
        fund_movements,
        wallet,
        valuation_date=as_of,
        period_start=period_start,
        start_prices=start_prices,
    )
    _add_inr_metrics(all_time_result, fund_movements, fx)
    _add_inr_metrics(
        current_year_result,
        fund_movements,
        fx,
        period_start=period_start,
    )

    history_start = min(
        trades["execution_time"].min(),
        wallet["date"].min(),
    )
    LOGGER.info("Fetching historical portfolio and QQQ benchmark prices")
    benchmark_history = fetch_benchmark_history(
        trades["ticker"].tolist(),
        history_start,
        as_of,
    )
    LOGGER.info("Calculating portfolio versus QQQ comparisons")
    benchmark_comparison = build_benchmark_comparison(
        trades=trades,
        wallet=wallet,
        close_prices=benchmark_history["portfolio_closes"],
        benchmark_prices=benchmark_history["benchmark_prices"],
        current_year_result=current_year_result,
        all_time_result=all_time_result,
        current_year=current_year,
        as_of=as_of,
    )

    validations = build_validations(trades, holdings, brokerage)
    inventory = build_source_inventory(reports)

    LOGGER.info("Writing workbook to %s", output_path)
    return write_workbook(
        output_path=output_path,
        all_time_result=all_time_result,
        current_year_result=current_year_result,
        trades=trades,
        wallet=wallet,
        fund_movements=fund_movements,
        validations=validations,
        inventory=inventory,
        as_of=as_of,
        current_year=current_year,
        fx=fx,
        benchmark_comparison=benchmark_comparison,
    )
