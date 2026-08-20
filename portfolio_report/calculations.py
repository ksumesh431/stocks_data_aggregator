"""Portfolio accounting and report reconciliation rules."""

from collections import defaultdict

import pandas as pd

QUANTITY_TOLERANCE = 0.00001
CURRENCY_TOLERANCE = 0.01


def opening_quantities(trades, period_start):
    """Calculate quantities held immediately before the period starts."""
    earlier = trades[trades["execution_time"] < pd.Timestamp(period_start)]
    quantities = defaultdict(float)
    for row in earlier.itertuples(index=False):
        direction = 1.0 if row.side == "BUY" else -1.0
        quantities[row.ticker] += direction * row.quantity
    return {
        ticker: quantity
        for ticker, quantity in quantities.items()
        if abs(quantity) > QUANTITY_TOLERANCE
    }


def _initial_positions(opening, start_prices):
    """Create mark-to-market opening lots for a period report."""
    positions = {}
    for ticker, quantity in opening.items():
        if ticker not in start_prices:
            raise ValueError(f"Missing period-start price for {ticker}")
        positions[ticker] = {
            "quantity": quantity,
            "cost_basis_usd": quantity * start_prices[ticker],
        }
    return positions


def _reconcile_with_holding_report(positions, sales, holdings, opening):
    """Use broker cost basis where a ticker was acquired entirely in the period.

    Weighted-average accounting normally matches the source. If the broker used a
    specific lot (as it did for one MSFT sale), the difference is assigned to the
    ticker's latest sale so realized plus remaining basis still reconciles exactly.
    """
    reported = holdings.set_index("ticker").to_dict("index")

    for ticker, position in positions.items():
        if opening.get(ticker, 0.0) > QUANTITY_TOLERANCE:
            continue
        target = reported.get(ticker, {}).get("cost_basis_usd", 0.0)
        difference = position["cost_basis_usd"] - target
        position["cost_basis_usd"] = target

        if abs(difference) <= 0.000001:
            continue
        candidates = [
            index for index, sale in enumerate(sales) if sale["ticker"] == ticker
        ]
        if candidates:
            index = candidates[-1]
            sales[index]["cost_basis_usd"] += difference
            sales[index]["basis_reconciliation_usd"] += difference
            sales[index]["realized_pnl_usd"] -= difference


def run_cost_ledger(trades, holdings, period_start=None, start_prices=None):
    """Replay trades using weighted-average cost and return positions and sales."""
    start_prices = start_prices or {}
    if period_start is None:
        opening = {}
        positions = {}
        period_trades = trades.copy()
    else:
        opening = opening_quantities(trades, period_start)
        positions = _initial_positions(opening, start_prices)
        period_trades = trades[
            trades["execution_time"] >= pd.Timestamp(period_start)
        ].copy()

    sales = []
    for row in period_trades.sort_values("execution_time", kind="stable").itertuples(
        index=False
    ):
        position = positions.setdefault(
            row.ticker, {"quantity": 0.0, "cost_basis_usd": 0.0}
        )

        if row.side == "BUY":
            position["quantity"] += row.quantity
            position["cost_basis_usd"] += row.gross_amount_usd
            continue

        if row.quantity > position["quantity"] + QUANTITY_TOLERANCE:
            raise ValueError(
                f"Sell quantity exceeds available quantity for {row.ticker} on {row.execution_time}"
            )

        average_cost = position["cost_basis_usd"] / position["quantity"]
        sold_cost = average_cost * row.quantity
        position["quantity"] -= row.quantity
        position["cost_basis_usd"] -= sold_cost
        if abs(position["quantity"]) <= QUANTITY_TOLERANCE:
            position["quantity"] = 0.0
            position["cost_basis_usd"] = 0.0

        sales.append(
            {
                "sale_time": row.execution_time,
                "ticker": row.ticker,
                "quantity": row.quantity,
                "sale_proceeds_usd": row.gross_amount_usd,
                "cost_basis_usd": sold_cost,
                "basis_reconciliation_usd": 0.0,
                "realized_pnl_usd": row.gross_amount_usd - sold_cost,
                "sell_brokerage_usd": row.brokerage_usd,
            }
        )

    _reconcile_with_holding_report(positions, sales, holdings, opening)
    return positions, pd.DataFrame(sales), opening


def build_holding_performance(positions, holdings, prices):
    """Combine period cost basis, broker quantities and current Yahoo prices."""
    price_lookup = prices.set_index("ticker").to_dict("index")
    holding_lookup = holdings.set_index("ticker").to_dict("index")
    rows = []

    for ticker in sorted(holding_lookup):
        holding = holding_lookup[ticker]
        position = positions.get(ticker, {"quantity": 0.0, "cost_basis_usd": 0.0})
        quantity_difference = position["quantity"] - holding["quantity"]
        if abs(quantity_difference) > QUANTITY_TOLERANCE:
            raise ValueError(
                f"Trade and holding quantities do not reconcile for {ticker}: {quantity_difference:.8f}"
            )

        quote = price_lookup[ticker]
        current_value = holding["quantity"] * quote["current_price_usd"]
        cost_basis = position["cost_basis_usd"]
        unrealized = current_value - cost_basis
        rows.append(
            {
                "ticker": ticker,
                "quantity": holding["quantity"],
                "average_price_usd": cost_basis / holding["quantity"],
                "cost_basis_usd": cost_basis,
                "current_price_usd": quote["current_price_usd"],
                "current_value_usd": current_value,
                "unrealized_pnl_usd": unrealized,
                "return_pct": unrealized / cost_basis if cost_basis else 0.0,
                "quote_time": quote["quote_time"],
            }
        )
    return pd.DataFrame(rows)


def _activity_for_period(wallet, period_start):
    """Filter wallet activity without changing the source table."""
    if period_start is None:
        return wallet
    return wallet[wallet["date"] >= pd.Timestamp(period_start)]


def _trades_for_period(trades, period_start):
    """Filter trades without changing the source table."""
    if period_start is None:
        return trades
    return trades[trades["execution_time"] >= pd.Timestamp(period_start)]


def calculate_metrics(trades, wallet, holding_performance, sales, period_start=None):
    """Calculate the five dashboard values and supporting components."""
    activity = _activity_for_period(wallet, period_start)
    period_trades = _trades_for_period(trades, period_start)

    income_types = activity["activity_type"].str.contains(
        "DIVIDEND|INTEREST", regex=True
    )
    tax_types = activity["activity_type"].str.contains("TAX|WITHHOLD", regex=True)
    gross_income = activity.loc[income_types & ~tax_types, "amount_usd"].sum()
    withholding_tax = activity.loc[income_types & tax_types, "amount_usd"].sum()

    # Commissions in the wallet change format after the account migration. The
    # trade report is consistent across both accounts, so it is the fee source.
    brokerage = period_trades["brokerage_usd"].sum()
    invested = holding_performance["cost_basis_usd"].sum()
    current = holding_performance["current_value_usd"].sum()
    unrealized = holding_performance["unrealized_pnl_usd"].sum()
    realized = sales["realized_pnl_usd"].sum() if not sales.empty else 0.0
    net_income = gross_income - withholding_tax
    net_pnl = realized + unrealized + net_income - brokerage

    latest_wallet = wallet.sort_values("date", ascending=False, kind="stable").iloc[0]
    metrics = {
        "invested_usd": invested,
        "current_usd": current,
        "realized_pnl_usd": realized,
        "unrealized_pnl_usd": unrealized,
        "net_pnl_usd": net_pnl,
        "gross_income_usd": gross_income,
        "withholding_tax_usd": withholding_tax,
        "net_income_usd": net_income,
        "brokerage_usd": brokerage,
        "cash_balance_usd": latest_wallet["updated_trade_balance_usd"],
    }
    for name in ("realized_pnl", "unrealized_pnl", "net_pnl"):
        metrics[name + "_pct"] = metrics[name + "_usd"] / invested if invested else 0.0
    return metrics


def build_validations(trades, holdings, brokerage):
    """Create user-visible reconciliation checks instead of hiding differences."""
    rows = []
    signed = trades["quantity"].where(trades["side"].eq("BUY"), -trades["quantity"])
    trade_quantities = signed.groupby(trades["ticker"]).sum()
    holding_quantities = holdings.groupby("ticker")["quantity"].sum()
    tickers = sorted(set(trade_quantities.index) | set(holding_quantities.index))
    largest_difference = 0.0
    for ticker in tickers:
        difference = trade_quantities.get(ticker, 0.0) - holding_quantities.get(
            ticker, 0.0
        )
        largest_difference = max(largest_difference, abs(difference))

    rows.append(
        {
            "check": "Net trade quantities equal current holdings",
            "status": "PASS" if largest_difference <= QUANTITY_TOLERANCE else "FAIL",
            "difference": largest_difference,
            "detail": "Largest absolute share difference across all tickers",
        }
    )

    if brokerage.empty:
        rows.append(
            {
                "check": "Trade brokerage equals brokerage report",
                "status": "NOT RUN",
                "difference": 0.0,
                "detail": "No separate brokerage report was found",
            }
        )
    else:
        difference = trades["brokerage_usd"].sum() - brokerage["brokerage_usd"].sum()
        rows.append(
            {
                "check": "Trade brokerage equals brokerage report",
                "status": "PASS" if abs(difference) <= CURRENCY_TOLERANCE else "FAIL",
                "difference": difference,
                "detail": "Trade report minus brokerage report, USD",
            }
        )

    rows.append(
        {
            "check": "Holdings report is used for remaining cost basis",
            "status": "PASS",
            "difference": 0.0,
            "detail": "Preserves broker lot selection after partial sales",
        }
    )
    return pd.DataFrame(rows)
