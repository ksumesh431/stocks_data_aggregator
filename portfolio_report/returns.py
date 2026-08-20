"""Money-weighted return calculations based only on external cash flows."""

import pandas as pd


def _xnpv(rate, cash_flows):
    """Calculate irregular-date net present value at one annual rate."""
    first_date = cash_flows["date"].min()
    total = 0.0
    for row in cash_flows.itertuples(index=False):
        years = (row.date - first_date).days / 365.0
        total += row.cash_flow_usd / ((1.0 + rate) ** years)
    return total


def calculate_xirr(cash_flows):
    """Return the annualized rate that makes irregular cash-flow NPV equal zero.

    Portfolio cash flows normally have one sign change: deposits are negative
    investor flows and withdrawals/ending value are positive. Bisection is used
    because it is deterministic and safer than an unconstrained Newton iteration.
    """
    cash_flows = cash_flows.copy()
    cash_flows["date"] = pd.to_datetime(cash_flows["date"]).dt.normalize()
    cash_flows = cash_flows.groupby("date", as_index=False)["cash_flow_usd"].sum()

    if cash_flows.empty:
        return None
    if not (cash_flows["cash_flow_usd"] < 0).any():
        return None
    if not (cash_flows["cash_flow_usd"] > 0).any():
        return None

    lower = -0.999999
    upper = 1.0
    lower_value = _xnpv(lower, cash_flows)
    upper_value = _xnpv(upper, cash_flows)

    # Very strong short-period returns can be above 100%, so expand the upper
    # bound until a root is bracketed or the rate is no longer meaningful.
    while lower_value * upper_value > 0 and upper < 1_000_000:
        upper *= 2.0
        upper_value = _xnpv(upper, cash_flows)

    if lower_value * upper_value > 0:
        return None

    for _iteration in range(200):
        midpoint = (lower + upper) / 2.0
        midpoint_value = _xnpv(midpoint, cash_flows)
        if abs(midpoint_value) < 0.0000001:
            return midpoint
        if lower_value * midpoint_value <= 0:
            upper = midpoint
        else:
            lower = midpoint
            lower_value = midpoint_value
    return (lower + upper) / 2.0


def opening_cash_balance(wallet, period_start):
    """Return the final reported wallet balance before the period starts."""
    earlier = wallet[wallet["date"] < pd.Timestamp(period_start)]
    if earlier.empty:
        return 0.0
    latest = earlier.sort_values("date", ascending=False, kind="stable").iloc[0]
    return float(latest["updated_trade_balance_usd"])


def build_xirr_cash_flows(
    fund_movements,
    terminal_date,
    terminal_value,
    period_start=None,
    opening_value=0.0,
):
    """Build investor-perspective cash flows for an XIRR calculation."""
    rows = []
    if period_start is not None and opening_value:
        rows.append(
            {
                "date": pd.Timestamp(period_start).normalize(),
                "description": "Opening portfolio value",
                "cash_flow_usd": -abs(opening_value),
            }
        )

    movements = fund_movements.copy()
    if period_start is not None:
        movements = movements[movements["date"] >= pd.Timestamp(period_start)]
    movements = movements[movements["date"] <= pd.Timestamp(terminal_date)]

    for row in movements.itertuples(index=False):
        is_deposit = row.movement.strip().lower() == "deposit"
        rows.append(
            {
                "date": pd.Timestamp(row.date).normalize(),
                "description": row.movement,
                "cash_flow_usd": -abs(row.usd_amount)
                if is_deposit
                else abs(row.usd_amount),
            }
        )

    rows.append(
        {
            "date": pd.Timestamp(terminal_date).normalize(),
            "description": "Ending securities plus cash",
            "cash_flow_usd": abs(terminal_value),
        }
    )
    return pd.DataFrame(rows).sort_values("date", kind="stable").reset_index(drop=True)
