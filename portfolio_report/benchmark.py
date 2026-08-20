"""Compare portfolio performance with a cash-flow-consistent QQQ benchmark."""

import pandas as pd

from portfolio_report.returns import calculate_xirr

VALUE_TOLERANCE = 0.000001


def _normalized_dates(values):
    """Return timezone-free calendar dates for a pandas date-like collection."""
    dates = pd.to_datetime(values)
    if getattr(dates.dtype, "tz", None) is not None:
        dates = dates.dt.tz_localize(None)
    return dates.dt.normalize()


def _new_york_trade_dates(execution_times):
    """Convert the INDmoney execution timestamps from IST to US market dates."""
    times = pd.to_datetime(execution_times)
    if getattr(times.dtype, "tz", None) is None:
        times = times.dt.tz_localize("Asia/Kolkata")
    return (
        times.dt.tz_convert("America/New_York")
        .dt.tz_localize(None)
        .dt.normalize()
    )


def _align_prices(prices, calendar, tickers):
    """Forward-fill market closes onto calendar days without hiding missing data."""
    prices = prices.copy()
    prices.index = pd.to_datetime(prices.index).tz_localize(None).normalize()
    prices = prices[~prices.index.duplicated(keep="last")].sort_index()
    prices = prices.reindex(columns=tickers)

    combined_index = prices.index.union(calendar)
    return prices.reindex(combined_index).sort_index().ffill().reindex(calendar)


def _daily_positions(trades, calendar, tickers):
    """Replay signed trade quantities into an end-of-day position matrix."""
    trades = trades.copy()
    trades["market_date"] = _new_york_trade_dates(trades["execution_time"])
    trades["signed_quantity"] = trades["quantity"].where(
        trades["side"].eq("BUY"), -trades["quantity"]
    )

    opening = (
        trades[trades["market_date"] < calendar.min()]
        .groupby("ticker")["signed_quantity"]
        .sum()
        .reindex(tickers, fill_value=0.0)
    )
    changes = (
        trades[
            trades["market_date"].between(calendar.min(), calendar.max())
        ]
        .groupby(["market_date", "ticker"])["signed_quantity"]
        .sum()
        .unstack(fill_value=0.0)
        .reindex(index=calendar, columns=tickers, fill_value=0.0)
    )
    return changes.cumsum().add(opening, axis="columns")


def _daily_cash_balance(wallet, calendar):
    """Forward-fill the broker's reported end-of-day trade cash balance."""
    wallet = wallet.copy()
    wallet["calendar_date"] = _normalized_dates(wallet["date"])

    # INDmoney exports wallet rows newest-first, including within the same day.
    # The first row for a date therefore contains that day's ending balance.
    balances = (
        wallet.drop_duplicates("calendar_date", keep="first")
        .set_index("calendar_date")["updated_trade_balance_usd"]
        .sort_index()
    )
    combined_index = balances.index.union(calendar)
    return (
        balances.reindex(combined_index)
        .sort_index()
        .ffill()
        .reindex(calendar)
        .fillna(0.0)
    )


def _daily_external_flows(wallet, calendar):
    """Return deposits as positive account flows and withdrawals as negative."""
    wallet = wallet.copy()
    wallet["calendar_date"] = _normalized_dates(wallet["date"])
    external = wallet[wallet["activity_type"].isin(["DEPOSIT", "WITHDRAWAL"])].copy()
    amounts = external["amount_usd"].abs()
    external["external_flow_usd"] = amounts.where(
        external["activity_type"].eq("DEPOSIT"), -amounts
    )
    return (
        external.groupby("calendar_date")["external_flow_usd"]
        .sum()
        .reindex(calendar, fill_value=0.0)
    )


def calculate_portfolio_twr(
    trades,
    wallet,
    close_prices,
    period_start,
    period_end,
):
    """Calculate a daily-linked TWR using broker cash and reconstructed positions.

    External deposits and withdrawals are removed from daily performance. Trades,
    dividends, taxes and brokerage remain inside account value and therefore affect
    the return exactly as they affected the investor's account.
    """
    period_start = pd.Timestamp(period_start).normalize()
    period_end = pd.Timestamp(period_end).normalize()
    calendar = pd.date_range(period_start, period_end, freq="D")
    tickers = sorted(trades["ticker"].dropna().unique())

    positions = _daily_positions(trades, calendar, tickers)
    prices = _align_prices(close_prices, calendar, tickers)
    required_prices = positions.abs().gt(VALUE_TOLERANCE)
    missing_prices = prices.isna() & required_prices
    if missing_prices.any().any():
        missing_tickers = missing_prices.any()[missing_prices.any()].index.tolist()
        raise RuntimeError(
            "Historical Yahoo prices are missing while these positions were held: "
            + ", ".join(missing_tickers)
        )

    securities = (positions * prices.fillna(0.0)).sum(axis="columns")
    cash = _daily_cash_balance(wallet, calendar)
    external_flows = _daily_external_flows(wallet, calendar)
    account_value = securities + cash

    active_dates = account_value[account_value.abs().gt(VALUE_TOLERANCE)].index
    if active_dates.empty:
        return {
            "return": None,
            "start_date": None,
            "end_date": period_end,
            "daily": pd.DataFrame(),
        }

    first_active_date = active_dates.min()
    daily = pd.DataFrame(
        {
            "account_value_usd": account_value,
            "securities_value_usd": securities,
            "cash_balance_usd": cash,
            "external_flow_usd": external_flows,
        }
    ).loc[first_active_date:]

    previous_value = daily["account_value_usd"].shift(1)
    valid_denominator = previous_value.abs().gt(VALUE_TOLERANCE)
    daily["daily_return"] = 0.0
    daily.loc[valid_denominator, "daily_return"] = (
        (
            daily.loc[valid_denominator, "account_value_usd"]
            - daily.loc[valid_denominator, "external_flow_usd"]
        )
        / previous_value[valid_denominator]
        - 1.0
    )
    daily["cumulative_return"] = (1.0 + daily["daily_return"]).cumprod() - 1.0

    return {
        "return": float(daily["cumulative_return"].iloc[-1]),
        "start_date": first_active_date,
        "end_date": period_end,
        "daily": daily,
    }


def calculate_same_period_benchmark_twr(benchmark_prices, portfolio_daily):
    """Link QQQ returns only for days when the portfolio had capital at work."""
    if portfolio_daily.empty:
        return None

    prices = benchmark_prices.copy()
    prices.index = pd.to_datetime(prices.index).tz_localize(None).normalize()
    prices = prices[~prices.index.duplicated(keep="last")].sort_index()
    calendar = portfolio_daily.index
    combined_index = prices.index.union(calendar)
    aligned = prices.reindex(combined_index).sort_index().ffill().reindex(calendar)
    if aligned.isna().any():
        raise RuntimeError("Historical QQQ prices do not cover the portfolio period")

    qqq_return = aligned.pct_change(fill_method=None).fillna(0.0)
    portfolio_was_funded = (
        portfolio_daily["account_value_usd"]
        .shift(1)
        .abs()
        .gt(VALUE_TOLERANCE)
    )
    comparable_return = qqq_return.where(portfolio_was_funded, 0.0)
    return float((1.0 + comparable_return).prod() - 1.0)


def _price_on_or_after(prices, date):
    """Return the first available total-return price on or after a date."""
    available = prices[prices.index >= pd.Timestamp(date).normalize()]
    if available.empty:
        raise RuntimeError(f"QQQ has no price on or after {pd.Timestamp(date).date()}")
    return float(available.iloc[0])


def _price_on_or_before(prices, date):
    """Return the final available total-return price on or before a date."""
    available = prices[prices.index <= pd.Timestamp(date).normalize()]
    if available.empty:
        raise RuntimeError(f"QQQ has no price on or before {pd.Timestamp(date).date()}")
    return float(available.iloc[-1])


def build_cash_flow_matched_benchmark(portfolio_cash_flows, benchmark_prices):
    """Invest the portfolio's exact external cash flows into synthetic QQQ units."""
    prices = benchmark_prices.copy()
    prices.index = pd.to_datetime(prices.index).tz_localize(None).normalize()
    prices = prices[~prices.index.duplicated(keep="last")].sort_index()

    cash_flows = portfolio_cash_flows.copy()
    cash_flows["date"] = _normalized_dates(cash_flows["date"])
    terminal_rows = cash_flows[
        cash_flows["description"].str.startswith("Ending", na=False)
    ]
    if terminal_rows.empty:
        raise ValueError("Portfolio cash flows are missing an ending value")

    terminal_date = terminal_rows["date"].max()
    investor_flows = cash_flows[~cash_flows.index.isin(terminal_rows.index)].copy()
    units = 0.0
    for row in investor_flows.sort_values("date", kind="stable").itertuples(
        index=False
    ):
        if row.description == "Opening portfolio value":
            price = _price_on_or_before(prices, row.date)
        else:
            price = _price_on_or_after(prices, row.date)

        if row.cash_flow_usd < 0:
            units += abs(row.cash_flow_usd) / price
        else:
            units -= abs(row.cash_flow_usd) / price

    terminal_price = _price_on_or_before(prices, terminal_date)
    terminal_value = units * terminal_price
    benchmark_cash_flows = pd.concat(
        [
            investor_flows[["date", "description", "cash_flow_usd"]],
            pd.DataFrame(
                [
                    {
                        "date": terminal_date,
                        "description": "Ending cash-flow-matched QQQ value",
                        "cash_flow_usd": terminal_value,
                    }
                ]
            ),
        ],
        ignore_index=True,
    ).sort_values("date", kind="stable")

    return {
        "xirr": calculate_xirr(benchmark_cash_flows),
        "terminal_value_usd": terminal_value,
        "cash_flows": benchmark_cash_flows.reset_index(drop=True),
    }


def _display_money_weighted_return(xirr, cash_flows):
    """De-annualize a sub-year XIRR and retain annualized XIRR after one year."""
    if xirr is None:
        return {"return": None, "annualized": False, "days": 0}

    negative_flows = cash_flows[cash_flows["cash_flow_usd"] < 0]
    if negative_flows.empty:
        return {"return": None, "annualized": False, "days": 0}

    start_date = pd.Timestamp(negative_flows["date"].min()).normalize()
    end_date = pd.Timestamp(cash_flows["date"].max()).normalize()
    days = max((end_date - start_date).days, 0)
    if days < 365:
        period_return = (1.0 + xirr) ** (days / 365.0) - 1.0
        return {"return": period_return, "annualized": False, "days": days}
    return {"return": xirr, "annualized": True, "days": days}


def build_period_comparison(
    label,
    result,
    trades,
    wallet,
    close_prices,
    benchmark_prices,
    period_start,
    period_end,
):
    """Build professional and investor comparisons for one reporting period."""
    portfolio_twr = calculate_portfolio_twr(
        trades,
        wallet,
        close_prices,
        period_start,
        period_end,
    )
    benchmark_twr = calculate_same_period_benchmark_twr(
        benchmark_prices,
        portfolio_twr["daily"],
    )
    benchmark_mwr = build_cash_flow_matched_benchmark(
        result["cash_flows"],
        benchmark_prices,
    )
    portfolio_display_mwr = _display_money_weighted_return(
        result["metrics"]["xirr"], result["cash_flows"]
    )
    benchmark_display_mwr = _display_money_weighted_return(
        benchmark_mwr["xirr"], benchmark_mwr["cash_flows"]
    )

    portfolio_mwr = portfolio_display_mwr["return"]
    qqq_mwr = benchmark_display_mwr["return"]
    return {
        "label": label,
        "start_date": portfolio_twr["start_date"],
        "end_date": portfolio_twr["end_date"],
        "portfolio_twr": portfolio_twr["return"],
        "benchmark_twr": benchmark_twr,
        "twr_excess": portfolio_twr["return"] - benchmark_twr,
        "portfolio_mwr": portfolio_mwr,
        "benchmark_mwr": qqq_mwr,
        "mwr_excess": portfolio_mwr - qqq_mwr,
        "mwr_is_annualized": portfolio_display_mwr["annualized"],
        "portfolio_xirr": result["metrics"]["xirr"],
        "benchmark_xirr": benchmark_mwr["xirr"],
        "benchmark_terminal_value_usd": benchmark_mwr["terminal_value_usd"],
        "benchmark_cash_flows": benchmark_mwr["cash_flows"],
        "daily": portfolio_twr["daily"],
    }


def build_benchmark_comparison(
    trades,
    wallet,
    close_prices,
    benchmark_prices,
    current_year_result,
    all_time_result,
    current_year,
    as_of,
):
    """Build current-year and all-time QQQ comparison results."""
    all_time_start = min(
        pd.Timestamp(wallet["date"].min()).normalize(),
        _new_york_trade_dates(trades["execution_time"]).min(),
    )
    return {
        "ticker": "QQQ",
        "current_year": build_period_comparison(
            str(current_year),
            current_year_result,
            trades,
            wallet,
            close_prices,
            benchmark_prices,
            pd.Timestamp(year=current_year, month=1, day=1),
            as_of,
        ),
        "all_time": build_period_comparison(
            "All Time",
            all_time_result,
            trades,
            wallet,
            close_prices,
            benchmark_prices,
            all_time_start,
            as_of,
        ),
    }
