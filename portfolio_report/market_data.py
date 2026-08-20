"""Retrieve no-key Yahoo Finance prices through yfinance."""

import logging
from datetime import timedelta

import pandas as pd
import yfinance as yf

USD_INR_TICKER = "INR=X"
LOGGER = logging.getLogger(__name__)


def _close_series(download, ticker):
    """Extract a ticker's Close series from single- or multi-ticker results."""
    if download.empty:
        return pd.Series(dtype="float64")

    if isinstance(download.columns, pd.MultiIndex):
        first_level = download.columns.get_level_values(0)
        if ticker in first_level:
            return download[ticker]["Close"].dropna()
        second_level = download.columns.get_level_values(1)
        if ticker in second_level:
            return download["Close"][ticker].dropna()

    if "Close" in download.columns:
        close = download["Close"]
        if isinstance(close, pd.DataFrame):
            close = close.iloc[:, 0]
        return close.dropna()
    return pd.Series(dtype="float64")


def _download_latest(symbols, interval):
    """Download latest prices in one non-threaded request for reliability."""
    return yf.download(
        symbols,
        period="5d",
        interval=interval,
        auto_adjust=False,
        progress=False,
        threads=False,
        group_by="ticker",
        timeout=20,
    )


def _individual_latest_price(ticker):
    """Retry a missing symbol individually, first intraday and then daily."""
    for interval in ("1m", "1d"):
        try:
            download = _download_latest([ticker], interval)
            closes = _close_series(download, ticker)
            if not closes.empty:
                return closes
        # yfinance can surface transport, parsing, or SQLite cache exceptions.
        except Exception as error:  # noqa: BLE001
            LOGGER.warning("Yahoo retry failed for %s: %s", ticker, error)
    return pd.Series(dtype="float64")


def fetch_latest_market_data(tickers):
    """Fetch latest stock prices and USD/INR without an API key."""
    symbols = sorted(set(tickers) | {USD_INR_TICKER})
    try:
        download = _download_latest(symbols, "1m")
    # Fall back below regardless of which yfinance dependency failed.
    except Exception as error:  # noqa: BLE001
        LOGGER.warning("Intraday Yahoo download failed: %s", error)
        download = pd.DataFrame()

    rows = []
    missing = []
    for ticker in symbols:
        closes = _close_series(download, ticker)
        if closes.empty:
            closes = _individual_latest_price(ticker)
        if closes.empty:
            missing.append(ticker)
            continue

        quote_time = pd.Timestamp(closes.index[-1])
        rows.append(
            {
                "ticker": ticker,
                "last_price": float(closes.iloc[-1]),
                "quote_time": quote_time.strftime("%Y-%m-%d %H:%M %Z").strip(),
            }
        )

    if missing:
        raise RuntimeError(
            "Yahoo Finance did not return prices for: {}".format(", ".join(missing))
        )

    quotes = pd.DataFrame(rows)
    fx_row = quotes[quotes["ticker"].eq(USD_INR_TICKER)].iloc[0]
    prices = quotes[~quotes["ticker"].eq(USD_INR_TICKER)].copy()
    prices = prices.rename(columns={"last_price": "current_price_usd"})
    fx = {
        "rate": float(fx_row["last_price"]),
        "quote_time": fx_row["quote_time"],
        "ticker": USD_INR_TICKER,
    }
    return prices.reset_index(drop=True), fx


def fetch_period_start_prices(tickers, period_start):
    """Fetch the adjusted close immediately before a reporting period starts."""
    tickers = sorted(set(tickers))
    if not tickers:
        return {}

    start = pd.Timestamp(period_start) - timedelta(days=14)
    end = pd.Timestamp(period_start)
    download = yf.download(
        tickers,
        start=start.strftime("%Y-%m-%d"),
        end=end.strftime("%Y-%m-%d"),
        interval="1d",
        auto_adjust=True,
        progress=False,
        threads=False,
        group_by="ticker",
        timeout=20,
    )

    prices = {}
    missing = []
    for ticker in tickers:
        closes = _close_series(download, ticker)
        if closes.empty:
            missing.append(ticker)
        else:
            prices[ticker] = float(closes.iloc[-1])

    if missing:
        raise RuntimeError(
            "Could not find a pre-period Yahoo close for: {}".format(", ".join(missing))
        )
    return prices
