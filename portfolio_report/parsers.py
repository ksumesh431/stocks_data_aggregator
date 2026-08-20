"""Parse each source report into consistently named pandas tables."""

from pathlib import Path

import pandas as pd


def _read_table_at_marker(path, marker):
    """Find a header label in column A and read the table below that row."""
    raw = pd.read_excel(path, header=None)
    first_column = raw.iloc[:, 0].astype(str).str.strip()
    matches = raw.index[first_column.eq(marker)].tolist()
    if not matches:
        raise ValueError(f"Could not find header '{marker}' in {path}")
    return pd.read_excel(path, header=matches[0])


def _parse_datetime(series, preferred_format=None):
    """Parse dates with a strict known format and a safe general fallback."""
    if preferred_format:
        parsed = pd.to_datetime(series, format=preferred_format, errors="coerce")
        missing = parsed.isna() & series.notna()
        if missing.any():
            parsed.loc[missing] = pd.to_datetime(
                series.loc[missing], format="mixed", errors="coerce"
            )
        return parsed
    return pd.to_datetime(series, format="mixed", errors="coerce")


def _number(series):
    """Convert report values such as '-' and blanks to numeric values."""
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def _metadata_value(path, label):
    """Read a label/value pair from the report preamble."""
    raw = pd.read_excel(path, header=None)
    matches = raw.index[raw.iloc[:, 0].astype(str).str.strip().eq(label)].tolist()
    if not matches or raw.shape[1] < 2:
        return None
    return raw.iloc[matches[0], 1]


def parse_trade_report(path):
    """Parse the INDmoney order report into one row per completed trade."""
    table = _read_table_at_marker(path, "Stock Name")
    table = table[table["Transaction Type"].isin(["BUY", "SELL"])].copy()

    result = pd.DataFrame(
        {
            "execution_time": _parse_datetime(
                table["Order Execution Time"], "%d %b %Y, %I:%M %p"
            ),
            "stock_name": table["Stock Name"].astype(str).str.strip(),
            "ticker": table["Stock Symbol"].astype(str).str.strip().str.upper(),
            "side": table["Transaction Type"].astype(str).str.strip().str.upper(),
            "order_type": table["Order Type"].astype(str).str.strip().str.lower(),
            "quantity": _number(table["Quantity"]),
            "price_usd": _number(table["Price ($)"]),
            "gross_amount_usd": _number(table["Order Amount ($)"]),
            "brokerage_usd": _number(table["Brokerage ($)"]),
            "broker_reference": table["Broker Reference Id"].astype(str).str.strip(),
        }
    )
    result = result[result["execution_time"].notna()].copy()
    return result.sort_values("execution_time", kind="stable").reset_index(drop=True)


def parse_holding_report(path):
    """Parse current holdings and retain the broker-reported cost basis."""
    table = _read_table_at_marker(path, "Stock Symbol")
    valid_rows = pd.to_numeric(table["Quantity"], errors="coerce").notna()
    table = table[valid_rows].copy()

    holdings = pd.DataFrame(
        {
            "ticker": table["Stock Symbol"].astype(str).str.strip().str.upper(),
            "holding_since": _parse_datetime(
                table["Holding Since"], "%d %b %Y, %I:%M %p"
            ),
            "quantity": _number(table["Quantity"]),
            "average_price_usd": _number(table["Avg. Price ($)"]),
            # Despite its label, this column equals quantity x average price.
            "cost_basis_usd": _number(table["Total Value ($)"]),
        }
    )
    as_of = pd.to_datetime(_metadata_value(path, "Holdings as on"), errors="coerce")
    return {"data": holdings.reset_index(drop=True), "as_of": as_of}


def parse_wallet_report(path):
    """Parse dividends, taxes, funding, trades and running cash balances."""
    table = _read_table_at_marker(path, "Date")
    dates = _parse_datetime(table["Date"])
    table = table[dates.notna()].copy()
    dates = dates[dates.notna()]

    result = pd.DataFrame(
        {
            "date": dates,
            "activity_type": table["Type"].astype(str).str.strip().str.upper(),
            "description": table["Description"].astype(str).str.strip(),
            "money_movement": table["Money Movement"].astype(str).str.strip(),
            "amount_usd": _number(table["($) Amount"]),
            "ticker": table["Ticker"].fillna("").astype(str).str.strip().str.upper(),
            "quantity": _number(table["Quantity"]),
            "price_usd": _number(table["($) Price"]),
            "commission_usd": _number(table["($) Commission"]),
            "updated_trade_balance_usd": _number(table["($) Updated Trade Balance"]),
            "updated_withdrawable_balance_usd": _number(
                table["($) Updated Withdrawable Balance"]
            ),
        }
    )
    return result.reset_index(drop=True)


def parse_brokerage_report(path):
    """Parse the separate brokerage export used as a reconciliation check."""
    if path is None:
        return pd.DataFrame()
    table = _read_table_at_marker(path, "Order Placed Time")
    valid_rows = table["Transaction Type"].isin(["BUY", "SELL"])
    table = table[valid_rows].copy()
    return pd.DataFrame(
        {
            "execution_time": _parse_datetime(
                table["Order Execution Time"], "%d %b %Y, %I:%M %p"
            ),
            "ticker": table["Stock Symbol"].astype(str).str.strip().str.upper(),
            "side": table["Transaction Type"].astype(str).str.strip().str.upper(),
            "gross_amount_usd": _number(table["Order Value ($)"]),
            "brokerage_usd": _number(table["Brokerage Charged ($)"]),
        }
    ).reset_index(drop=True)


def parse_fund_movement_reports(paths):
    """Combine all financial-year deposit and withdrawal exports."""
    frames = []
    for path in paths:
        table = _read_table_at_marker(path, "Transaction Date")
        dates = _parse_datetime(table["Transaction Date"])
        table = table[dates.notna()].copy()
        dates = dates[dates.notna()]
        if table.empty:
            continue

        frames.append(
            pd.DataFrame(
                {
                    "date": dates,
                    "bank": table["Bank Used"].astype(str).str.strip(),
                    "movement": table["Money Movement"].astype(str).str.strip(),
                    "inr_amount": _number(table["INR Amount"]),
                    "gst_inr": _number(table["GST Charged"]),
                    "tcs_inr": _number(table["TCS Charged(Indicative)"]),
                    "bank_fee_inr": _number(table["Bank Fixed Fee"]),
                    "exchange_rate": _number(table["Exchange Rate"]),
                    "usd_amount": _number(table["USD Amount"]),
                    "source_file": Path(path).name,
                }
            )
        )

    if not frames:
        return pd.DataFrame(
            columns=[
                "date",
                "bank",
                "movement",
                "inr_amount",
                "gst_inr",
                "tcs_inr",
                "bank_fee_inr",
                "exchange_rate",
                "usd_amount",
                "source_file",
            ]
        )
    return (
        pd.concat(frames, ignore_index=True).sort_values("date").reset_index(drop=True)
    )
