# Portfolio performance workbook

This project reads the reports in `reports_directory`, fetches no-key Yahoo Finance market data, and creates one Excel workbook with matching current-year and all-time dashboards in both USD and INR.

Before running the generator, see the [INDmoney report download and folder guide](docs/REPORT_DOWNLOAD_GUIDE.md) for direct download links, mobile-app sources, and the required directory layout.

## Run it

```powershell
uv sync
uv run python generate_report.py
```

The default output is `output/portfolio_performance.xlsx`. Optional arguments:

```powershell
uv run python generate_report.py --reports-dir reports_directory --output output/my_report.xlsx --year 2026
```

If the destination workbook is open in Excel or otherwise locked, the command exits before loading reports or calling Yahoo. It explains the likely cause and prints a ready-to-run `--output` command with an alternative filename instead of displaying a Python traceback.

`uv sync` creates the project-local `.venv` and installs the exact versions recorded in `uv.lock`. The Python version is pinned in `.python-version`; no global packages are used by these commands.

No API key is required. Stock prices and `INR=X` (USD/INR) are downloaded through [yfinance](https://ranaroussi.github.io/yfinance/reference/api/yfinance.download.html).

## Function flow

1. `generate_report.py` handles command-line arguments only.
2. `report_discovery.py` finds the latest required exports and inventories every source file.
3. `parsers.py` converts each spreadsheet layout into normalized tables.
4. `market_data.py` retrieves the latest available intraday prices, historical closes, dividend-adjusted QQQ prices, USD/INR, and any required year-opening prices from Yahoo.
5. `calculations.py` replays trades, reconciles to broker holdings, and calculates both reporting periods.
6. `returns.py` builds external investor cash flows and calculates XIRR.
7. `benchmark.py` reconstructs daily account values and creates the fair portfolio-versus-QQQ comparisons.
8. `excel_writer.py` creates the styled dashboards, the visual QQQ comparison, and audit sheets.
9. `workflow.py` connects those modules in that order.

There are intentionally no Python type annotations, per the project requirement. Functions still use descriptive names, docstrings, narrow responsibilities, and normalized column names.

## How each report is used

| Report | Role |
|---|---|
| Trade report | Executions, quantities, trade amounts, and brokerage |
| Holdings report | Authoritative current quantity and remaining broker cost basis |
| Running balance | Dividends, withholding tax, and latest cash balance |
| Brokerage report | Reconciliation against trade-report commissions |
| Deposit/withdrawal reports | External cash flows for XIRR; deposits are never counted as profit |
| Monthly PDF statements | Validation snapshots; not parsed because the newer structured exports contain the required data |

The May 2026 broker account migration changes how commissions appear in the wallet statement. For that reason, brokerage always comes from the trade report, which remains consistent across both accounts.

## Dashboard formulas

- **XIRR (primary highlighted percentage):** annualized, money-weighted return using the exact dates of external deposits and withdrawals plus the ending value of securities and cash.
- **Invested:** remaining cost basis of open holdings.
- **Current:** open quantity multiplied by the latest available Yahoo price.
- **Realized P&L:** gross sale proceeds minus disposed cost basis, before brokerage.
- **Unrealized P&L:** current market value minus remaining cost basis.
- **Net P&L:** realized + unrealized + dividends/interest - withholding tax - brokerage.
- **INR invested:** USD cost basis multiplied by the USD-weighted average of the actual bank conversion rates in the deposit reports.
- **INR current:** current USD market value multiplied by the latest `INR=X` quote.
- **INR unrealized P&L:** current INR value minus historical INR cost basis, so the return includes both stock performance and the USD/INR movement.

For current-year reporting, positions already held on January 1 are marked to the last Yahoo close before the year began. This prevents gains from an earlier year from leaking into the current-year result. The funding reports do not identify which remittance dollar funded each stock lot, so historical FX is allocated at portfolio level. This is suitable for performance tracking but is not a tax calculation.

XIRR is the headline percentage because this is a personal portfolio with deposits occurring on different dates. The net P&L percentage remains visible as a useful cost-basis comparison, but it does not account for how long each contribution was invested. The `XIRR Cash Flows` sheet makes every input auditable. For the current-year calculation, any opening securities and cash are treated as a synthetic investment on January 1.

## QQQ comparison sheet

The `QQQ Comparison` sheet gives a current-year and an all-time answer to two different questions without requiring you to interpret an accounting table:

- **Did the strategy beat QQQ?** Portfolio time-weighted return (TWR) is compared with dividend-adjusted QQQ over the same funded days. TWR removes the effect of deposit and withdrawal timing, so this is the standard comparison for judging the investment choices.
- **Did the actual money beat QQQ?** Portfolio money-weighted return is compared with a synthetic QQQ investment that receives the exact same deposits and withdrawals on matching dates. This measures the combined effect of choices and contribution timing.

Each period includes highlighted percentage-point differences, a plain-English verdict, and a portfolio-versus-QQQ chart. Current-year money-weighted figures are displayed as period returns when less than one year has elapsed; all-time figures of at least one year remain annualized. Net P&L is intentionally absent from the benchmark chart because profit dollars divided by cost basis is not a standardized benchmark return.

## Tests

```powershell
uv run python -m unittest discover -s tests -v
uv run ruff check .
```
