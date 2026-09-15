# INDmoney report download and folder guide

Use this checklist when refreshing the input data. Download each report and place it in the exact folder shown below; the generator discovers the files automatically.

## Report locations

| Report | Where to download it | Destination folder | Usage |
|---|---|---|---|
| Holdings report | [INDmoney US taxation documents](https://www.indmoney.com/widget/page?page=usDocDownloadPage&documentCode=HOLDINGS_REPORT&category=REPORTS) | `reports_directory/holding_reports/` | Required; provides current quantities and broker cost basis |
| Trade / order report | [INDmoney order report download](https://www.indmoney.com/widget/page?page=usDocDownloadPage&documentCode=ORDER_REPORT&category=REPORTS), or the INDmoney mobile app | `reports_directory/trade_reports/` | Required; provides buys, sells, quantities, prices, and brokerage |
| Monthly account statements | [INDmoney US account statements](https://www.indmoney.com/investments/us-stocks/related-documents/account-statements) | `reports_directory/monthly_account_statements/` | Optional validation history; download the available PDF statements |
| Deposit and withdrawal report | INDmoney app → Buying Power → question mark (top left) → US Stock Reports | `reports_directory/deposit_and_withdrawl_reports/` | Required for historical USD/INR funding rates and XIRR cash flows |
| Running balance / wallet statement | INDmoney app → Buying Power → question mark (top left) → US Stock Reports | `reports_directory/running_balance_reports/` | Required for dividends, withholding tax, and current cash balance |
| Brokerage report | INDmoney app → Buying Power → question mark (top left) → US Stock Reports | `reports_directory/brokerage_reports/` | Optional reconciliation against trade-report brokerage |

The folder name `deposit_and_withdrawl_reports` intentionally retains the existing spelling. Do not rename it unless the matching path in the code is also changed.

## Expected structure

```text
reports_directory/
├── brokerage_reports/
│   └── IND-BROKERAGE-REPORT....xls
├── deposit_and_withdrawl_reports/
│   └── financial-year reports....xlsx
├── holding_reports/
│   └── IND-HOLDINGS-REPORT....xls
├── monthly_account_statements/
│   └── YYYY-MM.pdf
├── running_balance_reports/
│   └── wallet statement....xlsx
└── trade_reports/
    └── IND-ORDER-REPORT....xls
```

## Refresh checklist

1. Download a holdings report with the latest available date.
2. Download the trade/order report from the beginning of the account through the same date.
3. Download the running balance report for the full available account period.
4. Download every applicable financial-year deposit and withdrawal report. These contain the actual bank conversion rates used for INR performance.
5. Add newly available monthly account-statement PDFs for reference.
6. Close any previously generated Excel workbook.
7. Run:

```powershell
uv sync
uv run python generate_report.py
```

If several exports exist in a required single-report folder, the generator selects the most recently modified file. Financial-year deposit/withdrawal files and monthly PDFs are combined across all available files.
