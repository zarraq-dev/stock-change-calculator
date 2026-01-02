# Stock Change Calculator

A Python command-line tool that calculates the percentage price movement of stocks between two dates.

## Features

- Calculate stock price changes between any two dates
- Input via CSV file or command-line arguments
- Automatic ticker resolution from company names or ISINs
- Outputs to both terminal and CSV file
- Handles weekends and holidays by adjusting to the next trading day
- Supports multiple stock exchanges (NYSE, NASDAQ, LSE, etc.)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/stock-change-calculator.git
cd stock-change-calculator
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Option 1: CSV File Input

```bash
python src/stock_calculator.py --file "path/to/input.csv"
```

### Option 2: Command-Line Input

```bash
python src/stock_calculator.py --stocks "Apple,Microsoft,Shell" --start "01-Jan-25" --end "01-Apr-25"
```

### Optional: Specify Output Directory

```bash
python src/stock_calculator.py --file "input.csv" --output "path/to/output/"
```

If no output directory is specified, the CSV output file is saved in the same location as the script.

## Input File Format

The CSV input file must follow this structure:

```
Start Date,01-Oct-25,End Date,01-Jan-26
,,,
Stocks,,,
Name,Ticker,ISIN,
National Grid PLC,,,
Microsoft Corp,MSFT,,
Apple Inc,,US0378331005,
```

### Structure Breakdown

| Row | Content |
|-----|---------|
| 1 | Start Date, `<date>`, End Date, `<date>` |
| 2 | Blank row |
| 3 | "Stocks" header |
| 4 | Column headers: Name, Ticker, ISIN |
| 5+ | Stock data (one per row) |

### Notes on Input

- **Date format:** dd-mmm-yy (e.g., 01-Jan-25)
- **Ticker and ISIN are optional.** If not provided, the program will look them up using the stock name.
- If Ticker is provided, it will be used directly for the lookup.
- If only ISIN is provided, it will be converted to a Ticker.

## Output Format

### Terminal Output

```
Start Date: 01-Jan-25, End Date: 01-Apr-25

Note: Start date adjusted to 02-Jan-25 (next trading day) for: Apple Inc, Microsoft Corp

Stock Name,Ticker,ISIN,Start Price,End Price,Percentage,Currency
Apple Inc,AAPL,US0378331005,150.25,175.50,16.8,USD
Microsoft Corp,MSFT,US5949181045,380.00,420.75,10.7,USD
National Grid PLC,NG.L,GB00BDR05C01,10.50,11.20,6.7,GBP
```

### CSV Output

The same format is saved to `stock_changes_output.csv` in the specified output directory.

If a file with that name already exists, the program creates a versioned file:
- `stock_changes_output_v1.csv`
- `stock_changes_output_v2.csv`
- etc.

## Output Columns

| Column | Description |
|--------|-------------|
| Stock Name | Full company name |
| Ticker | Stock ticker symbol |
| ISIN | International Securities Identification Number |
| Start Price | Closing price on start date |
| End Price | Closing price on end date |
| Percentage | Percentage change (number only, no % symbol) |
| Currency | Three-letter currency code (USD, GBP, EUR, etc.) |

## Error Handling

| Scenario | Behaviour |
|----------|-----------|
| Stock not found | Row displays "Stock details not found" |
| Stock delisted | Row displays "Delisted" |
| Date is weekend/holiday | Adjusts to next trading day, note displayed |
| API unreachable | Displays error message and exits |
| Invalid input file | Displays specific error message |

## Dependencies

- Python 3.8+
- yfinance - Third-party library for fetching stock data from Yahoo Finance. Handles public holidays implicitly by returning data for the next available trading day.
- requests - HTTP library for OpenFIGI API calls
- pandas - Data manipulation library (used internally by yfinance)

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

## Author

Zarraq Ahmed
