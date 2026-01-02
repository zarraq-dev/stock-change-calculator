"""
Stock Change Calculator

A CLI tool that calculates the percentage price movement of stocks between two dates.

USAGE
    python src/stock_calculator.py --file "input.csv"
    python src/stock_calculator.py --stocks "Apple,Microsoft" --start "01-Jan-25" --end "01-Apr-25"

ARGUMENTS
    --file (optional)
        Path to CSV input file containing stock list and date range.

    --stocks (optional)
        Comma-separated list of stock names.

    --start (required with --stocks)
        Start date in dd-mmm-yy format (e.g., 01-Jan-25).

    --end (required with --stocks)
        End date in dd-mmm-yy format (e.g., 01-Apr-25).

    --output (optional)
        Directory path for output CSV file. Defaults to script location.

DEPENDENCIES
    yfinance - Third-party library for fetching stock data from Yahoo Finance.
               Handles public holidays implicitly by returning data for the next
               available trading day when a holiday is requested.
    requests - HTTP library for OpenFIGI API calls.
    pandas   - Data manipulation library (used internally by yfinance).
"""

import argparse
import csv
import os
import re
import sys
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple, Union

import requests
import yfinance

# Exception Classes

class CsvParsingError(Exception):
    """Raised when CSV file parsing fails."""
    pass


class CliArgumentError(Exception):
    """Raised when CLI arguments are invalid."""
    pass


class ApiError(Exception):
    """Raised when external API is unreachable."""
    pass


class StockDelistedError(Exception):
    """Raised when stock is delisted and has no data."""
    pass


# Constants

DATE_FORMAT: str = "%d-%b-%y"  # Format for date parsing (e.g., 01-Jan-25)
DATE_PATTERN: str = r"^\d{2}-[A-Za-z]{3}-\d{2}$"  # Regex pattern for date validation
OPENFIGI_URL: str = "https://api.openfigi.com/v3/mapping"  # OpenFIGI API endpoint
DEFAULT_OUTPUT_FILENAME: str = "stock_changes_output.csv"  # Default output file name

# Exchange code to yfinance suffix mapping
DICT_EXCHANGE_SUFFIX: Dict[str, str] = {
    "US": "",  # US exchanges (NYSE, NASDAQ)
    "UN": "",  # NYSE
    "UQ": "",  # NASDAQ
    "LN": ".L",  # London Stock Exchange
    "GY": ".DE",  # Germany (Xetra)
    "FP": ".PA",  # France (Euronext Paris)
    "JP": ".T",  # Japan (Tokyo)
    "HK": ".HK",  # Hong Kong
    "AU": ".AX",  # Australia
    "CN": ".TO",  # Canada (Toronto)
}


# CSV Parsing Functions

def parse_csv_file(s_input_file_path: str) -> Dict[str, Any]:
    """
    Parse CSV input file and extract dates and stock list.

    Args:
        s_input_file_path: Path to the CSV file to parse.

    Returns:
        Dictionary containing:
            - s_start_date: Start date string
            - s_end_date: End date string
            - list_dict_stocks: List of stock dictionaries with name, ticker, isin

    Raises:
        CsvParsingError: If file structure is invalid or dates are missing/malformed.
    """
    list_list_rows: List[List[str]] = []  # Raw rows from CSV file

    # Read CSV file
    with open(s_input_file_path, 'r', encoding='utf-8') as file_input:
        reader_csv = csv.reader(file_input)
        # Loop through each row in the CSV file
        for list_row in reader_csv:
            list_list_rows.append(list_row)

    # Validate minimum row count
    if len(list_list_rows) < 5:
        raise CsvParsingError("Invalid CSV structure. Missing required rows including date row.")

    # Parse dates from row 1
    list_first_row: List[str] = list_list_rows[0]  # First row containing dates

    if len(list_first_row) < 4:
        raise CsvParsingError("Missing Start Date or End Date in row 1.")

    s_start_date_label: str = list_first_row[0].strip()  # Should be "Start Date"
    s_start_date: str = list_first_row[1].strip()  # Start date value
    s_end_date_label: str = list_first_row[2].strip()  # Should be "End Date"
    s_end_date: str = list_first_row[3].strip()  # End date value

    # Validate date labels
    if s_start_date_label.lower() != "start date" or s_end_date_label.lower() != "end date":
        raise CsvParsingError("Missing Start Date or End Date labels in row 1.")

    # Validate date format
    if not re.match(DATE_PATTERN, s_start_date):
        raise CsvParsingError(f"Invalid date format for Start Date. Expected dd-mmm-yy (e.g., 01-Jan-25), got: {s_start_date}")

    if not re.match(DATE_PATTERN, s_end_date):
        raise CsvParsingError(f"Invalid date format for End Date. Expected dd-mmm-yy (e.g., 01-Jan-25), got: {s_end_date}")

    # Parse stocks from row 5 onwards (index 4+)
    list_dict_stocks: List[Dict[str, str]] = []  # List of stock dictionaries

    # Loop through each stock row starting from row 5
    for n_row_index in range(4, len(list_list_rows)):
        list_stock_row: List[str] = list_list_rows[n_row_index]  # Current stock row

        # Skip empty rows
        if not list_stock_row or not list_stock_row[0].strip():
            continue

        s_name: str = list_stock_row[0].strip() if len(list_stock_row) > 0 else ""  # Stock name
        s_ticker: str = list_stock_row[1].strip() if len(list_stock_row) > 1 else ""  # Stock ticker
        s_isin: str = list_stock_row[2].strip() if len(list_stock_row) > 2 else ""  # Stock ISIN

        dict_stock: Dict[str, str] = {"s_name": s_name, "s_ticker": s_ticker, "s_isin": s_isin}
        list_dict_stocks.append(dict_stock)

    # Validate stock list is not empty
    if len(list_dict_stocks) == 0:
        raise CsvParsingError("No stocks found in file. Stock list is empty.")

    dict_result: Dict[str, Any] = {"s_start_date": s_start_date, "s_end_date": s_end_date, "list_dict_stocks": list_dict_stocks}

    return dict_result


# CLI Argument Parsing Functions

def validate_date_format(s_date: str) -> bool:
    """
    Validate that a date string matches the expected format.

    Args:
        s_date: Date string to validate.

    Returns:
        True if valid, False otherwise.
    """
    return bool(re.match(DATE_PATTERN, s_date))


def parse_arguments(list_s_args: List[str]) -> Dict[str, Any]:
    """
    Parse command-line arguments.

    Args:
        list_s_args: List of command-line argument strings.

    Returns:
        Dictionary containing parsed arguments.

    Raises:
        CliArgumentError: If arguments are invalid or missing.
    """
    parser_args = argparse.ArgumentParser(description="Calculate stock price percentage changes.", add_help=False)

    parser_args.add_argument("--file", dest="s_input_file_path", type=str, help="Path to CSV input file")
    parser_args.add_argument("--stocks", dest="s_stocks", type=str, help="Comma-separated stock names")
    parser_args.add_argument("--start", dest="s_start_date", type=str, help="Start date (dd-mmm-yy)")
    parser_args.add_argument("--end", dest="s_end_date", type=str, help="End date (dd-mmm-yy)")
    parser_args.add_argument("--output", dest="s_output_directory_path", type=str, help="Output directory path")

    try:
        namespace_args = parser_args.parse_args(list_s_args)
    except SystemExit:
        raise CliArgumentError("Invalid command-line arguments provided.")

    s_input_file_path: Optional[str] = namespace_args.s_input_file_path  # Path to input CSV file
    s_stocks: Optional[str] = namespace_args.s_stocks  # Comma-separated stock names
    s_start_date: Optional[str] = namespace_args.s_start_date  # Start date
    s_end_date: Optional[str] = namespace_args.s_end_date  # End date
    s_output_directory_path: Optional[str] = namespace_args.s_output_directory_path  # Output directory

    # Check for missing required arguments
    if not s_input_file_path and not s_stocks:
        raise CliArgumentError("Missing required arguments. Provide --file or --stocks with --start and --end.")

    # Check mutual exclusivity
    if s_input_file_path and s_stocks:
        raise CliArgumentError("--file and --stocks are mutually exclusive. Use one or the other, not both.")

    # Check stocks requires dates
    if s_stocks:
        if not s_start_date:
            raise CliArgumentError("When using --stocks, --start date is required.")
        if not s_end_date:
            raise CliArgumentError("When using --stocks, --end date is required.")

        # Validate date formats
        if not validate_date_format(s_start_date):
            raise CliArgumentError(f"Invalid date format for --start. Expected dd-mmm-yy (e.g., 01-Jan-25), got: {s_start_date}")
        if not validate_date_format(s_end_date):
            raise CliArgumentError(f"Invalid date format for --end. Expected dd-mmm-yy (e.g., 01-Jan-25), got: {s_end_date}")

    dict_args: Dict[str, Any] = {"s_input_file_path": s_input_file_path, "s_stocks": s_stocks, "s_start_date": s_start_date, "s_end_date": s_end_date, "s_output_directory_path": s_output_directory_path}

    return dict_args


# Percentage Calculation Functions

def calculate_percentage_change(n_start_price: float, n_end_price: float) -> float:
    """
    Calculate the percentage change between two prices.

    Args:
        n_start_price: Starting price.
        n_end_price: Ending price.

    Returns:
        Percentage change as a float (e.g., 50.0 for 50% increase).
    """
    n_percentage: float = ((n_end_price - n_start_price) / n_start_price) * 100  # Percentage change calculation

    return n_percentage


# OpenFIGI Lookup Functions

def lookup_ticker_from_openfigi(s_stock_name: str = "", s_isin: str = "") -> Dict[str, Any]:
    """
    Look up stock ticker using Bloomberg OpenFIGI API.

    Args:
        s_stock_name: Company name to search. Required if s_isin not provided.
        s_isin: ISIN to search. Takes priority over s_stock_name if provided.

    Returns:
        Dictionary containing:
            - s_ticker: Stock ticker symbol (empty if not found)
            - s_exchange_code: Exchange code (empty if not found)
            - b_not_found: True if stock was not found

    Raises:
        ApiError: If API is unreachable.
        ValueError: If neither s_stock_name nor s_isin is provided.
    """
    if not s_stock_name and not s_isin:
        raise ValueError("Either s_stock_name or s_isin must be provided to lookup_ticker_from_openfigi()")

    list_dict_api_query: List[Dict[str, str]] = []  # Query payload for OpenFIGI API

    if s_isin:
        dict_query: Dict[str, str] = {"idType": "ID_ISIN", "idValue": s_isin}
        list_dict_api_query.append(dict_query)
    else:
        dict_query: Dict[str, str] = {"idType": "NAME", "idValue": s_stock_name}
        list_dict_api_query.append(dict_query)

    dict_headers: Dict[str, str] = {"Content-Type": "application/json"}  # HTTP request headers

    try:
        response_api = requests.post(OPENFIGI_URL, json=list_dict_api_query, headers=dict_headers, timeout=30)
    except Exception as e:
        raise ApiError(f"OpenFIGI API unreachable. Error: {str(e)}")

    if response_api.status_code != 200:
        raise ApiError(f"OpenFIGI API returned error status: {response_api.status_code}")

    list_dict_response: List[Dict[str, Any]] = response_api.json()  # API response data

    # Check for no match
    if not list_dict_response or "warning" in list_dict_response[0]:
        return {"s_ticker": "", "s_exchange_code": "", "b_not_found": True}

    # Extract data from response - check if "data" key exists and has items
    if "data" not in list_dict_response[0] or len(list_dict_response[0]["data"]) == 0:
        return {"s_ticker": "", "s_exchange_code": "", "b_not_found": True}

    dict_first_result: Dict[str, Any] = list_dict_response[0]["data"][0]  # First matching result

    s_ticker: str = dict_first_result.get("ticker", "")  # Ticker symbol
    s_exchange_code: str = dict_first_result.get("exchCode", "")  # Exchange code

    return {"s_ticker": s_ticker, "s_exchange_code": s_exchange_code, "b_not_found": False}


# Exchange Suffix Mapping Functions

def map_exchange_to_suffix(s_exchange_code: str) -> str:
    """
    Map exchange code to yfinance ticker suffix.

    Args:
        s_exchange_code: Exchange code from OpenFIGI (e.g., "LN", "US").

    Returns:
        Suffix string for yfinance (e.g., ".L" for London, "" for US).
    """
    s_suffix: str = DICT_EXCHANGE_SUFFIX.get(s_exchange_code, "")  # Look up suffix or default to empty

    return s_suffix


# Date Adjustment Functions

def adjust_to_trading_day(s_date: str, b_return_flag: bool = False) -> Union[str, Tuple[str, bool]]:
    """
    Adjust a date to the next trading day if it falls on a weekend.

    Note: This function only handles weekends (Saturday/Sunday). Public holidays
    are handled implicitly by the yfinance library, which returns data for the
    next available trading day when a holiday is requested.

    Args:
        s_date: Date string in dd-mmm-yy format.
        b_return_flag: If True, also return whether date was adjusted.

    Returns:
        If b_return_flag is False: Adjusted date string.
        If b_return_flag is True: Tuple of (adjusted date string, was_adjusted boolean).
    """
    dt_date: datetime = datetime.strptime(s_date, DATE_FORMAT)  # Parsed datetime object
    n_weekday: int = dt_date.weekday()  # Day of week (0=Monday, 6=Sunday)
    b_was_adjusted: bool = False  # Flag indicating if date was changed

    # Saturday (5) -> Monday (add 2 days)
    if n_weekday == 5:
        dt_date = dt_date + timedelta(days=2)
        b_was_adjusted = True
    # Sunday (6) -> Monday (add 1 day)
    elif n_weekday == 6:
        dt_date = dt_date + timedelta(days=1)
        b_was_adjusted = True

    s_adjusted_date: str = dt_date.strftime(DATE_FORMAT)  # Formatted adjusted date

    if b_return_flag:
        return s_adjusted_date, b_was_adjusted
    else:
        return s_adjusted_date


# Output File Versioning Functions

def generate_output_filename(s_output_directory_path: str) -> str:
    """
    Generate output filename with versioning if file already exists.

    Args:
        s_output_directory_path: Directory path for output file.

    Returns:
        Full path to output file with appropriate version suffix.
    """
    s_output_base_name_path: str = os.path.join(s_output_directory_path, DEFAULT_OUTPUT_FILENAME)  # Default output file path before versioning

    # Check if default file exists
    if not os.path.exists(s_output_base_name_path):
        return s_output_base_name_path

    # Find next available version number
    n_version: int = 1  # Starting version number

    while True:
        s_versioned_name: str = f"stock_changes_output_v{n_version}.csv"  # Versioned filename
        s_versioned_path: str = os.path.join(s_output_directory_path, s_versioned_name)  # Full versioned path

        if not os.path.exists(s_versioned_path):
            return s_versioned_path

        n_version += 1


# Stock Data Fetching Functions

def fetch_stock_price(s_ticker: str, s_date: str) -> float:
    """
    Fetch closing stock price for a given date using yfinance library.

    This function fetches a 5-day range starting from the requested date to handle
    cases where the exact date is a public holiday. The yfinance library (Yahoo Finance)
    automatically handles holidays by only returning data for actual trading days.
    We take the first available closing price from the returned data.

    Args:
        s_ticker: Stock ticker symbol (including exchange suffix if needed).
        s_date: Date string in dd-mmm-yy format.

    Returns:
        Closing price as float.

    Raises:
        StockDelistedError: If no price data is available (stock may be delisted).
    """
    dt_date: datetime = datetime.strptime(s_date, DATE_FORMAT)  # Parsed date
    dt_end_date: datetime = dt_date + timedelta(days=5)  # End date for price range query (covers holidays)

    s_yfinance_start_date: str = dt_date.strftime("%Y-%m-%d")  # Start date in yfinance format (YYYY-MM-DD)
    s_yfinance_end_date: str = dt_end_date.strftime("%Y-%m-%d")  # End date in yfinance format (YYYY-MM-DD)

    ticker_stock = yfinance.Ticker(s_ticker)  # yfinance Ticker object for the stock
    df_history = ticker_stock.history(start=s_yfinance_start_date, end=s_yfinance_end_date)  # Price history dataframe from Yahoo Finance

    if df_history.empty:
        raise StockDelistedError(f"Stock {s_ticker} appears to be delisted. No price data available.")

    n_close_price: float = df_history["Close"].iloc[0]  # First available closing price

    return n_close_price


def fetch_stock_currency(s_ticker: str) -> str:
    """
    Fetch currency for a stock using yfinance library.

    Args:
        s_ticker: Stock ticker symbol.

    Returns:
        Three-letter currency code (e.g., "USD", "GBP").
    """
    ticker_stock = yfinance.Ticker(s_ticker)  # yfinance Ticker object for the stock
    dict_info: Dict[str, Any] = ticker_stock.info  # Stock info dictionary from Yahoo Finance

    s_currency: str = dict_info.get("currency", "N/A")  # Currency code

    return s_currency


# Output Functions

def format_output_row(dict_stock_result: Dict[str, Any]) -> List[str]:
    """
    Format a stock result dictionary as a CSV row.

    Args:
        dict_stock_result: Dictionary containing stock calculation results.

    Returns:
        List of strings representing the CSV row.
    """
    s_name: str = dict_stock_result.get("s_name", "")  # Stock name
    s_ticker: str = dict_stock_result.get("s_ticker", "")  # Ticker symbol
    s_isin: str = dict_stock_result.get("s_isin", "")  # ISIN
    s_start_price: str = str(dict_stock_result.get("n_start_price", ""))  # Start price as string
    s_end_price: str = str(dict_stock_result.get("n_end_price", ""))  # End price as string
    s_percentage: str = str(dict_stock_result.get("n_percentage", ""))  # Percentage change as string
    s_currency: str = dict_stock_result.get("s_currency", "")  # Currency code
    s_error: str = dict_stock_result.get("s_error", "")  # Error message if any

    if s_error:
        return [s_name, s_ticker, s_isin, s_error, "", "", ""]

    return [s_name, s_ticker, s_isin, s_start_price, s_end_price, s_percentage, s_currency]


def write_output_csv(s_output_file_path: str, s_start_date: str, s_end_date: str, list_dict_results: List[Dict[str, Any]], list_s_date_adjustment_notes: List[str]) -> None:
    """
    Write results to CSV file.

    Args:
        s_output_file_path: Path to output CSV file.
        s_start_date: Start date string.
        s_end_date: End date string.
        list_dict_results: List of stock result dictionaries.
        list_s_date_adjustment_notes: List of date adjustment notes.
    """
    with open(s_output_file_path, 'w', newline='', encoding='utf-8') as file_output:
        writer_csv = csv.writer(file_output)

        # Write header row with dates
        writer_csv.writerow(["Start Date", s_start_date, "End Date", s_end_date])
        writer_csv.writerow([])  # Blank row

        # Write date adjustment notes if any. Loop through each date adjustment note
        for s_date_adjustment_note in list_s_date_adjustment_notes:
            writer_csv.writerow([s_date_adjustment_note])

        if list_s_date_adjustment_notes:
            writer_csv.writerow([])  # Blank row after notes

        # Write column headers
        writer_csv.writerow(["Stock Name", "Ticker", "ISIN", "Start Price", "End Price", "Percentage", "Currency"])

        # Write stock data rows. Loop through each stock result
        for dict_result in list_dict_results:
            list_row: List[str] = format_output_row(dict_result)  # Formatted row
            writer_csv.writerow(list_row)


def print_output_terminal(s_start_date: str, s_end_date: str, list_dict_results: List[Dict[str, Any]], list_s_date_adjustment_notes: List[str]) -> None:
    """
    Print results to terminal.

    Args:
        s_start_date: Start date string.
        s_end_date: End date string.
        list_dict_results: List of stock result dictionaries.
        list_s_date_adjustment_notes: List of date adjustment notes.
    """
    print(f"Start Date: {s_start_date}, End Date: {s_end_date}")
    print()

    # Print date adjustment notes if any. Loop through each date adjustment note
    for s_date_adjustment_note in list_s_date_adjustment_notes:
        print(f"Note: {s_date_adjustment_note}")

    if list_s_date_adjustment_notes:
        print()

    # Print column headers
    print("Stock Name,Ticker,ISIN,Start Price,End Price,Percentage,Currency")

    # Print stock data rows. Loop through each stock result
    for dict_result in list_dict_results:
        list_row: List[str] = format_output_row(dict_result)  # Formatted row
        print(",".join(list_row))


# Main Processing Functions

def process_stock(dict_stock: Dict[str, str], s_start_date: str, s_end_date: str) -> Tuple[Dict[str, Any], List[str]]:
    """
    Process a single stock and calculate percentage change.

    Args:
        dict_stock: Stock dictionary with name, ticker, isin.
        s_start_date: Start date string.
        s_end_date: End date string.

    Returns:
        Tuple of (result dictionary, list of date adjustment notes).
    """
    s_name: str = dict_stock.get("s_name", "")  # Stock name
    s_ticker: str = dict_stock.get("s_ticker", "")  # Stock ticker
    s_isin: str = dict_stock.get("s_isin", "")  # Stock ISIN
    list_s_date_adjustment_notes: List[str] = []  # Date adjustment notes for this stock

    dict_result: Dict[str, Any] = {"s_name": s_name, "s_ticker": "", "s_isin": s_isin, "n_start_price": None, "n_end_price": None, "n_percentage": None, "s_currency": "", "s_error": ""}

    # Resolve ticker if not provided
    if not s_ticker:
        try:
            dict_lookup: Dict[str, Any] = lookup_ticker_from_openfigi(s_stock_name=s_name, s_isin=s_isin)

            if dict_lookup.get("b_not_found", False):
                dict_result["s_error"] = "Stock details not found"
                return dict_result, list_s_date_adjustment_notes

            s_ticker = dict_lookup.get("s_ticker", "")
            s_exchange_code: str = dict_lookup.get("s_exchange_code", "")  # Exchange code from OpenFIGI
            s_suffix: str = map_exchange_to_suffix(s_exchange_code)  # yfinance suffix for exchange
            s_ticker = s_ticker + s_suffix

        except ApiError as e:
            raise e

    dict_result["s_ticker"] = s_ticker

    # Adjust dates to trading days
    s_adjusted_start_date: str
    b_start_adjusted: bool
    s_adjusted_start_date, b_start_adjusted = adjust_to_trading_day(s_start_date, b_return_flag=True)

    s_adjusted_end_date: str
    b_end_adjusted: bool
    s_adjusted_end_date, b_end_adjusted = adjust_to_trading_day(s_end_date, b_return_flag=True)

    if b_start_adjusted:
        list_s_date_adjustment_notes.append(f"Start date adjusted to {s_adjusted_start_date} (next trading day) for: {s_name}")

    if b_end_adjusted:
        list_s_date_adjustment_notes.append(f"End date adjusted to {s_adjusted_end_date} (next trading day) for: {s_name}")

    # Fetch stock prices
    try:
        n_start_price: float = fetch_stock_price(s_ticker, s_adjusted_start_date)
        n_end_price: float = fetch_stock_price(s_ticker, s_adjusted_end_date)
    except StockDelistedError:
        dict_result["s_error"] = "Delisted"
        return dict_result, list_s_date_adjustment_notes

    # Calculate percentage change
    n_percentage: float = calculate_percentage_change(n_start_price, n_end_price)

    # Fetch currency
    s_currency: str = fetch_stock_currency(s_ticker)

    # Update result
    dict_result["n_start_price"] = round(n_start_price, 2)
    dict_result["n_end_price"] = round(n_end_price, 2)
    dict_result["n_percentage"] = round(n_percentage, 2)
    dict_result["s_currency"] = s_currency

    return dict_result, list_s_date_adjustment_notes


def main() -> None:
    """Main entry point for the stock calculator."""
    list_s_args: List[str] = sys.argv[1:]  # Command line arguments

    try:
        dict_args: Dict[str, Any] = parse_arguments(list_s_args)
    except CliArgumentError as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

    s_input_file_path: Optional[str] = dict_args.get("s_input_file_path")  # Input CSV file path
    s_stocks: Optional[str] = dict_args.get("s_stocks")  # Comma-separated stocks
    s_start_date_arg: Optional[str] = dict_args.get("s_start_date")  # Start date from CLI
    s_end_date_arg: Optional[str] = dict_args.get("s_end_date")  # End date from CLI
    s_output_directory_path: Optional[str] = dict_args.get("s_output_directory_path")  # Output directory

    # Set default output directory to script location
    if not s_output_directory_path:
        s_output_directory_path = os.path.dirname(os.path.abspath(__file__))

    list_dict_stocks: List[Dict[str, str]] = []  # List of stocks to process
    s_start_date: str = ""  # Start date for calculations
    s_end_date: str = ""  # End date for calculations

    # Parse input source
    if s_input_file_path:
        try:
            dict_parsed: Dict[str, Any] = parse_csv_file(s_input_file_path)
            s_start_date = dict_parsed["s_start_date"]
            s_end_date = dict_parsed["s_end_date"]
            list_dict_stocks = dict_parsed["list_dict_stocks"]
        except CsvParsingError as e:
            print(f"Error: {str(e)}")
            sys.exit(1)
        except FileNotFoundError:
            print(f"Error: File not found: {s_input_file_path}")
            sys.exit(1)
    else:
        s_start_date = s_start_date_arg
        s_end_date = s_end_date_arg

        # Parse comma-separated stock names
        list_s_stock_names: List[str] = [s.strip() for s in s_stocks.split(",")]

        # Loop through each stock name and create dictionary
        for s_stock_name in list_s_stock_names:
            dict_stock: Dict[str, str] = {"s_name": s_stock_name, "s_ticker": "", "s_isin": ""}
            list_dict_stocks.append(dict_stock)

    # Process each stock
    list_dict_results: List[Dict[str, Any]] = []  # Results for all stocks
    list_s_all_date_adjustment_notes: List[str] = []  # All date adjustment notes

    # Loop through each stock to process
    for dict_stock in list_dict_stocks:
        try:
            dict_result: Dict[str, Any]
            list_s_date_adjustment_notes: List[str]
            dict_result, list_s_date_adjustment_notes = process_stock(dict_stock, s_start_date, s_end_date)

            list_dict_results.append(dict_result)
            list_s_all_date_adjustment_notes.extend(list_s_date_adjustment_notes)

        except ApiError as e:
            print(f"Error: {str(e)}")
            sys.exit(1)

    # Generate output filename
    s_output_file_path: str = generate_output_filename(s_output_directory_path)

    # Output results
    print_output_terminal(s_start_date, s_end_date, list_dict_results, list_s_all_date_adjustment_notes)
    write_output_csv(s_output_file_path, s_start_date, s_end_date, list_dict_results, list_s_all_date_adjustment_notes)

    print()
    print(f"Results saved to: {s_output_file_path}")


if __name__ == "__main__":
    main()
