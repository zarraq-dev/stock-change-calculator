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
import math
import os
import re
import sys
import time
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
OPENFIGI_MAPPING_URL: str = "https://api.openfigi.com/v3/mapping"  # OpenFIGI mapping API endpoint (for ISIN lookups)
OPENFIGI_SEARCH_URL: str = "https://api.openfigi.com/v3/search"  # OpenFIGI search API endpoint (for name lookups)
DEFAULT_OUTPUT_FILENAME: str = "stock_changes_output.csv"  # Default output file name
OPENFIGI_BATCH_SIZE: int = 10  # Maximum jobs per OpenFIGI mapping API request (anonymous limit)
OPENFIGI_MAPPING_DELAY_SECONDS: float = 2.5  # Delay between Mapping API requests (25 req/min limit)
OPENFIGI_SEARCH_DELAY_SECONDS: float = 13.0  # Delay between Search API requests (5 req/min limit, +1s buffer)

# Valid security types for filtering OpenFIGI results
LIST_S_VALID_SECURITY_TYPES: List[str] = [
    "Common Stock",  # Standard equities
    "REIT",  # Real Estate Investment Trusts
    "ETP",  # Exchange Traded Products (covers ETFs)
]

# Exchange priority order for result selection (UK first, US second, rest alphabetically)
LIST_S_EXCHANGE_PRIORITY: List[str] = [
    "LN",  # London Stock Exchange (UK)
    "US",  # US exchanges (generic)
    "UN",  # NYSE
    "UQ",  # NASDAQ
    "UM",  # US mutual funds
    "CN",  # Canada (Toronto)
    "CT",  # Canada (TSX Venture)
    "GF",  # Germany (Frankfurt)
    "GR",  # Germany (XETRA)
    "GY",  # Germany (generic)
    "NA",  # Netherlands (Amsterdam/Euronext)
]

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


# Ticker Sanitisation Functions

def sanitise_ticker(s_ticker_raw: str) -> str:
    """
    Sanitise a ticker symbol by removing invalid characters.

    OpenFIGI sometimes returns tickers with trailing slashes (e.g., "NG/")
    which are invalid for yfinance lookups. This function strips such characters.

    Args:
        s_ticker_raw: Raw ticker symbol from API.

    Returns:
        Sanitised ticker symbol with trailing slashes removed.
    """
    s_ticker_sanitised: str = s_ticker_raw.rstrip("/")  # Remove trailing slashes

    return s_ticker_sanitised


# OpenFIGI Lookup Functions

def lookup_ticker_from_openfigi(s_stock_name: str = "", s_isin: str = "") -> Dict[str, Any]:
    """
    Look up stock ticker using Bloomberg OpenFIGI API.

    Uses the Mapping API for ISIN lookups and the Search API for name lookups.

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

    dict_headers: Dict[str, str] = {"Content-Type": "application/json"}  # HTTP request headers

    if s_isin:
        # Use Mapping API for ISIN lookup
        list_dict_api_query: List[Dict[str, str]] = [{"idType": "ID_ISIN", "idValue": s_isin}]

        try:
            response_api = requests.post(OPENFIGI_MAPPING_URL, json=list_dict_api_query, headers=dict_headers, timeout=30)
        except Exception as e:
            raise ApiError(f"OpenFIGI API unreachable. Error: {str(e)}")

        if response_api.status_code == 429:
            raise ApiError("OpenFIGI API rate limit exceeded. Too many requests. Please wait a few minutes and try again.")

        if response_api.status_code != 200:
            raise ApiError(f"OpenFIGI API returned error status: {response_api.status_code}")

        list_dict_response: List[Dict[str, Any]] = response_api.json()  # API response data

        # Check for no match
        if not list_dict_response or "warning" in list_dict_response[0] or "error" in list_dict_response[0]:
            return {"s_ticker": "", "s_exchange_code": "", "b_not_found": True}

        # Extract data from response
        if "data" not in list_dict_response[0] or len(list_dict_response[0]["data"]) == 0:
            return {"s_ticker": "", "s_exchange_code": "", "b_not_found": True}

        dict_first_result: Dict[str, Any] = list_dict_response[0]["data"][0]  # First matching result

        s_ticker_raw: str = dict_first_result.get("ticker", "")  # Raw ticker symbol from API
        s_ticker: str = sanitise_ticker(s_ticker_raw)  # Sanitised ticker symbol
        s_exchange_code: str = dict_first_result.get("exchCode", "")  # Exchange code

        return {"s_ticker": s_ticker, "s_exchange_code": s_exchange_code, "b_not_found": False}

    # Use Search API for name lookup
    dict_query: Dict[str, str] = {"query": s_stock_name}

    try:
        response_api = requests.post(OPENFIGI_SEARCH_URL, json=dict_query, headers=dict_headers, timeout=30)
    except Exception as e:
        raise ApiError(f"OpenFIGI API unreachable. Error: {str(e)}")

    if response_api.status_code == 429:
        raise ApiError("OpenFIGI API rate limit exceeded. Too many requests. Please wait a few minutes and try again.")

    if response_api.status_code != 200:
        raise ApiError(f"OpenFIGI API returned error status: {response_api.status_code}")

    dict_response: Dict[str, Any] = response_api.json()  # API response data

    # Check for no match
    if "data" not in dict_response or len(dict_response.get("data", [])) == 0:
        return {"s_ticker": "", "s_exchange_code": "", "b_not_found": True}

    dict_first_result: Dict[str, Any] = dict_response["data"][0]  # First matching result

    s_ticker_raw: str = dict_first_result.get("ticker", "")  # Raw ticker symbol from API
    s_ticker: str = sanitise_ticker(s_ticker_raw)  # Sanitised ticker symbol
    s_exchange_code: str = dict_first_result.get("exchCode", "")  # Exchange code

    return {"s_ticker": s_ticker, "s_exchange_code": s_exchange_code, "b_not_found": False}


def select_and_validate_best_result(
    list_dict_api_results: List[Dict[str, Any]],
    s_search_query: str,
    s_start_date: str,
    s_end_date: str
) -> Optional[Dict[str, Any]]:
    """
    Select the best result from OpenFIGI Search API and validate it exists in yfinance.

    Uses AND filter logic with exchange priority ordering:
    1. Loop through exchanges in priority order (UK first, US second, rest alphabetically)
    2. For each exchange, loop through all results
    3. When a result matches ALL filter conditions:
       - Security type is valid (Common Stock, REIT, or ETP)
       - Exchange code matches current priority exchange
       - Search query appears in the result name
    4. Validate the ticker exists in yfinance by fetching price data
    5. If valid, return the result with cached prices
    6. If invalid, continue to next exchange in priority order

    Args:
        list_dict_api_results: List of result dictionaries from OpenFIGI Search API.
        s_search_query: Original search query string (used for name matching).
        s_start_date: Start date string in dd-mmm-yy format (for price validation).
        s_end_date: End date string in dd-mmm-yy format (for price validation).

    Returns:
        Dictionary containing API result fields plus cached prices, or None if no valid result found.
        If valid, includes: ticker, exchCode, name, s_full_ticker, n_start_price, n_end_price, s_currency
    """
    s_search_query_lowercase: str = s_search_query.lower()  # Lowercase search query for case-insensitive matching

    # Loop through exchanges in priority order
    for s_priority_exchange_code in LIST_S_EXCHANGE_PRIORITY:
        # Loop through all API results for current exchange
        for dict_api_result in list_dict_api_results:
            s_security_type: str = dict_api_result.get("securityType", "")  # Primary security type
            s_security_type2: str = dict_api_result.get("securityType2", "")  # Secondary security type
            s_exchange_code: str = dict_api_result.get("exchCode", "")  # Exchange code
            s_result_name: str = dict_api_result.get("name", "")  # Result name
            s_result_name_lowercase: str = s_result_name.lower()  # Lowercase name for matching

            # Check ALL filter conditions (AND logic)
            b_has_valid_security_type: bool = (
                s_security_type in LIST_S_VALID_SECURITY_TYPES or
                s_security_type2 in LIST_S_VALID_SECURITY_TYPES
            )  # True if security type is valid
            b_matches_priority_exchange: bool = (s_exchange_code == s_priority_exchange_code)  # True if exchange matches current priority
            b_query_in_name: bool = (s_search_query_lowercase in s_result_name_lowercase)  # True if query appears in name

            if b_has_valid_security_type and b_matches_priority_exchange and b_query_in_name:
                # Build full ticker with exchange suffix
                s_ticker_raw: str = dict_api_result.get("ticker", "")  # Raw ticker from API
                s_ticker_sanitised: str = sanitise_ticker(s_ticker_raw)  # Sanitised ticker
                s_exchange_suffix: str = map_exchange_to_suffix(s_exchange_code)  # Exchange suffix for yfinance
                s_full_ticker: str = s_ticker_sanitised + s_exchange_suffix  # Full ticker with suffix

                # Validate ticker exists in yfinance by fetching prices
                dict_validation: Dict[str, Any] = validate_ticker_and_fetch_prices(
                    s_full_ticker,
                    s_start_date,
                    s_end_date
                )  # Validation result with cached prices

                if dict_validation.get("b_valid", False):
                    # Return API result with cached prices
                    return {
                        "ticker": s_ticker_sanitised,
                        "exchCode": s_exchange_code,
                        "name": s_result_name,
                        "s_full_ticker": s_full_ticker,
                        "n_start_price": dict_validation.get("n_start_price"),
                        "n_end_price": dict_validation.get("n_end_price"),
                        "s_currency": dict_validation.get("s_currency")
                    }
                # If validation failed, continue to next exchange in priority order

    return None  # No valid result found


def resolve_ticker_with_prices(s_stock_name: str, s_start_date: str, s_end_date: str) -> Dict[str, Any]:
    """
    Resolve a stock ticker using OpenFIGI Search API and validate with yfinance.

    Uses exchange priority algorithm to select the best result. Each candidate
    ticker is validated against yfinance by fetching price data. If valid,
    the prices are cached in the result to avoid duplicate API calls later.

    Args:
        s_stock_name: Company name to search.
        s_start_date: Start date string in dd-mmm-yy format.
        s_end_date: End date string in dd-mmm-yy format.

    Returns:
        Dictionary containing:
            - s_ticker: Stock ticker symbol (empty if not found)
            - s_exchange_code: Exchange code (empty if not found)
            - s_full_ticker: Full ticker with exchange suffix (empty if not found)
            - b_not_found: True if stock was not found
            - n_start_price: Cached start price (if valid)
            - n_end_price: Cached end price (if valid)
            - s_currency: Currency code (if valid)

    Raises:
        ApiError: If API is unreachable or returns an error.
    """
    dict_headers: Dict[str, str] = {"Content-Type": "application/json"}  # HTTP request headers
    dict_query: Dict[str, str] = {"query": s_stock_name}  # Search query

    try:
        response_api = requests.post(OPENFIGI_SEARCH_URL, json=dict_query, headers=dict_headers, timeout=30)
    except Exception as e:
        raise ApiError(f"OpenFIGI API unreachable. Error: {str(e)}")

    if response_api.status_code == 429:
        # Extract rate limit headers for debugging
        s_rate_limit: str = response_api.headers.get("ratelimit-limit", "unknown")  # Max requests allowed
        s_rate_remaining: str = response_api.headers.get("ratelimit-remaining", "unknown")  # Requests remaining
        s_rate_reset: str = response_api.headers.get("ratelimit-reset", "unknown")  # Seconds until reset
        s_error_message: str = (
            f"OpenFIGI Search API rate limit exceeded.\n"
            f"  Endpoint: {OPENFIGI_SEARCH_URL}\n"
            f"  Rate Limit: {s_rate_limit} requests\n"
            f"  Remaining: {s_rate_remaining}\n"
            f"  Reset in: {s_rate_reset} seconds\n"
            f"Please wait and try again."
        )
        raise ApiError(s_error_message)

    if response_api.status_code != 200:
        raise ApiError(f"OpenFIGI API returned error status: {response_api.status_code}")

    dict_response: Dict[str, Any] = response_api.json()  # API response data

    # Check for no results from API
    if "data" not in dict_response or len(dict_response.get("data", [])) == 0:
        return {"s_ticker": "", "s_exchange_code": "", "s_full_ticker": "", "b_not_found": True}

    list_dict_api_results: List[Dict[str, Any]] = dict_response["data"]  # All results from API

    # Select best result using exchange priority algorithm with yfinance validation
    dict_best_result: Optional[Dict[str, Any]] = select_and_validate_best_result(
        list_dict_api_results,
        s_stock_name,
        s_start_date,
        s_end_date
    )  # Best matching result after filtering and validation

    # If no valid result found, return not found (no fallback)
    if dict_best_result is None:
        return {"s_ticker": "", "s_exchange_code": "", "s_full_ticker": "", "b_not_found": True}

    s_ticker: str = dict_best_result.get("ticker", "")  # Sanitised ticker symbol
    s_exchange_code: str = dict_best_result.get("exchCode", "")  # Exchange code
    s_full_ticker: str = dict_best_result.get("s_full_ticker", "")  # Full ticker with suffix

    return {
        "s_ticker": s_ticker,
        "s_exchange_code": s_exchange_code,
        "s_full_ticker": s_full_ticker,
        "b_not_found": False,
        "n_start_price": dict_best_result.get("n_start_price"),
        "n_end_price": dict_best_result.get("n_end_price"),
        "s_currency": dict_best_result.get("s_currency")
    }


def resolve_tickers_batch_with_prices(
    list_dict_stocks: List[Dict[str, str]],
    s_start_date: str,
    s_end_date: str
) -> List[Dict[str, Any]]:
    """
    Resolve multiple stock tickers using Bloomberg OpenFIGI API and validate with yfinance.

    Uses the Mapping API (with batching) for ISIN lookups and the Search API for name lookups.
    The Search API does not support batching, so name lookups are done one at a time with delays.
    All resolved tickers are validated against yfinance, and prices are cached.

    Args:
        list_dict_stocks: List of stock dictionaries with s_name, s_ticker, s_isin keys.
                          Only stocks without a ticker will be looked up.
        s_start_date: Start date string in dd-mmm-yy format.
        s_end_date: End date string in dd-mmm-yy format.

    Returns:
        List of result dictionaries in the same order as input, each containing:
            - s_ticker: Stock ticker symbol (empty if not found)
            - s_exchange_code: Exchange code (empty if not found)
            - s_full_ticker: Full ticker with exchange suffix (empty if not found)
            - b_not_found: True if stock was not found
            - b_skipped: True if stock already had a ticker (no lookup needed)
            - n_start_price: Cached start price (if valid)
            - n_end_price: Cached end price (if valid)
            - s_currency: Currency code (if valid)

    Raises:
        ApiError: If API is unreachable or returns an error.
    """
    list_dict_results: List[Dict[str, Any]] = []  # Results for all stocks
    list_n_isin_indices: List[int] = []  # Indices of stocks with ISINs (can be batched)
    list_dict_isin_stocks: List[Dict[str, str]] = []  # Stocks with ISINs
    list_n_name_indices: List[int] = []  # Indices of stocks with only names (cannot be batched)
    list_dict_name_stocks: List[Dict[str, str]] = []  # Stocks with only names

    # Categorize stocks
    for n_index, dict_stock in enumerate(list_dict_stocks):
        s_ticker: str = dict_stock.get("s_ticker", "")  # Existing ticker
        s_isin: str = dict_stock.get("s_isin", "")  # Stock ISIN

        if s_ticker:
            # Stock already has ticker, skip lookup (prices will be fetched later in process_stock)
            list_dict_results.append({
                "s_ticker": s_ticker,
                "s_exchange_code": "",
                "s_full_ticker": "",
                "b_not_found": False,
                "b_skipped": True,
                "n_start_price": None,
                "n_end_price": None,
                "s_currency": None
            })
        elif s_isin:
            # Stock has ISIN, can be batched via Mapping API
            list_dict_results.append(None)  # Placeholder
            list_n_isin_indices.append(n_index)
            list_dict_isin_stocks.append(dict_stock)
        else:
            # Stock has only name, use Search API (no batching)
            list_dict_results.append(None)  # Placeholder
            list_n_name_indices.append(n_index)
            list_dict_name_stocks.append(dict_stock)

    n_isin_count: int = len(list_dict_isin_stocks)  # Number of ISIN lookups
    n_name_count: int = len(list_dict_name_stocks)  # Number of name lookups

    # If no stocks need lookup, return early
    if n_isin_count == 0 and n_name_count == 0:
        return list_dict_results

    print(f"Looking up {n_isin_count + n_name_count} stock tickers ({n_isin_count} by ISIN, {n_name_count} by name)...")

    # Process ISIN lookups using Mapping API with batching
    if n_isin_count > 0:
        n_isin_batch_count: int = math.ceil(n_isin_count / OPENFIGI_BATCH_SIZE)  # Number of ISIN batches

        for n_batch_index in range(n_isin_batch_count):
            n_start_index: int = n_batch_index * OPENFIGI_BATCH_SIZE  # Start index for this batch
            n_end_index: int = min(n_start_index + OPENFIGI_BATCH_SIZE, n_isin_count)  # End index for this batch

            # Build query for this batch
            list_dict_api_query: List[Dict[str, str]] = []  # Query payload for OpenFIGI API

            for n_stock_index in range(n_start_index, n_end_index):
                dict_stock: Dict[str, str] = list_dict_isin_stocks[n_stock_index]  # Current stock
                s_isin: str = dict_stock.get("s_isin", "")  # Stock ISIN
                dict_query: Dict[str, str] = {"idType": "ID_ISIN", "idValue": s_isin}
                list_dict_api_query.append(dict_query)

            # Make API request
            dict_headers: Dict[str, str] = {"Content-Type": "application/json"}  # HTTP request headers

            try:
                response_api = requests.post(OPENFIGI_MAPPING_URL, json=list_dict_api_query, headers=dict_headers, timeout=30)
            except Exception as e:
                raise ApiError(f"OpenFIGI API unreachable. Error: {str(e)}")

            if response_api.status_code == 429:
                # Extract rate limit headers for debugging
                s_rate_limit: str = response_api.headers.get("ratelimit-limit", "unknown")  # Max requests allowed
                s_rate_remaining: str = response_api.headers.get("ratelimit-remaining", "unknown")  # Requests remaining
                s_rate_reset: str = response_api.headers.get("ratelimit-reset", "unknown")  # Seconds until reset
                s_error_message: str = (
                    f"OpenFIGI Mapping API rate limit exceeded.\n"
                    f"  Endpoint: {OPENFIGI_MAPPING_URL}\n"
                    f"  Rate Limit: {s_rate_limit} requests\n"
                    f"  Remaining: {s_rate_remaining}\n"
                    f"  Reset in: {s_rate_reset} seconds\n"
                    f"Please wait and try again."
                )
                raise ApiError(s_error_message)

            if response_api.status_code != 200:
                raise ApiError(f"OpenFIGI API returned error status: {response_api.status_code}")

            list_dict_response: List[Dict[str, Any]] = response_api.json()  # API response data

            # Parse response and update results
            for n_response_index, dict_response in enumerate(list_dict_response):
                n_original_index: int = list_n_isin_indices[n_start_index + n_response_index]  # Original index in input list

                # Check for no match or error
                if "warning" in dict_response or "error" in dict_response or "data" not in dict_response or len(dict_response.get("data", [])) == 0:
                    list_dict_results[n_original_index] = {
                        "s_ticker": "",
                        "s_exchange_code": "",
                        "s_full_ticker": "",
                        "b_not_found": True,
                        "b_skipped": False,
                        "n_start_price": None,
                        "n_end_price": None,
                        "s_currency": None
                    }
                else:
                    dict_first_result: Dict[str, Any] = dict_response["data"][0]  # First matching result
                    s_ticker_raw: str = dict_first_result.get("ticker", "")  # Raw ticker symbol from API
                    s_ticker: str = sanitise_ticker(s_ticker_raw)  # Sanitised ticker symbol
                    s_exchange_code: str = dict_first_result.get("exchCode", "")  # Exchange code
                    s_exchange_suffix: str = map_exchange_to_suffix(s_exchange_code)  # Exchange suffix for yfinance
                    s_full_ticker: str = s_ticker + s_exchange_suffix  # Full ticker with suffix

                    # Validate ticker with yfinance and get cached prices
                    dict_validation: Dict[str, Any] = validate_ticker_and_fetch_prices(
                        s_full_ticker,
                        s_start_date,
                        s_end_date
                    )

                    if dict_validation.get("b_valid", False):
                        list_dict_results[n_original_index] = {
                            "s_ticker": s_ticker,
                            "s_exchange_code": s_exchange_code,
                            "s_full_ticker": s_full_ticker,
                            "b_not_found": False,
                            "b_skipped": False,
                            "n_start_price": dict_validation.get("n_start_price"),
                            "n_end_price": dict_validation.get("n_end_price"),
                            "s_currency": dict_validation.get("s_currency")
                        }
                    else:
                        # Ticker not valid in yfinance, mark as not found
                        list_dict_results[n_original_index] = {
                            "s_ticker": "",
                            "s_exchange_code": "",
                            "s_full_ticker": "",
                            "b_not_found": True,
                            "b_skipped": False,
                            "n_start_price": None,
                            "n_end_price": None,
                            "s_currency": None
                        }

            print(f"  ISIN batch {n_batch_index + 1}/{n_isin_batch_count} complete ({n_end_index - n_start_index} stocks)")

            # Add delay between batches (Mapping API: 25 req/min)
            if n_batch_index < n_isin_batch_count - 1 or n_name_count > 0:
                time.sleep(OPENFIGI_MAPPING_DELAY_SECONDS)

    # Process name lookups using Search API (one at a time)
    if n_name_count > 0:
        # Calculate estimated time for user feedback
        n_estimated_minutes: int = (n_name_count * int(OPENFIGI_SEARCH_DELAY_SECONDS)) // 60  # Estimated minutes
        print(f"  Note: Name lookups are rate-limited to 5/min. Estimated time: ~{n_estimated_minutes} minutes")

        for n_name_index, dict_stock in enumerate(list_dict_name_stocks):
            s_name: str = dict_stock.get("s_name", "")  # Stock name
            n_original_index: int = list_n_name_indices[n_name_index]  # Original index in input list

            dict_lookup_result: Dict[str, Any] = resolve_ticker_with_prices(s_name, s_start_date, s_end_date)

            list_dict_results[n_original_index] = {
                "s_ticker": dict_lookup_result.get("s_ticker", ""),
                "s_exchange_code": dict_lookup_result.get("s_exchange_code", ""),
                "s_full_ticker": dict_lookup_result.get("s_full_ticker", ""),
                "b_not_found": dict_lookup_result.get("b_not_found", True),
                "b_skipped": False,
                "n_start_price": dict_lookup_result.get("n_start_price"),
                "n_end_price": dict_lookup_result.get("n_end_price"),
                "s_currency": dict_lookup_result.get("s_currency")
            }

            # Progress update every 5 stocks (more frequent due to longer delays)
            if (n_name_index + 1) % 5 == 0 or n_name_index == n_name_count - 1:
                print(f"  Name lookup {n_name_index + 1}/{n_name_count} complete")

            # Add delay between requests (Search API: 5 req/min)
            if n_name_index < n_name_count - 1:
                time.sleep(OPENFIGI_SEARCH_DELAY_SECONDS)

    return list_dict_results


def resolve_all_tickers(
    list_dict_stocks: List[Dict[str, str]],
    s_start_date: str,
    s_end_date: str
) -> List[Dict[str, Any]]:
    """
    Resolve tickers for all stocks using batch OpenFIGI lookup with yfinance validation.

    This function looks up tickers for stocks that don't have one provided,
    validates them against yfinance, applies exchange suffixes, and caches
    price data to avoid duplicate API calls.

    Args:
        list_dict_stocks: List of stock dictionaries with s_name, s_ticker, s_isin keys.
        s_start_date: Start date string in dd-mmm-yy format.
        s_end_date: End date string in dd-mmm-yy format.

    Returns:
        Updated list of stock dictionaries with resolved tickers and cached prices.

    Raises:
        ApiError: If OpenFIGI API is unreachable or returns an error.
    """
    # Perform batch lookup with yfinance validation
    list_dict_lookup_results: List[Dict[str, Any]] = resolve_tickers_batch_with_prices(
        list_dict_stocks,
        s_start_date,
        s_end_date
    )

    # Update stock dictionaries with lookup results
    for n_index, dict_lookup_result in enumerate(list_dict_lookup_results):
        if dict_lookup_result.get("b_skipped", False):
            # Stock already had ticker, no changes needed (prices will be fetched later)
            continue

        if dict_lookup_result.get("b_not_found", False):
            # Stock not found, leave ticker empty (will be handled later as error)
            list_dict_stocks[n_index]["s_ticker"] = ""
            list_dict_stocks[n_index]["b_not_found"] = True
        else:
            # Use the full ticker already validated and constructed
            s_full_ticker: str = dict_lookup_result.get("s_full_ticker", "")  # Full ticker with suffix

            list_dict_stocks[n_index]["s_ticker"] = s_full_ticker
            list_dict_stocks[n_index]["b_not_found"] = False

            # Store cached prices to avoid re-fetching later
            list_dict_stocks[n_index]["n_start_price"] = dict_lookup_result.get("n_start_price")
            list_dict_stocks[n_index]["n_end_price"] = dict_lookup_result.get("n_end_price")
            list_dict_stocks[n_index]["s_currency"] = dict_lookup_result.get("s_currency")

    return list_dict_stocks


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

def validate_ticker_and_fetch_prices(s_ticker: str, s_start_date: str, s_end_date: str) -> Dict[str, Any]:
    """
    Validate a ticker exists in yfinance by attempting to fetch price data.

    This function serves dual purpose:
    1. Validates that a ticker is recognised by yfinance
    2. Fetches and caches price data to avoid duplicate API calls later

    Args:
        s_ticker: Stock ticker symbol (including exchange suffix if needed).
        s_start_date: Start date string in dd-mmm-yy format.
        s_end_date: End date string in dd-mmm-yy format.

    Returns:
        Dictionary containing:
            - b_valid: True if ticker has price data, False otherwise
            - n_start_price: Start closing price (if valid)
            - n_end_price: End closing price (if valid)
            - s_currency: Currency code (if valid)
    """
    try:
        # Adjust dates to trading days (skip weekends)
        s_adjusted_start_date: str = adjust_to_trading_day(s_start_date)  # Adjusted start date
        s_adjusted_end_date: str = adjust_to_trading_day(s_end_date)  # Adjusted end date

        # Parse dates for yfinance
        dt_start: datetime = datetime.strptime(s_adjusted_start_date, DATE_FORMAT)  # Parsed start date
        dt_start_end: datetime = dt_start + timedelta(days=5)  # End of start date range (covers holidays)
        dt_end: datetime = datetime.strptime(s_adjusted_end_date, DATE_FORMAT)  # Parsed end date
        dt_end_end: datetime = dt_end + timedelta(days=5)  # End of end date range (covers holidays)

        s_yf_start_begin: str = dt_start.strftime("%Y-%m-%d")  # yfinance format for start date begin
        s_yf_start_end: str = dt_start_end.strftime("%Y-%m-%d")  # yfinance format for start date end
        s_yf_end_begin: str = dt_end.strftime("%Y-%m-%d")  # yfinance format for end date begin
        s_yf_end_end: str = dt_end_end.strftime("%Y-%m-%d")  # yfinance format for end date end

        # Create yfinance ticker object
        ticker_stock = yfinance.Ticker(s_ticker)  # yfinance Ticker object

        # Fetch start price
        df_start_history = ticker_stock.history(start=s_yf_start_begin, end=s_yf_start_end)  # Start date price history
        if df_start_history.empty:
            return {"b_valid": False}

        n_start_price: float = df_start_history["Close"].iloc[0]  # First available closing price for start

        # Fetch end price
        df_end_history = ticker_stock.history(start=s_yf_end_begin, end=s_yf_end_end)  # End date price history
        if df_end_history.empty:
            return {"b_valid": False}

        n_end_price: float = df_end_history["Close"].iloc[0]  # First available closing price for end

        # Fetch currency
        dict_info: Dict[str, Any] = ticker_stock.info  # Stock info dictionary
        s_currency: str = dict_info.get("currency", "N/A")  # Currency code

        return {
            "b_valid": True,
            "n_start_price": round(n_start_price, 2),
            "n_end_price": round(n_end_price, 2),
            "s_currency": s_currency
        }

    except Exception:
        # Any error means ticker is invalid or has no data
        return {"b_valid": False}


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

def process_stock(dict_stock: Dict[str, Any], s_start_date: str, s_end_date: str) -> Tuple[Dict[str, Any], List[str]]:
    """
    Process a single stock and calculate percentage change.

    Expects ticker to be pre-resolved via resolve_all_tickers() before calling this function.
    If prices were cached during ticker resolution, they will be used directly.
    Otherwise, prices will be fetched from yfinance.

    Args:
        dict_stock: Stock dictionary with name, ticker, isin, optional cached prices, and optional b_not_found flag.
        s_start_date: Start date string.
        s_end_date: End date string.

    Returns:
        Tuple of (result dictionary, list of date adjustment notes).
    """
    s_name: str = dict_stock.get("s_name", "")  # Stock name
    s_ticker: str = dict_stock.get("s_ticker", "")  # Stock ticker (pre-resolved)
    s_isin: str = dict_stock.get("s_isin", "")  # Stock ISIN
    b_not_found: bool = dict_stock.get("b_not_found", False)  # Flag indicating ticker lookup failed
    list_s_date_adjustment_notes: List[str] = []  # Date adjustment notes for this stock

    # Check for cached prices from ticker resolution
    n_cached_start_price: Optional[float] = dict_stock.get("n_start_price")  # Cached start price
    n_cached_end_price: Optional[float] = dict_stock.get("n_end_price")  # Cached end price
    s_cached_currency: Optional[str] = dict_stock.get("s_currency")  # Cached currency

    dict_result: Dict[str, Any] = {"s_name": s_name, "s_ticker": s_ticker, "s_isin": s_isin, "n_start_price": None, "n_end_price": None, "n_percentage": None, "s_currency": "", "s_error": ""}

    # Check if ticker lookup failed
    if b_not_found or not s_ticker:
        dict_result["s_error"] = "Stock details not found"
        return dict_result, list_s_date_adjustment_notes

    # Check if we have cached prices from ticker resolution
    if n_cached_start_price is not None and n_cached_end_price is not None and s_cached_currency is not None:
        # Use cached prices - no need to fetch again
        n_start_price: float = n_cached_start_price
        n_end_price: float = n_cached_end_price
        s_currency: str = s_cached_currency

        # Calculate percentage change
        n_percentage: float = calculate_percentage_change(n_start_price, n_end_price)

        # Update result
        dict_result["n_start_price"] = round(n_start_price, 2)
        dict_result["n_end_price"] = round(n_end_price, 2)
        dict_result["n_percentage"] = round(n_percentage, 2)
        dict_result["s_currency"] = s_currency

        return dict_result, list_s_date_adjustment_notes

    # No cached prices - fetch from yfinance (for user-provided tickers)
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
        n_start_price_fetched: float = fetch_stock_price(s_ticker, s_adjusted_start_date)
        n_end_price_fetched: float = fetch_stock_price(s_ticker, s_adjusted_end_date)
    except StockDelistedError:
        dict_result["s_error"] = "Delisted"
        return dict_result, list_s_date_adjustment_notes

    # Calculate percentage change
    n_percentage_calc: float = calculate_percentage_change(n_start_price_fetched, n_end_price_fetched)

    # Fetch currency
    s_currency_fetched: str = fetch_stock_currency(s_ticker)

    # Update result
    dict_result["n_start_price"] = round(n_start_price_fetched, 2)
    dict_result["n_end_price"] = round(n_end_price_fetched, 2)
    dict_result["n_percentage"] = round(n_percentage_calc, 2)
    dict_result["s_currency"] = s_currency_fetched

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

    # Resolve all tickers using batch OpenFIGI lookup with yfinance validation
    try:
        list_dict_stocks = resolve_all_tickers(list_dict_stocks, s_start_date, s_end_date)
    except ApiError as e:
        print(f"Error: {str(e)}")
        sys.exit(1)

    # Process each stock
    list_dict_results: List[Dict[str, Any]] = []  # Results for all stocks
    list_s_all_date_adjustment_notes: List[str] = []  # All date adjustment notes

    print("Fetching stock prices...")

    # Loop through each stock to process
    for dict_stock in list_dict_stocks:
        dict_result: Dict[str, Any]
        list_s_date_adjustment_notes: List[str]
        dict_result, list_s_date_adjustment_notes = process_stock(dict_stock, s_start_date, s_end_date)

        list_dict_results.append(dict_result)
        list_s_all_date_adjustment_notes.extend(list_s_date_adjustment_notes)

    # Generate output filename
    s_output_file_path: str = generate_output_filename(s_output_directory_path)

    # Output results
    print_output_terminal(s_start_date, s_end_date, list_dict_results, list_s_all_date_adjustment_notes)
    write_output_csv(s_output_file_path, s_start_date, s_end_date, list_dict_results, list_s_all_date_adjustment_notes)

    print()
    print(f"Results saved to: {s_output_file_path}")


if __name__ == "__main__":
    main()
