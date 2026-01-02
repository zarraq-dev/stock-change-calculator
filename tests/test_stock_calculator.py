"""
Test suite for Stock Change Calculator

USAGE
    pytest tests/test_stock_calculator.py -v

TEST CATEGORIES
    1. CSV Parsing
    2. Percentage Calculation
    3. OpenFIGI Lookup
    4. Date Adjustment
    5. Output File Versioning
    6. CLI Argument Parsing
"""

import pytest
import os
import tempfile
from datetime import datetime
from unittest.mock import Mock, patch, MagicMock
from typing import Dict, List, Any


# CSV Parsing Tests

class TestCsvParsing:
    """Tests for CSV file parsing functionality."""

    def test_parse_valid_csv_with_correct_structure(self) -> None:
        """Test parsing a valid CSV file with correct structure."""
        # Arrange
        s_csv_content: str = """Start Date,01-Oct-25,End Date,01-Jan-26
,,,
Stocks,,,
Name,Ticker,ISIN,
Apple Inc,AAPL,,
Microsoft Corp,MSFT,US5949181045,
"""
        # Act & Assert
        # Should return parsed data with dates and stock list
        from src.stock_calculator import parse_csv_file

        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as file_temp:
            file_temp.write(s_csv_content)
            s_temp_path: str = file_temp.name  # Path to temporary test file

        try:
            dict_result: Dict[str, Any] = parse_csv_file(s_temp_path)

            assert dict_result['s_start_date'] == '01-Oct-25'
            assert dict_result['s_end_date'] == '01-Jan-26'
            assert len(dict_result['list_dict_stocks']) == 2
            assert dict_result['list_dict_stocks'][0]['s_name'] == 'Apple Inc'
            assert dict_result['list_dict_stocks'][0]['s_ticker'] == 'AAPL'
            assert dict_result['list_dict_stocks'][1]['s_isin'] == 'US5949181045'
        finally:
            os.unlink(s_temp_path)

    def test_parse_csv_missing_dates_row_raises_exception(self) -> None:
        """Test that missing dates row raises a clear exception."""
        # Arrange
        s_csv_content: str = """Stocks,,,
Name,Ticker,ISIN,
Apple Inc,AAPL,,
"""
        from src.stock_calculator import parse_csv_file, CsvParsingError

        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as file_temp:
            file_temp.write(s_csv_content)
            s_temp_path: str = file_temp.name  # Path to temporary test file

        try:
            # Act & Assert
            with pytest.raises(CsvParsingError) as exc_info:
                parse_csv_file(s_temp_path)

            assert 'date' in str(exc_info.value).lower()
        finally:
            os.unlink(s_temp_path)

    def test_parse_csv_invalid_date_format_raises_exception(self) -> None:
        """Test that invalid date format raises a clear exception."""
        # Arrange
        s_csv_content: str = """Start Date,2025-10-01,End Date,2026-01-01
,,,
Stocks,,,
Name,Ticker,ISIN,
Apple Inc,AAPL,,
"""
        from src.stock_calculator import parse_csv_file, CsvParsingError

        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as file_temp:
            file_temp.write(s_csv_content)
            s_temp_path: str = file_temp.name  # Path to temporary test file

        try:
            # Act & Assert
            with pytest.raises(CsvParsingError) as exc_info:
                parse_csv_file(s_temp_path)

            assert 'date format' in str(exc_info.value).lower()
        finally:
            os.unlink(s_temp_path)

    def test_parse_csv_empty_stock_list_raises_exception(self) -> None:
        """Test that empty stock list raises a clear exception."""
        # Arrange
        s_csv_content: str = """Start Date,01-Oct-25,End Date,01-Jan-26
,,,
Stocks,,,
Name,Ticker,ISIN,
"""
        from src.stock_calculator import parse_csv_file, CsvParsingError

        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as file_temp:
            file_temp.write(s_csv_content)
            s_temp_path: str = file_temp.name  # Path to temporary test file

        try:
            # Act & Assert
            with pytest.raises(CsvParsingError) as exc_info:
                parse_csv_file(s_temp_path)

            assert 'empty' in str(exc_info.value).lower() or 'no stocks' in str(exc_info.value).lower()
        finally:
            os.unlink(s_temp_path)

    def test_parse_csv_with_only_stock_names(self) -> None:
        """Test parsing CSV where only stock names are provided (no ticker/ISIN)."""
        # Arrange
        s_csv_content: str = """Start Date,01-Oct-25,End Date,01-Jan-26
,,,
Stocks,,,
Name,Ticker,ISIN,
National Grid PLC,,,
Shell PLC,,,
"""
        from src.stock_calculator import parse_csv_file

        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as file_temp:
            file_temp.write(s_csv_content)
            s_temp_path: str = file_temp.name  # Path to temporary test file

        try:
            # Act
            dict_result: Dict[str, Any] = parse_csv_file(s_temp_path)

            # Assert
            assert len(dict_result['list_dict_stocks']) == 2
            assert dict_result['list_dict_stocks'][0]['s_name'] == 'National Grid PLC'
            assert dict_result['list_dict_stocks'][0]['s_ticker'] == ''
            assert dict_result['list_dict_stocks'][0]['s_isin'] == ''
        finally:
            os.unlink(s_temp_path)


# Percentage Calculation Tests

class TestPercentageCalculation:
    """Tests for percentage change calculation."""

    def test_calculate_positive_percentage_change(self) -> None:
        """Test calculation when stock price increased."""
        from src.stock_calculator import calculate_percentage_change

        # Arrange
        n_start_price: float = 100.0  # Starting price
        n_end_price: float = 150.0  # Ending price (50% increase)

        # Act
        n_result: float = calculate_percentage_change(n_start_price, n_end_price)

        # Assert
        assert n_result == 50.0

    def test_calculate_negative_percentage_change(self) -> None:
        """Test calculation when stock price decreased."""
        from src.stock_calculator import calculate_percentage_change

        # Arrange
        n_start_price: float = 100.0  # Starting price
        n_end_price: float = 75.0  # Ending price (25% decrease)

        # Act
        n_result: float = calculate_percentage_change(n_start_price, n_end_price)

        # Assert
        assert n_result == -25.0

    def test_calculate_zero_percentage_change(self) -> None:
        """Test calculation when stock price unchanged."""
        from src.stock_calculator import calculate_percentage_change

        # Arrange
        n_start_price: float = 100.0  # Starting price
        n_end_price: float = 100.0  # Ending price (no change)

        # Act
        n_result: float = calculate_percentage_change(n_start_price, n_end_price)

        # Assert
        assert n_result == 0.0

    def test_calculate_percentage_with_decimal_precision(self) -> None:
        """Test calculation returns appropriate decimal precision."""
        from src.stock_calculator import calculate_percentage_change

        # Arrange
        n_start_price: float = 100.0  # Starting price
        n_end_price: float = 116.78  # Ending price

        # Act
        n_result: float = calculate_percentage_change(n_start_price, n_end_price)

        # Assert
        assert round(n_result, 2) == 16.78

    def test_calculate_percentage_with_small_numbers(self) -> None:
        """Test calculation with small stock prices (penny stocks)."""
        from src.stock_calculator import calculate_percentage_change

        # Arrange
        n_start_price: float = 0.05  # Starting price (5 cents)
        n_end_price: float = 0.10  # Ending price (10 cents, 100% increase)

        # Act
        n_result: float = calculate_percentage_change(n_start_price, n_end_price)

        # Assert
        assert n_result == 100.0

    def test_calculate_percentage_with_large_numbers(self) -> None:
        """Test calculation with large stock prices."""
        from src.stock_calculator import calculate_percentage_change

        # Arrange
        n_start_price: float = 500000.0  # Starting price (Berkshire Hathaway style)
        n_end_price: float = 550000.0  # Ending price (10% increase)

        # Act
        n_result: float = calculate_percentage_change(n_start_price, n_end_price)

        # Assert
        assert n_result == 10.0


# OpenFIGI Lookup Tests

class TestOpenFigiLookup:
    """Tests for OpenFIGI API ticker resolution."""

    @patch('src.stock_calculator.requests.post')
    def test_lookup_ticker_by_name_success(self, mock_post: Mock) -> None:
        """Test successful ticker lookup using company name."""
        from src.stock_calculator import lookup_ticker_from_openfigi

        # Arrange
        s_company_name: str = 'Apple Inc'  # Company name to search

        mock_response: Mock = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = [
            {
                'data': [
                    {
                        'ticker': 'AAPL',
                        'exchCode': 'US',
                        'name': 'APPLE INC'
                    }
                ]
            }
        ]
        mock_post.return_value = mock_response

        # Act
        dict_result: Dict[str, str] = lookup_ticker_from_openfigi(s_stock_name=s_company_name)

        # Assert
        assert dict_result['s_ticker'] == 'AAPL'

    @patch('src.stock_calculator.requests.post')
    def test_lookup_ticker_by_isin_success(self, mock_post: Mock) -> None:
        """Test successful ticker lookup using ISIN."""
        from src.stock_calculator import lookup_ticker_from_openfigi

        # Arrange
        s_isin: str = 'US0378331005'  # Apple ISIN

        mock_response: Mock = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = [
            {
                'data': [
                    {
                        'ticker': 'AAPL',
                        'exchCode': 'US',
                        'name': 'APPLE INC'
                    }
                ]
            }
        ]
        mock_post.return_value = mock_response

        # Act
        dict_result: Dict[str, str] = lookup_ticker_from_openfigi(s_isin=s_isin)

        # Assert
        assert dict_result['s_ticker'] == 'AAPL'

    @patch('src.stock_calculator.requests.post')
    def test_lookup_ticker_not_found(self, mock_post: Mock) -> None:
        """Test handling when stock is not found in OpenFIGI."""
        from src.stock_calculator import lookup_ticker_from_openfigi

        # Arrange
        s_company_name: str = 'Nonexistent Company XYZ'  # Invalid company name

        mock_response: Mock = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = [
            {
                'warning': 'No match found'
            }
        ]
        mock_post.return_value = mock_response

        # Act
        dict_result: Dict[str, str] = lookup_ticker_from_openfigi(s_stock_name=s_company_name)

        # Assert
        assert dict_result['s_ticker'] == ''
        assert dict_result['b_not_found'] == True

    @patch('src.stock_calculator.requests.post')
    def test_lookup_ticker_api_error(self, mock_post: Mock) -> None:
        """Test handling when OpenFIGI API returns an error."""
        from src.stock_calculator import lookup_ticker_from_openfigi, ApiError

        # Arrange
        s_company_name: str = 'Apple Inc'  # Company name to search

        mock_post.side_effect = Exception('Connection refused')

        # Act & Assert
        with pytest.raises(ApiError) as exc_info:
            lookup_ticker_from_openfigi(s_stock_name=s_company_name)

        assert 'api' in str(exc_info.value).lower() or 'unreachable' in str(exc_info.value).lower()

    @patch('src.stock_calculator.requests.post')
    def test_lookup_returns_exchange_code_for_suffix(self, mock_post: Mock) -> None:
        """Test that lookup returns exchange code for yfinance suffix mapping."""
        from src.stock_calculator import lookup_ticker_from_openfigi

        # Arrange
        s_company_name: str = 'National Grid PLC'  # UK listed company

        mock_response: Mock = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = [
            {
                'data': [
                    {
                        'ticker': 'NG',
                        'exchCode': 'LN',
                        'name': 'NATIONAL GRID PLC'
                    }
                ]
            }
        ]
        mock_post.return_value = mock_response

        # Act
        dict_result: Dict[str, str] = lookup_ticker_from_openfigi(s_stock_name=s_company_name)

        # Assert
        assert dict_result['s_ticker'] == 'NG'
        assert dict_result['s_exchange_code'] == 'LN'


# Exchange Suffix Mapping Tests

class TestExchangeSuffixMapping:
    """Tests for mapping exchange codes to yfinance ticker suffixes."""

    def test_map_us_exchange_no_suffix(self) -> None:
        """Test that US exchanges (NYSE, NASDAQ) have no suffix."""
        from src.stock_calculator import map_exchange_to_suffix

        # Act & Assert
        assert map_exchange_to_suffix('US') == ''
        assert map_exchange_to_suffix('UN') == ''
        assert map_exchange_to_suffix('UQ') == ''

    def test_map_london_exchange_suffix(self) -> None:
        """Test that London Stock Exchange maps to .L suffix."""
        from src.stock_calculator import map_exchange_to_suffix

        # Act & Assert
        assert map_exchange_to_suffix('LN') == '.L'

    def test_map_unknown_exchange_returns_empty(self) -> None:
        """Test that unknown exchange codes return empty suffix."""
        from src.stock_calculator import map_exchange_to_suffix

        # Act & Assert
        assert map_exchange_to_suffix('UNKNOWN') == ''


# Date Adjustment Tests

class TestDateAdjustment:
    """Tests for adjusting dates to valid trading days."""

    def test_adjust_weekend_saturday_to_monday(self) -> None:
        """Test that Saturday is adjusted to the following Monday."""
        from src.stock_calculator import adjust_to_trading_day

        # Arrange
        s_date: str = '04-Jan-25'  # Saturday

        # Act
        s_adjusted: str = adjust_to_trading_day(s_date)

        # Assert
        assert s_adjusted == '06-Jan-25'  # Monday

    def test_adjust_weekend_sunday_to_monday(self) -> None:
        """Test that Sunday is adjusted to the following Monday."""
        from src.stock_calculator import adjust_to_trading_day

        # Arrange
        s_date: str = '05-Jan-25'  # Sunday

        # Act
        s_adjusted: str = adjust_to_trading_day(s_date)

        # Assert
        assert s_adjusted == '06-Jan-25'  # Monday

    def test_weekday_unchanged(self) -> None:
        """Test that a weekday remains unchanged."""
        from src.stock_calculator import adjust_to_trading_day

        # Arrange
        s_date: str = '06-Jan-25'  # Monday

        # Act
        s_adjusted: str = adjust_to_trading_day(s_date)

        # Assert
        assert s_adjusted == '06-Jan-25'  # Same date

    def test_returns_adjustment_flag_when_changed(self) -> None:
        """Test that function indicates when date was adjusted."""
        from src.stock_calculator import adjust_to_trading_day

        # Arrange
        s_date: str = '04-Jan-25'  # Saturday

        # Act
        s_adjusted: str
        b_was_adjusted: bool
        s_adjusted, b_was_adjusted = adjust_to_trading_day(s_date, b_return_flag=True)

        # Assert
        assert b_was_adjusted == True

    def test_returns_no_adjustment_flag_when_unchanged(self) -> None:
        """Test that function indicates when date was not adjusted."""
        from src.stock_calculator import adjust_to_trading_day

        # Arrange
        s_date: str = '06-Jan-25'  # Monday

        # Act
        s_adjusted: str
        b_was_adjusted: bool
        s_adjusted, b_was_adjusted = adjust_to_trading_day(s_date, b_return_flag=True)

        # Assert
        assert b_was_adjusted == False


# Output File Versioning Tests

class TestOutputFileVersioning:
    """Tests for output file versioning logic."""

    def test_no_existing_file_uses_default_name(self) -> None:
        """Test that default filename is used when no file exists."""
        from src.stock_calculator import generate_output_filename

        # Arrange
        with tempfile.TemporaryDirectory() as s_temp_dir:
            # Act
            s_filename: str = generate_output_filename(s_temp_dir)

            # Assert
            assert s_filename == os.path.join(s_temp_dir, 'stock_changes_output.csv')

    def test_existing_file_creates_v1(self) -> None:
        """Test that v1 suffix is added when default file exists."""
        from src.stock_calculator import generate_output_filename

        # Arrange
        with tempfile.TemporaryDirectory() as s_temp_dir:
            s_existing_path: str = os.path.join(s_temp_dir, 'stock_changes_output.csv')
            with open(s_existing_path, 'w') as file_existing:
                file_existing.write('dummy content')  # Create existing file

            # Act
            s_filename: str = generate_output_filename(s_temp_dir)

            # Assert
            assert s_filename == os.path.join(s_temp_dir, 'stock_changes_output_v1.csv')

    def test_existing_v1_creates_v2(self) -> None:
        """Test that v2 suffix is used when v1 already exists."""
        from src.stock_calculator import generate_output_filename

        # Arrange
        with tempfile.TemporaryDirectory() as s_temp_dir:
            s_default_path: str = os.path.join(s_temp_dir, 'stock_changes_output.csv')
            s_v1_path: str = os.path.join(s_temp_dir, 'stock_changes_output_v1.csv')

            with open(s_default_path, 'w') as file_default:
                file_default.write('dummy content')  # Create default file
            with open(s_v1_path, 'w') as file_v1:
                file_v1.write('dummy content')  # Create v1 file

            # Act
            s_filename: str = generate_output_filename(s_temp_dir)

            # Assert
            assert s_filename == os.path.join(s_temp_dir, 'stock_changes_output_v2.csv')

    def test_increments_version_correctly(self) -> None:
        """Test that version number increments correctly with multiple existing files."""
        from src.stock_calculator import generate_output_filename

        # Arrange
        with tempfile.TemporaryDirectory() as s_temp_dir:
            # Create default, v1, v2, v3 files
            for s_suffix in ['', '_v1', '_v2', '_v3']:
                s_path: str = os.path.join(s_temp_dir, f'stock_changes_output{s_suffix}.csv')
                with open(s_path, 'w') as file_versioned:
                    file_versioned.write('dummy content')

            # Act
            s_filename: str = generate_output_filename(s_temp_dir)

            # Assert
            assert s_filename == os.path.join(s_temp_dir, 'stock_changes_output_v4.csv')


# CLI Argument Parsing Tests

class TestCliArgumentParsing:
    """Tests for command-line argument parsing."""

    def test_parse_valid_file_argument(self) -> None:
        """Test parsing valid --file argument."""
        from src.stock_calculator import parse_arguments

        # Arrange
        list_s_args: List[str] = ['--file', 'input.csv']  # Command line arguments

        # Act
        dict_args: Dict[str, Any] = parse_arguments(list_s_args)

        # Assert
        assert dict_args['s_input_file_path'] == 'input.csv'

    def test_parse_valid_stocks_arguments(self) -> None:
        """Test parsing valid --stocks, --start, --end arguments."""
        from src.stock_calculator import parse_arguments

        # Arrange
        list_s_args: List[str] = [
            '--stocks', 'Apple,Microsoft,Shell',
            '--start', '01-Jan-25',
            '--end', '01-Apr-25'
        ]

        # Act
        dict_args: Dict[str, Any] = parse_arguments(list_s_args)

        # Assert
        assert dict_args['s_stocks'] == 'Apple,Microsoft,Shell'
        assert dict_args['s_start_date'] == '01-Jan-25'
        assert dict_args['s_end_date'] == '01-Apr-25'

    def test_parse_optional_output_argument(self) -> None:
        """Test parsing optional --output argument."""
        from src.stock_calculator import parse_arguments

        # Arrange
        list_s_args: List[str] = [
            '--file', 'input.csv',
            '--output', 'C:/output/folder/'
        ]

        # Act
        dict_args: Dict[str, Any] = parse_arguments(list_s_args)

        # Assert
        assert dict_args['s_output_directory_path'] == 'C:/output/folder/'

    def test_missing_required_arguments_raises_exception(self) -> None:
        """Test that missing required arguments raises exception."""
        from src.stock_calculator import parse_arguments, CliArgumentError

        # Arrange
        list_s_args: List[str] = []  # Empty arguments

        # Act & Assert
        with pytest.raises(CliArgumentError):
            parse_arguments(list_s_args)

    def test_stocks_without_dates_raises_exception(self) -> None:
        """Test that --stocks without --start and --end raises exception."""
        from src.stock_calculator import parse_arguments, CliArgumentError

        # Arrange
        list_s_args: List[str] = ['--stocks', 'Apple,Microsoft']  # Missing dates

        # Act & Assert
        with pytest.raises(CliArgumentError) as exc_info:
            parse_arguments(list_s_args)

        assert 'start' in str(exc_info.value).lower() or 'end' in str(exc_info.value).lower()

    def test_file_and_stocks_mutually_exclusive(self) -> None:
        """Test that --file and --stocks cannot be used together."""
        from src.stock_calculator import parse_arguments, CliArgumentError

        # Arrange
        list_s_args: List[str] = [
            '--file', 'input.csv',
            '--stocks', 'Apple,Microsoft',
            '--start', '01-Jan-25',
            '--end', '01-Apr-25'
        ]

        # Act & Assert
        with pytest.raises(CliArgumentError) as exc_info:
            parse_arguments(list_s_args)

        assert 'mutually exclusive' in str(exc_info.value).lower() or 'both' in str(exc_info.value).lower()

    def test_invalid_date_format_in_cli_raises_exception(self) -> None:
        """Test that invalid date format in CLI raises exception."""
        from src.stock_calculator import parse_arguments, CliArgumentError

        # Arrange
        list_s_args: List[str] = [
            '--stocks', 'Apple',
            '--start', '2025-01-01',  # Wrong format (should be dd-mmm-yy)
            '--end', '01-Apr-25'
        ]

        # Act & Assert
        with pytest.raises(CliArgumentError) as exc_info:
            parse_arguments(list_s_args)

        assert 'date format' in str(exc_info.value).lower()


# Stock Data Fetching Tests

class TestStockDataFetching:
    """Tests for fetching stock price data from yfinance."""

    @patch('src.stock_calculator.yfinance.Ticker')
    def test_fetch_stock_price_success(self, mock_ticker_class: Mock) -> None:
        """Test successful stock price fetch."""
        from src.stock_calculator import fetch_stock_price

        # Arrange
        s_ticker: str = 'AAPL'  # Stock ticker symbol
        s_date: str = '06-Jan-25'  # Date to fetch price for

        mock_ticker: Mock = Mock()
        mock_history: Mock = Mock()
        mock_history.empty = False
        mock_history.__getitem__ = Mock(return_value=Mock(__getitem__=Mock(return_value=150.25)))
        mock_ticker.history.return_value = mock_history
        mock_ticker_class.return_value = mock_ticker

        # Act
        n_price: float = fetch_stock_price(s_ticker, s_date)

        # Assert
        assert n_price == 150.25

    @patch('src.stock_calculator.yfinance.Ticker')
    def test_fetch_stock_price_delisted(self, mock_ticker_class: Mock) -> None:
        """Test handling of delisted stock."""
        from src.stock_calculator import fetch_stock_price, StockDelistedError

        # Arrange
        s_ticker: str = 'FRC'  # First Republic Bank (delisted)
        s_date: str = '06-Jan-25'  # Date to fetch price for

        mock_ticker: Mock = Mock()
        mock_history: Mock = Mock()
        mock_history.empty = True  # No data available
        mock_ticker.history.return_value = mock_history
        mock_ticker_class.return_value = mock_ticker

        # Act & Assert
        with pytest.raises(StockDelistedError):
            fetch_stock_price(s_ticker, s_date)

    @patch('src.stock_calculator.yfinance.Ticker')
    def test_fetch_stock_currency(self, mock_ticker_class: Mock) -> None:
        """Test fetching stock currency."""
        from src.stock_calculator import fetch_stock_currency

        # Arrange
        s_ticker: str = 'AAPL'  # Stock ticker symbol

        mock_ticker: Mock = Mock()
        mock_ticker.info = {'currency': 'USD'}
        mock_ticker_class.return_value = mock_ticker

        # Act
        s_currency: str = fetch_stock_currency(s_ticker)

        # Assert
        assert s_currency == 'USD'
