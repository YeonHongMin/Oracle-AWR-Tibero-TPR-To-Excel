#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AWR HTML to Excel Converter with Charts.
Converts Oracle AWR HTML spool reports into formatted Excel workbooks,
automatically applying cell formats and charts defined in INI configuration files.

Made By, yeonhong.min@gmail.com
Version 1.0 - 2025.11.21
"""

import re
import io
import os
import sys
import argparse
import codecs
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any
import xlsxwriter


# Application-wide constants
class Constants:
    """Constants shared across the converter."""
    
    # File encodings
    ENCODING_EUC_KR = 'euc-kr'
    ENCODING_UTF8 = 'utf-8'
    
    # File extensions
    EXCEL_EXTENSION = '.xlsx'
    
    # HTML parsing markers and tags
    INI_NAME_PATTERN = r'^SQL&gt; rem INI_NAME=(.*)'
    TABLE_START_TAG = '<table'
    TABLE_END_TAG = '</table>'
    ROW_TAG = '<tr'
    CELL_TAG_HEADER = '<th'
    CELL_TAG_DATA = '<td'
    
    # Excel limits and defaults
    # Maximum rows per .xlsx worksheet (legacy .xls was 65,536); also used when expanding 'E' (end row) in FORMAT ranges
    MAX_EXCEL_ROWS = 1048576
    DEFAULT_CHART_WIDTH = 15
    DEFAULT_CHART_HEIGHT = 10
    
    # INI directive prefixes and patterns
    INI_FORMAT_PREFIX = 'FORMAT'
    INI_CHART_PREFIX = 'CHART'
    INI_SQL_PATTERN = r'^SQL\d+='

    # Timestamp column pattern (SQL: to_char(..., 'yyyy-mm-dd hh24:mi:ss')) and default chart X-axis format
    TIMESTAMP_PATTERN = r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$'
    X_TIME_FORMAT_DEFAULT = 'mm-dd hh:mm'

    # Paste direction in SQLn directives (matching the original OSRAnalyzer macro):
    #   T2D (Top to Down) : Default. Writes query rows normally down the worksheet.
    #   Other (e.g. LTOR) : Transposed paste. Each query record becomes one column on the target sheet.
    DIRECTION_T2D = 'T2D'
    
    # Chart types
    CHART_TYPE_LINE = 'LINE'
    CHART_TYPE_BAR = 'BAR'
    
    # Date and datetime formats
    DATE_FORMAT_SHORT = 'yyyy/mm/dd'
    DATE_FORMAT_LONG = 'yyyy/mm/dd hh:mm'
    EXCEL_DATE_FORMAT = '%Y-%m-%dT%H:%M:%S'
    EXCEL_DATE_SHORT_FORMAT = '%Y-%m-%d'


class AWRHtmlToExcelConverter:
    """
    Main converter class that turns AWR HTML spool files into Excel workbooks.

    Reads an Oracle AWR HTML spool file, builds worksheets from the spooled HTML tables,
    and applies number formats, transposed header overlays, and charts defined in the INI file.

    Attributes:
        debug (bool): Whether debug logging is enabled.
        ini_format_config (Dict): Per-sheet cell number format definitions from the INI file.
        ini_chart_config (Dict): Per-sheet chart definitions from the INI file.
        excel_format_cache (Dict): Cache of created xlsxwriter Format objects.
        current_sheet_name (str): Name of the worksheet currently being written.
    """
    
    def __init__(self, debug: bool = False):
        """
        Initialize converter state.

        Args:
            debug: Whether to enable verbose debug logging.
        """
        self.debug = debug
        
        # Parsed INI configuration
        self.ini_format_config: Dict[str, List] = {}  # Per-sheet cell format ranges
        self.ini_chart_config: Dict[str, List] = {}   # Per-sheet chart definitions
        # SQLn directives: [(index, query_section_name, target_sheet, start_row, start_col, include_title_row, direction), ...]
        self.ini_sql_config: List[Tuple[int, str, str, int, int, bool, str]] = []
        # Overlay cells per target sheet produced by transposed helper queries: {sheet: [(row0, col0, value), ...]}
        self.sheet_overlays: Dict[str, List[Tuple[int, int, str]]] = {}
        self.excel_format_cache: Dict[str, Any] = {}  # Cache of xlsxwriter Format objects
        
        # Active worksheet tracking
        self.current_sheet_name: str = ''
        self.current_header_row: List[str] = []
        
        # Excel workbook and active worksheet handles
        self.workbook: Optional[xlsxwriter.Workbook] = None
        self.current_worksheet: Optional[xlsxwriter.worksheet.Worksheet] = None
        
        # Default Excel cell format objects
        self.number_format: Optional[Any] = None
        self.string_format: Optional[Any] = None
        self.date_format: Optional[Any] = None
        self.datetime_format: Optional[Any] = None
        
        # File paths
        self.input_html_path: str = ''
        self.output_excel_path: str = ''
        self.ini_config_path: str = ''
        
        # Current cell coordinates (for debugging)
        self.current_row_index: int = 0

        # Chart X-axis (time) label rotation angle in degrees; default is vertical (-90), override via --x-rotation
        self.x_label_rotation: int = -90
        # Chart X-axis time number format; override via --x-format, or pass empty string to keep INI setting
        self.x_time_format: str = Constants.X_TIME_FORMAT_DEFAULT
        self.current_col_index: int = 0
    
    def parse_command_line_arguments(self) -> None:
        """
        Parse command-line arguments and configure input/output paths and chart options.

        CLI Options:
            -i, --input: Input AWR HTML file path (required).
            -o, --output: Output Excel file path (optional; defaults to <input>.xlsx).
            -n, --ini: INI configuration file path (optional; auto-extracted from HTML if omitted).
            --debug: Enable debug logging.

        Raises:
            SystemExit: If required arguments are missing.
        """
        parser = argparse.ArgumentParser(
            description='Convert Oracle AWR HTML report to Excel with formatting and charts',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog='''
Examples:
  %(prog)s -i awr_report.html
  %(prog)s -i awr_report.html -o output.xlsx
  %(prog)s -i awr_report.html -n config.ini --debug
            '''
        )
        
        parser.add_argument(
            '-i', '--input',
            required=True,
            help='Input AWR HTML file path'
        )
        
        parser.add_argument(
            '-o', '--output',
            help='Output Excel file path (default: input filename with .xlsx extension)'
        )
        
        parser.add_argument(
            '-n', '--ini',
            help='INI configuration file path (default: auto-extract from HTML)'
        )
        
        parser.add_argument(
            '--debug',
            action='store_true',
            help='Enable debug output for troubleshooting'
        )
        
        parser.add_argument(
            '--x-rotation',
            type=int,
            default=-90,
            metavar='DEG',
            help='Rotation of the chart X axis (time) labels in degrees, -90..90 '
                 '(default: -90 = vertical, 0 = horizontal)'
        )

        parser.add_argument(
            '--x-format',
            default=Constants.X_TIME_FORMAT_DEFAULT,
            metavar='FMT',
            help='Number format of the chart X axis time labels '
                 "(default: 'mm-dd hh:mm'; empty string keeps the INI setting)"
        )

        args = parser.parse_args()

        # Clamp X-axis label rotation to Excel's valid range (-90 .. 90)
        self.x_label_rotation = max(-90, min(90, args.x_rotation))

        # X-axis time label number format
        self.x_time_format = args.x_format

        # Required input HTML file path
        self.input_html_path = args.input
        
        # Optional output Excel file path
        if args.output:
            self.output_excel_path = args.output
        else:
            # Replace input extension with .xlsx
            base_name = os.path.splitext(self.input_html_path)[0]
            self.output_excel_path = base_name + Constants.EXCEL_EXTENSION
        
        # Optional INI configuration file path
        self.ini_config_path = args.ini if args.ini else ''
        
        # Enable debug mode if requested
        if args.debug:
            self.debug = True
    
    def _read_text(self, path: str) -> str:
        """
        Read an entire file into a string, trying candidate encodings strictly in order.

        Tries utf-8 -> cp949 -> euc-kr and returns the first decoding that succeeds without error.
        (Previously opening euc-kr first with errors='replace' caused UTF-8 files from AL32UTF8 DB
        SQL*Plus spools to be misdecoded as EUC-KR, corrupting non-ASCII text.)
        Falls back to utf-8 with errors='replace' if all strict decodings fail.

        Args:
            path: Path to the file to read.

        Returns:
            str: Decoded file contents.
        """
        with open(path, 'rb') as f:
            raw = f.read()
        for encoding in (Constants.ENCODING_UTF8, 'cp949', Constants.ENCODING_EUC_KR):
            try:
                text = raw.decode(encoding)
                if self.debug:
                    print(f"Read {path} as {encoding}")
                return text
            except UnicodeDecodeError:
                continue
        return raw.decode(Constants.ENCODING_UTF8, errors='replace')

    def extract_ini_filename_from_html(self) -> str:
        """
        Extract the source INI filename recorded inside the HTML spool file.

        Searches the HTML file for the marker line "SQL> rem INI_NAME=xxx.ini"
        and returns the referenced INI path.

        Returns:
            str: Extracted INI file path, or an empty string if not found.
        """
        try:
            text = self._read_text(self.input_html_path)
        except Exception as e:
            if self.debug:
                print(f"Error reading file: {e}")
            return ""

        for line in text.splitlines():
            match = re.match(Constants.INI_NAME_PATTERN, line)
            if match:
                ini_filename = "./" + match.group(1).strip()
                if self.debug:
                    print(f"Found INI file reference: {ini_filename}")
                return ini_filename

        return ""

    def _resolve_ini_path(self, ini_path: str) -> str:
        """
        Search fallback directories for the INI filename when the recorded path is not found.

        Because the HTML records the literal INI path used at SQL generation time (e.g. ../a.ini),
        running the converter from a different working directory may not find it directly.
        Checks candidate directories in order using the basename:
        HTML file directory -> current working directory -> script directory -> each parent directory.

        Args:
            ini_path: INI path extracted from HTML or passed on the CLI.

        Returns:
            str: Resolved existing file path, or the original path if not found anywhere.
        """
        base_name = os.path.basename(ini_path.replace('\\', '/'))
        html_dir = os.path.dirname(os.path.abspath(self.input_html_path))
        script_dir = os.path.dirname(os.path.abspath(__file__))

        search_dirs = []
        for directory in (html_dir, os.getcwd(), script_dir):
            for candidate in (directory, os.path.dirname(directory)):
                if candidate not in search_dirs:
                    search_dirs.append(candidate)

        for directory in search_dirs:
            candidate = os.path.join(directory, base_name)
            if os.path.exists(candidate):
                if self.debug:
                    print(f"Resolved INI file: {candidate}")
                return candidate

        return ini_path

    def load_and_parse_ini_configuration(self) -> None:
        """
        Load and parse SQL mapping, cell format, and chart directives from the INI file.

        INI Directive Examples:
            FORMAT: Cell number format range
                e.g. FORMAT1=sheet^[1.2:E.4]^###,##0.0

            CHART: Chart definition
                e.g. CHART1=sheet,[range],ACTIVE,sql_name,title,row,col,type,legend,x_format,x_title

        Raises:
            SystemExit: If the INI configuration file does not exist or cannot be read.
        """
        if not os.path.exists(self.ini_config_path):
            print(f"Error: INI configuration file not found: {self.ini_config_path}")
            sys.exit(1)
        
        if self.debug:
            print(f"\nParsing INI file: {self.ini_config_path}")
        
        try:
            with io.StringIO(self._read_text(self.ini_config_path)) as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    
                    # Skip blank lines and lines that do not start with SQLn=, FORMAT, or CHART
                    is_sql_line = bool(re.match(Constants.INI_SQL_PATTERN, line))
                    if not line or not (
                        is_sql_line or line.startswith(
                            (Constants.INI_FORMAT_PREFIX, Constants.INI_CHART_PREFIX)
                        )
                    ):
                        continue

                    # Strip any trailing newline characters
                    line = line.rstrip('\r\n')

                    if is_sql_line:
                        self._parse_sql_configuration(line, line_num)
                    elif line.startswith(Constants.INI_FORMAT_PREFIX):
                        self._parse_format_configuration(line, line_num)
                    elif line.startswith(Constants.INI_CHART_PREFIX):
                        self._parse_chart_configuration(line, line_num)
        
        except Exception as e:
            print(f"Error reading INI file: {e}")
            sys.exit(1)
        
        if self.debug:
            print(f"Loaded {len(self.ini_format_config)} format configurations")
            print(f"Loaded {len(self.ini_chart_config)} chart configurations")
    
    def _parse_sql_configuration(self, line: str, line_num: int) -> None:
        """
        Parse a SQLn directive line (following the original OSRAnalyzer macro rules).

        Syntax: SQLn=query_name,sheet_name,[start_row.start_col],autofit[,include_header[,direction]]
        Examples:
            SQL1=load_sql,load,[1.1],FALSE
            SQL2=load_sql2,load,[1.3],TRUE,FALSE,LTOR

        - include_header: Defaults to TRUE (writes column headers along with data).
        - direction: Defaults to T2D (top-to-down standard table layout). Any non-T2D value
          (such as LTOR) transposes the result so each record occupies one column on the sheet.

        Args:
            line: Raw SQLn directive line.
            line_num: 1-based line number in the INI file (for diagnostics).
        """
        try:
            key, value = line.split('=', 1)
            number = int(key[3:])
            parts = [p.strip() for p in value.split(',')]

            if len(parts) < 3:
                return

            title_flag = (parts[4].upper() != 'FALSE') if len(parts) > 4 else True
            direction = parts[5].upper() if len(parts) > 5 and parts[5] else Constants.DIRECTION_T2D

            query_name, sheet_name = parts[0], parts[1]
            row_str, col_str = parts[2].strip('[]').split('.')
            self.ini_sql_config.append(
                (number, query_name, sheet_name, int(row_str), int(col_str),
                 title_flag, direction)
            )

            if self.debug:
                print(f"  SQL (line {line_num}): {query_name} -> sheet={sheet_name} "
                      f"at ({row_str},{col_str}) title={title_flag} {direction}")

        except (ValueError, IndexError) as e:
            if self.debug:
                print(f"  Warning: Failed to parse SQL line {line_num}: {e}")

    def _parse_format_configuration(self, line: str, line_num: int) -> None:
        """
        Parse a FORMAT directive line.

        Syntax: FORMAT1=sheet_name^[range]^number_format
        Example: FORMAT1=load^[1.3:E.36]^###,##0

        Args:
            line: Raw FORMAT directive line.
            line_num: 1-based line number in the INI file (for diagnostics).
        """
        parts = re.split(r'[=\^]', line)
        
        if len(parts) >= 4:
            format_id = parts[0]  # FORMAT1, FORMAT2, ...
            sheet_name = parts[1]
            range_str = parts[2]
            number_format = parts[3]
            
            # Strip brackets and replace 'E' (end row) with MAX_EXCEL_ROWS
            range_str = range_str.strip('[]').replace('E', str(Constants.MAX_EXCEL_ROWS))
            
            # Parse range string into coordinate tuples
            cell_ranges = self._parse_excel_cell_range(range_str, number_format)
            
            # Store format ranges under target worksheet name
            if sheet_name not in self.ini_format_config:
                self.ini_format_config[sheet_name] = []
            self.ini_format_config[sheet_name].extend(cell_ranges)
            
            if self.debug:
                print(f"  FORMAT (line {line_num}): sheet={sheet_name}, "
                      f"range={range_str}, format={number_format}")
    
    def _parse_chart_configuration(self, line: str, line_num: int) -> None:
        """
        Parse a CHART directive line.

        Syntax: CHART1=sheet_name,[range],ACTIVE,sql_name,title,row,col,type,legend_bool,x_format,x_title
        Example: CHART1=load,[1.2:E.2/1.31:E.32],ACTIVE,load,Transactions,2,56,LINE,TRUE,dd hh:mm,Time

        Args:
            line: Raw CHART directive line.
            line_num: 1-based line number in the INI file (for diagnostics).
        """
        try:
            if '=' not in line:
                print(f"  Warning: Failed to parse CHART line {line_num} (missing '='): {line}")
                return
            parts = line.split('=', 1)[1].split(',')

            if len(parts) < 8:
                print(f"  Warning: Failed to parse CHART line {line_num} (too few fields): {line}")
                return

            sheet_name = parts[0].strip()
            range_str = parts[1].strip().strip('[]')
            active_flag = parts[2].strip() if len(parts) > 2 else 'ACTIVE'
            if active_flag.upper() != 'ACTIVE':
                if self.debug:
                    print(f"  CHART (line {line_num}): skipped non-ACTIVE ({active_flag}) on sheet={sheet_name}")
                return

            if len(parts) >= 11:
                chart_title = ','.join(parts[4:-6]).strip()
                position_row = int(parts[-6].strip())
                position_col = int(parts[-5].strip())
                chart_type = parts[-4].strip() or Constants.CHART_TYPE_LINE
                x_axis_format = parts[-2].strip()
                x_axis_title = parts[-1].strip()
            else:
                chart_title = parts[4].strip() if len(parts) > 4 else ''
                position_row = int(parts[5].strip()) if len(parts) > 5 else 2
                position_col = int(parts[6].strip()) if len(parts) > 6 else 2
                chart_type = parts[7].strip() if len(parts) > 7 else Constants.CHART_TYPE_LINE
                x_axis_format = parts[9].strip() if len(parts) > 9 else ''
                x_axis_title = parts[10].strip() if len(parts) > 10 else ''

            # Append parsed chart definition for the target worksheet
            if sheet_name not in self.ini_chart_config:
                self.ini_chart_config[sheet_name] = []

            chart_config = (
                range_str, chart_title, position_row, position_col,
                chart_type, x_axis_format, x_axis_title
            )
            self.ini_chart_config[sheet_name].append(chart_config)

            if self.debug:
                print(f"  CHART (line {line_num}): sheet={sheet_name}, "
                      f"title={chart_title}, pos=({position_row},{position_col}), "
                      f"type={chart_type}")

        except (IndexError, ValueError) as e:
            print(f"  Warning: Failed to parse CHART line {line_num}: {e} -> {line}")
    
    def _parse_excel_cell_range(self, range_str: str, format_str: str) -> List[List]:
        """
        Parse a slash-delimited range specification into 1-based coordinate lists.

        Args:
            range_str: Range specification (e.g. '1.2:E.4' or '1.2:E.4/1.5:E.6').
            format_str: Excel number format string to apply to those ranges.

        Returns:
            List[List]: Coordinate entries formatted as [[r1, c1, r2, c2, format_str], ...].

        Examples:
            >>> _parse_excel_cell_range('1.2:5.4', '###,##0')
            [[1, 2, 5, 4, '###,##0']]

            >>> _parse_excel_cell_range('1.2:5.4/1.6:5.8', '###,##0')
            [[1, 2, 5, 4, '###,##0'], [1, 6, 5, 8, '###,##0']]
        """
        result = []
        
        # Split multiple sub-ranges separated by '/'
        ranges = range_str.split('/')
        
        for range_part in ranges:
            range_part = range_part.strip()
            
            # Parse sub-range coordinates (row1.col1:row2.col2)
            parts = re.split(r'[.:]', range_part)
            
            if len(parts) == 4:
                try:
                    row1, col1, row2, col2 = map(int, parts)
                    result.append([row1, col1, row2, col2, format_str])
                except ValueError as e:
                    if self.debug:
                        print(f"Warning: Failed to parse range '{range_part}': {e}")
        
        return result
    
    def _convert_column_number_to_letter(self, col_num: int) -> str:
        """
        Convert a 1-based column number into an Excel column letter (1 -> 'A', 27 -> 'AA').

        Args:
            col_num: 1-based column number.

        Returns:
            str: Excel column letter(s) (A, B, ..., Z, AA, AB, ...).

        Examples:
            >>> _convert_column_number_to_letter(1)
            'A'
            >>> _convert_column_number_to_letter(26)
            'Z'
            >>> _convert_column_number_to_letter(27)
            'AA'
        """
        result = ''
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + ord('A')) + result
            col_num //= 26
        return result
    
    def _convert_column_index_to_letter(self, col_index: int) -> str:
        """
        Convert a 0-based column index into an Excel column letter.

        Args:
            col_index: 0-based column index.

        Returns:
            str: Excel column letter(s).

        Examples:
            >>> _convert_column_index_to_letter(0)
            'A'
            >>> _convert_column_index_to_letter(25)
            'Z'
        """
        return self._convert_column_number_to_letter(col_index + 1)
    
    def _is_numeric_string(self, text: str) -> bool:
        """
        Check whether a string represents a plain decimal number.

        Args:
            text: Candidate string to test.

        Returns:
            bool: True if the string is a valid decimal number, False otherwise.

        Note:
            Supports thousands-separated numbers (e.g. "1,234.56") while rejecting
            scientific-notation lookalikes (e.g. SQL_IDs matching "1234e56").
        """
        if not text:
            return False

        # Allow plain decimal numbers only after stripping commas (prevents SQL_IDs like '1234e56' from becoming floats/inf)
        clean_text = text.replace(',', '').strip()
        if not re.match(r'^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$', clean_text):
            return False
        try:
            float(clean_text)
            return True
        except ValueError:
            return False
    
    def _is_date_string(self, text: str) -> bool:
        """
        Check whether a string starts with an ISO date ('YYYY-MM-DD').

        Args:
            text: Candidate string to test.

        Returns:
            bool: True if the string starts with 'YYYY-MM-DD', False otherwise.

        Note:
            Matches both 'YYYY-MM-DD' and 'YYYY-MM-DD HH:MM:SS'.
        """
        if not text or len(text) < 10:
            return False
        
        # Match leading YYYY-MM-DD date pattern
        date_pattern = r'^\d{4}-\d{2}-\d{2}'
        return bool(re.match(date_pattern, text))
    
    def _convert_date_to_excel_format(self, date_str: str) -> str:
        """
        Convert a space-separated date/timestamp string into ISO-8601 'T' format.

        Args:
            date_str: Input date string (e.g. '2021-01-26 19:30:00').

        Returns:
            str: ISO-8601 formatted string (e.g. '2021-01-26T19:30:00').

        Note:
            Replaces spaces with 'T' for strptime parsing.
        """
        # Replace space between date and time with 'T'
        return date_str.replace(' ', 'T')
    
    def process_html_and_create_excel(self) -> None:
        """
        Parse the input HTML spool file and build the output Excel workbook.

        Processing Steps:
        1. Create the xlsxwriter Workbook.
        2. Register default cell formats.
        3. Extract per-section HTML tables.
        4. Resolve SQLn sheet renames and transposed helper overlays.
        5. Write each worksheet, apply formats/overlays, and embed charts.
        6. Close and save the workbook.
        """
        if self.debug:
            print(f"\nProcessing HTML file: {self.input_html_path}")
        
        # Create output Excel workbook
        try:
            self.workbook = xlsxwriter.Workbook(self.output_excel_path)
        except Exception as e:
            print(f"Error: Failed to create Excel file: {e}")
            sys.exit(1)
        
        # Initialize default cell formats
        self._define_default_excel_formats()
        
        # Read HTML spool file and populate worksheets
        try:
            html_content = self._read_html_file()
            tables = self._extract_all_tables_from_html(html_content)

            if self.debug:
                print(f"Found {len(tables)} tables in HTML")

            # Reorganize tables according to SQLn directives (sheet renaming and transposed overlays)
            tables = self._collect_transposed_sections(tables)
            
            # Write each remaining table into its own worksheet
            for sheet_name, table_html in tables.items():
                print(f"Processing sheet: [{sheet_name}]")
                
                self.current_sheet_name = sheet_name
                table_data = self._extract_table_data_from_html(table_html)
                
                if table_data:
                    self._write_table_data_to_worksheet(table_data)
        
        except Exception as e:
            print(f"Error processing HTML: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            self.workbook.close()
            sys.exit(1)
        
        # Save and close the workbook
        try:
            self.workbook.close()
            if self.debug:
                print(f"\nExcel file saved: {self.output_excel_path}")
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            sys.exit(1)
    
    def _collect_transposed_sections(self, tables: Dict[str, str]) -> Dict[str, str]:
        """
        Collect cell overlays from transposed helper queries (e.g. load_sql2) onto their
        target worksheets and remove the consumed helper tables from standalone sheet creation.

        Automates Excel's "Paste Special -> Transpose" step using the same rules as the
        original OSRAnalyzer macro (Get_Data_From_ORACLE_To_EXCEL2):
        - Only SQLn directives whose direction is not T2D (e.g. LTOR) are transposed.
        - Row 0 of the helper table contains column headers (<th>). Each data record is written
          as one column on the target sheet (fields stacked top-to-bottom) starting at [row.col].
        - If include_header is TRUE, the header names are written vertically in the starting column
          first and records start in the next column.
        - Missing source/target tables are skipped safely. Main T2D queries whose section name differs
          from their target sheet name are renamed. A single helper query may feed multiple sheets.

        Args:
            tables: Mapping of {section_name: concatenated_table_html}.

        Returns:
            Dict[str, str]: Filtered mapping excluding consumed helper sections.
        """
        self.sheet_overlays = {}
        consumed = set()
        tables = dict(tables)
        normal = Constants.DIRECTION_T2D

        # Case-insensitive lookup helper (Excel sheet names are case-insensitive)
        def find(name: str) -> Optional[str]:
            if name in tables:
                return name
            lowered = name.lower()
            return next((n for n in tables if n.lower() == lowered), None)

        # Pass 1: Rename main T2D sections when query_name differs from target sheet_name
        #         and no sheet with sheet_name exists yet.
        for number, query_name, sheet_name, row, col, title, direction in self.ini_sql_config:
            if direction != normal:
                continue
            source = find(query_name)
            if source is None or find(sheet_name) is not None:
                continue
            tables = {(sheet_name if n == source else n): h for n, h in tables.items()}
            if self.debug:
                print(f"  Rename sheet: '{source}' -> '{sheet_name}' (SQL{number})")

        # Pass 2: Collect transposed paste cells for non-T2D helper queries.
        for number, query_name, sheet_name, row, col, title, direction in self.ini_sql_config:
            if direction == normal:
                continue
            source, target = find(query_name), find(sheet_name)
            if source is None or target is None:
                if self.debug:
                    print(f"  Skip SQL{number}: '{query_name}' -> '{sheet_name}' "
                          f"(not found in HTML)")
                continue

            rows = self._extract_table_data_from_html(tables[source])
            if not rows:
                continue
            header, records = rows[0], rows[1:]
            cells = self.sheet_overlays.setdefault(target, [])
            start_col = col - 1

            if title:
                for j, name in enumerate(header):
                    cells.append((row - 1 + j, start_col, name))
                start_col += 1
            for k, record in enumerate(records):
                for j, value in enumerate(record):
                    cells.append((row - 1 + j, start_col + k, value))
            consumed.add(source)

            if self.debug:
                print(f"  Transpose: {source} ({len(records)} records x {len(header)} cols) "
                      f"-> {target} ({row},{col}) {direction}")

        # Remove consumed helper sections at the end so one helper query can serve multiple sheets
        return {name: html for name, html in tables.items() if name not in consumed}

    def _apply_sheet_overlays(self) -> None:
        """
        Write collected transposed overlay cells onto the active worksheet.

        Invoked after writing the main table data so overlay labels overwrite placeholder
        column headers at the target coordinates (which chart series names reference for legends).
        """
        for row0, col0, value in self.sheet_overlays.get(self.current_sheet_name, []):
            self.current_worksheet.write_string(row0, col0, value, self.string_format)

    def _define_default_excel_formats(self) -> None:
        """
        Register default cell formats on the active workbook:
        - Integer number format with thousands separator (###,##0)
        - Left-aligned text string format
        - Date format (yyyy/mm/dd)
        - Datetime format (yyyy/mm/dd hh:mm)
        """
        # Default numeric format with thousands separator
        self.number_format = self.workbook.add_format({
            'num_format': '###,##0',
            'align': 'right'
        })
        
        # Default text format
        self.string_format = self.workbook.add_format({
            'align': 'left'
        })
        
        # Date-only format (yyyy/mm/dd)
        self.date_format = self.workbook.add_format({
            'num_format': Constants.DATE_FORMAT_SHORT,
            'align': 'left'
        })
        
        # Date + time format (yyyy/mm/dd hh:mm)
        self.datetime_format = self.workbook.add_format({
            'num_format': Constants.DATE_FORMAT_LONG,
            'align': 'left'
        })
        
        if self.debug:
            print("Default Excel formats defined")
    
    def _read_html_file(self) -> str:
        """
        Read and decode the input HTML file (see _read_text for encoding fallback order).

        Returns:
            str: Decoded HTML file content.
        """
        return self._read_text(self.input_html_path)

    def _extract_all_tables_from_html(self, html_content: str) -> Dict[str, str]:
        """
        Extract all HTML tables grouped by section marker into a dictionary.

        Args:
            html_content: Full HTML spool text.

        Returns:
            Dict[str, str]: Mapping of {section_name: concatenated_table_html}.

        Note:
            Section names are extracted from "SQL> rem [section_name]" markers.
            Multiple <table> blocks within the same section (caused by SQL*Plus PAGESIZE breaks)
            are concatenated in order rather than overwriting earlier pages.
        """
        table_chunks: Dict[str, List[str]] = {}
        current_sheet_name = ''
        buf: List[str] = []
        in_table = False

        # Use splitlines() so CRLF '\r' characters are stripped immediately and never leak as _x000D_
        lines = html_content.splitlines()

        for line in lines:
            # Detect section marker: SQL> rem [section_name]
            sheet_match = re.search(r'SQL&gt; rem \[([^\]]+)\]', line)
            if sheet_match:
                # Flush any unfinished table buffer from the previous section
                if current_sheet_name and buf:
                    table_chunks.setdefault(current_sheet_name, []).append('\n'.join(buf))
                    buf = []

                current_sheet_name = sheet_match.group(1)
                in_table = False
                continue

            # Start of an HTML <table> block
            if Constants.TABLE_START_TAG in line:
                in_table = True
                buf = [line]
                continue

            # End of an HTML <table> block (append to section chunks to preserve multi-page tables)
            if in_table and Constants.TABLE_END_TAG in line:
                buf.append(line)
                if current_sheet_name:
                    table_chunks.setdefault(current_sheet_name, []).append('\n'.join(buf))
                buf = []
                in_table = False
                continue

            # Accumulate lines inside <table> ... </table>
            if in_table:
                buf.append(line)

        # Flush trailing table buffer if EOF was reached mid-table
        if current_sheet_name and buf:
            table_chunks.setdefault(current_sheet_name, []).append('\n'.join(buf))

        return {name: '\n'.join(chunks) for name, chunks in table_chunks.items()}

    def _extract_table_data_from_html(self, table_html: str) -> List[List[str]]:
        """
        Parse HTML table markup into a 2D list of cell strings.

        Args:
            table_html: HTML string containing one or more concatenated <table> blocks.

        Returns:
            List[List[str]]: Rows of cleaned cell strings [[col1, col2, ...], ...].

        Note:
            - Parses both <th> and <td> cells, skipping repeated <th>-only header rows
              on subsequent SQL*Plus page tables.
            - Unescapes HTML entities and strips CR ('\r') characters inside cells.
            - Empty cells become empty strings.
        """
        table_data = []

        blocks = re.findall(r'<table[^>]*>(.*?)</table>', table_html, re.DOTALL | re.IGNORECASE)
        if not blocks:
            blocks = [table_html]

        for block_idx, block_html in enumerate(blocks):
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', block_html, re.DOTALL | re.IGNORECASE)

            for row_html in rows:
                # Skip repeated column header rows (<th>-only rows) emitted on page 2+ tables
                if block_idx > 0 and re.search(r'<th\b', row_html, re.IGNORECASE) and not re.search(r'<td\b', row_html, re.IGNORECASE):
                    continue

                row_data = []

                # Extract cells (<th> or <td>) from the row
                cells = re.findall(
                    r'<t[hd][^>]*>(.*?)</t[hd]>',
                    row_html,
                    re.DOTALL | re.IGNORECASE
                )

                for cell_html in cells:
                    # Strip inner HTML tags
                    cell_text = re.sub(r'<[^>]+>', '', cell_html)

                    # Unescape HTML character entities
                    cell_text = (cell_text
                                .replace('&nbsp;', ' ')
                                .replace('&lt;', '<')
                                .replace('&gt;', '>')
                                .replace('&amp;', '&')
                                .replace('&quot;', '"')
                                .replace('&#39;', "'")
                                .replace('&apos;', "'"))

                    # Normalize intra-cell line breaks (strip '\r' to prevent Excel _x000D_ corruption)
                    cell_text = cell_text.replace('\r\n', '\n').replace('\r', '\n').strip()

                    row_data.append(cell_text)

                if row_data:  # Append non-empty rows only
                    table_data.append(row_data)

        return table_data
    
    def _write_table_data_to_worksheet(self, table_data: List[List[str]]) -> None:
        """
        Write parsed 2D table data to a new Excel worksheet, apply overlays, and add charts.

        Args:
            table_data: 2D list of cell strings (row 0 = column headers).

        Note:
            - Applies custom INI cell formats when configured.
            - Automatically converts dates, timestamps, and decimal numbers.
            - Embeds configured charts on the worksheet.
        """
        # Truncate at MAX_EXCEL_ROWS if exceeded (xlsxwriter silently drops rows past the limit)
        if len(table_data) > Constants.MAX_EXCEL_ROWS:
            print(f"  Warning: sheet [{self.current_sheet_name}] has {len(table_data)} rows; "
                  f"truncated to {Constants.MAX_EXCEL_ROWS} (xlsx limit)")
            table_data = table_data[:Constants.MAX_EXCEL_ROWS]
        
        # Add worksheet to workbook
        try:
            self.current_worksheet = self.workbook.add_worksheet(self.current_sheet_name)
        except Exception as e:
            if self.debug:
                print(f"Warning: Failed to create worksheet '{self.current_sheet_name}': {e}")
            return
        
        # Ensure xlsxwriter Format objects exist for this sheet's INI FORMAT rules
        self._create_format_objects_for_current_sheet()
        self.current_header_row = table_data[0] if table_data else []
        
        # Write rows and cells
        for row_idx, row_data in enumerate(table_data):
            self.current_row_index = row_idx
            
            for col_idx, cell_value in enumerate(row_data):
                self.current_col_index = col_idx
                
                # Check whether an INI FORMAT rule covers this cell
                has_custom_format, custom_format_key = self._check_cell_has_custom_format(
                    row_idx, col_idx
                )
                
                # Write typed cell value with the appropriate format
                self._write_cell_value(
                    row_idx, col_idx, cell_value,
                    has_custom_format, custom_format_key
                )
        
        # Apply transposed helper query overlays (SQLn LTOR) onto target cells
        self._apply_sheet_overlays()

        # Create and insert charts configured for this sheet
        self._add_charts_to_worksheet(len(table_data))
        
        if self.debug:
            print(f"  Wrote {len(table_data)} rows to worksheet '{self.current_sheet_name}'")
    
    def _create_format_objects_for_current_sheet(self) -> None:
        """
        Create and cache xlsxwriter Format objects required by the active worksheet's INI rules.

        Note:
            Format objects are cached in excel_format_cache for reuse across sheets.
        """
        if self.current_sheet_name not in self.ini_format_config:
            return
        
        for format_config in self.ini_format_config[self.current_sheet_name]:
            if len(format_config) >= 5:
                format_str = format_config[4]
                
                # Register format object if not already cached
                if format_str not in self.excel_format_cache:
                    self.excel_format_cache[format_str] = self.workbook.add_format({
                        'num_format': format_str,
                        'align': 'right'
                    })
    
    def _check_cell_has_custom_format(
        self, row: int, col: int
    ) -> Tuple[bool, str]:
        """
        Check whether a 0-based (row, col) coordinate falls within an INI FORMAT range.

        Args:
            row: 0-based row index.
            col: 0-based column index.

        Returns:
            Tuple[bool, str]: (has_format, format_string).
        """
        if self.current_sheet_name not in self.ini_format_config:
            return False, ''
        
        # Convert to 1-based coordinates used in INI FORMAT specifications
        row_1based = row + 1
        col_1based = col + 1
        
        for format_config in self.ini_format_config[self.current_sheet_name]:
            if len(format_config) >= 5:
                r1, c1, r2, c2, format_str = format_config
                
                # Check inclusive coordinate bounds
                if r1 <= row_1based <= r2 and c1 <= col_1based <= c2:
                    return True, format_str
        
        return False, ''
    
    def _write_cell_value(
        self,
        row: int,
        col: int,
        value: str,
        has_custom_format: bool,
        custom_format_key: str
    ) -> None:
        """
        Write a single cell value using the appropriate Excel data type and cell format.

        Args:
            row: 0-based row index.
            col: 0-based column index.
            value: Raw cell string extracted from HTML.
            has_custom_format: Whether a custom INI FORMAT applies to this cell.
            custom_format_key: Key into excel_format_cache when has_custom_format is True.
        """
        if not value:
            # Write a true Excel blank cell so empty/NULL values are not plotted as 0 on line charts
            self.current_worksheet.write_blank(row, col, None, self.string_format)
            return

        # Always write the header row (row 0) as text
        if row == 0:
            self.current_worksheet.write_string(row, col, value, self.string_format)
            return

        # Write full timestamps ('yyyy-mm-dd hh24:mi:ss') as Excel datetime regardless of column position
        if re.match(Constants.TIMESTAMP_PATTERN, value):
            self._write_date_value(row, col, value)
            return

        # Write date-only values ('yyyy-mm-dd') as Excel date values regardless of column position
        if re.match(r'^\d{4}-\d{2}-\d{2}$', value):
            self._write_date_value(row, col, value)
            return

        # Preserve identifier columns such as SQL_ID as strings even if all characters are digits
        if col < len(self.current_header_row):
            col_key = re.sub(r'[\s_]+', '', self.current_header_row[col]).upper()
            if col_key in {'SQLID', 'MAXQUERYSQLID', 'TOPSQLID', 'PREVSQLID'}:
                self.current_worksheet.write_string(row, col, value, self.string_format)
                return

        # Write numeric strings as numbers across all columns (including Col A: SNAP_ID, INST_ID, RANK)
        if self._is_numeric_string(value):
            try:
                num_value = float(value.replace(',', ''))
                if has_custom_format and custom_format_key in self.excel_format_cache:
                    self.current_worksheet.write_number(
                        row, col, num_value,
                        self.excel_format_cache[custom_format_key]
                    )
                else:
                    self.current_worksheet.write_number(
                        row, col, num_value,
                        self.number_format
                    )
                return
            except ValueError:
                pass

        # Fallback: write as plain text string
        self.current_worksheet.write_string(row, col, value, self.string_format)

    def _write_date_value(self, row: int, col: int, date_str: str) -> None:
        """
        Parse a date or timestamp string and write it as an Excel datetime cell.

        Args:
            row: 0-based row index.
            col: 0-based column index.
            date_str: Date ('YYYY-MM-DD') or timestamp ('YYYY-MM-DD HH:MM:SS') string.
        """
        excel_date_str = self._convert_date_to_excel_format(date_str)

        try:
            # Keep datetime format whenever an explicit time component (HH:MM:SS) is present, including midnight (00:00:00)
            has_time = len(date_str.strip()) > 10

            if has_time:
                # Date + time value
                dt_obj = datetime.strptime(
                    excel_date_str[:19],
                    Constants.EXCEL_DATE_FORMAT
                )
                self.current_worksheet.write_datetime(row, col, dt_obj, self.datetime_format)
            else:
                # Date-only value
                dt_obj = datetime.strptime(
                    excel_date_str[:10],
                    Constants.EXCEL_DATE_SHORT_FORMAT
                )
                self.current_worksheet.write_datetime(row, col, dt_obj, self.date_format)

        except ValueError:
            # Fall back to text string if date parsing fails
            self.current_worksheet.write_string(row, col, date_str, self.string_format)
    
    def _add_charts_to_worksheet(self, num_rows: int) -> None:
        """
        Create and insert all INI-configured charts for the active worksheet.

        Args:
            num_rows: Total worksheet row count including header (used to replace 'E').

        Chart Creation Steps:
        1. Look up chart definitions for the active sheet.
        2. Parse series/category ranges and substitute 'E' with num_rows.
        3. Instantiate an xlsxwriter Chart object (line, bar, etc.).
        4. Add X-axis categories and Y-axis data series.
        5. Configure chart title, axes, legend, and border.
        6. Anchor the chart on the worksheet.
        """
        # Skip if no charts are configured for this sheet or if there are no data rows
        if self.current_sheet_name not in self.ini_chart_config or num_rows < 2:
            return
        
        if self.debug:
            chart_count = len(self.ini_chart_config[self.current_sheet_name])
            print(f"  Adding {chart_count} charts to worksheet")
        
        # Build and insert each configured chart
        for idx, chart_config in enumerate(
            self.ini_chart_config[self.current_sheet_name], 1
        ):
            try:
                self._create_and_insert_chart(chart_config, num_rows, idx)
            except Exception as e:
                if self.debug:
                    print(f"  Warning: Failed to create chart {idx}: {e}")
    
    def _create_and_insert_chart(
        self,
        chart_config: Tuple,
        num_rows: int,
        chart_num: int
    ) -> None:
        """
        Instantiate, configure, and insert a single chart into the active worksheet.

        Args:
            chart_config: Parsed chart configuration tuple.
            num_rows: Total row count on the worksheet.
            chart_num: 1-based chart sequence number (for diagnostics).
        """
        # Unpack chart configuration tuple
        (range_str, title, pos_row, pos_col,
         chart_type, x_format, x_title) = chart_config
        title = self._resolve_title_vars(title)
        
        if self.debug:
            print(f"    Chart {chart_num}: {title} at ({pos_row}, {pos_col})")
        
        # Create embedded xlsxwriter chart object
        chart = self.workbook.add_chart({
            'type': chart_type.lower(),
            'embedded': True
        })
        
        # Parse range specification and attach data series
        self._add_series_to_chart(chart, range_str, num_rows)
        
        # Configure title, axis formats/rotation, and legend
        self._configure_chart_properties(chart, title, x_format, x_title)
        
        # Insert chart at 0-based anchor coordinates
        self.current_worksheet.insert_chart(pos_row - 1, pos_col - 1, chart)
    
    def _add_series_to_chart(
        self,
        chart: Any,
        range_str: str,
        num_rows: int
    ) -> None:
        """
        Parse a slash-delimited chart range specification and add series to the chart.

        Args:
            chart: Target xlsxwriter Chart object.
            range_str: Chart range specification (first segment = categories, remaining = series).
            num_rows: Total row count on the worksheet.
        """
        # Split range specification into '/' segments
        ranges = range_str.split('/')
        
        categories_range = None
        
        for idx, range_part in enumerate(ranges):
            # Substitute 'E' with the actual row count
            range_part = range_part.replace('E', str(num_rows))
            range_part = range_part.strip().rstrip(']')
            
            # Parse segment coordinates (row1.col1:row2.col2)
            parts = re.split(r'[.:]', range_part)
            
            if len(parts) != 4:
                continue
            
            try:
                r1, c1, r2, c2 = map(int, parts)
            except ValueError:
                continue
            
            # Convert 1-based INI coordinates to 0-based xlsxwriter indices
            r1 -= 1
            c1 -= 1
            r2 -= 1
            c2 -= 1
            
            if idx == 0:
                # First segment defines X-axis category labels
                categories_range = (r1, c1, r2, c2)
            else:
                # Subsequent segments define Y-axis data series columns
                self._add_data_series(chart, r1, c1, r2, c2, categories_range)
    
    def _add_data_series(
        self,
        chart: Any,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        categories_range: Optional[Tuple[int, int, int, int]]
    ) -> None:
        """
        Add one chart series per column across (c1 .. c2).

        Args:
            chart: Target xlsxwriter Chart object.
            r1, c1, r2, c2: 0-based data range coordinates.
            categories_range: Optional 0-based category range tuple.
        """
        for col in range(c1, c2 + 1):
            # Determine series legend name and starting data row
            if r1 == 0:
                # Use row 0 header cell as the series name and start values at row 1
                series_name = [self.current_sheet_name, 0, col]
                values_r1 = 1
            else:
                series_name = ''
                values_r1 = r1
            
            # Build series definition
            series_config = {
                'name': series_name,
                'values': [self.current_sheet_name, values_r1, col, r2, col]
            }
            
            # Exclude row 0 header from categories as well so category rows align with data rows
            if categories_range:
                cat_r1, cat_c1, cat_r2, cat_c2 = categories_range
                if cat_r1 == 0:
                    cat_r1 = 1
                series_config['categories'] = [
                    self.current_sheet_name,
                    cat_r1, cat_c1, cat_r2, cat_c2
                ]
            
            chart.add_series(series_config)
    
    def _resolve_title_vars(self, title: str) -> str:
        """
        Substitute @@hostname@@ and @@instance@@ placeholders in chart titles from the HTML filename.

        Reproduces the legacy DPI macro behavior (`VARIABLE5=@@hostname@@`) without Excel macros.
        Supported filename patterns:
            AWR_<host>_<inst>_<yyyymmdd_hhmi>.html (legacy)
            AWR_<NODE|RAC>_<initag>_<host>_<inst>_<yyyymmdd_hhmiss>.html
        Because <initag> may contain underscores, host and inst are taken as the last two tokens
        preceding the date/time suffix. Both hhmi (legacy) and hhmiss (current) timestamps are accepted.

        Args:
            title: Raw chart title string from the INI CHART definition.

        Returns:
            str: Title with placeholders resolved (or stripped cleanly if filename cannot be parsed).
        """
        if '@@' not in title:
            return title
        base = os.path.basename(self.input_html_path)
        m = re.match(r'(AWR_(?:(?:NODE|RAC)_)?.*)_\d{8}_\d{4}(?:\d{2})?', base)
        host, inst = '', ''
        if m:
            tokens = m.group(1).split('_')
            if len(tokens) >= 3:
                host, inst = tokens[-2], tokens[-1]
        return (title.replace('@@hostname@@', host)
                     .replace('@@instance@@', inst)
                     .replace(' -  ', ' - ').rstrip(' -'))

    def _configure_chart_properties(
        self,
        chart: Any,
        title: str,
        x_format: str,
        x_title: str
    ) -> None:
        """
        Configure chart title, X-axis formatting/rotation, legend placement, and chart border.

        Args:
            chart: Target xlsxwriter Chart object.
            title: Resolved chart title.
            x_format: X-axis number format from the INI file.
            x_title: X-axis title.
        """
        # Set chart title
        if title:
            chart.set_title({
                'name': title,
                'name_font': {'name': 'Arial', 'size': 10}
            })
        
        # Configure X-axis (title, number format, and label rotation)
        x_axis_config = {}
        if x_title:
            x_axis_config['name'] = x_title
        # Unify time axes (empty x_format or containing 'hh') to --x-format (default 'mm-dd hh:mm')
        # and force text_axis=True so 5-minute intervals are not collapsed onto day buckets by Excel.
        if self.x_time_format and (not x_format or 'hh' in x_format.lower()):
            x_axis_config['num_format'] = self.x_time_format
            x_axis_config['text_axis'] = True
        elif x_format:
            x_axis_config['num_format'] = x_format
        if self.x_label_rotation:
            # Rotate timestamp labels so they do not overlap (-90 = vertical, read bottom-to-top)
            x_axis_config['num_font'] = {'rotation': self.x_label_rotation}
        if x_axis_config:
            chart.set_x_axis(x_axis_config)
        
        # Configure legend placement
        if self.current_sheet_name == 'dbsize':
            chart.set_legend({'position': 'none'})
        else:
            chart.set_legend({'position': 'overlay_right'})
        
        # Remove outer chart border
        chart.set_chartarea({'border': {'none': True}})
    
    def run(self) -> None:
        """
        Execute the end-to-end HTML-to-Excel conversion workflow.

        Workflow:
        1. Parse CLI arguments.
        2. Determine INI file path (CLI option or auto-extracted from HTML).
        3. Resolve and verify INI file existence.
        4. Load INI SQL/FORMAT/CHART configuration.
        5. Parse HTML tables and generate the formatted Excel workbook.
        6. Print completion summary.

        Raises:
            SystemExit: On missing input files or unrecoverable conversion errors.
        """
        print("=" * 70)
        print("AWR HTML to Excel Converter")
        print("=" * 70)
        
        # 1. Parse command-line arguments
        self.parse_command_line_arguments()
        
        # 2. Determine INI configuration path
        if not self.ini_config_path:
            # Auto-extract INI filename from HTML marker when -n/--ini is omitted
            print("\nExtracting INI file name from HTML...")
            self.ini_config_path = self.extract_ini_filename_from_html()
            
            if not self.ini_config_path:
                print("Error: HTML file does not contain INI file reference!")
                print("Please specify INI file using -n option.")
                sys.exit(1)
        
        # 3. Verify INI file existence (searching fallback directories if needed)
        if not os.path.exists(self.ini_config_path):
            self.ini_config_path = self._resolve_ini_path(self.ini_config_path)

        if not os.path.exists(self.ini_config_path):
            print(f"Error: INI configuration file not found: {self.ini_config_path}")
            sys.exit(1)
        
        # Print resolved file paths
        print(f"\nInput HTML file : {self.input_html_path}")
        print(f"Output Excel file: {self.output_excel_path}")
        print(f"INI config file  : {self.ini_config_path}")
        
        # 4. Parse INI configuration
        print("\nLoading INI configuration...")
        self.load_and_parse_ini_configuration()
        
        # 5. Convert HTML tables into Excel workbook
        print("\nConverting HTML to Excel...")
        self.process_html_and_create_excel()
        
        # 6. Print completion message
        print("\n" + "=" * 70)
        print(f"[OK] Conversion completed successfully!")
        print(f"[OK] Excel file created: {self.output_excel_path}")
        print("=" * 70)


def main():
    """
    CLI entry point: instantiate AWRHtmlToExcelConverter and run it.
    """
    try:
        converter = AWRHtmlToExcelConverter()
        converter.run()
    except KeyboardInterrupt:
        print("\n\nConversion interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nUnexpected error occurred: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
