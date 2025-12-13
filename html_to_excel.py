#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AWR HTML to Excel Converter with Charts.
Oracle AWR HTML 리포트를 Excel 파일로 변환하는 프로그램.
INI 설정 파일을 기반으로 셀 형식과 차트를 자동으로 생성.

Made By, yeonhong.min@gmail.com
Version 1.0 - 2025.11.21
"""

import re
import os
import sys
import argparse
import codecs
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Any
import xlsxwriter


# 상수 정의
class Constants:
    """프로그램 전체에서 사용되는 상수 정의"""
    
    # 파일 인코딩
    ENCODING_EUC_KR = 'euc-kr'
    ENCODING_UTF8 = 'utf-8'
    
    # 파일 확장자
    EXCEL_EXTENSION = '.xlsx'
    
    # HTML 패턴
    INI_NAME_PATTERN = r'^SQL&gt; rem INI_NAME=(.*)'
    TABLE_START_TAG = '<table'
    TABLE_END_TAG = '</table>'
    ROW_TAG = '<tr'
    CELL_TAG_HEADER = '<th'
    CELL_TAG_DATA = '<td'
    
    # Excel 관련
    MAX_EXCEL_ROWS = 65535
    DEFAULT_CHART_WIDTH = 15
    DEFAULT_CHART_HEIGHT = 10
    
    # INI 파일 섹션
    INI_FORMAT_PREFIX = 'FORMAT'
    INI_CHART_PREFIX = 'CHART'
    
    # 차트 타입
    CHART_TYPE_LINE = 'LINE'
    CHART_TYPE_BAR = 'BAR'
    
    # 날짜 형식
    DATE_FORMAT_SHORT = 'yyyy/mm/dd'
    DATE_FORMAT_LONG = 'yyyy/mm/dd hh:mm'
    EXCEL_DATE_FORMAT = '%Y-%m-%dT%H:%M:%S'
    EXCEL_DATE_SHORT_FORMAT = '%Y-%m-%d'


class AWRHtmlToExcelConverter:
    """
    AWR HTML 파일을 Excel로 변환하는 메인 클래스
    
    이 클래스는 Oracle AWR HTML 리포트를 읽어서 Excel 파일로 변환하며,
    INI 설정 파일에 정의된 형식과 차트를 자동으로 적용.
    
    Attributes:
        debug (bool): 디버그 모드 활성화 여부
        ini_format_config (Dict): INI 파일에서 읽은 셀 형식 설정
        ini_chart_config (Dict): INI 파일에서 읽은 차트 설정
        excel_format_cache (Dict): Excel 형식 객체 캐시
        current_sheet_name (str): 현재 처리 중인 시트 이름
    """
    
    def __init__(self, debug: bool = False):
        """
        클래스 초기화
        
        Args:
            debug: 디버그 모드 활성화 여부
        """
        self.debug = debug
        
        # INI 파일에서 읽은 설정 저장
        self.ini_format_config: Dict[str, List] = {}  # 시트별 셀 형식 설정
        self.ini_chart_config: Dict[str, List] = {}   # 시트별 차트 설정
        self.excel_format_cache: Dict[str, Any] = {}  # Excel 형식 객체 캐시
        
        # 현재 처리 중인 시트 정보
        self.current_sheet_name: str = ''
        
        # Excel 워크북 및 워크시트 객체
        self.workbook: Optional[xlsxwriter.Workbook] = None
        self.current_worksheet: Optional[xlsxwriter.worksheet.Worksheet] = None
        
        # Excel 기본 형식 객체들
        self.number_format: Optional[Any] = None
        self.string_format: Optional[Any] = None
        self.date_format: Optional[Any] = None
        self.datetime_format: Optional[Any] = None
        
        # 파일 경로
        self.input_html_path: str = ''
        self.output_excel_path: str = ''
        self.ini_config_path: str = ''
        
        # 현재 처리 위치 (디버깅용)
        self.current_row_index: int = 0
        self.current_col_index: int = 0
    
    def parse_command_line_arguments(self) -> None:
        """
        명령행 인수를 파싱하여 입력/출력 파일명 설정
        
        명령행 옵션:
            -i, --input: 입력 HTML 파일 (필수)
            -o, --output: 출력 Excel 파일 (선택, 기본값: 입력파일명.xlsx)
            -n, --ini: INI 설정 파일 (선택, 기본값: HTML에서 자동 추출)
            --debug: 디버그 출력 활성화
        
        Raises:
            SystemExit: 필수 인수가 누락된 경우
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
        
        args = parser.parse_args()
        
        # 입력 파일 설정 (필수)
        self.input_html_path = args.input
        
        # 출력 파일 설정 (옵션)
        if args.output:
            self.output_excel_path = args.output
        else:
            # 입력 파일의 확장자를 .xlsx로 변경
            base_name = os.path.splitext(self.input_html_path)[0]
            self.output_excel_path = base_name + Constants.EXCEL_EXTENSION
        
        # INI 파일 설정 (옵션)
        self.ini_config_path = args.ini if args.ini else ''
        
        # 디버그 모드 설정
        if args.debug:
            self.debug = True
    
    def extract_ini_filename_from_html(self) -> str:
        """
        HTML 파일의 첫 번째 줄에서 INI 파일명 추출
        
        HTML 파일 내의 "SQL> rem INI_NAME=xxx.ini" 패턴을 찾아서
        INI 설정 파일명을 추출합니다.
        
        Returns:
            str: INI 파일 경로 (찾지 못한 경우 빈 문자열)
        
        Note:
            - 먼저 EUC-KR 인코딩으로 시도
            - 실패시 UTF-8 인코딩으로 재시도
        """
        ini_filename = ""
        encodings = [Constants.ENCODING_EUC_KR, Constants.ENCODING_UTF8]
        
        for encoding in encodings:
            try:
                if encoding == Constants.ENCODING_EUC_KR:
                    file_handle = codecs.open(
                        self.input_html_path,
                        'r',
                        encoding=encoding,
                        errors='replace'
                    )
                else:
                    file_handle = open(
                        self.input_html_path,
                        'r',
                        encoding=encoding,
                        errors='replace'
                    )
                
                with file_handle as f:
                    for line in f:
                        # INI 파일명 패턴 매칭
                        match = re.match(Constants.INI_NAME_PATTERN, line)
                        if match:
                            ini_filename = "./" + match.group(1).strip()
                            if self.debug:
                                print(f"Found INI file reference: {ini_filename}")
                            break
                
                if ini_filename:
                    break
                    
            except UnicodeDecodeError:
                if self.debug:
                    print(f"Failed to read with {encoding} encoding, trying next...")
                continue
            except Exception as e:
                if self.debug:
                    print(f"Error reading file with {encoding}: {e}")
                continue
        
        return ini_filename
    
    def load_and_parse_ini_configuration(self) -> None:
        """
        INI 설정 파일을 파싱하여 형식 및 차트 설정 로드
        
        INI 파일 형식:
            FORMAT: 셀 형식 지정
                예) FORMAT1=sheet^[1.2:E.4]^###,##0.0
            
            CHART: 차트 생성 설정
                예) CHART1=sheet,[range],ACTIVE,SQL명,제목,행,열,타입,불린,X형식,X제목
        
        Raises:
            SystemExit: INI 파일이 존재하지 않는 경우
        """
        if not os.path.exists(self.ini_config_path):
            print(f"Error: INI configuration file not found: {self.ini_config_path}")
            sys.exit(1)
        
        if self.debug:
            print(f"\nParsing INI file: {self.ini_config_path}")
        
        try:
            with open(self.ini_config_path, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    
                    # 빈 줄이거나 FORMAT/CHART로 시작하지 않으면 건너뛰기
                    if not line or not line.startswith(
                        (Constants.INI_FORMAT_PREFIX, Constants.INI_CHART_PREFIX)
                    ):
                        continue
                    
                    # 줄바꿈 문자 제거
                    line = line.rstrip('\r\n')
                    
                    if line.startswith(Constants.INI_FORMAT_PREFIX):
                        self._parse_format_configuration(line, line_num)
                    elif line.startswith(Constants.INI_CHART_PREFIX):
                        self._parse_chart_configuration(line, line_num)
        
        except Exception as e:
            print(f"Error reading INI file: {e}")
            sys.exit(1)
        
        if self.debug:
            print(f"Loaded {len(self.ini_format_config)} format configurations")
            print(f"Loaded {len(self.ini_chart_config)} chart configurations")
    
    def _parse_format_configuration(self, line: str, line_num: int) -> None:
        """
        FORMAT 설정 라인 파싱
        
        형식: FORMAT1=시트명^[범위]^숫자형식
        예: FORMAT1=load^[1.3:E.36]^###,##0
        
        Args:
            line: FORMAT 설정 라인
            line_num: 라인 번호 (디버깅용)
        """
        parts = re.split(r'[=\^]', line)
        
        if len(parts) >= 4:
            format_id = parts[0]  # FORMAT1, FORMAT2, ...
            sheet_name = parts[1]
            range_str = parts[2]
            number_format = parts[3]
            
            # 대괄호 제거하고 'E'를 최대 행 번호로 변경
            range_str = range_str.strip('[]').replace('E', str(Constants.MAX_EXCEL_ROWS))
            
            # 범위를 파싱하여 좌표 리스트로 변환
            cell_ranges = self._parse_excel_cell_range(range_str, number_format)
            
            # 시트별로 형식 설정 저장
            if sheet_name not in self.ini_format_config:
                self.ini_format_config[sheet_name] = []
            self.ini_format_config[sheet_name].extend(cell_ranges)
            
            if self.debug:
                print(f"  FORMAT (line {line_num}): sheet={sheet_name}, "
                      f"range={range_str}, format={number_format}")
    
    def _parse_chart_configuration(self, line: str, line_num: int) -> None:
        """
        CHART 설정 라인 파싱
        
        형식: CHART1=시트명,[범위],ACTIVE,SQL명,제목,행,열,타입,불린,X형식,X제목
        예: CHART1=load,[1.2:E.2/1.31:E.32],ACTIVE,load,Transactions,2,56,LINE,TRUE,dd hh:mm,Time
        
        Args:
            line: CHART 설정 라인
            line_num: 라인 번호 (디버깅용)
        """
        try:
            parts = line.split('=', 1)[1].split(',')
            
            if len(parts) >= 8:
                sheet_name = parts[0]
                range_str = parts[1].strip('[]')
                # active_flag = parts[2] if len(parts) > 2 else 'ACTIVE'  # 현재 미사용
                # sql_name = parts[3] if len(parts) > 3 else ''  # 현재 미사용
                chart_title = parts[4] if len(parts) > 4 else ''
                position_row = int(parts[5]) if len(parts) > 5 else 2
                position_col = int(parts[6]) if len(parts) > 6 else 2
                chart_type = parts[7] if len(parts) > 7 else Constants.CHART_TYPE_LINE
                x_axis_format = parts[9] if len(parts) > 9 else ''
                x_axis_title = parts[10] if len(parts) > 10 else ''
                
                # 시트별로 차트 설정 저장
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
            if self.debug:
                print(f"  Warning: Failed to parse CHART line {line_num}: {e}")
    
    def _parse_excel_cell_range(self, range_str: str, format_str: str) -> List[List]:
        """
        범위 문자열을 파싱하여 Excel 좌표 리스트로 변환
        
        Args:
            range_str: 범위 문자열 (예: '1.2:E.4' 또는 '1.2:E.4/1.5:E.6')
            format_str: 적용할 Excel 숫자 형식 문자열
        
        Returns:
            List[List]: [[r1, c1, r2, c2, format_str], ...] 형태의 좌표 리스트
        
        Examples:
            >>> _parse_excel_cell_range('1.2:5.4', '###,##0')
            [[1, 2, 5, 4, '###,##0']]
            
            >>> _parse_excel_cell_range('1.2:5.4/1.6:5.8', '###,##0')
            [[1, 2, 5, 4, '###,##0'], [1, 6, 5, 8, '###,##0']]
        """
        result = []
        
        # '/'로 구분된 여러 범위 처리
        ranges = range_str.split('/')
        
        for range_part in ranges:
            range_part = range_part.strip()
            
            # 범위를 좌표로 파싱 (행1.열1:행2.열2)
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
        열 번호를 Excel 열 문자로 변환 (1 -> 'A', 27 -> 'AA')
        
        Args:
            col_num: 열 번호 (1부터 시작)
        
        Returns:
            str: Excel 열 문자 (A, B, ..., Z, AA, AB, ...)
        
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
        0 기반 열 인덱스를 Excel 열 문자로 변환
        
        Args:
            col_index: 열 인덱스 (0부터 시작)
        
        Returns:
            str: Excel 열 문자
        
        Examples:
            >>> _convert_column_index_to_letter(0)
            'A'
            >>> _convert_column_index_to_letter(25)
            'Z'
        """
        return self._convert_column_number_to_letter(col_index + 1)
    
    def _is_numeric_string(self, text: str) -> bool:
        """
        문자열이 숫자를 나타내는지 확인
        
        Args:
            text: 확인할 문자열
        
        Returns:
            bool: 숫자 문자열 여부
        
        Note:
            쉼표(,)가 포함된 숫자도 처리 (예: "1,234.56")
        """
        if not text:
            return False
        
        # 쉼표 제거 후 숫자 변환 시도
        clean_text = text.replace(',', '')
        try:
            float(clean_text)
            return True
        except ValueError:
            return False
    
    def _is_date_string(self, text: str) -> bool:
        """
        문자열이 날짜 형식인지 확인
        
        Args:
            text: 확인할 문자열
        
        Returns:
            bool: 날짜 문자열 여부
        
        Note:
            'YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM:SS' 형식 지원
        """
        if not text or len(text) < 10:
            return False
        
        # 날짜 패턴 매칭: YYYY-MM-DD 형식
        date_pattern = r'^\d{4}-\d{2}-\d{2}'
        return bool(re.match(date_pattern, text))
    
    def _convert_date_to_excel_format(self, date_str: str) -> str:
        """
        날짜 문자열을 Excel 호환 형식으로 변환
        
        Args:
            date_str: 입력 날짜 문자열 (예: '2021-01-26 19:30:00')
        
        Returns:
            str: Excel 날짜 형식 (예: '2021-01-26T19:30:00')
        
        Note:
            공백을 'T'로 변환하여 ISO 8601 형식으로 만듭니다.
        """
        # 날짜와 시간 사이의 공백을 'T'로 변경
        return date_str.replace(' ', 'T')
    
    def process_html_and_create_excel(self) -> None:
        """
        HTML 파일을 처리하여 Excel 워크북 생성
        
        주요 처리 과정:
        1. Excel 워크북 생성
        2. 기본 형식 정의
        3. HTML에서 테이블 추출
        4. 각 테이블을 워크시트로 변환
        5. 형식 및 차트 적용
        6. 워크북 저장
        """
        if self.debug:
            print(f"\nProcessing HTML file: {self.input_html_path}")
        
        # Excel 워크북 생성
        try:
            self.workbook = xlsxwriter.Workbook(self.output_excel_path)
        except Exception as e:
            print(f"Error: Failed to create Excel file: {e}")
            sys.exit(1)
        
        # 기본 Excel 형식 정의
        self._define_default_excel_formats()
        
        # HTML 파일 읽기 및 처리
        try:
            html_content = self._read_html_file()
            tables = self._extract_all_tables_from_html(html_content)
            
            if self.debug:
                print(f"Found {len(tables)} tables in HTML")
            
            # 각 테이블을 워크시트로 변환
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
        
        # 워크북 저장
        try:
            self.workbook.close()
            if self.debug:
                print(f"\nExcel file saved: {self.output_excel_path}")
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            sys.exit(1)
    
    def _define_default_excel_formats(self) -> None:
        """
        Excel에서 사용할 기본 형식 정의
        
        정의되는 형식:
        - 숫자 형식 (###,##0)
        - 문자열 형식 (텍스트)
        - 날짜 형식 (yyyy/mm/dd)
        - 날짜+시간 형식 (yyyy/mm/dd hh:mm)
        """
        # 숫자 형식: 천 단위 구분 기호 사용
        self.number_format = self.workbook.add_format({
            'num_format': '###,##0',
            'align': 'right'
        })
        
        # 문자열 형식
        self.string_format = self.workbook.add_format({
            'align': 'left'
        })
        
        # 날짜 형식 (yyyy/mm/dd)
        self.date_format = self.workbook.add_format({
            'num_format': Constants.DATE_FORMAT_SHORT,
            'align': 'left'
        })
        
        # 날짜+시간 형식 (yyyy/mm/dd hh:mm)
        self.datetime_format = self.workbook.add_format({
            'num_format': Constants.DATE_FORMAT_LONG,
            'align': 'left'
        })
        
        if self.debug:
            print("Default Excel formats defined")
    
    def _read_html_file(self) -> str:
        """
        HTML 파일을 읽어서 문자열로 반환
        
        Returns:
            str: HTML 파일 내용
        
        Raises:
            Exception: 파일 읽기 실패 시
        """
        encodings = [Constants.ENCODING_EUC_KR, Constants.ENCODING_UTF8]
        
        for encoding in encodings:
            try:
                if encoding == Constants.ENCODING_EUC_KR:
                    with codecs.open(
                        self.input_html_path,
                        'r',
                        encoding=encoding,
                        errors='replace'
                    ) as f:
                        return f.read()
                else:
                    with open(
                        self.input_html_path,
                        'r',
                        encoding=encoding,
                        errors='replace'
                    ) as f:
                        return f.read()
            except UnicodeDecodeError:
                if self.debug:
                    print(f"Failed to read with {encoding} encoding, trying next...")
                continue
        
        raise Exception("Failed to read HTML file with supported encodings")
    
    def _extract_all_tables_from_html(self, html_content: str) -> Dict[str, str]:
        """
        HTML에서 모든 테이블을 추출하여 딕셔너리로 반환
        
        Args:
            html_content: HTML 파일 내용
        
        Returns:
            Dict[str, str]: {시트명: 테이블HTML} 형태의 딕셔너리
        
        Note:
            시트명은 "SQL> rem [시트명]" 패턴에서 추출
        """
        tables = {}
        current_sheet_name = ''
        table_html = ''
        in_table = False
        
        lines = html_content.split('\n')
        
        for line in lines:
            # 시트 이름 찾기: SQL> rem [시트명]
            sheet_match = re.search(r'SQL&gt; rem \[([^\]]+)\]', line)
            if sheet_match:
                # 이전 테이블이 있으면 저장
                if current_sheet_name and table_html:
                    tables[current_sheet_name] = table_html
                    table_html = ''
                
                current_sheet_name = sheet_match.group(1)
                in_table = False
                continue
            
            # 테이블 시작
            if Constants.TABLE_START_TAG in line:
                in_table = True
                table_html = line + '\n'
                continue
            
            # 테이블 종료
            if in_table and Constants.TABLE_END_TAG in line:
                table_html += line + '\n'
                if current_sheet_name:
                    tables[current_sheet_name] = table_html
                table_html = ''
                in_table = False
                continue
            
            # 테이블 내용 수집
            if in_table:
                table_html += line + '\n'
        
        # 마지막 테이블 처리
        if current_sheet_name and table_html:
            tables[current_sheet_name] = table_html
        
        return tables
    
    def _extract_table_data_from_html(self, table_html: str) -> List[List[str]]:
        """
        HTML 테이블에서 데이터를 추출하여 2차원 리스트로 반환
        
        Args:
            table_html: 테이블 HTML 문자열
        
        Returns:
            List[List[str]]: [[셀1, 셀2, ...], [셀1, 셀2, ...], ...] 형태의 데이터
        
        Note:
            - <th> 태그와 <td> 태그를 모두 셀로 처리
            - HTML 엔티티를 일반 문자로 변환
            - 빈 셀은 빈 문자열로 처리
        """
        table_data = []
        
        # HTML에서 행 추출
        rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL | re.IGNORECASE)
        
        for row_html in rows:
            row_data = []
            
            # 행에서 셀 추출 (th 또는 td 태그)
            cells = re.findall(
                r'<t[hd][^>]*>(.*?)</t[hd]>',
                row_html,
                re.DOTALL | re.IGNORECASE
            )
            
            for cell_html in cells:
                # HTML 태그 제거
                cell_text = re.sub(r'<[^>]+>', '', cell_html)
                
                # HTML 엔티티 변환
                cell_text = (cell_text
                            .replace('&nbsp;', ' ')
                            .replace('&lt;', '<')
                            .replace('&gt;', '>')
                            .replace('&amp;', '&')
                            .replace('&quot;', '"'))
                
                # 앞뒤 공백 제거
                cell_text = cell_text.strip()
                
                row_data.append(cell_text)
            
            if row_data:  # 빈 행이 아니면 추가
                table_data.append(row_data)
        
        return table_data
    
    def _write_table_data_to_worksheet(self, table_data: List[List[str]]) -> None:
        """
        테이블 데이터를 Excel 워크시트에 작성
        
        Args:
            table_data: 2차원 리스트 형태의 테이블 데이터
        
        Note:
            - INI 설정에 따라 셀 형식 자동 적용
            - 숫자/날짜 자동 감지 및 형식 적용
            - 차트 자동 생성
        """
        # 워크시트 생성
        try:
            self.current_worksheet = self.workbook.add_worksheet(self.current_sheet_name)
        except Exception as e:
            if self.debug:
                print(f"Warning: Failed to create worksheet '{self.current_sheet_name}': {e}")
            return
        
        # INI 설정에 정의된 형식 객체 생성
        self._create_format_objects_for_current_sheet()
        
        # 데이터 쓰기
        for row_idx, row_data in enumerate(table_data):
            self.current_row_index = row_idx
            
            for col_idx, cell_value in enumerate(row_data):
                self.current_col_index = col_idx
                
                # 셀에 적용할 형식 확인
                has_custom_format, custom_format_key = self._check_cell_has_custom_format(
                    row_idx, col_idx
                )
                
                # 셀 값 쓰기 (형식에 따라 처리)
                self._write_cell_value(
                    row_idx, col_idx, cell_value,
                    has_custom_format, custom_format_key
                )
        
        # 차트 추가
        self._add_charts_to_worksheet(len(table_data))
        
        if self.debug:
            print(f"  Wrote {len(table_data)} rows to worksheet '{self.current_sheet_name}'")
    
    def _create_format_objects_for_current_sheet(self) -> None:
        """
        현재 시트에 대한 INI 형식 설정을 Excel 형식 객체로 생성
        
        Note:
            형식 객체는 excel_format_cache에 캐싱되어 재사용됩니다.
        """
        if self.current_sheet_name not in self.ini_format_config:
            return
        
        for format_config in self.ini_format_config[self.current_sheet_name]:
            if len(format_config) >= 5:
                format_str = format_config[4]
                
                # 이미 생성된 형식이 아니면 새로 생성
                if format_str not in self.excel_format_cache:
                    self.excel_format_cache[format_str] = self.workbook.add_format({
                        'num_format': format_str,
                        'align': 'right'
                    })
    
    def _check_cell_has_custom_format(
        self, row: int, col: int
    ) -> Tuple[bool, str]:
        """
        특정 셀에 INI 설정 형식이 적용되는지 확인
        
        Args:
            row: 행 인덱스 (0부터 시작)
            col: 열 인덱스 (0부터 시작)
        
        Returns:
            Tuple[bool, str]: (형식 적용 여부, 형식 키)
        """
        if self.current_sheet_name not in self.ini_format_config:
            return False, ''
        
        # 1-based 인덱스로 변환 (INI 파일은 1부터 시작)
        row_1based = row + 1
        col_1based = col + 1
        
        for format_config in self.ini_format_config[self.current_sheet_name]:
            if len(format_config) >= 5:
                r1, c1, r2, c2, format_str = format_config
                
                # 범위 내에 있는지 확인
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
        셀에 값을 쓰기 (타입에 따라 적절한 형식 적용)
        
        Args:
            row: 행 인덱스
            col: 열 인덱스
            value: 셀 값
            has_custom_format: 사용자 정의 형식 적용 여부
            custom_format_key: 사용자 정의 형식 키
        """
        if not value:
            # 빈 셀
            self.current_worksheet.write_string(row, col, '', self.string_format)
            return
        
        # 사용자 정의 형식이 있는 경우
        if has_custom_format and custom_format_key in self.excel_format_cache:
            custom_format = self.excel_format_cache[custom_format_key]
            
            # 숫자로 변환 시도
            if self._is_numeric_string(value):
                try:
                    num_value = float(value.replace(',', ''))
                    self.current_worksheet.write_number(row, col, num_value, custom_format)
                    return
                except ValueError:
                    pass
        
        # 첫 번째 열의 날짜 확인
        if col == 0 and self._is_date_string(value):
            self._write_date_value(row, col, value)
            return
        
        # 일반 숫자 확인
        if col > 0 and self._is_numeric_string(value):
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
        
        # 기본: 문자열로 처리
        self.current_worksheet.write_string(row, col, value, self.string_format)
    
    def _write_date_value(self, row: int, col: int, date_str: str) -> None:
        """
        날짜 값을 Excel 날짜 형식으로 쓰기
        
        Args:
            row: 행 인덱스
            col: 열 인덱스
            date_str: 날짜 문자열
        """
        excel_date_str = self._convert_date_to_excel_format(date_str)
        
        try:
            # 시간 정보가 있는지 확인
            has_time = len(date_str) > 10 and not date_str.endswith(' 00:00:00')
            
            if has_time:
                # 날짜+시간 형식
                dt_obj = datetime.strptime(
                    excel_date_str[:19],
                    Constants.EXCEL_DATE_FORMAT
                )
                self.current_worksheet.write_datetime(row, col, dt_obj, self.datetime_format)
            else:
                # 날짜만
                dt_obj = datetime.strptime(
                    excel_date_str[:10],
                    Constants.EXCEL_DATE_SHORT_FORMAT
                )
                self.current_worksheet.write_datetime(row, col, dt_obj, self.date_format)
        
        except ValueError:
            # 날짜 파싱 실패시 문자열로 처리
            self.current_worksheet.write_string(row, col, date_str, self.string_format)
    
    def _add_charts_to_worksheet(self, num_rows: int) -> None:
        """
        INI 설정에 따라 차트를 생성하고 워크시트에 추가
        
        Args:
            num_rows: 테이블의 전체 행 수 ('E'를 실제 행 수로 치환용)
        
        차트 생성 과정:
        1. INI 파일의 차트 설정 읽기
        2. 데이터 범위 파싱 및 'E'를 실제 행 수로 치환
        3. Excel 차트 객체 생성 (선형, 막대 등)
        4. 데이터 시리즈 추가 (X축, Y축 데이터)
        5. 차트 속성 설정 (제목, 축, 범례)
        6. 워크시트에 차트 삽입
        """
        # 현재 시트에 차트 설정이 없으면 종료
        if self.current_sheet_name not in self.ini_chart_config:
            return
        
        if self.debug:
            chart_count = len(self.ini_chart_config[self.current_sheet_name])
            print(f"  Adding {chart_count} charts to worksheet")
        
        # 각 차트 설정에 대해 처리
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
        개별 차트를 생성하고 워크시트에 삽입
        
        Args:
            chart_config: 차트 설정 튜플
            num_rows: 총 행 수
            chart_num: 차트 번호 (디버깅용)
        """
        # 차트 설정 언패킹
        (range_str, title, pos_row, pos_col,
         chart_type, x_format, x_title) = chart_config
        
        if self.debug:
            print(f"    Chart {chart_num}: {title} at ({pos_row}, {pos_col})")
        
        # Excel 차트 객체 생성
        chart = self.workbook.add_chart({
            'type': chart_type.lower(),
            'embedded': True
        })
        
        # 데이터 범위 파싱 및 시리즈 추가
        self._add_series_to_chart(chart, range_str, num_rows)
        
        # 차트 속성 설정
        self._configure_chart_properties(chart, title, x_format, x_title)
        
        # 워크시트에 차트 삽입 (1-based를 0-based로 변환)
        self.current_worksheet.insert_chart(pos_row - 1, pos_col - 1, chart)
    
    def _add_series_to_chart(
        self,
        chart: Any,
        range_str: str,
        num_rows: int
    ) -> None:
        """
        차트에 데이터 시리즈 추가
        
        Args:
            chart: Excel 차트 객체
            range_str: 데이터 범위 문자열
            num_rows: 총 행 수
        """
        # 데이터 범위 파싱 (여러 범위를 '/'로 구분)
        ranges = range_str.split('/')
        
        categories_range = None
        
        for idx, range_part in enumerate(ranges):
            # 'E'를 실제 행 수로 치환
            range_part = range_part.replace('E', str(num_rows))
            range_part = range_part.strip().rstrip(']')
            
            # 범위를 좌표로 파싱 (행1.열1:행2.열2)
            parts = re.split(r'[.:]', range_part)
            
            if len(parts) != 4:
                continue
            
            try:
                r1, c1, r2, c2 = map(int, parts)
            except ValueError:
                continue
            
            # 0-based 인덱스로 조정
            r1 -= 1
            c1 -= 1
            r2 -= 1
            c2 -= 1
            
            if idx == 0:
                # 첫 번째 범위는 카테고리(X축) 데이터
                categories_range = (r1, c1, r2, c2)
            else:
                # 이후 범위는 데이터 시리즈
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
        차트에 개별 데이터 시리즈 추가
        
        Args:
            chart: Excel 차트 객체
            r1, c1, r2, c2: 데이터 범위 좌표
            categories_range: 카테고리 범위 좌표
        """
        for col in range(c1, c2 + 1):
            # 시리즈 이름 설정
            if r1 == 0:
                # 첫 행을 시리즈 이름으로 사용
                series_name = [self.current_sheet_name, 0, col]
                values_r1 = 1
            else:
                series_name = ''
                values_r1 = r1
            
            # 시리즈 추가
            series_config = {
                'name': series_name,
                'values': [self.current_sheet_name, values_r1, col, r2, col]
            }
            
            # 카테고리 범위 추가
            if categories_range:
                cat_r1, cat_c1, cat_r2, cat_c2 = categories_range
                series_config['categories'] = [
                    self.current_sheet_name,
                    cat_r1, cat_c1, cat_r2, cat_c2
                ]
            
            chart.add_series(series_config)
    
    def _configure_chart_properties(
        self,
        chart: Any,
        title: str,
        x_format: str,
        x_title: str
    ) -> None:
        """
        차트 속성 설정 (제목, 축, 범례 등)
        
        Args:
            chart: Excel 차트 객체
            title: 차트 제목
            x_format: X축 숫자 형식
            x_title: X축 제목
        """
        # 차트 제목 설정
        if title:
            chart.set_title({
                'name': title,
                'name_font': {'name': 'Arial', 'size': 10}
            })
        
        # X축 설정
        if x_title or x_format:
            x_axis_config = {}
            if x_title:
                x_axis_config['name'] = x_title
            if x_format:
                x_axis_config['num_format'] = x_format
            chart.set_x_axis(x_axis_config)
        
        # 범례 위치 설정
        if self.current_sheet_name == 'dbsize':
            chart.set_legend({'position': 'none'})
        else:
            chart.set_legend({'position': 'overlay_right'})
        
        # 차트 테두리 제거
        chart.set_chartarea({'border': {'none': True}})
    
    def run(self) -> None:
        """
        변환 프로세스 전체를 실행하는 메인 메서드
        
        실행 순서:
        1. 명령행 인수 파싱
        2. INI 파일 경로 확인 (자동 추출 또는 지정)
        3. INI 파일 존재 여부 확인
        4. INI 설정 파싱
        5. HTML 처리 및 Excel 생성
        6. 완료 메시지 출력
        
        Raises:
            SystemExit: 오류 발생 시
        """
        print("=" * 70)
        print("AWR HTML to Excel Converter")
        print("=" * 70)
        
        # 1. 명령행 인수 파싱
        self.parse_command_line_arguments()
        
        # 2. INI 파일 경로 확인
        if not self.ini_config_path:
            # INI 파일이 지정되지 않았으면 HTML에서 자동 추출
            print("\nExtracting INI file name from HTML...")
            self.ini_config_path = self.extract_ini_filename_from_html()
            
            if not self.ini_config_path:
                print("Error: HTML file does not contain INI file reference!")
                print("Please specify INI file using -n option.")
                sys.exit(1)
        
        # 3. INI 파일 존재 여부 확인
        if not os.path.exists(self.ini_config_path):
            print(f"Error: INI configuration file not found: {self.ini_config_path}")
            sys.exit(1)
        
        # 파일 정보 출력
        print(f"\nInput HTML file : {self.input_html_path}")
        print(f"Output Excel file: {self.output_excel_path}")
        print(f"INI config file  : {self.ini_config_path}")
        
        # 4. INI 설정 파싱
        print("\nLoading INI configuration...")
        self.load_and_parse_ini_configuration()
        
        # 5. HTML 처리 및 Excel 생성
        print("\nConverting HTML to Excel...")
        self.process_html_and_create_excel()
        
        # 6. 완료 메시지
        print("\n" + "=" * 70)
        print(f"[OK] Conversion completed successfully!")
        print(f"[OK] Excel file created: {self.output_excel_path}")
        print("=" * 70)


def main():
    """
    프로그램 진입점
    
    AWRHtmlToExcelConverter 인스턴스를 생성하고 실행합니다.
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
