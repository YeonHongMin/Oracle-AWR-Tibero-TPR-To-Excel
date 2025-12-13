#!/usr/bin/env python3
"""
init_to_sql.py - INI 파일을 SQL 파일로 변환하는 프로그램
AWR(Automatic Workload Repository) 설정 INI 파일을 Oracle SQL 스크립트로 변환

Made By, yeonhong.min@gmail.com
Version 1.0 2025.11.21
"""

import sys
import os
import re
import argparse

# 디버그 모드 플래그 (True로 설정하면 디버그 정보 출력)
debug = False

def write_sql_header(outfile, ini_file):
    """
    SQL 파일 헤더 작성 함수
    
    Parameters:
    -----------
    outfile : file object
        출력할 SQL 파일 객체
    ini_file : str
        입력 INI 파일명 (메타 정보 출력에 사용)
    
    설명:
    -----
    - AWR 리포트 생성을 위한 SQL 헤더를 작성
    - 변수 선언, 날짜 설정, 스냅샷 ID 조회 등의 초기 설정 포함
    """
    
    # 시간값 입력 헤더
    header_common = """Prompt ############### Define Date : YYYYMMDDHH24MI #################################
var snapdate_from number
exec :snapdate_from:=202601010000

var snapdate_to number
exec :snapdate_to:=202601150000
Prompt ##############################################################################

var inst_no number
var snap_fr number
var snap_to number
var inst_no1 number
var inst_no2 number
var inst_no3 number
var inst_no4 number
var inst_no5 number
var inst_no6 number
var dbid number

exec select instance_number into :inst_no from v$instance;
exec SELECT NVL(MAX(SNAP_ID),1)+1 into :snap_fr FROM DBA_HIST_SNAPSHOT WHERE BEGIN_INTERVAL_TIME<=TO_DATE(:snapdate_from,'YYYYMMDDHH24MI') and dbid in (select dbid  from v$database);
exec SELECT NVL(MAX(SNAP_ID),1)+1 SNAP_TO into :snap_to FROM DBA_HIST_SNAPSHOT WHERE END_INTERVAL_TIME<=TO_DATE(:snapdate_to,'YYYYMMDDHH24MI') and dbid in (select dbid  from v$database);
exec SELECT dbid into :dbid from v$database;
exec :inst_no1:=1
exec :inst_no2:=2
exec :inst_no3:=3
exec :inst_no4:=4
exec :inst_no5:=5
exec :inst_no6:=6

print :snap_fr
print :snap_to
print :dbid
print :inst_no

set markup html on
set linesize 32767 trimspool on
set pages 50000
set feedback off
set termout off
set echo off"""
    
    # Spool 파일 이름 헤더
    header_specific = """				
prompt SQL>
set linesize 1000 pagesize 9999 trimspool on
set termout off time off
set sqlprompt "SQL> "
alter session set nls_date_format='RR/MM/DD HH24:MI';
column report_name   new_value report_name   format a30;
select 'AWR_' || host_name || '_' || instance_name|| '_'|| to_char(sysdate, 'yyyymmdd_hh24mi') || '.html' report_name from v$instance;

set echo on
set markup html on

spool &report_name
rem INI_NAME=""" + ini_file
    
    # 헤더를 파일에 작성
    outfile.write(header_common + header_specific + '\n')

def write_sql_footer(outfile):
    """
    SQL 파일 종료 부분 작성 함수
    
    Parameters:
    -----------
    outfile : file object
        출력할 SQL 파일 객체
    
    설명:
    -----
    SQL*Plus 세션을 종료하는 exit 명령어 추가
    """
    outfile.write("exit;\n")

def write_section_sql(outfile, section, sql_lines):
    """
    SQL 구문을 파일에 작성하는 함수
    
    Parameters:
    -----------
    outfile : file object
        출력할 SQL 파일 객체
    section : str
        현재 섹션명 (INI 파일의 [section] 부분)
    sql_lines : list
        SQL 구문 라인들의 리스트
    
    설명:
    -----
    - SELECT 문인 경우에만 처리
    - 섹션명을 주석(rem)으로 추가
    - SQL 구문 마지막에 세미콜론(;) 추가
    """
    if sql_lines:
        first_line = sql_lines[0]
        
        # 첫 번째 라인에 'select'가 포함되어 있는지 확인 (대소문자 무시)
        if re.search(r'select', first_line, re.IGNORECASE):
            # 섹션명을 주석으로 추가
            outfile.write(f"rem [{section}]\n")
            # 첫 번째 라인 작성
            outfile.write(first_line + "\n")
            # 나머지 라인들 작성
            if len(sql_lines) > 1:
                outfile.write("\n".join(sql_lines[1:]) + "\n")
            # SQL 구문 종료 세미콜론 추가
            outfile.write(";\n")

def convert_ini_to_sql(infile_name, outfile_name):
    """
    INI 파일을 파싱하여 SQL 파일 생성하는 메인 처리 함수
    
    Parameters:
    -----------
    infile_name : str
        입력 INI 파일 경로
    outfile_name : str
        출력 SQL 파일 경로
    
    처리 과정:
    ----------
    1. INI 파일을 한 줄씩 읽기
    2. 주석(#)과 빈 줄 무시
    3. [섹션명] 형태의 섹션 헤더 감지
    4. SQL 구문 수집 및 섹션별로 처리
    5. SQL 파일로 출력
    
    INI 파일 형식:
    -------------
    [섹션명]
    SQL 구문...
    
    [다음섹션]
    SQL 구문...
    """
    # SQL 라인들을 임시 저장할 리스트
    sql_lines = []
    # 현재 처리 중인 섹션명
    section = ""
    
    # 입력 파일을 UTF-8 인코딩으로 열기
    with open(infile_name, 'r', encoding='utf-8') as ini_file:
        # 출력 파일을 UTF-8 인코딩으로 열기
        with open(outfile_name, 'w', encoding='utf-8') as outfile:
            # SQL 파일 헤더 작성
            write_sql_header(outfile, infile_name)
            
            # INI 파일을 한 줄씩 처리
            for line in ini_file:
                # 줄 끝의 개행 문자(\n)와 캐리지 리턴(\r) 제거
                line = line.rstrip('\n\r')
                # 윈도우 형식의 캐리지 리턴 문자(0x0D) 완전 제거
                line = line.replace('\x0D', '')
                
                # === 무시할 라인들 처리 ===
                
                # '#'로 시작하는 주석 라인 무시
                if line.startswith('#'):
                    continue
                # '[#'로 시작하는 비활성 섹션 무시
                if line.startswith('[#'):
                    continue
                # 공백만 있는 라인 무시
                if re.match(r'^\s*$', line):
                    continue
                # 빈 라인 무시
                if line == "":
                    continue
                
                # === 섹션 헤더 처리 ===
                # [섹션명] 형태의 패턴 매칭 (\w+는 알파벳, 숫자, 언더스코어)
                section_match = re.match(r'^\[(\w+)\]', line)
                if section_match:
                    # 이전 섹션의 SQL이 있으면 파일에 작성
                    if sql_lines:
                        write_section_sql(outfile, section, sql_lines)
                        sql_lines = []  # 리스트 초기화
                    # 새로운 섹션명 저장
                    section = section_match.group(1)
                else:
                    # === SQL 라인 처리 ===
                    # 섹션 헤더가 아닌 모든 라인은 SQL 구문으로 간주
                    sql_lines.append(line)
                    # 디버그 모드일 때 추가된 라인 출력
                    if debug:
                        print(f"push ({line})")
            
            # 파일 끝에 도달했을 때 남은 SQL 라인들 처리
            if sql_lines:
                write_section_sql(outfile, section, sql_lines)
            
            # SQL 파일 종료 부분 작성
            write_sql_footer(outfile)

def run_cli():
    """
    프로그램 메인 함수
    
    기능:
    -----
    1. 명령줄 인자 파싱
    2. 입력 파일 존재 여부 확인
    3. 출력 파일명 결정
    4. INI 파일을 SQL 파일로 변환
    5. 결과 메시지 출력
    
    사용법:
    ------
    python init_to_sql.py -i input.ini [-o output.sql]
    
    옵션:
    -----
    -i, --input  : 필수. 입력 INI 파일 경로
    -o, --output : 선택. 출력 SQL 파일 경로 (생략시 .ini를 .sql로 변경)
    """
    
    # ArgumentParser 객체 생성 - 명령줄 인자 파싱용
    parser = argparse.ArgumentParser(
        description='Convert INI file to SQL file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example:
  %(prog)s -i inifile.ini
  %(prog)s -i inifile.ini -o output.sql
        """
    )
    
    # -i/--input 인자 정의 (필수 인자)
    parser.add_argument('-i', '--input', 
                        required=True,
                        help='Configuration Review ini files')
    
    # -o/--output 인자 정의 (선택 인자)
    parser.add_argument('-o', '--output', 
                        help='specify extracted file name (default: replace ini extension with .sql)')
    
    # 명령줄 인자 파싱
    args = parser.parse_args()
    
    # 출력 파일명 결정
    if args.output:
        # -o 옵션으로 지정된 파일명 사용
        outfile = args.output
    else:
        # 입력 파일명에서 확장자를 .sql로 변경
        # 예: Full_AwrOne_Single.ini → Full_AwrOne_Single.sql
        base_name = os.path.splitext(args.input)[0]
        outfile = base_name + '.sql'
    
    # 입력 파일 존재 여부 확인
    if not os.path.exists(args.input):
        print(f"Error: Input file '{args.input}' not found")
        sys.exit(1)
    
    # INI 파일 파싱 및 SQL 파일 생성
    try:
        convert_ini_to_sql(args.input, outfile)
        # 성공 메시지 출력
        print(f"Successfully converted '{args.input}' to '{outfile}'")
    except Exception as e:
        # 오류 발생시 메시지 출력 후 종료
        print(f"Error processing file: {e}")
        sys.exit(1)

# 스크립트가 직접 실행될 때만 main() 함수 호출
# (모듈로 import될 때는 실행되지 않음)
if __name__ == "__main__":
    run_cli()
