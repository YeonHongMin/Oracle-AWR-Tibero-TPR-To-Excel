# 사용 방법

## 1. 준비

1) Python 3.13.11 설치 (예: `C:\Python313`)
- https://www.python.org/downloads/

2) 환경 변수 등록
- 윈도우에서 "시스템 환경 변수 편집"을 엽니다.
- "환경 변수"에서 "새로 만들기"를 해서 아래 2개를 등록합니다.
  - `C:\Python313`
  - `C:\Python313\Scripts`

3) xlsxwriter 설치
- https://pypi.org/project/xlsxwriter/#files
- 다운로드 링크 (예: `C:\Python313`에 다운로드):
  - https://files.pythonhosted.org/packages/3a/0c/3662f4a66880196a590b202f0db82d919dd2f89e99a27fadef91c4a33d41/xlsxwriter-3.2.9-py3-none-any.whl
- PowerShell에서 설치 (pip는 Python 설치 시 함께 설치됨)

```powershell
cd C:\Python313
pip install xlsxwriter-3.2.9-py3-none-any.whl
```

## 2. Init File을 SQL 파일로 변경

```powershell
c:\html_to_excel> python .\AwrONE_ini_to_sql.py -i .\SLIM_AwrOne_2node.ini
```

```text
Successfully converted '.\SLIM_AwrOne_2node.ini' to '.\SLIM_AwrOne_2node.sql'
```

## 3. SQL 파일에서 AWR 수집 시간 지정

- 수정할 대상 파일
  - `SLIM_AwrOne_2node.sql`

수정 전:
```text
Prompt ############### Define Date : YYYYMMDDHH24MI #################################
var snapdate_from number
exec :snapdate_from:=202601010000
var snapdate_to number
exec :snapdate_to:=202601150000
Prompt ##############################################################################
```

수정 후 (시간 값만 변경, 예: 12월 10일 09:00~18:00):
```text
Prompt ############### Define Date : YYYYMMDDHH24MI #################################
var snapdate_from number
exec :snapdate_from:=202512100900
var snapdate_to number
exec :snapdate_to:=202512101800
Prompt ##############################################################################
```

## 4. Oracle DB 접속 후 SQL 수행

- 각 노드에서 RAC 관련 쿼리가 중복되는데, 추후에는 한 번만 수행하도록 ini 파일을 분리할 예정입니다.
- `SELECT ANY DICTIONARY` 권한이 있는 유저면 모두 가능합니다.
- 주의: `rac msg`, `rac load` 부분은 오래 걸릴 수 있으니 해당 쿼리는 제외하세요. (쿼리 튜닝 예정)

1번 노드:
```powershell
c:\> sqlplus system/xxx@rac_node1
SQL> @SLIM_AwrOne_2node.sql
```

```text
==> AWR snap이 나오고 백그라운드로 쿼리가 수행되며 HTML 파일이 생성됩니다.
    멈춘 것이 아닙니다. 출력이 끝나면 로그아웃됩니다.
```

2번 노드:
```powershell
c:\> sqlplus system/xxx@rac_node2
SQL> @SLIM_AwrOne_2node.sql
```

## 5. HTML 파일을 Excel로 변환 (핵심)

1) PowerShell을 엽니다.  
2) HTML 파일, INI 파일, Python 코드를 동일 디렉터리에 둡니다.  
3) 아래처럼 실행합니다.

```powershell
c:\html_to_excel> python .\AwrONE_html_to_excel.py -i .\AWR_node1_prod_20251217_1930.html
```

```text
======================================================================
AWR HTML to Excel Converter
======================================================================
Extracting INI file name from HTML...
Input HTML file : .\AWR_node1_prod_20251217_1930.html
Output Excel file: .\AWR_node1_prod_20251217_1930.xlsx
INI config file  : ./SLIM_AwrOne_2node.ini
Loading INI configuration...
Converting HTML to Excel...
Processing sheet: [dbinfo]
Processing sheet: [snapinfo]
Processing sheet: [param_sga]
Processing sheet: [param_sga_sql2]
<중략>
Processing sheet: [sga_raw]
Processing sheet: [sga_raw_sql2]
Processing sheet: [sga]
======================================================================
[OK] Conversion completed successfully!
[OK] Excel file created: .\AWR_node1_prod_20251217_1930.xlsx
======================================================================
```

4) 출력 파일명을 지정할 수도 있습니다.

```powershell
c:\html_to_excel> python .\AwrONE_html_to_excel.py -i .\AWR_node1_prod_20251217_1930.html -o out1.xlsx
```

```text
<중략>
Processing sheet: [libcache_raw]
Processing sheet: [libcache]
Processing sheet: [sga_raw]
Processing sheet: [sga_raw_sql2]
Processing sheet: [sga]
======================================================================
[OK] Conversion completed successfully!
[OK] Excel file created: out1.xlsx
======================================================================
```

5) Debug 모드로 그래프 위치도 확인할 수 있습니다.

```powershell
c:\html_to_excel> python .\AwrONE_html_to_excel.py -i .\AWR_node1_prod_20251217_1930.html -o out2.xlsx --debug
```

```text
======================================================================
AWR HTML to Excel Converter
======================================================================
Extracting INI file name from HTML...
Found INI file reference: ./SLIM_AwrOne_2node.ini
Input HTML file : .\AWR_node1_prod_20251217_1930.html
Output Excel file: out2.xlsx
INI config file  : ./SLIM_AwrOne_2node.ini
Loading INI configuration...
Parsing INI file: ./SLIM_AwrOne_2node.ini
  FORMAT (line 28): sheet=snapinfo, range=1.5:65535.5, format=###,##0
  FORMAT (line 47): sheet=param_sga, range=1.2:65535.3, format=###,##0
  FORMAT (line 77): sheet=param, range=1.2:65535.3, format=###,##0
  FORMAT (line 78): sheet=param, range=1.1:65535.1, format=###,##0
  FORMAT (line 112): sheet=load, range=1.3:65535.36, format=###,##0
  FORMAT (line 113): sheet=load, range=1.37:65535.54, format=##0%
  CHART (line 114): sheet=load, title=Transactions (per sec), pos=(2,56), type=LINE
  CHART (line 115): sheet=load, title=Redo Size (bytes/s), pos=(22,56), type=LINE
  CHART (line 116): sheet=load, title=Logical Reads (per sec), pos=(42,56), type=LINE
  CHART (line 117): sheet=load, title=Block Changes (per sec), pos=(62,56), type=LINE
<중략>
  CHART (line 5590): sheet=sga, title=library cache detail, pos=(102,20), type=LINE
Loaded 65 format configurations
Loaded 45 chart configurations
Converting HTML to Excel...
Processing HTML file: .\AWR_node1_prod_20251217_1930.html
Default Excel formats defined
Found 124 tables in HTML
Processing sheet: [dbinfo]
  Wrote 3 rows to worksheet 'dbinfo'
Processing sheet: [snapinfo]
  Wrote 41 rows to worksheet 'snapinfo'
<중략>
  CHART (line 5590): sheet=sga, title=library cache detail, pos=(102,20), type=LINE
Loaded 65 format configurations
Loaded 45 chart configurations
Converting HTML to Excel...
Processing HTML file: .\AWR_node1_prod_20251217_1930.html
Default Excel formats defined
Found 124 tables in HTML
Processing sheet: [dbinfo]
  Wrote 3 rows to worksheet 'dbinfo'
Processing sheet: [snapinfo]
  Wrote 41 rows to worksheet 'snapinfo'
```

6) 엑셀 결과 확인
- `RANK1~`로 시작되는 컬럼의 일부는 `~_sql2` 탭을 "행열 변환"으로 붙여야 합니다.
