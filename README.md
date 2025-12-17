1. 준비
1) python 3.13.11 을 설치. 예> C:\Python313 에 설치
   https://www.python.org/downloads/
2) 환경 변수 등록
 - 윈도우에서 "시스템 환경 변수 편집"을 열음
 - "환경 변수"에서 "새로 만들기"를 해서 
    C:\Python313
    C:\Python313\Scripts 
    2개를 환경변수 등록
3) xlsxwriter를 설치
 - https://pypi.org/project/xlsxwriter/#files 으로 갑니다. 
 - 다운로드 링크는 아래와 같습니다. 저는 C:\Python313 디렉터리에 다운로드 받았습니다.
   https://files.pythonhosted.org/packages/3a/0c/3662f4a66880196a590b202f0db82d919dd2f89e99a27fadef91c4a33d41/xlsxwriter-3.2.9-py3-none-any.whl
 - powershell을 열어서 xlsxwriter를 설치 합니다. pip는 python을 설치하면 자동 설치 됩니다. (버전이 안맞거나 해서 업그레이드/다운그레이드 해야할 수도 있습니다.)
   cd C:\Python313
   pip install xlsxwriter-3.2.9-py3-none-any.whl

2. Init File을 SQL 파일로 변경
- powershell 을 열어서 ini 파일을 sql으로 변경
  c:\ html_to_excel> python .\AwrONE_ini_to_sql.py -i .\SLIM_AwrOne_2node.ini
         
     Successfully converted '.\SLIM_AwrOne_2node.ini' to '.\SLIM_AwrOne_2node.sql'

3. SQL 파일의 맨 위에서 AWR의 수집 시간을 지정
- 수정할 대상 파일
    ==> SLIM_AwrOne_2node.sql
- 수정 전
        Prompt ############### Define Date : YYYYMMDDHH24MI #################################
        var snapdate_from number
        exec :snapdate_from:=202601010000
        var snapdate_to number
        exec :snapdate_to:=202601150000
        Prompt ##############################################################################
- 수정 후. 시간 값만 변경 합니다. 예를들어서 12월 10일 09시00~18시00분까지를 수집하면 아래와 같습니다.
        Prompt ############### Define Date : YYYYMMDDHH24MI #################################
        var snapdate_from number
        exec :snapdate_from:=202512100900
        var snapdate_to number
        exec :snapdate_to:=202512101800
        Prompt ##############################################################################

4. SQL을 Oracle DB에 접속해서 수행 합니다.
  - 각 노드에서 RAC 관련 쿼리가 중복되는데나중에는 한번만 수행되도록 ini 파일을 나눌 예정입니다.
  - select any dictionary 권한이 있는 유저면 모두 가능 합니다. 
   주의할 것은, rac msg, rac load 부분이 오래 걸릴 수 있으니 해당 쿼리는 제외하세요.  => 쿼리 튜닝 예정입니다. 
 -- 1번 노드
   c:\ sqlplus system/xxx@rac_node1
   SQL> @SLIM_AwrOne_2node.sql
  ==> awr snap이 나오고 백그라운드로 쿼리가 수행되고 html 파일이 생성됩니다.
      멈춘것이 아닙니다. 백그라운드로 출력되는 중이고 끝나면 로그아웃 됩니다.
 -- 2번 노드
   c:\ sqlplus system/xxx@rac_node2
   SQL> @SLIM_AwrOne_2node.sql

5. html 파일을 excel으로 변환 합니다. 이것이 핵심 입니다. 
 1) powershell을 엽니다.
 2) html 파일과 ini 파일, python code를 하나의 디렉터리에 놓습니다.
 3) 아래와 같이 수행합니다.
  c:\ html_to_excel> python .\AwrONE_html_to_excel.py -i .\AWR_node1_prod_20251217_1930.html

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
 4) output을 지정해서 변환도 가능 합니다.
   c:\ html_to_excel> python .\AwrONE_html_to_excel.py -i .\AWR_node1_prod_20251217_1930.html -o out1.xlsx

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
 5) debug mode를 사용해서 graph가 어디에 붙는지도 확인 가능 합니다.
   c:\ html_to_excel> python .\AwrONE_html_to_excel.py -i .\AWR_node1_prod_20251217_1930.html -o out2.xlsx --debug

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
5) 엑셀 결과 확인
- 단, RANK1~로 시작되는 컬럼의 일부는 ~_sql2 탭을 "행열 변환"으로 붙여야 합니다.
