# TprOne — Tibero TPR to Excel

Tibero TPR(Tibero Performance Repository) 데이터를 차트가 포함된 Excel 워크북으로 변환하는 도구입니다.

수집할 내용(SQL 쿼리, 시트 구성, 차트)은 INI 파일로 정의하고,
`tbsql` HTML 스풀로 DB에서 수집한 뒤 지표별 라인 차트가 들어간
`.xlsx`로 변환합니다.

> English version: [README.md](README.md)
> (Oracle AWR용 형제 도구 **AwrOne**은 [`../Oracle/`](../Oracle/)를 참조하세요.)

## 파이프라인

```
INI ──TprONE_ini_to_sql.py──▶ .sql ──tbsql(spool HTML)──▶ .html ──TprONE_html_to_excel.py──▶ .xlsx(차트 포함)
```

## 사용법

### 1. 수집용 SQL 생성

```bash
python TprONE_ini_to_sql.py -i Simple_TprOne_2node.ini -o out.sql
```

- 리포트 종류(`NODE` vs `TAC`)는 INI 파일명(`*Single*` → `NODE`, 그 외 `TAC`)과 본문 바인드 변수로 자동 판별하며, `-k {NODE,TAC}` 옵션으로 직접 지정할 수도 있습니다.

### 2. 수집 시간대 지정 (필수)

생성된 `out.sql` 상단의 두 줄을 원하는 구간으로 바꿉니다:

```sql
exec :snapdate_from:=202611171100
exec :snapdate_to:=202611172359
```

- 형식: `YYYYMMDDHH24MI` (년월일시분, DB 서버 로컬 시각).
- 기본값(`202601010000`~`202601150000`)은 자리 표시자이므로 반드시 원하는 구간
  (예: 장애 시간대, 최근 하루)으로 바꿔서 실행하세요.
- 시작/종료 시각이 지정 구간 안에 포함되는 스냅샷(`begin_interval_time >= :snapdate_from`
  및 `end_interval_time <= :snapdate_to`)을 수집합니다.
- `SYS._TPR_*` 테이블은 이미 구간 증분(delta)을 저장하므로 첫 스냅샷부터 그대로
  포함되며(Oracle처럼 `snap_id - 1` 차감 불필요), 구간 내 스냅샷이 없으면
  `ERROR-20001: No TPR ... snapshots found ...` 오류로 멈춥니다.
- `:inst_no`, `:dbid`는 접속 정보에서 자동으로 얻으므로 날짜 두 줄만 바꾸면 됩니다.
- Tibero의 `THREAD#` / `INSTANCE_NUMBER`는 **0-based**(`0..N-1`)이므로 TAC ini는
  `:inst_no1:=0` … `:inst_no8:=7`로 미리 설정돼 있습니다. 인스턴스 번호가 다르면
  이 값들을 수정하면 됩니다.

### 3. `tbsql`로 수집

`tbsql`과 DB 네트워크 접속만 있으면 아무 곳에서나 실행할 수 있습니다. 고정된
디렉터리는 없으며, 결과 HTML은 현재 디렉터리에 생깁니다.

- Single 계열 ini: 각 인스턴스에 접속해 **노드별로 1회씩** 실행.
- TAC(Tibero Active Cluster) 계열 ini: 아무 인스턴스에나 **1회** 실행 — 쿼리가 글로벌 스냅샷 ID(`snap_gid`)
  기준으로 클러스터 전체를 커버합니다.
  (2노드 클러스터에서 `4node`나 `8node` ini를 써도 동작하며, 없는 노드는 `inst#<N>`으로
  표시되고 0으로 채워집니다.)

```bash
# Single 계열 — 노드별 1회
tbsql sys/<비밀번호>@<노드1-접속별칭> @out.sql
tbsql sys/<비밀번호>@<노드2-접속별칭> @out.sql

# TAC 계열 — 아무 노드에 1회
tbsql sys/<비밀번호>@<아무노드-접속별칭> @out.sql
```

### 4. Excel로 변환

```bash
python TprONE_html_to_excel.py -i TPR_TAC_Simple_2node_tac1_prod1_20261117_201453.html
```

변환기는 HTML 안에 기록된 ini 이름을 자동으로 찾아 차트가 포함된 `.xlsx`를 생성합니다.

- 탐색 순서: 리포트 디렉터리 → 현재 디렉터리 → 스크립트 디렉터리 (`-n <ini>`로 직접 지정 가능).
- 출력 경로: 기본값은 입력 HTML과 같은 위치의 `.xlsx` (`-o <out.xlsx>`로 변경 가능).
- 차트 X축 옵션: `--x-rotation DEG`(`-90..90`, 기본 `-90`), `--x-format FMT`(기본 `'mm-dd hh:mm'`, 빈 문자열 `""` 전달 시 INI 설정 유지).
- ASH 변환: `*_with_ash.ini`로 수집했다면 추가로 생긴 `..._ASH.html`도 같은 방법으로 변환합니다.

## 출력 예시

| `load` — 노드별 DBTime(s) | `tac_load` — Global Cache blocks received |
|---|---|
| ![load 차트](docs/images/load_1.png) | ![tac_load 차트](docs/images/tac_load_1.png) |

## 요구 사항

- Python 3.x 표준 라이브러리 + `xlsxwriter`
- Tibero **7 이상**; `tbsql` 클라이언트
- `SYS._TPR_*`, `SYS._VT_*`, `V$*`, `GV$*` 읽기 권한 (`SYS` 또는 해당 TPR 테이블·동적 뷰 SELECT 권한이 있는 계정)
- 선택적 DB 파라미터:
  - `ACTIVE_SESSION_HISTORY=Y`: `*_with_ash.ini`에서 ASH 데이터가 쌓이려면 활성화 필요
  - `TPR_SEGMENT_STATISTICS=Y`: `Full_*`의 세그먼트 통계 시트(`seg_*`)가 수집되려면 활성화 필요

## 수집 INI 종류

| ini | 용도 | 실행 |
|---|---|---|
| `Simple_TprOne_Single.ini` | 단일 인스턴스 핵심 지표 | TAC 각 노드에서 실행 |
| `Simple_TprOne_Single_with_ash.ini` | 위 + ASH (별도 `_ASH.html`) | 〃 |
| `Full_TprOne_Single.ini` | 단일 인스턴스 전체 (Top SQL, Segment 등) | 〃 |
| `Full_TprOne_Single_with_ash.ini` | 위 + ASH | 〃 |
| `Simple_TprOne_{2..8}node.ini` | TAC 노드 비교 핵심 지표 (2~8노드) | 1회 (클러스터 단위) |
| `Full_TprOne_{2..8}node.ini` | TAC 노드 비교 전체 (2~8노드) | 〃 |

### Simple vs Full — 전체 추출 대상 `[섹션]` 상세 목록

#### 1. 단일 인스턴스 INI (`Simple_TprOne_Single.ini` vs `Full_TprOne_Single.ini`)

- **`Simple` (35개 시트 — `Simple`과 `Full` 공통 포함)**:
  - **환경 및 파라미터**:
    - `[dbinfo]` — 호스트/DB/인스턴스명, 버전, TAC 클러스터 여부, 수집 스냅샷 시작/종료 시각, DBID, 0-based `Instance#`(`THREAD#`)
    - `[snapinfo]` — 스냅샷 ID(`snap_id`), 스냅 종료 시각, 구간 길이(초), 인스턴스 기동 시각
    - `[param_sga]` — 공유 메모리/풀/캐시 크기 관련 핵심 파라미터(`TOTAL_SHM_SIZE`, `MEMORY_TARGET`, `DB_BLOCK_SIZE`, `LOG_BUFFER_SIZE` 등)
    - `[param]` — 전체 DB 파라미터 목록(`v$parameters`)
    - `[param_diff]` — 스냅샷 구간 내 변경된 파라미터 이력(`sys._tpr_modified_param`)
  - **부하 프로필 및 TAC/시스템 지표**:
    - `[load]` — Load Profile 시계열(구간 초, DB Time, DB CPU, Redo KB/s, 논리 읽기, 블록 변경, 물리 읽기/쓰기 및 MB/s, User Calls, Parse/Hard Parse, Sort, Logon, Exec, Rollback/Tx, Buffer/Redo/Lib/Sort/Soft-Parse/Latch 적중률 %)
    - `[tac_load]` — TAC Cluster Cache(CCC) 초당 블록 수신/전송, 인터커넥트(INC) 초당 메시지 및 트래픽(MB/s), 물리 읽기/쓰기(MB/s), BAST 수신, 블록 락 실패율
    - `[tac_gc_efficiency]` — 버퍼 접근 효율(`local %`, `remote %`, `disk %`)
    - `[time_model]` — Time Model 상세(DB Time, DB CPU, SQL 처리, 파싱, 하드파싱, 옵티마이저, SQL 실행 시간 및 DB Time 대비 비중 %)
    - `[osstat]` — OS CPU 사용 시간(`Busy`, `User`, `Sys`, `IOWait`, `Nice`, `Idle` 초 및 `%`), CPU/코어/소켓 수, 물리 메모리(GB)
    - `[inst_activity]` — 핵심 인스턴스 활동 지표(논리/물리 읽기, 물리 쓰기, 블록 변경, User Calls, 실행 수, Commit, Rollback, Redo Entries, Free 블록 스캔)
    - `[sysstat_key]` — 주요 내부 카운터(`sys._tpr_jcntstat`: 인덱스 분할, 블록 클린아웃, 버퍼 스캔, DBWR 플러시 요청, Redo Entries, Logon, INC 수신)
    - `[xd_sysstat]` — 스토리지 서버(`sys._tpr_ssvr_jcntstat`) 통계 건수, 크기, 소요시간
    - `[sysmetric]` — 스냅샷별 시스템 메트릭(`sys._tpr_metric`)
  - **대기 이벤트**:
    - `[wait5]` — 스냅샷별 Top 5 대기 이벤트(`DB CPU time` 포함) 시계열(대기 횟수, 총 시간 s, 평균 대기 ms, `% DB time`, Call/Exec당 DB Time)
    - `[wait20]` — 스냅샷별 Top 20 대기 이벤트 시계열(대기 횟수, 총 시간 s, 평균 대기 ms, `% DB time`)
    - `[wait15_raw]` — 스냅샷별 Non-Idle 대기 이벤트 원본 목록(Class, 대기 횟수, 타임아웃, 총 대기 ms, 최대 대기 ms)
    - `[wait_disk]` — 블록 유형별(단일/멀티블록, 데이터, 인덱스 리프/브랜치, 세그먼트 헤더, Undo, LOB) 디스크 읽기 대기 빈도(`waits/s`) 및 평균 응답시간(`avgms`)
    - `[redo_wait]` — Redo 로그 플러시/스위치/LGWR 대기시간(`s`) 및 평균 지연(`ms`), Redo 생성량(MB/s) 및 쓰기 시간
    - `[wait_buffer]` — 버퍼 Pin/Write/CR·CUR Busy(로컬·글로벌)/스핀락/체크포인트 대기시간(`s`), 평균 지연(`ms`), DBWR 버퍼 스캔 통계
    - `[waitstat]` — 버퍼 대기 횟수, 대기시간, 클래스 수(`sys._tpr_waitstat`)
  - **메모리, Undo, Redo 및 동시성**:
    - `[sga]` — 공유 메모리 세부 영역별 사용량(MB: `Free Space`, `PP`, `DD_CACHE`, `LIBRARY_CACHE Used`, `SLAB_*`, `CCC/CWS_RSB_CHUNK`)
    - `[pga]` — PGA 및 실행 메모리 사용량(MB: `ALLOCATED`, `FIXED`, `USED`, `Ex memory` 한도/할당량, Workarea 메모리 및 임시 세그먼트 사용량)
    - `[pga_hit]` — Workarea 메모리 내 처리율(`%W/A in-mem`), PGA 대비 Ex 메모리 비중, Workarea 메모리/Temp 사용량(MB)
    - `[libcache]` — 네임스페이스별(`SQL AREA`, `TBL`, `VIEW`, `PSMUNIT`, `OBJ`, `SEQ`, `SYN`, `USER`) Library Cache Gets 및 전체 `Pin Hit %`
    - `[latch]` — 대기시간 기준 Top 5 Latch 시계열(대기시간, Miss 비율 %, Miss 횟수)
    - `[enq]` — 총 대기시간 기준 Top 5 락/Wlock/트랜잭션 대기 시계열(초)
    - `[undo]` — Undo 블록 수, 트랜잭션 수, 최대 동시성, Snapshot Too Old, 공간 부족 에러, 활성 블록(`sys._tpr_undostat`)
    - `[redo]` — Redo 로그 그룹·멤버 구성(`v$log` / `v$logfile`)
    - `[log_ckpt]` — 스레드별 Redo 로그 스위치 간격 요약(스위치 횟수, 최소/최대/평균 분·초)
    - `[logswitch]` — 스레드별(`Thread#0..3`) 시간대별 Redo 로그 스위치 횟수
    - `[log_switch_f]` — 스레드별 로그 스위치 간격 분포(`<= 5 min`, `5 ~ 10 min`, `10 ~ 30 min`, `> 30 min`, `Total`)
    - `[log_switch]` — 일자별×24시간(`00`..`23`시) Redo 로그 스위치 히트맵
    - `[archsize]` — 스레드별(`Thread#0..3`) 일자별 아카이브 로그 생성 개수 및 용량(MB)
    - `[sqlstat]` — 스냅샷별 캡처된 SQL 요약(고유 SQL 수, 1회성 SQL 수, 실행 수, Buffer Gets, Disk Reads, Elapsed/CPU 시간, `%Single-Use`)

- **`Full` (`Simple` 35개에 아래 52개 상세 시트 추가 = 총 87개 시트)**:
  - **TAC 인터커넥트 및 클러스터 캐시 심층 분석**:
    - `[tac_char]` — Cluster Cache(`JC_CCC_%`) 카운트, 전송량(MB), 소요시간(s) 시계열
    - `[tac_msg]` — 인터커넥트 Ping(`sys._tpr_misc`) 평균/최소/최대 지연(ms) 및 `<1ms`/`<2ms`/`<4ms` 응답 건수 분포
    - `[tac_gesstat]` — 클러스터 메시징(`JC_CCC_%`, `JC_CWS_%`, `JC_GES_%`, `JC_GCS_%`) 영역별 건수·용량(MB)·시간(s)
    - `[tac_gccr_srv]` — Cluster Cache CR 서버(`JC_CCC_CR_%`) 건수·용량(MB)·시간(s)
    - `[tac_gccu_srv]` — Cluster Cache Current 서버(`JC_CCC_CUR_%`) 건수·용량(MB)·시간(s)
    - `[tac_gctran]` — Cluster Cache/CWS 전송(`JC_CCC_%`, `JC_CWS_%`) 건수·용량(MB)·시간(s)
  - **7대 지표별 Top 10 SQL (지표당 3개 시트 = 21개 시트 + SQL 분석 3개 시트)**:
    - `[sql_ela]` / `[sql_ela2]` / `[sql_ela_raw]` — **Elapsed Time(수행시간)** Top 10 SQL (시계열 차트 / Module·SQL 텍스트 포함 요약표 / 스냅샷별 원본)
    - `[sql_cpu]` / `[sql_cpu2]` / `[sql_cpu_raw]` — **CPU Time** Top 10 SQL
    - `[sql_get]` / `[sql_get2]` / `[sql_get_raw]` — **Buffer Gets(논리적 읽기)** Top 10 SQL
    - `[sql_read]` / `[sql_read2]` / `[sql_read_raw]` — **Disk Reads(물리적 읽기)** Top 10 SQL
    - `[sql_exe]` / `[sql_exe2]` / `[sql_exe_raw]` — **Executions(실행 횟수)** Top 10 SQL
    - `[sql_parse]` / `[sql_parse2]` / `[sql_parse_raw]` — **Parse Calls(파싱 횟수)** Top 10 SQL
    - `[sql_clwait]` / `[sql_clwait2]` / `[sql_clwait_raw]` — **I/O 및 Temp 세그먼트 대기시간** Top 10 SQL
    - `[sql_stat]` — 스냅샷별 `sys._tpr_sqlstats` 행 수, 고유 SQL 수, 실행 수, Elapsed/CPU 시간 추이
    - `[sql_literal_cnt]` — 앞 60자(`substr(sql_text, 1, 60)`) 기준 리터럴 중복 SQL 상위 목록(중복 건수 순)
    - `[sql_literal_mem]` — 앞 60자 기준 리터럴 중복 SQL 상위 목록(텍스트 총 용량 KB 순)
  - **세그먼트 통계 및 테이블스페이스/데이터파일 I/O**:
    - `[seg_stat_lio]` — 논리적 읽기(Logical Reads) Top 5 세그먼트(`sys._tpr_segmentstat`)
    - `[seg_stat_pio]` — 물리적 읽기(Physical Reads) Top 5 세그먼트
    - `[seg_stat_bbw]` — Buffer Busy Waits Top 5 세그먼트
    - `[tbsstat_total]` / `[filestat_total]` — 스냅샷별 전체 테이블스페이스·데이터파일 물리 읽기/쓰기·블록·시간 원본
    - `[tbsstat]` / `[filestat]` — 물리적 읽기 상위 5개 테이블스페이스·데이터파일 시계열
    - `[filestat_avg_blks_per_rd]` — 1회 읽기당 평균 블록 수 상위 5개 데이터파일 시계열
    - `[filestat_avg_time_per_rd]` / `[filestat_avg_time_per_rd_topio]` — 평균 읽기 시간 상위 5개 데이터파일 시계열 및 읽기 시간 순 원본
    - `[filestat_avg_bwt]` / `[filestat_avg_bwt_topio]` — 평균 쓰기 시간 상위 5개 데이터파일 시계열 및 쓰기 시간 순 원본
  - **대기, 메모리, 락 및 아카이브 상세**:
    - `[wait5_raw]` — 스냅샷별 Top 5 Non-Idle 대기 이벤트 원본 목록(대기 횟수, 시간 s, 평균 ms)
    - `[workarea]` / `[workarea_raw]` — SQL Workarea 크기 구간별 Optimal/OnePass/MultiPass 실행 히스토그램(`sys._tpr_sqlwa_hist`)
    - `[sess_cache]` — 스냅샷별 활성 ASH 샘플링 고유 세션 수(`count(distinct sid)`)
    - `[bufpool]` — 버퍼 캐시 Free 버퍼 요청/Victim, Dirty LRU 스캔, 버퍼 스캔량, Free 버퍼 대기(ms), Buffer Busy 대기(ms)
    - `[enq_av]` — 평균 대기시간(ms) 기준 Top 5 락/Wlock 대기 시계열
    - `[latch_sleep]` / `[latch_miss]` — Sleep 횟수 및 Miss 비율 기준 Top 5 Latch 시계열
    - `[rowcache]` — 딕셔너리·스키마 캐시(`JC_DD_%`, `JC_SC_%`) 건수·용량(MB)·시간(s)
    - `[arch]` — 수집 구간 내 아카이브 로그 상세 목록(`v$archived_log`: 스레드, 시퀀스, 파일명, 시각, 용량 MB)

- **`*_with_ash.ini` (별도 `..._ASH.html`로 스풀되는 1개 시트 추가)**:
  - `[ash]` — 수집 구간 내 Active Session History(`sys._tpr_active_session_history`) 원본 샘플 (27개 컬럼, 최대 150,000행 상한)

---

#### 2. 멀티노드 TAC INI (`Simple_TprOne_{2..8}node.ini` vs `Full_TprOne_{2..8}node.ini`)

- **`Simple` (9개 시트 — 글로벌 스냅샷 `snap_gid` 기준 클러스터 전체 노드 병렬 비교)**:
  - `[dbinfo]` — DB명, DBID, TAC 클러스터 표시, 전체 인스턴스 수, 접속 호스트/인스턴스/버전/상태
  - `[snapinfo]` — 스레드(`THREAD#`)별 시작/종료 `SnapGID`, 로컬 `SnapID` 범위 및 스냅 시각
  - `[param]` — 전 노드(`#0`..`#(N-1)`) 시스템 파라미터 병렬 비교표(`gv$system_parameters`)
  - `[load]` — 노드별 Load Profile 비교(`DBTime(s)`, `DBCPU(s)`, `Redo(KB)`, `LReads`, `PReads`, `UCalls`, `Execs`, `Parses`)
  - `[tac_load]` — 노드별 TAC 클러스터 트래픽 비교(`GC recv`, `GC srv`, `INC recv`, `INC sent`, `INC KB`, `INC sntKB`, `Prd`, `Pwr`)
  - `[tac_char]` — 노드별 Cluster Cache(`JC_CCC_%`) 처리 건수 비교
  - `[tac_msg]` — 노드별 인터커넥트(`JC_INC_%`) 메시지 건수 비교
  - `[tac_gccr_srv]` — 노드별 Cluster Cache CR(`JC_CCC_CR_%`) 처리 건수 비교
  - `[pga_hit]` — 노드별 Workarea `in-mem %` 및 `Ex/PGA %` 비교

- **`Full` (`Simple` 9개에 아래 14개 클러스터 시트 추가 = 총 23개 시트)**:
  - `[param_sga]` — 전 노드 메모리/SGA 핵심 파라미터 병렬 비교표
  - `[tac_gc_efficiency]` — 노드별 버퍼 접근 효율(`local %`, `remote %`, `disk %`) 비교
  - `[wait5]` — 노드별 `DBT(s)`, `DBCPU(s)`, Non-Idle `Waits(s)` 비교
  - `[bufpool]` — 노드별 버퍼 캐시 세부 카운터(`LReads`, `PReads`, `PWrts`, `Chngs`, `FreeReq`, `FreeVct`, `DirtyScn`, `FBWt(ms)`, `BsyWt(ms)`) 비교
  - `[undo]` — 노드별 Undo 블록(`UndoBlks`), 트랜잭션 수(`TxCnt`), 최대 동시성(`MaxConc`) 비교
  - `[ckpt_redo]` — 노드별 체크포인트·인스턴스 복구·Redo 지표(`IncCkpt`, `InstRecv`, `Recv(ms)`, `RedoMB`, `LgwrWrt`, `RdoWait(ms)`) 비교
  - `[sess_cache]` — 노드별 활성 ASH 샘플링 고유 세션 수 비교
  - `[redo]` — 클러스터 전체 스레드 Redo 로그 그룹·멤버 구성
  - `[log_ckpt]` — 스레드별 Redo 로그 스위치 간격(스위치 수, 최소/최대/평균 분·초) 요약
  - `[logswitch]` — 전 스레드 시간대별 Redo 로그 스위치 횟수
  - `[log_switch_f]` — 스레드별 로그 스위치 간격 분포(`<= 5 min`, `5 ~ 10 min`, `10 ~ 30 min`, `> 30 min`, `Total`)
  - `[log_switch]` — 클러스터 전체 일자별×24시간(`00`..`23`시) Redo 로그 스위치 매트릭스
  - `[archsize]` — 전 스레드 일자별 아카이브 로그 생성 개수 및 용량(MB)
  - `[arch]` — 수집 구간 내 클러스터 전체 아카이브 로그 상세 목록

## 출력 파일명 규칙

```
TPR_<NODE|TAC>_<ini태그>_<host>_<inst>_<yyyymmdd_hhmiss>.html
```

- `NODE` = Single 계열 ini, `TAC` = 2..8node 계열 ini
- `<ini태그>` = ini 파일명에서 `TprOne_`를 뺀 것. 길이를 줄이기 위해 NODE 계열의 `_Single`은
  생략하고 `_with_ash`는 `_ash`로 압축한다
  (예: `Simple_TprOne_Single.ini` → `Simple`, `Full_TprOne_Single_with_ash.ini` → `Full_ash`,
  `Simple_TprOne_4node.ini` → `Simple_4node`)
- 시각은 초까지(`hhmiss`) 들어가 같은 분에 재실행해도 파일이 덮어쓰이지 않는다
- 예: `TPR_TAC_Simple_2node_tac1_prod1_20261117_201453.html`,
  `TPR_NODE_Full_ash_tac1_prod1_20261117_181040.html`
- ASH 섹션(`SPOOL=ash`)은 `..._ASH.html`로 별도 spool

## 차트 제목 변수

Single 계열 ini의 CHART 제목에 쓰인 `@@hostname@@`, `@@instance@@`는 변환 시
입력 파일명에서 치환된다 (예: `DB Time (s) - @@hostname@@` → `... - tac1`).
멀티노드 TAC 차트는 클러스터 전체 노드를 함께 그리므로 호스트명을 붙이지 않는다.

## INI 문법 요약

- `[section]`이 두 번 나온다: ① 설정 블록(`SQLn=`, `FORMAT`, `CHART`, `SPOOL=`),
  ② 쿼리 블록(`select`로 시작). 쿼리 블록만 SQL 출력에 들어가고 앞에 `rem [section]`
  마커가 붙는다.
- 바인드 변수: `:snapdate_from` / `:snapdate_to`(요청 구간), `:snap_fr` / `:snap_to`
  (확정된 로컬 스냅샷 id), `:gid_fr` / `:gid_to`(확정된 TAC 글로벌 스냅샷 id),
  `:inst_no`, `:inst_no1..8`(`0..7`), `:dbid`.
- 차트 앵커 규칙: CHART의 열(COL) 필드 = 해당 시트 데이터 끝 열 + 2.
  컬럼을 추가/삭제할 때는 CHART/FORMAT 열 참조도 함께 밀어줘야 한다.
