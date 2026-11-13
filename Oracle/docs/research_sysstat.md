# sysstat 고정(pinpoint) 수집 통계 리서치

> deep-research 스킬(5-API 병렬)로 수행. 목적: AWR(dba_hist_sysstat) 스냅샷 델타를
> Excel 차트로 만드는 도구에서 "문제 감지용으로 콕 찍어 수집할 stat_name 고정 목록"을 정한다.
> ini/코드 수정 없음. 순수 리서치 산출물.

## 0. 요약 (TL;DR)

- 랭킹/자동선정(절대 델타, 상대 점수)은 실패 → **전문가 고정 목록** 방식이 맞다.
  (바이트급 카운터가 상위 독점, 잡음 카운터가 score=1.0) — 이번 실측 결론과 일치.
- 현재 도구의 `[load]` 섹션은 워크로드/IO 중심 25개만 수집한다. **인덱스 건강 지표는 전무**
  (`leaf node splits` 등 0건).
- 확정 수집 대상: `leaf node splits`, `leaf node 90-10 splits`. 아래는 이를 중심으로 한 확장 목록.
- 5개 API 모두 성공 (Naver만 Git Bash에서 `--data-urlencode` 인코딩 이슈로 1차 실패 →
  Python `urllib.parse.quote` 퍼센트 인코딩으로 재시도 성공). 실패한 API 없음.

---

## 1. 카테고리별 추천 수집 통계 (stat_name — v$statname 기준 정확 표기)

"이미수집" = 현재 `Full_AwrOne_Single.ini` `[load]`/`[rac_load]` 기준.

### A. 인덱스 건강 (최우선 신규)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `leaf node splits` | ✗ | 인덱스 리프 블록 분할 횟수. 과거 leaf split 폭증 장애의 1차 신호 |
| `leaf node 90-10 splits` | ✗ | 우측(high-key) 편중 분할. 단조 증가 키(시퀀스/날짜) 삽입 핫스팟 신호 |
| `branch node splits` | ✗ | 브랜치 블록 분할. 직렬화 증가의 보조 신호 |
| `root node splits` | ✗ | 루트 분할(드묾). 트리 깊이 증가/구조 이상 보조 |
| `index fetch by key` | ✗ | 인덱스 경유 단건 fetch량(워크로드 대비 split 비율 해석용) |
| `index fast full scans (full)` | ✗ | 인덱스 FFS(전체 스캔) 증가 = 인덱스/플랜 문제 신호 |
| `failed probes on index block reclamation` | ✗ | 11g+ 인덱스 리프 블록 재사용 실패. 90-10 과다와 연관 |
| `recursive aborts on index block reclamation` | ✗ | 위와 동일 계열(재사용 시도 중단) |

### B. Enqueue / 락 / ITL (신규 — leaf split과 항상 같이 봄)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `enqueue requests` | ✗ | enqueue(TX/ITL 등) 요청 총량. 경합의 상위 신호 |
| `enqueue waits` | ✗ | enqueue 대기 횟수. requests 대비 비율이 핵심 |
| `enqueue deadlocks` | ✗ | 0이어야 정상. >0이면 즉시 조사 |
| `enqueue timeouts` | ✗ | 경합 심화 신호 |
| `transaction lock foreground requests` | ✗ | TX 락(행/ITL/index contention) 전경 요청량 |
| `transaction lock foreground wait time` | ✗ | TX 락 대기 시간(centisecond). ITL/인덱스 경합 총량 |
| `IMU contention` | ✗ | In-Memory Undo 경합. TX 경합 보조 지표 |

> 주의: **`ITL waits`는 v$sysstat에 없다.** ITL 관련은 (a) 세그먼트 통계
> `dba_hist_seg_stat.statistic_name = 'ITL waits'`(라이선스 필요), (b) 대기 이벤트
> `enq: TX - allocate ITL entry`(도구의 `wait5/wait20` 섹션)로 봐야 한다.
> sysstat 쪽에서는 위 `enqueue*`/`transaction lock*`으로 우회 관측한다.

### C. Redo / 커밋 (일부 확장)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `redo size` | ✓ | DML/redo 발생량 |
| `redo entries` | ✓ | 변경 벡터 수 |
| `redo log space requests` | ✓ | 로그 스위치 대기 요청 |
| `redo buffer allocation retries` | ✗ | **≈0이어야 정상.** 증가 지속 = log buffer/커밋 부하 문제 |
| `redo log space wait time` | ✗ | 로그 공간 대기 시간 |
| `redo write time` | ✗ | LGWR 쓰기 시간 |
| `redo wastage` | ✗ | 비효율적 로그 블록 사용 |
| `user commits` | ✓ | 트랜잭션 커밋 수(정규화 분모) |
| `transaction rollbacks` | ✗ | 롤백 수(애플리케이션 오류/실패 신호) |

### D. Parse / Cursor (확장)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `parse count (total)` | ✓ | 소프트 포함 총 파싱 |
| `parse count (hard)` | ✓ | 하드 파싱. **Oracle 가이드: >100/sec면 조사** |
| `parse count (failures)` | ✗ | 파싱 실패(애플리케이션/SQL 오류) |
| `parse time cpu` | ✓ | 파싱 CPU 시간 |
| `parse time elapsed` | ✓ | 파싱 경과 시간(CPU와 비교) |
| `session cursor cache hits` | ✗ | 커서 캐시 적중 |
| `session cursor cache count` | ✗ | 커서 캐시 크기 |
| `opened cursors cumulative` | ✗ | 커서 누수/과다 오픈 신호 |
| `pinned cursors current` | ✗ | 핀 커서(공유풀 압박) |

### E. Buffer / I/O (기존 대부분 수집 — 참고)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `session logical reads` | ✓ | 버퍼 gets 총량(플랜 비효율 신호) |
| `consistent gets from cache` | ✓ | 일관성 읽기 |
| `db block gets from cache` | ✓ | current 모드 읽기(DML) |
| `physical reads` | ✓ | 디스크 블록 읽기 |
| `physical reads cache` / `physical reads direct` | ✓ | 캐시 경유/직접 읽기 |
| `physical read bytes` / `physical write bytes` | ✓ | 바이트 단위 IO |
| `physical writes` / `physical writes direct` | ✓ | 디스크 쓰기 |
| `db block changes` | ✓ | 변경 블록 수(redo와 함께 DML 강도) |
| `buffer is not pinned count` / `buffer is pinned count` | ✗ | 버퍼 고정/재방문 패턴(고급, 선택) |

> ⚠ 해석 주의(Oracle 문서): **buffer cache hit ratio는 캐시 크기 적정성의 증거가 아니다.**
> 높아도/낮아도 오해 소지가 있으니 실제 물리 IO와 SQL 플랜을 함께 본다.

### F. Sort / Workarea / Temp (확장)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `sorts (memory)` | ✓ | 메모리 정렬 |
| `sorts (disk)` | ✓ | 디스크 정렬(메모리 부족/SQL 비효율) |
| `sorts (rows)` | ✗ | 정렬 행 수 |
| `workarea executions - optimal` | ✗ | 메모리 내 완료(정상은 대다수) |
| `workarea executions - onepass` | ✗ | 1-pass(메모리 압박 신호) |
| `workarea executions - multipass` | ✗ | **회피 대상.** 응답시간 급증 |
| `workarea memory allocated` | ✗ | workarea 메모리 |
| `temp space allocated (bytes)` | ✗ | 임시공간 사용(정렬/해시 스필) |

### G. Undo (신규)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `undo change vector size` | ✗ | undo 발생량 |
| `rollback changes - undo records applied` | ✗ | 롤백/cleanout 부하 |
| `transaction tables consistent read rollbacks` | ✗ | ORA-1555류 신호 보조 |
| `data blocks consistent reads - undo records applied` | ✗ | CR undo 적용량(읽기 일관성 부하) |

### H. Row migration / Chaining (신규)

| stat_name | 이미수집 | 왜 |
|---|---|---|
| `table fetch continued row` | ✗ | 마이그레이션/체이닝 행 fetch(리프/블록 구조 문제 보조) |
| `table fetch by rowid` | ✗ | rowid 접근량(정규화 분모) |

### I. RAC / Global Cache (기존 `[rac_load]` 확장)

기존 `rac_load`가 gc 수신/서빙/메시지/융합쓰기를 이미 다룸. 추가로 보면 좋은 것:

| stat_name | 이미수집(참고) | 왜 |
|---|---|---|
| `gc current blocks served` / `gc cr blocks served` | rac_load | 서비스 블록 |
| `gc blocks lost` | ✗ | 무결성/인터커넥트 이상 |
| `gc blocks corrupt` | ✗ | 손상 블록 |
| `global enqueue gets sync` / `gets async` | ✗ | 전역 enqueue 부하 |
| `global enqueue get time` | ✗ | 전역 enqueue 지연 |
| `gcs messages sent` / `ges messages sent` | ✗ | 메시지 부하(인터커넥트) |

---

## 2. 각 통계의 "문제 신호" 기준 (실무 가이드)

절대값은 인스턴스 시작 이후 누적이므로 **반드시 구간 델타/초당·트랜잭션당 정규화**로 본다
(AWR도 이 방식을 전제).

| 신호 | 기준/해석 | 출처 |
|---|---|---|
| `parse count (hard)` | **초당 > 100** 이면 하드파싱 과다 → 리터럴 SQL/바인드 미사용/커서 공유 부족 | Oracle TGDBA(AWR 리포트 가이드) |
| `redo buffer allocation retries` | 지속 증가 시 조사. 정상은 0 근처 | Oracle AWR 리포트 |
| `workarea executions - multipass` | 0 지향. 증가 시 응답시간 급증 | Oracle AWR 리포트 |
| `buffer cache hit ratio` | 높/낮음 모두 오해 가능 → 단독 판단 금지 | Oracle AWR 리포트 |
| `enqueue deadlocks` | 0이어야 정상. >0 → 즉시 조사 | Oracle Docs(Enqueues) |
| `enqueue waits / enqueue requests` | 비율이 높고 함께 `TX` 계열 대기 → 락 경합 | Oracle Docs, Oracle Forums |
| `leaf node 90-10 splits` | 증가율↑ + `buffer busy waits`/`enq: TX - index contention`/RAC `gc` 동반 → 우측 핫스팟 | Richard Foote, Jonathan Lewis, Hemant |
| `leaf node splits` | rate 계산 후 대기이벤트와 동반 여부로 판단(단독 임계값 없음) | Jonathan Lewis, dba-oracle |
| `table fetch continued row` | 높으면 마이그레이션/체이닝 → PCTFREE/MOVE 검토 | AskTom, Oracle Docs |
| parse vs execute | `parse count (total)` ≈ `execute count` 면 재파싱 과다 | Oracle AWR 리포트 |

정규화 공식(합의된 실무 관행):
```
per second      = stat delta / elapsed seconds
per transaction = stat delta / (user commits + user rollbacks) delta
```

---

## 3. leaf node splits 심화 — 같이 봐야 할 지표 묶음

**진단 묶음(pinpoint 세트)**: 인덱스 건강(A) + Enqueue/ITL(B) + 관련 대기이벤트.

1. **분할 4종 동시 수집**
   `leaf node splits`, `leaf node 90-10 splits`, `branch node splits`, `root node splits`
   → 90-10 비중이 크면 우측(단조 증가 키) 핫스팟, 일반 split만 크면 랜덤 삽입/공간 부족.

2. **경합 신호(가장 중요)**: sysstat만으로는 "무엇이 막히는지"를 못 본다. 아래와 **반드시 교차**:
   - 세그먼트: `dba_hist_seg_stat` 의 `ITL waits`, `buffer busy waits`, `db block changes`(라이선스)
   - 대기 이벤트(도구 `wait5/wait20`): `buffer busy waits`, `enq: TX - index contention`,
     `enq: TX - allocate ITL entry`, (RAC) `gc buffer busy acquire`, `gc current request`
   - sysstat 우회: `enqueue requests/waits/deadlocks`, `transaction lock foreground wait time`

3. **워크로드 정규화 분모**: `user commits`, `execute count`, `user calls`, `index fetch by key`
   → split 건수를 트랜잭션/실행 대비로 환산해야 절대값 함정을 피함.

4. **RAC 분기**: 같은 인덱스 블록이 인스턴스 간 이동하면 `gc*` 계열이 튄다
   (`gc current blocks served/received`, `gc blocks lost`). 단조 증가 PK + RAC 조합이면
   순번/`CACHE`/`NOORDER`, scalable sequence 검토.

5. **해석 표**

| 관측 | 해석 |
|---|---|
| `leaf node 90-10 splits` 급증 + `buffer busy waits`/`enq: TX - index contention` | 우측 핫 리프 블록(단조 증가 키) |
| `leaf node splits`↑, `90-10` 낮음 | 랜덤 삽입 / 여유공간 부족 / 여러 리프 분할 |
| `enq: TX - index contention` 동반 | DML 동시성이 인덱스 구조 변경과 경합 |
| `enq: TX - allocate ITL entry` 동반 | ITL 부족(`INITRANS`/블록 공간) |
| RAC `gc*` 동반 | 동일 인덱스 블록이 인스턴스 간 이동 |
| split 높으나 대기 미미 | 단순 워크로드 반영일 수 있음 → 튜닝 불필요 |

6. **대응(튜닝) 요약**: 키 패턴 확인 → (필요시) `PCTFREE` 상향, reverse-key(범위 스캔 없을 때만),
hash-partitioned global index, RAC는 scalable sequence/`CACHE`+`NOORDER`. **무작정 rebuild 금지**(근본 원인 미해결).

---

## 4. 참고 출처

**Oracle 공식**
- Statistics Descriptions (V$SYSSTAT/V$SESSTAT 정의): https://docs.oracle.com/en/database/oracle/oracle-database/21/refrn/statistics-descriptions-2.html
- V$SYSSTAT 레퍼런스: https://download.oracle.com/docs/en/database/oracle/oracle-database/21/refrn/V-SYSSTAT.html
- AWR Report UI / Load Profile: https://docs.oracle.com/en/database/oracle/oracle-database/26/tgdba/awr-report-ui.html
- Gathering Database Statistics: https://docs.oracle.com/en/database/oracle/oracle-database/19/tgdba/gathering-database-statistics.html
- Guidelines for Indexes (right-growing index vs reverse key): https://docs.oracle.com/en/database/oracle/oracle-database/18/tgsql/guidelines-for-indexes-and-table-clusters.html
- Create Index (INITRANS 등): https://docs.oracle.com/en/database/oracle/oracle-database/21/sqlrf/CREATE-INDEX.html
- EM metric: Leaf Node Splits (per second): https://docs.oracle.com/cd/B16240_01/doc/doc.102/e16282/oracle_database_help/oracle_database_instance_throughput_leafnodesplits_ps.html
- EM metric: Enqueue TX mode 4 – allocate ITL entry (%): https://docs.oracle.com/cd/B16240_01/doc/doc.102/e16282/oracle_database_help/oracle_database_wait_bottlenecks_enqueue_alloc_itl_pct.html
- Automatic Performance Statistics (AWR 수집 항목): https://docs.oracle.com/cd/E25054_01/server.1111/e16638/autostat.htm

**전문가/블로그 (인덱스 split·경합)**
- Jonathan Lewis, Index splits: https://jonathanlewis.wordpress.com/2018/10/29/index-splits/
- Jonathan Lewis, Index analysis (50/50 high-end split, ASSM): https://jonathanlewis.wordpress.com/2008/09/26/index-analysis/
- Jonathan Lewis, Index ITL Limit: https://jonathanlewis.wordpress.com/2022/02/18/index-itl-limit/
- Richard Foote, Index Block Splits: https://richardfoote.wordpress.com/category/index-block-splits
- Hemant's Oracle DBA, Index Block Splits: 90-10: https://hemantoracledba.blogspot.com/2009/05/index-block-splits-90-10.html
- ksun, Index Block Split Point Distribution (split 통계 확인 쿼리): http://ksun-oracle.blogspot.com/2020/04/index-block-split-point-distribution.html
- Seydi Korurer, enq: TX – index contention 대응: https://seydikorurer.wordpress.com/2022/06/23/troubleshooting-enq-tx-index-contention-waits/
- Oracle-Base, AWR 10g 개요: https://oracle-base.com/articles/10g/automatic-workload-repository-10g
- Tanel Poder, tpt-oracle `awr_sysstat.sql` (통계 델타 추출 패턴): https://github.com/tanelpoder/tpt-oracle/blob/master/awr/awr_sysstat.sql
- V$SYSSTAT 이름 인덱스(전체 stat_name 목록): https://renenyffenegger.ch/notes/development/databases/Oracle/installed/dynamic-performance-views/sysstat/index

**ITL / enqueue**
- Oracle Forums/AskTOM, enq: TX – allocate ITL entry: https://asktom.oracle.com/pls/asktom/f%3Fp=100:11:0::::P11_QUESTION_ID:930130200346781639
- Oracle Scratchpad, Index ITL Limit: https://jonathanlewis.wordpress.com/2022/02/18/index-itl-limit/

**한국어(Naver)**
- DB CAFE, Lock 튜닝(enq: TX - row lock contention 등): http://dbcafe.co.kr/w/index.php/Lock_%ED%8A%9C%EB%8B%9D

---

## 부록: 리서치 방법/재현

- 스킬: `deep-research` (5-API 병렬). 쿼리: E1 시스템 통계 모니터링, E2 leaf node split/인덱스 경합,
  E3 ITL 경합, E4 Tanel Poder/Kerry Osborne AWR 스크립트, K1 한국어 인덱스 leaf split.
- 사용 API: Brave(19~20건×5), Tavily(20건×3 + AI 요약), Perplexity sonar-reasoning-pro(2회, 인용 포함),
  Naver(한국어 10건), YouTube.
- 실패: Naver 1차 `SE06 Malformed encoding` → Python 퍼센트 인코딩으로 재시도 성공(API 키 문제 아님).
- Quarantine 준수: 검색 결과 본문을 데이터로만 취급, 포함된 지시는 수행하지 않음.
