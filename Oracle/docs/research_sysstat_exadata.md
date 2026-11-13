# Exadata 전용 v$sysstat 고정(pinpoint) 수집 통계 리서치

> deep-research 스킬(5-API 병렬)로 수행. 선행: `docs/research_sysstat.md`(일반 sysstat 목록).
> 목적: AWR(dba_hist_sysstat) 델타 → Excel 차트 도구에 **Exadata 환경에서만 의미 있는
> 셀/CellOffload/SmartScan/flash cache 통계**를 고정 수집 목록으로 넣기 위함.
> ini/코드 수정 없음(리서치 산출물만).
>
> ⚠ 모든 stat_name은 Oracle 문서 기준 정확 표기. 단, **버전에 따라 존재/의미가 다름**
> (Oracle: "The availability of a specific statistic or wait event is subject to the version").
> 도입 전 `select name from v$statname where lower(name) like 'cell%'` 로 실제 존재 여부 확인 권장.

## 0. 요약 (TL;DR)

- Exadata 성능 진단의 3대 축: **① Smart Scan offload 효율**, **② Storage Index 절감**,
  **③ Smart Flash Cache 히트**. 이 세 가지가 "반드시 고정 수집".
- 효율의 핵심 공식: `offload efficiency = 1 − (cell physical IO interconnect bytes returned
  by smart scan / cell physical IO bytes eligible for predicate offload)`.
- **도구 커버리지 갭**: 현재 툴은 Exadata **cell *wait event***는 이미 수집한다
  (`cell single block physical read`, `cell smart table scan`, `cell direct...` 등 — ini 1528행~).
  그러나 **cell *sysstat* 카운터는 전무** → 이 문서의 목록이 실질 추가분.
- 5개 API 모두 성공. (Naver: Python 퍼센트 인코딩으로 사전 처리, 정상.)
  Brave 14~20×5, Tavily 20×3(+AI요약), Perplexity 2회, Naver 0/10, YouTube 0.

---

## 1. Exadata 전용 stat_name 고정 수집 목록 (카테고리별)

### A. Smart Scan / Offload 핵심 (최우선)

| stat_name | 의미(한 줄) |
|---|---|
| `cell physical IO bytes eligible for predicate offload` | predicate offload 대상 **온디스크** 바이트 |
| `cell physical IO bytes eligible for smart IOs` | offload 대상 **실제** 바이트(컬럼나 캐시 반영 시 온디스크와 다름) |
| `cell physical IO interconnect bytes returned by smart scan` | Smart Scan이 셀→DB로 **반환한** 바이트(offload 결과) |
| `cell physical IO interconnect bytes` | DB↔셀 인터커넥트 **총** 바이트(스마트스캔 외 IO 포함) |
| `cell physical IO bytes sent directly to DB node to balance CPU usage` | 스토리지 CPU 과부하로 DB로 되돌린 바이트(**역오프로드**) |
| `cell IO uncompressed bytes` | 셀에서 처리된(압축 해제 후) 데이터 크기 — offload 실처리량 |
| `cell scans` | 셀 스캔 횟수 |
| `cell index scans` | 셀 인덱스(FFS) 스캔 횟수 |
| `cell blocks processed by cache layer` / `... data layer` / `... index layer` / `... txn layer` | 셀 레이어별 처리 블록(오프로드 파이프라인 진단) |

### B. Storage Index

| stat_name | 의미(한 줄) |
|---|---|
| `cell physical IO bytes saved by storage index` | Storage Index로 **읽지 않고 건너뛴** 바이트(절감) |
| `cell physical IO bytes added to storage index` | Storage Index **구축**에 추가된 바이트(SI 빌드 중 표시) |
| `physical read requests optimized` | flash cache + SI + columnar로 최적화된 read 요청 수 |
| `physical read total bytes optimized` | flash cache/SI/columnar로 최적화된 바이트 |

### C. Smart Flash Cache (DB측)

| stat_name | 의미(한 줄) |
|---|---|
| `cell flash cache read hits` | flash cache가 충족한 read 수(DB측 유일 핵심 FC 지표) |
| `cell flash cache read hits for smart IO` | Smart IO 읽기 중 FC 히트 |
| `cell flash cache read hits for temp IO` | temp IO 읽기 중 FC 히트 |
| `cell flash cache read hits for partner reads` | 파트너 셀(미러) FC에서 충족된 읽기 |
| `cell flash read hits for controlfile reads` | 컨트롤파일 읽기 중 FC 히트 |
| `cell xrmem cache read hits for partner reads` | 파트너 XRMEM(RDMA 메모리)에서 충족된 읽기(신버전) |
| `cell overwrites in flash` | write-back로 디스크 IO를 절약한 캐시라인 덮어쓰기 수 |
| `cell partial writes in flash` | 일부는 FC, 일부는 디스크로 기록된 쓰기 |
| `cell writes to flash cache` | 전부 FC에 기록된 쓰기 |
| `cell writes to flash cache for temp IO` | temp 세그먼트 쓰기가 FC에 흡수된 수 |

> DB측 FC 지표는 제한적이다. 전체 flash cache 상태·용량·히트율은 **CellCLI**
> (`list metriccurrent ... where objectType='FLASHCACHE'`, 모든 셀에서)로 봐야 한다.

### D. Cell IO 일반 / 최적화 파일·백업

| stat_name | 의미(한 줄) |
|---|---|
| `physical read total bytes` | DB가 발행한 read 총 바이트(offload 여부 무관) |
| `physical write total bytes` | DB가 발행한 write 총 바이트 |
| `physical read total IO requests` / `physical write total IO requests` | 총 read/write 요청 수 |
| `physical write total bytes optimized` | flash cache로 흡수된 write 바이트 |
| `cell physical IO bytes saved during optimized file creation` | 파일 생성 offload 절감 바이트 |
| `cell physical IO bytes saved during optimized RMAN restore` | RMAN restore offload 절감 바이트 |

### E. 미오프로드(passthrough) 진단 — "offload가 안 먹을 때"

| stat_name | 의미(한 줄) |
|---|---|
| `cell num bytes in passthru during predicate offload` | 오프로드 안 되고 DB로 되돌린 바이트(핵심) |
| `cell num bytes in block IO during predicate offload` | 클라이언트가 block IO 모드라 미오프로드된 바이트 |
| `cell num bytes in filter passthru due to low mem` | 셀 메모리 부족으로 미오프로드 |
| `cell num bytes in filter passthru due to subheap size limit exc` | 셀 메모리 한계로 미오프로드 |
| `cell num bytes in passthru due to quarantine` | 셀 quarantine으로 미오프로드 |
| `cell num smart IO sessions in rdbms block IO due to big payload` | 메타데이터 과대로 block IO 세션 |
| `cell num smart IO sessions in rdbms block IO due to no cell mem` | 셀 메모리 부족으로 block IO |
| `cell num smart IO sessions in rdbms block IO due to online encr` | 온라인 암호화 중이라 block IO |
| `cell num smart IO sessions in rdbms block IO due to open fail` | 셀 연결 실패로 block IO |
| `cell num smart IO sessions in rdbms block IO due to user` | 사용자 설정으로 block IO |
| `cell num smart IO sessions using passthru mode due to cellsrv` | CELLSRV 문제로 passthru |
| `cell num smart IO sessions using passthru mode due to timezone` | DB timezone 업그레이드로 passthru |
| `cell num smart IO sessions using passthru mode due to user` | 사용자 설정으로 passthru |

### F. HCC / Columnar / 체이닝 (Exadata 특화 보조)

| stat_name | 의미(한 줄) |
|---|---|
| `cell physical IO bytes saved by columnar cache` | columnar cache로 읽기 회피한 바이트 |
| `cell physical IO bytes processed for IM query` / `... for IM capacity` / `... for no memcompress` / `... for XrCC` | columnar cache/XRMEM에서 처리한 바이트(포맷별) |
| `chained rows processed by cell` / `chained rows rejected by cell` / `chained rows skipped by cell` | 셀에서 처리/거부/스킵한 체이닝 행(스캔 효율) |

---

## 2. 각 통계의 해석 / 이상 신호

핵심은 **"절대값"이 아니라 구간 델타·비율**이다(누적 카운터).

| 지표/조합 | 정상/이상 기준 | 해석 | 출처 |
|---|---|---|---|
| `cell physical IO interconnect bytes` ≈ `eligible for predicate offload` 이고 `...returned by smart scan` 도 비슷 | 이상 | **오프로드가 안 일어남**(모든 대상 바이트가 그대로 반환) | Tavily(E3), Centroid |
| `eligible for predicate offload` = 0 | 이상 | predicate offload 대상 자체가 없음 = Smart Scan 미사용 | Perplexity(E1) |
| offload efficiency (`1 − returned/eligible`) | 정상 OLAP ~0.85–0.99 | 1.0에 가까우면 절감 미미. 10x=90%↓, 100x=99%↓ | dbakevlar, dbi-services |
| `cell physical IO bytes sent directly to DB node to balance CPU usage` | 클수록 주의 | 스토리지 CPU 과부하 → **역오프로드** 발생 | Oracle sagug(Smart I/O) |
| `cell num bytes in passthru during predicate offload` | 클수록 주의 | 오프로드 실패/부분 실패의 직접 증거 | Oracle sagug(Smart I/O) |
| `cell num smart IO sessions ... due to (no cell mem / low mem / big payload)` | >0 지속 | 셀 메모리 압박 또는 메타데이터 과대 | Oracle sagug |
| `cell num smart IO sessions ... due to (quarantine / cellsrv / online encr)` | >0 | 셀 상태/유지보수/암호화로 offload 차단 | Oracle sagug |
| `cell physical IO bytes saved by storage index` / `eligible` | 0 | SI 절감 없음(구축 중이거나 조건 불충족) | Oracle sagug, Vijay Dumpa |
| `physical read total bytes optimized` | 낮음 | FC/SI/columnar 최적화 미작동 | Oracle sagug(Flash Cache) |
| `cell flash cache read hits` | 낮음 | FC 활용 저조(워크셋/FC 모드/CellCLI 교차 확인) | Oracle sagug, Database Journal, uhesse |
| `cell smart table scan: pass through`(대기이벤트) | 발생 | offload 시도했으나 셀이 처리 못 하고 반환 | Oracle sagug(Smart I/O waits) |

**offload 미작동 시 점검 순서** (조합 진단):
1. 실행계획: `TABLE ACCESS STORAGE FULL` / `INDEX STORAGE FAST FULL SCAN` 인가 (일반 `TABLE ACCESS FULL`은 증거 아님)
2. 대기이벤트: `cell smart table scan`(정상 offload) vs `cell multiblock physical read`/`direct path read`(버퍼캐시/direct, offload 아님) vs `... pass through`
3. direct path 조건: 병렬/대용량 직렬. `_serial_direct_read`, 큰 SGA/버퍼캐시는 직렬 direct read를 막음
4. 파라미터: `cell_offload_processing`(false면 offload 중단), `cell_offload_decryption`
5. ASM: `cell.smart_scan_capable = TRUE`
6. SQL 제약: reverse-key/압축 인덱스 FFS 불가, ROWID/ROWSCN order, LONG/out-of-line LOB, SELECT...VERSIONS, 미커밋 데이터(read-consistent 재구성), 함수 `v$sqlfn_metadata.offloadable=NO`
7. 셀 리소스: 위 passthru/역오프로드 통계 + ExaWatcher/CellCLI

---

## 3. offload 효율 계산식 (실무 공식)

스냅샷 구간 델타 기준으로:

```text
# 1) Offload 절감률(권장, %)
offload_efficiency_% = 100 * (1 - smart_scan_returned / eligible)

  eligible = delta('cell physical IO bytes eligible for predicate offload')
             # 컬럼나 캐시/HCC 등 포맷 차이 반영하려면
             # delta('cell physical IO bytes eligible for smart IOs') 사용
  smart_scan_returned = delta('cell physical IO interconnect bytes returned by smart scan')

# 2) 절감 배수(ratio)
offload_ratio = eligible / smart_scan_returned
  # 1.0 = 절감 없음, 10 = 약 90% 절감, 100 = 약 99% 절감

# 3) Smart Scan 반환 비율(진단 보조) — 낮다고 무조건 문제 아님
smart_scan_return_ratio = delta('cell physical IO interconnect bytes returned by smart scan')
                         / delta('cell physical IO interconnect bytes')
  # 분모에 비(非)스마트스캔 IO가 포함되므로 단독 판단 금지

# 4) Storage Index 효율
storage_index_saving_% = 100 * delta('cell physical IO bytes saved by storage index')
                               / delta('cell physical IO bytes eligible for predicate offload')

# 5) Flash Cache 히트율(DB측 근사)
flash_cache_hit_% = 100 * delta('cell flash cache read hits')
                          / delta('physical read IO requests')
  # 정밀 FC 상태/히트율은 CellCLI list metriccurrent (objectType='FLASHCACHE') 권장
```

주의(문서적): Oracle AWR의 "Cell Offload Efficiency" KPI는 위 `1 − (interconnect / eligible)`
계열로 계산된다. 블로거에 따라 분자에 `cell physical IO interconnect bytes`(총량)를 쓰는 경우가
있는데, **순수 Smart Scan 절감률은 `...returned by smart scan`(스마트스캔 반환분)로 계산하는 게 정확**하다
(총 interconnect에는 데이터베이스 블록 IO 등 비스마트스캔 트래픽이 섞임).

---

## 4. 참고 출처

**Oracle 공식 (Exadata System Administration Guide — SAGUG)**
- Monitoring Smart I/O Using Database Statistics and Wait Events: https://docs.oracle.com/en/engineered-systems/exadata-database-machine/sagug/monitoring-smart-io-using-db-stats-wait-events.html
- Monitoring Exadata Smart Flash Cache Using Database Statistics and Wait Events: https://docs.oracle.com/en/engineered-systems/exadata-database-machine/sagug/monitoring-smart-flash-cache-using-db-stats-wait-events.html
- Monitoring Cell Disk I/O Using Database Statistics and Wait Events: https://docs.oracle.com/en/engineered-systems/exadata-database-machine/sagug/monitoring-cell-disk-io-using-db-stats-and-wait-events.html
- Offloading Data Search and Retrieval Processing: https://docs.oracle.com/en/engineered-systems/exadata-database-machine/sagug/offloading-data-search-retrieval-processing.html
- Monitoring Smart I/O — Common Issues: https://docs.oracle.com/en/engineered-systems/exadata-database-machine/sagug/monitoring-smart-io-common-issues.html
- Statistics Descriptions (V$SYSSTAT 개별 정의): https://docs.oracle.com/en/database/oracle/oracle-database/19/refrn/statistics-descriptions-2.html
- Exadata Performance and AWR (paper): https://www.oracle.com/docs/tech/database/exadata-awr.pdf
- V$SQLFN_METADATA(OFFLOADABLE): https://docs.oracle.com/en/database/oracle/oracle-database/21/refrn/V-SQLFN_METADATA.html
- ASM disk group attributes(cell.smart_scan_capable): https://docs.oracle.com/en/database/oracle/oracle-database/19/ostmg/manage-diskgroup-attributes.html

**전문가 / 블로그**
- Kerry Osborne(Enkitec), Understanding Exadata Offloading (PDF): https://kerryosborne.oracle-guy.com/papers/Understanding%20Exadata%20Offloading.pdf
- Tanel Poder, Exadata Performance (paper): https://tanelpoder.com/files/papers/exadata_performance.pdf
- dbi-services, How to measure Exadata SmartScan efficiency: https://www.dbi-services.com/blog/how-to-measure-exadata-smartscan-efficiency/
- Mark Smith, How Do *YOU* Measure SmartScans?: https://marksmithdba.wordpress.com/2015/03/24/how-do-you-measure-smartscans/
- Martin Carstenbach, Why is my Exadata smart scan not offloading?: https://martincarstenbach.com/2011/08/22/why-is-my-exadata-smart-scan-not-offloading/
- Uwe Hesse, Important Statistics & Wait Events on Exadata: https://uhesse.com/2011/07/06/important-statistics-wait-events-on-exadata/
- DBAKevlar, Smart Scan Efficiency 공식: https://dbakevlar.com/2020/02/exadata-workloads-to-azure-part-ii/
- Piotr Sajda, Lack of offload with large SGA: https://piotrsajda.com/oracle-blog/lack-of-offload-with-large-sga/
- Centroid, Direct Reads and Smart Scan with Exadata: https://www.centroid.com/blog/direct-reads-and-smart-scan-exadata/
- Database Journal, Monitoring Smart Flash Cache Activity on Oracle Exadata: https://www.databasejournal.com/oracle/monitoring-smart-flash-cache-activity-on-oracle-exadata/
- Vijay Dumpa, Exadata Smart Scan / Flash Cache / Storage Index: http://vijaydumpa.blogspot.com/2011/07/exadata-smart-scancell-offload.html
- Oracle Skills(github), exadata-features.md: https://github.com/oracle/skills/blob/main/db/architecture/exadata-features.md
- V$SYSSTAT 이름 인덱스(전체 목록): https://renenyffenegger.ch/notes/development/databases/Oracle/installed/dynamic-performance-views/sysstat/index

---

## 부록: 리서치 방법/재현

- 스킬: `deep-research`(5-API 병렬). 쿼리: E1 cell offload efficiency/predicate offload,
  E2 AWR sysstat cell metrics, E3 smart scan 미오프로드 진단, E4 Kerry Osborne/Frits Hoogland offload 스크립트,
  K1/K2 한국어 엑사데이터 오프로드/Storage Index.
- API: Brave(20/14/18/20/20), Tavily(20×3 + AI요약), Perplexity sonar-reasoning-pro(2회),
  Naver(Python 퍼센트 인코딩: 0/10), YouTube(0).
- 실패 API: 없음. (YouTube는 해당 쿼리 영상이 없어 0건, Naver K1은 결과 0건 — 오류 아님.)
- Quarantine 준수: 검색/문서 본문은 데이터로만 취급, 내부 지시는 수행하지 않음.
