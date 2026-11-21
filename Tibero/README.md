# TprOne — Tibero TPR to Excel

Convert Tibero TPR (Tibero Performance Repository) data into Excel workbooks with charts.

An INI file defines what to collect (SQL queries, sheet layout, charts).
The pipeline collects from the database via `tbsql` HTML spool and turns the
result into a formatted `.xlsx` with line charts per metric.

> Korean version: [README_ko.md](README_ko.md)
> (For Oracle AWR, see the sibling tool **AwrOne** in [`../Oracle/`](../Oracle/).)

## Pipeline

```
INI ──TprONE_ini_to_sql.py──▶ .sql ──tbsql (spool HTML)──▶ .html ──TprONE_html_to_excel.py──▶ .xlsx (with charts)
```

## Usage

### 1. Generate the collection SQL

```bash
python TprONE_ini_to_sql.py -i Simple_TprOne_2node.ini -o out.sql
```

- Report kind (`NODE` vs `TAC`) is inferred from the INI filename (`*Single*` → `NODE`, otherwise `TAC`) and cross-checked against the INI body; pass `-k {NODE,TAC}` to override explicitly.

### 2. Set the collection time window (required)

Open the generated `out.sql` and edit the two lines near the top:

```sql
exec :snapdate_from:=202611171100
exec :snapdate_to:=202611172359
```

- Format: `YYYYMMDDHH24MI` (year, month, day, hour, minute — local DB time).
- The shipped defaults (`202601010000`–`202601150000`) are placeholders;
  always set the range you actually want (e.g. an incident window or the last day).
- Snapshots whose interval falls inside the window (`begin_interval_time >= :snapdate_from`
  and `end_interval_time <= :snapdate_to`) are collected.
- Since `SYS._TPR_*` tables already store per-interval deltas, the first snapshot
  inside the window is included directly (no baseline `snap_id - 1` subtraction).
  If no snapshots match, the script stops with `ERROR-20001: No TPR ... snapshots found ...`.
- `:inst_no` and `:dbid` are auto-detected from the connection — you only edit
  the two dates.
- Tibero `THREAD#` / `INSTANCE_NUMBER` is **0-based** (`0..N-1`), so TAC ini files
  preset `:inst_no1:=0` … `:inst_no8:=7`; adjust them if your instances are
  numbered differently.

### 3. Collect with `tbsql`

Run anywhere that has `tbsql` and network access to the DB — there is no fixed
directory; the HTML report is written to the current directory.

- Single-family ini: run **once per node**, connecting to each instance.
- TAC (Tibero Active Cluster) ini: run **once** against any instance — the queries already cover
  the whole cluster grouped by global snapshot ID (`snap_gid`). (A `4node` or
  `8node` ini on a 2-node cluster still works; missing nodes show as `inst#<N>`
  with zero-filled series.)

```bash
# Single family — once per node
tbsql sys/<password>@<node1-conn> @out.sql
tbsql sys/<password>@<node2-conn> @out.sql

# TAC family — once, any node
tbsql sys/<password>@<any-node-conn> @out.sql
```

### 4. Convert to Excel

```bash
python TprONE_html_to_excel.py -i TPR_TAC_Simple_2node_tac1_prod1_20261117_201453.html
```

- The converter reads the source ini name recorded inside the HTML and searches
  the report's directory, the working directory, and the script directory; use
  `-n <ini>` to override.
- The `.xlsx` (with charts) is written next to the HTML (or to `-o <out.xlsx>`).
- Optional chart flags: `--x-rotation DEG` (`-90..90`, default `-90`) and
  `--x-format FMT` (default `'mm-dd hh:mm'`; pass `""` to keep the INI's format).
- For `*_with_ash.ini` collections, convert the extra `..._ASH.html` the same way.

## Sample output

| `load` — DBTime(s) per node | `tac_load` — Global Cache blocks received |
|---|---|
| ![load chart](docs/images/load_1.png) | ![tac_load chart](docs/images/tac_load_1.png) |

## Requirements

- Python 3.x standard library + `xlsxwriter`
- Tibero **7 or later**; `tbsql` client
- Read access to `SYS._TPR_*`, `SYS._VT_*`, `V$*`, and `GV$*` views (e.g. `SYS`
  or an account granted SELECT on the TPR repository tables and dynamic views)
- Optional DB parameters:
  - `ACTIVE_SESSION_HISTORY=Y` is required for `*_with_ash.ini` to populate ASH rows
  - `TPR_SEGMENT_STATISTICS=Y` is required for `Full_*` segment statistics sheets (`seg_*`) to populate

## Collection INI set

| ini | Purpose | How to run |
|---|---|---|
| `Simple_TprOne_Single.ini` | Single-instance key metrics | Run on each TAC node |
| `Simple_TprOne_Single_with_ash.ini` | Same + ASH (separate `_ASH.html`) | 〃 |
| `Full_TprOne_Single.ini` | Single-instance full set (Top SQL, segments, ...) | 〃 |
| `Full_TprOne_Single_with_ash.ini` | Same + ASH | 〃 |
| `Simple_TprOne_{2..8}node.ini` | TAC node-comparison key metrics (2 to 8 nodes) | Once (cluster-wide) |
| `Full_TprOne_{2..8}node.ini` | TAC node-comparison full set (2 to 8 nodes) | 〃 |

### Simple vs Full — Collected `[section]` Reference

#### 1. Single-Instance INIs (`Simple_TprOne_Single.ini` vs `Full_TprOne_Single.ini`)

- **`Simple` (35 sheets — included in both `Simple` and `Full`)**:
  - **Environment & Parameters**:
    - `[dbinfo]` — Host, DB name, instance name, version, TAC cluster flag, snapshot window, DBID, 0-based `THREAD#`
    - `[snapinfo]` — TPR snapshot IDs (`snap_id`), timestamps, interval duration (s), instance startup time
    - `[param_sga]` — Shared memory / pool / cache sizing parameters (`TOTAL_SHM_SIZE`, `MEMORY_TARGET`, `DB_BLOCK_SIZE`, `LOG_BUFFER_SIZE`, ...)
    - `[param]` — All instance configuration parameters (`v$parameters`)
    - `[param_diff]` — Modified parameter history across snapshots (`sys._tpr_modified_param`)
  - **Workload & TAC Profile**:
    - `[load]` — Load Profile time series (Elapsed, DB Time, DB CPU, Redo KB/s, Logical Reads, Block Changes, Physical Reads/Writes & MB/s, User Calls, Parses/Hard Parses, Sorts, Logons, Executes, Rollbacks/Transactions, Buffer/Redo/Library/Sort/Soft-Parse/Latch hit %)
    - `[tac_load]` — TAC Cluster Cache (CCC) blocks received/served per sec, Interconnect (INC) messages & MB/s, physical read/write MB/s, BAST received, block lock get fails
    - `[tac_gc_efficiency]` — Buffer access efficiency breakdown (`local %`, `remote %`, `disk %`)
    - `[time_model]` — Time Model breakdown (DB Time, DB CPU, SQL processing, parse, hard parse, optimizer, SQL execute time and `% of DB Time`)
    - `[osstat]` — OS CPU time (`Busy`, `User`, `Sys`, `IOWait`, `Nice`, `Idle` seconds & `%`), CPUs/Cores/Sockets, physical memory (GB)
    - `[inst_activity]` — Key instance activity deltas (Logical/Physical Reads, Physical Writes, Block Changes, User Calls, Executes, Commits, Rollbacks, Redo Entries, Free Blocks Scanned)
    - `[sysstat_key]` — Key internal counters (`sys._tpr_jcntstat`: index splits, cleanouts, buffer scans, DBWR flush requests, redo entries, logons, INC messages)
    - `[xd_sysstat]` — Storage server (`sys._tpr_ssvr_jcntstat`) row count, size, and elapsed time
    - `[sysmetric]` — TPR system metrics (`sys._tpr_metric`) per snapshot
  - **Wait Events**:
    - `[wait5]` — Top 5 Timed Events (including `DB CPU time`) per snapshot (waits, time s, avg wait ms, `% DB time`, DB time per call/exec)
    - `[wait20]` — Top 20 Timed Events per snapshot (waits, time s, avg wait ms, `% DB time`)
    - `[wait15_raw]` — Non-idle wait event raw breakdown per snapshot (`Class`, `Waits`, `Timeouts`, `Time(ms)`, `Max(ms)`)
    - `[wait_disk]` — User I/O disk read wait frequency (`waits/s`) and latency (`avgms`) by block type (single/multi-block, data, index leaf/branch, segment header, undo, lob)
    - `[redo_wait]` — Redo & log flush/switch/LGWR wait time (`s`) and average latency (`ms`) plus redo generation/write stats
    - `[wait_buffer]` — Buffer pin/write/CR-CUR busy (local & global)/spinlock/checkpoint wait time (`s`), average wait (`ms`), and DBWR scan stats
    - `[waitstat]` — Buffer wait count, wait time, and distinct classes (`sys._tpr_waitstat`)
  - **Memory, Undo, Redo & Concurrency**:
    - `[sga]` — Shared memory pool/allocator usage in MB (`Free Space`, `PP`, `DD_CACHE`, `LIBRARY_CACHE Used`, `SLAB_*`, `CCC/CWS_RSB_CHUNK`)
    - `[pga]` — PGA and execution memory usage in MB (`ALLOCATED`, `FIXED`, `USED`, `Ex memory` limits/allocated, workarea memory & temp segment used)
    - `[pga_hit]` — Workarea in-memory `%`, execution memory `% of PGA`, and workarea memory vs temp segment (MB)
    - `[libcache]` — Library cache gets by namespace (`SQL AREA`, `TBL`, `VIEW`, `PSMUNIT`, `OBJ`, `SEQ`, `SYN`, `USER`) and overall `Pin Hit %`
    - `[latch]` — Top 5 latches by wait time (`wait_time`, `missratio %`, `misses`)
    - `[enq]` — Top 5 lock/wlock/transaction wait events by wait time (s)
    - `[undo]` — Undo blocks, transaction count, max concurrency, snapshot too old count, no-space errors, active blocks (`sys._tpr_undostat`)
    - `[redo]` — Redo log group and member layout (`v$log` / `v$logfile`)
    - `[log_ckpt]` — Per-thread redo log switch interval statistics (switch count, min/max/avg gap in minutes and seconds)
    - `[logswitch]` — Hourly redo log switch count matrix across threads (`Thread#0..3`)
    - `[log_switch_f]` — Per-thread log switch interval frequency buckets (`<= 5 min`, `5 ~ 10 min`, `10 ~ 30 min`, `> 30 min`, `Total`)
    - `[log_switch]` — Daily 24-hour (`00`..`23`) redo log switch heatmap matrix
    - `[archsize]` — Daily archived log count and size (MB) across threads (`Thread#0..3`)
    - `[sqlstat]` — Per-snapshot SQL summary (`sys._tpr_sqlstats`: distinct SQLs, single-use SQLs, executions, buffer gets, disk reads, elapsed/CPU s, `%Single-Use`)

- **`Full` (adds 52 detail sheets on top of `Simple` = 87 sheets total)**:
  - **TAC Interconnect & Cluster Cache Deep-Dive**:
    - `[tac_char]` — Cluster Cache (`JC_CCC_%`) total count, size (MB), and time (s) series
    - `[tac_msg]` — Interconnect (`sys._tpr_misc` INC ping) avg/min/max latency (ms) and `<1ms` / `<2ms` / `<4ms` histogram counts
    - `[tac_gesstat]` — Cluster messaging breakdown across `JC_CCC_%`, `JC_CWS_%`, `JC_GES_%`, and `JC_GCS_%` (count, MB, seconds)
    - `[tac_gccr_srv]` — Cluster Cache CR server (`JC_CCC_CR_%`) count, size (MB), and time (s)
    - `[tac_gccu_srv]` — Cluster Cache Current server (`JC_CCC_CUR_%`) count, size (MB), and time (s)
    - `[tac_gctran]` — Cluster Cache & CWS transfer (`JC_CCC_%`, `JC_CWS_%`) count, size (MB), and time (s)
  - **Top SQL by 7 Metrics (3 sheets per metric = 21 sheets + 3 SQL analysis sheets)**:
    - `[sql_ela]` / `[sql_ela2]` / `[sql_ela_raw]` — Top 10 SQL by **Elapsed Time (s)** (per-snapshot series, summary with module & SQL text, per-snapshot raw ranking)
    - `[sql_cpu]` / `[sql_cpu2]` / `[sql_cpu_raw]` — Top 10 SQL by **CPU Time (s)**
    - `[sql_get]` / `[sql_get2]` / `[sql_get_raw]` — Top 10 SQL by **Buffer Gets**
    - `[sql_read]` / `[sql_read2]` / `[sql_read_raw]` — Top 10 SQL by **Disk Reads**
    - `[sql_exe]` / `[sql_exe2]` / `[sql_exe_raw]` — Top 10 SQL by **Executions**
    - `[sql_parse]` / `[sql_parse2]` / `[sql_parse_raw]` — Top 10 SQL by **Parse Calls**
    - `[sql_clwait]` / `[sql_clwait2]` / `[sql_clwait_raw]` — Top 10 SQL by **I/O & Temp Segment Wait Time (s)**
    - `[sql_stat]` — Per-snapshot `sys._tpr_sqlstats` row count, distinct SQL count, executions, elapsed (s), CPU (s)
    - `[sql_literal_cnt]` — Duplicate literal SQL detector ranked by duplicate count (`substr(sql_text, 1, 60)`)
    - `[sql_literal_mem]` — Duplicate literal SQL detector ranked by total text volume (KB)
  - **Top Segments & Tablespace / Datafile I/O**:
    - `[seg_stat_lio]` — Top 5 segments by Logical Reads (`sys._tpr_segmentstat`)
    - `[seg_stat_pio]` — Top 5 segments by Physical Reads
    - `[seg_stat_bbw]` — Top 5 segments by Buffer Busy Waits
    - `[tbsstat_total]` / `[filestat_total]` — Full per-snapshot tablespace and datafile physical read/write count, block, and timing table
    - `[tbsstat]` / `[filestat]` — Top 5 tablespaces and datafiles by physical reads time series
    - `[filestat_avg_blks_per_rd]` — Top 5 datafiles by average blocks per physical read
    - `[filestat_avg_time_per_rd]` / `[filestat_avg_time_per_rd_topio]` — Top 5 datafiles by average read time and full per-snapshot table ordered by read time
    - `[filestat_avg_bwt]` / `[filestat_avg_bwt_topio]` — Top 5 datafiles by average write time and full per-snapshot table ordered by write time
  - **Wait, Memory, Lock & Archive Details**:
    - `[wait5_raw]` — Top 5 non-idle wait events per snapshot (raw list with waits, time s, avg ms)
    - `[workarea]` / `[workarea_raw]` — SQL workarea optimal/onepass/multipass execution histogram (`sys._tpr_sqlwa_hist`)
    - `[sess_cache]` — Distinct active sampled sessions per snapshot (`sys._tpr_active_session_history`)
    - `[bufpool]` — Buffer cache free-buffer requests/victims, dirty LRU scans, scanned buffers, free-buffer wait (ms), buffer busy wait (ms)
    - `[enq_av]` — Top 5 lock/wlock wait events by average wait time (ms)
    - `[latch_sleep]` / `[latch_miss]` — Top 5 latches by sleep count and miss ratio
    - `[rowcache]` — Dictionary & schema cache (`JC_DD_%`, `JC_SC_%`) count, size (MB), and time (s)
    - `[arch]` — Archived redo log details (`v$archived_log`: thread, sequence, file path, first/next time, size MB)

- **`*_with_ash.ini` (appends 1 sheet spooled to `..._ASH.html`)**:
  - `[ash]` — Active Session History (`sys._tpr_active_session_history`) raw samples within the window (27 columns, capped at 150,000 rows)

---

#### 2. Multi-Node TAC INIs (`Simple_TprOne_{2..8}node.ini` vs `Full_TprOne_{2..8}node.ini`)

- **`Simple` (9 sheets — cluster-wide node-by-node comparison grouped by `snap_gid`)**:
  - `[dbinfo]` — Database name, DBID, TAC cluster mode, instance count, connected host/instance/version/status
  - `[snapinfo]` — Per-thread (`THREAD#`) first/last `SnapGID`, first/last local `SnapID`, and begin/end timestamps
  - `[param]` — Side-by-side system parameter comparison across nodes (`#0`..`#(N-1)`)
  - `[load]` — Per-node Load Profile comparison (`DBTime(s)`, `DBCPU(s)`, `Redo(KB)`, `LReads`, `PReads`, `UCalls`, `Execs`, `Parses`)
  - `[tac_load]` — Per-node TAC cluster traffic comparison (`GC recv`, `GC srv`, `INC recv`, `INC sent`, `INC KB`, `INC sntKB`, `Prd`, `Pwr`)
  - `[tac_char]` — Per-node Cluster Cache (`JC_CCC_%`) operation count comparison
  - `[tac_msg]` — Per-node Interconnect (`JC_INC_%`) operation count comparison
  - `[tac_gccr_srv]` — Per-node Cluster Cache CR (`JC_CCC_CR_%`) count comparison
  - `[pga_hit]` — Per-node workarea `in-mem %` and `Ex/PGA %` comparison

- **`Full` (adds 14 cluster sheets on top of `Simple` = 23 sheets total)**:
  - `[param_sga]` — Side-by-side memory/SGA parameter comparison across nodes
  - `[tac_gc_efficiency]` — Per-node buffer access efficiency (`local %`, `remote %`, `disk %`)
  - `[wait5]` — Per-node `DBT(s)`, `DBCPU(s)`, and non-idle `Waits(s)` comparison
  - `[bufpool]` — Per-node buffer cache counters (`LReads`, `PReads`, `PWrts`, `Chngs`, `FreeReq`, `FreeVct`, `DirtyScn`, `FBWt(ms)`, `BsyWt(ms)`)
  - `[undo]` — Per-node undo blocks (`UndoBlks`), transaction counts (`TxCnt`), and max concurrency (`MaxConc`)
  - `[ckpt_redo]` — Per-node checkpoint, instance recovery, and redo metrics (`IncCkpt`, `InstRecv`, `Recv(ms)`, `RedoMB`, `LgwrWrt`, `RdoWait(ms)`)
  - `[sess_cache]` — Per-node distinct ASH sampled session count comparison
  - `[redo]` — Cluster-wide redo log group and member configuration across all threads
  - `[log_ckpt]` — Per-thread redo log switch interval statistics (switches, min/max/avg gap)
  - `[logswitch]` — Hourly redo log switch count matrix across all threads
  - `[log_switch_f]` — Per-thread log switch frequency buckets (`<= 5 min`, `5 ~ 10 min`, `10 ~ 30 min`, `> 30 min`, `Total`)
  - `[log_switch]` — Daily 24-hour (`00`..`23`) cluster-wide redo log switch matrix
  - `[archsize]` — Daily archived log count and size (MB) across all threads
  - `[arch]` — Cluster-wide archived redo log list within the collection window

## Output file naming

```
TPR_<NODE|TAC>_<tag>_<host>_<inst>_<yyyymmdd_hhmiss>.html
```

- `NODE` = Single-family ini, `TAC` = 2..8node ini
- `<tag>` = ini filename minus `TprOne_`, compressed for brevity: `_Single` is dropped
  for the NODE family and `_with_ash` becomes `_ash`
  (e.g. `Simple_TprOne_Single.ini` → `Simple`,
  `Full_TprOne_Single_with_ash.ini` → `Full_ash`,
  `Simple_TprOne_4node.ini` → `Simple_4node`)
- The timestamp includes seconds (`hhmiss`) so reruns within the same minute
  never overwrite earlier output
- Examples: `TPR_TAC_Simple_2node_tac1_prod1_20261117_201453.html`,
  `TPR_NODE_Full_ash_tac1_prod1_20261117_181040.html`
- ASH sections (`SPOOL=ash`) spool to a separate `..._ASH.html`

## Chart title variables

`@@hostname@@` and `@@instance@@` in a CHART title (used in Single-instance INIs)
are replaced from the input file name at conversion time (e.g.
`DB Time (s) - @@hostname@@` → `... - tac1`). Multi-node TAC charts omit
`@@hostname@@` because they plot all cluster nodes together.

## INI syntax quick reference

- Each `[section]` appears twice: a config block (`SQLn=`, `FORMAT`, `CHART`,
  `SPOOL=`) and a query block (starts with `select`). Only the query block goes
  into the SQL output; a `rem [section]` marker precedes it.
- Bind variables: `:snapdate_from` / `:snapdate_to` (requested range),
  `:snap_fr` / `:snap_to` (resolved local snapshot ids), `:gid_fr` / `:gid_to`
  (resolved TAC global snapshot ids), `:inst_no`, `:inst_no1..8` (`0..7`),
  `:dbid`.
- Chart anchor rule: the CHART column field = last data column of the sheet + 2.
  When you add/remove columns, shift the CHART/FORMAT column references too.
