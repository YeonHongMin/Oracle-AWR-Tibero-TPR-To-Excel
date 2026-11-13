# DB Performance Reports to Excel

DB 성능 리포트 데이터(Oracle AWR, Tibero TPR 예정)를 지표별 차트가 포함된
Excel 워크북으로 변환하는 프로젝트입니다.

> English version: [README.md](README.md)

## 하위 프로젝트

| 디렉터리 | 상태 | 설명 |
|---|---|---|
| [`Oracle/`](Oracle/) | **사용 가능** | AwrONE — Oracle AWR → Excel. INI로 수집 항목을 정의하고 SQL*Plus HTML 스풀로 수집한 뒤 차트 포함 `.xlsx`로 변환. 문서: [Oracle/README_ko.md](Oracle/README_ko.md) (English: [Oracle/README.md](Oracle/README.md)) |
| [`Tibero-Developing/`](Tibero-Developing/) | 개발 중 | TprONE — Tibero TPR → Excel. 플레이스홀더이며 아직 사용할 수 없습니다. |
