# 🏗️ ARCHITECTURE.md — 코드 구조 및 함수 레퍼런스

> **farmos-cropweather** 프로젝트의 모듈 구성, 데이터 흐름, 함수 레퍼런스를 정리한 개발자용 문서입니다.

---

## 📌 목차

1. [모듈 의존 관계](#1-모듈-의존-관계)
2. [데이터 흐름](#2-데이터-흐름)
3. [solar_core.py — ASOS 수집·캐싱 엔진](#3-solar_corepy)
4. [temp_core.py — 온도·GDD 계산 엔진](#4-temp_corepy)
5. [solar_multi.py — 일사·일조 엑셀 보고서](#5-solar_multipy)
6. [temp_multi.py — 온도·적산온도 엑셀 보고서](#6-temp_multipy)
7. [공유 모듈](#7-공유-모듈)
8. [엑셀 VLOOKUP 아키텍처](#8-엑셀-vlookup-아키텍처)
9. [확장 가이드](#9-확장-가이드)

---

## 1. 모듈 의존 관계

```
┌─────────────────┐    ┌─────────────────┐
│  solar_multi.py │    │  temp_multi.py  │   ← 엑셀 보고서 생성기
└────────┬────────┘    └───┬─────────┬───┘
         │                 │         │
         │     ┌───────────┘         │
         ▼     ▼                     ▼
   ┌──────────────┐          ┌──────────────┐
   │ solar_core.py│          │ temp_core.py  │   ← 계산 엔진
   └──────┬───────┘          └──────┬───────┘
          │                         │
          │       ┌─────────────────┘
          ▼       ▼
   ┌──────────────┐    ┌──────────────┐    ┌──────────────┐
   │ stations.py  │    │ excel_util.py│    │crops_gdd.csv │
   └──────────────┘    └──────────────┘    └──────────────┘
        지점·권역            엑셀 스타일         작물 Tbase
```

**핵심 의존성:**

- `temp_multi.py`는 **`solar_core.py`를 직접 import**하여 ASOS 데이터 수집·파싱·캐싱을 재사용한다. 별도의 API 호출 코드를 갖지 않는다.
- `temp_core.py`는 `solar_core.parse_rows()`가 반환하는 딕셔너리를 입력으로 받는다. `solar_core`를 직접 import하지는 않는다 (느슨한 결합).
- `stations.py`와 `excel_util.py`는 solar/temp 양쪽에서 공통 사용한다.

---

## 2. 데이터 흐름

### 2.1 수집 → 캐싱 → 파싱

```
공공데이터포털 ASOS API
        │
        ▼
solar_core.fetch_all_stations()
        │
        ├─ 확정 연도 → data/raw/{stnId}_{startYr}_{endYr}.csv (캐시)
        ├─ 올해      → API 재조회 → 캐시 덮어쓰기
        │
        ▼
solar_core.parse_rows()
        │
        ▼
  [{"year": 2026, "month": 7, "day": 15,
    "avg_ta": 25.3, "max_ta": 31.2, "min_ta": 21.0,
    "sum_gsr": 18.5, "sum_ss_hr": 8.2, ...}, ...]
```

### 2.2 solar 분석 흐름

```
parse_rows → _daily_average_rows → _monthly_metric_by_year
                                          │
                                          ▼
                                   _build_region_ref_sheet  →  _권역별REF (숨김)
                                   _build_station_ref_sheet →  _지점별REF (숨김)
                                          │
                                          ▼
                                   VLOOKUP 수식으로 표시 시트 구성
```

### 2.3 temp 분석 흐름

```
parse_rows → add_gdd_to_rows (gdd 키 추가)
                 │
                 ├─ _daily_average_rows → monthly 집계 → _권역별REF / _지점별REF
                 │
                 └─ gdd_reaching_dates → _store_reaching_to_ref
                         │                       │
                         ▼                       ▼
                   gdd_reaching_dates_normal → _지역달성REF / _지점달성REF
                                                    │
                                                    ▼
                                             VLOOKUP 수식으로 표시 시트 구성
```

---

## 3. solar_core.py

ASOS 데이터 수집·캐싱·파싱의 핵심 엔진. solar와 temp이 공유한다.

### 상수

| 상수 | 값 | 설명 |
|------|-----|------|
| `DATA_START` | `date(2016, 1, 1)` | 수집 시작일 |
| `MAX_QUERY_YEARS` | `2` | API 1회 최대 조회 기간(년) |
| `MAX_NUM_OF_ROWS` | `999` | API 1회 최대 응답 건수 |
| `CACHE_DIR` | `data/raw/` | CSV 캐시 디렉터리 |

### 함수 레퍼런스

#### 인증·유틸

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `load_apikeys` | `(path="apikey.txt")` | `dict[str, str]` | `라벨=값` 형식 파일에서 인증키 로드 |
| `yesterday` | `()` | `date` | D-1 날짜 |
| `safe_float` | `(val, default=None)` | `float\|None` | 안전한 float 변환 |

#### 데이터 수집

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `get_date_chunks` | `(start, end, max_years)` | `list[(date,date)]` | 날짜 범위를 N년 단위 청크로 분할 |
| `is_finalized_year` | `(year)` | `bool` | 확정 연도 여부 (year < 올해) |
| `fetch_station_data` | `(key, stn_id, start, end, use_cache, verbose)` | `list[dict]` | 단일 지점 전체 기간 수집 (자동 분할·캐싱) |
| `fetch_all_stations` | `(key, stn_ids, start, end, use_cache, verbose)` | `dict[int, list]` | 여러 지점 일괄 수집 |

#### 파싱·집계

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `parse_rows` | `(raw_rows)` | `list[dict]` | API/CSV 원시 데이터 → 분석용 dict 변환 |
| `aggregate_monthly` | `(parsed, metrics)` | `dict` | 일별 → 월별 집계 |
| `aggregate_yearly` | `(parsed, metrics)` | `dict` | 일별 → 연도별 집계 |
| `calc_normal` | `(monthly, base_year, n_years, metrics)` | `dict` | 평년(N년 평균) 산출 |

**`parse_rows()` 반환 딕셔너리 키:**

| 키 | 타입 | 출처 |
|----|------|------|
| `year`, `month`, `day` | `int` | 날짜 파싱 |
| `date` | `date` | 날짜 객체 |
| `stn_id` | `int` | 지점번호 |
| `avg_ta`, `max_ta`, `min_ta` | `float\|None` | 기온 (℃) |
| `sum_gsr` | `float\|None` | 합계 일사량 (MJ/m²) |
| `sum_ss_hr` | `float\|None` | 합계 일조시간 (hr) |
| `ss_dur` | `float\|None` | 가조시간 (hr) |
| `ss_rate` | `float\|None` | 일조율 (%) |
| `hr1_max_icsr` | `float\|None` | 1h 최다 일사 (MJ/m²) |
| `sum_rn` | `float\|None` | 일강수량 (mm) |

---

## 4. temp_core.py

온도·적산온도(GDD) 계산 전용 모듈. `solar_core.parse_rows()`의 반환값을 입력으로 받는다.

### 상수

| 상수 | 값 | 설명 |
|------|-----|------|
| `DEFAULT_T_BASE` | `10.0` | 기본 기준온도 (℃) |
| `DEFAULT_T_UPPER` | `35.0` | 기본 상한온도 (℃) |
| `DEFAULT_METHOD` | `"maxmin"` | 기본 GDD 계산 방법 |
| `DEFAULT_MILESTONES` | `[500, 1000, ..., 4000]` | 기본 도달일자 임계값 |
| `CROPS_GDD_CSV` | `crops_gdd.csv` | 작물 라이브러리 경로 |

### 함수 레퍼런스

#### 작물 라이브러리

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `load_crops_gdd` | `(csv_path=None)` | `dict[str, dict]` | crops_gdd.csv 로드. 키=crop_id |
| `list_crops` | `(csv_path=None)` | `list[tuple]` | (crop_id, 한글명, Tbase) 리스트 |

#### GDD 계산

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `calc_gdd` | `(avg_ta, max_ta, min_ta, t_base, t_upper, method)` | `float\|None` | 단일 일 GDD 계산 |
| `add_gdd_to_rows` | `(rows, t_base, t_upper, method)` | `list[dict]` | 행 리스트에 `gdd` 키 in-place 추가 |

**`calc_gdd` method 파라미터:**

| 값 | 계산식 | 비고 |
|----|--------|------|
| `"maxmin"` | `max((Tmax+Tmin)/2 − Tbase, 0)` | **기본값.** FAO56rev 표준 |
| `"avg"` | `max(avgTa − Tbase, 0)` | 기상청 관행 통계 비교 시 |
| `"aquacrop"` | 방법1 + Tupper 반영 | 열대·아열대 지역 |

#### 집계

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `daily_average_rows` | `(rows)` | `list[dict]` | 다중 지점 → 일별 지점평균 |
| `monthly_stats` | `(rows, metric_key, agg_type)` | `{year: {month: val}}` | 월별 집계 |
| `period_stats` | `(rows, metric_key, agg_type)` | `{year: val}` | 기간 전체 집계 |

`agg_type`: `"avg"` (평균), `"cumsum"` (합계), `"max"` (최댓값), `"min"` (최솟값)

#### 평년

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `calc_normal_value` | `(year_vals, base_year, n)` | `float\|None` | 연도별 → 평년 단일값 |
| `calc_normal_monthly` | `(monthly_by_year, base_year, n)` | `{month: val}` | 연도별 월별 → 평년 월별 |

#### 적산온도 도달

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `gdd_reaching_dates` | `(rows, t_base, t_upper, method, milestones, start_month, start_day)` | `{year: {ms: "M-D"\|None}}` | 연도별 도달일자 산출 |
| `gdd_reaching_dates_normal` | `(yearly_dates, base_year, n)` | `{ms: "M-D"\|None}` | 평년 도달일자 (DOY 평균) |

#### 편의 함수

| 함수 | 시그니처 | 반환 | 설명 |
|------|---------|------|------|
| `filter_same_period` | `(rows, start_md, end_md)` | `list[dict]` | MM-DD 동기간 필터 |
| `analyze_temperature` | `(parsed_rows, ...)` | `dict` | 한번에 전체 분석 (집계+도달일자+평년) |

---

## 5. solar_multi.py

7시트 일사·일조 엑셀 보고서 생성기.

### 엔트리포인트

| 함수 | 설명 |
|------|------|
| `generate_report(...)` | 데이터 수집 → 7시트 엑셀 생성 → 파일 저장 |
| `main()` | CLI 파서 → `generate_report` 호출 |

### 내부 함수 (시트별)

| 시트 | 생성 함수 | REF 시트 | 헬퍼 |
|------|----------|---------|------|
| 설정 | `_sheet_settings` | — | — |
| 권역별 | `_sheet_region_summary` | `_build_region_ref_sheet` | `_write_region_vlookup_block` |
| 지점별 | `_sheet_station_summary` | `_build_station_ref_sheet` | `_write_station_vlookup_block` |
| 기간별(월별) | `_sheet_period_analysis` | (권역별REF 공유) | `_write_period_vlookup_block` |
| 년도별 | `_sheet_yearly_analysis` | (권역별REF 공유) | `_write_yearly_vlookup_block` |
| 원데이터 | `_sheet_raw_data` | — | — |

---

## 6. temp_multi.py

9시트 온도·적산온도 엑셀 보고서 생성기. solar_multi.py와 동일한 VLOOKUP 아키텍처.

### 엔트리포인트

| 함수 | 설명 |
|------|------|
| `generate_report(...)` | 데이터 수집 → GDD 계산 → 9시트 엑셀 생성 → 파일 저장 |
| `main()` | CLI 파서 (`--crop`, `--tbase`, `--method` 등) → `generate_report` 호출 |

### 내부 함수 (시트별)

| 시트 | 생성 함수 | REF 시트 | 헬퍼 |
|------|----------|---------|------|
| 설정 | `_sheet_settings` | — | — |
| 권역별 | `_sheet_region_summary` | `_build_region_ref_sheet` | `_write_region_vlookup_block` |
| 지점별 | `_sheet_station_summary` | `_build_station_ref_sheet` | `_write_station_vlookup_block` |
| 기간별(월별) | `_sheet_period_analysis` | (권역별REF 공유) | `_write_period_vlookup_block` |
| 년도별 | `_sheet_yearly_analysis` | (권역별REF 공유) | `_write_yearly_vlookup_block` |
| 적산온도 도달(권역별) | `_sheet_reaching_dates_region` | `_build_reaching_ref_region` | `_write_reaching_vlookup_block` |
| 적산온도 도달(지점별) | `_sheet_reaching_dates_station` | `_build_reaching_ref_station` | `_write_reaching_vlookup_block` |
| 원데이터 | `_sheet_raw_data` | — | — |

### solar와의 차이

| 항목 | solar_multi | temp_multi |
|------|-------------|------------|
| 지표 | 일사·일조 5종 | 기온 4종 + 최고/최저 2종 = 6종 |
| 설정 파라미터 | 올해, 평년기간, 분석기간 | + Tbase, Tupper, method, crop |
| 기간별(월별) | 동일 서식 | 동일 서식 |
| 년도별 | 고정 시작년 | **슬라이딩 윈도우** (올해−N년 ~ 올해+5년) |
| 적산온도 도달 | 없음 | **2시트** (권역별/지점별, VLOOKUP 동적) |
| REF 시트 | 2개 | 4개 (기온 2 + 도달 2) |

---

## 7. 공유 모듈

### stations.py

ASOS 66개 관측 지점과 9개 권역의 매핑을 관리한다.

| 함수/상수 | 설명 |
|----------|------|
| `REGION_ORDER` | 권역 표시 순서 리스트: `["전국", "서울경기", "강원영동", ...]` |
| `station_name(stn_id)` | 지점번호 → 지점명 (예: `108` → `"서울"`) |
| `region_of(stn_id)` | 지점번호 → 소속 권역 (예: `108` → `"서울경기"`) |
| `stations_in_region(region)` | 권역 → 소속 지점번호 리스트 |
| `all_station_ids()` | 전체 66개 지점번호 리스트 |

### excel_util.py

openpyxl 기반 공통 엑셀 스타일·서식 유틸리티.

| 함수/상수 | 설명 |
|----------|------|
| `H(ws, row, col, text, ...)` | 헤더 셀 생성 (녹색 배경, 굵은 글꼴) |
| `C(ws, row, col, value, ...)` | 데이터 셀 생성 (테두리, 정렬) |
| `apply_title(ws, row, col, text, ...)` | 병합 타이틀 행 |
| `write_header_row(ws, row, headers)` | 헤더 행 일괄 작성 |
| `write_data_row(ws, row, values, ...)` | 데이터 행 일괄 작성 |
| `set_col_widths(ws, widths)` | 열 너비 일괄 설정 |
| `FONT_NAME`, `BORDER`, `FMT_NUM1`, ... | 공통 스타일 상수 |

### crops_gdd.csv

41개 작물의 GDD 파라미터. `temp_core.load_crops_gdd()`로 로드.

| 컬럼 | 설명 |
|------|------|
| `crop_id` | 식별자 (cropwater `crops_library.csv`와 공유 키) |
| `crop_name_ko` / `crop_name_en` | 작물명 |
| `t_base` | 기준온도 (℃) |
| `t_upper` | 상한온도 (℃), 빈칸 가능 |
| `gdd_milestones` | 도달일자 산출 임계값 (쉼표 구분) |
| `source` | 출처 약칭 |

---

## 8. 엑셀 VLOOKUP 아키텍처

### 8.1 설계 원칙

엑셀에서 "설정" 시트의 "올해" 셀(`B4`)을 바꾸면, **파이썬 재실행 없이** 모든 시트가 자동 갱신되어야 한다.

이를 위해 **숨긴 REF 시트 + VLOOKUP** 패턴을 사용한다.

### 8.2 구조

```
┌──────────────────┐        ┌─────────────────────────┐
│    설정 시트       │        │    _권역별REF (숨김)      │
│  B4 = 올해 (2026) │        │  A열: 복합키             │
│  B5 = B4-1       │        │  B열: 값                 │
│  B6 = 평년기간    │        │                          │
└──────────────────┘        │  평균기온_전국_2026_1      │
         │                  │  평균기온_전국_2026_2      │
         │                  │  ...                     │
         ▼                  │  평균기온_전국_평년2026_1   │
┌──────────────────┐        │  ...                     │
│    표시 시트       │        └─────────────────────────┘
│  =VLOOKUP(       │                    ▲
│   "평균기온_"     │────── VLOOKUP ─────┘
│   &B$1           │
│   &"_"&설정!$B$4 │
│   &"_1",         │
│   _권역별REF!A:B,│
│   2, FALSE)      │
└──────────────────┘
```

### 8.3 복합키 형식

**기온·일사 REF:**

```
{지표}_{subject}_{year}_{month}      → 월별 값
{지표}_{subject}_{year}_연           → 연간 값
{지표}_{subject}_평년{baseYear}_{month} → 평년 월별 값
```

예: `평균기온_전국_2026_7` = 2026년 7월 전국 평균기온

**적산온도 도달 REF:**

```
도달일_{ms}℃_{subject}_{year}       → "M-D" 문자열
도달DOY_{ms}℃_{subject}_{year}      → DOY 정수
도달일_{ms}℃_{subject}_평년{baseYear} → 평년 도달일
도달DOY_{ms}℃_{subject}_평년{baseYear} → 평년 DOY
```

예: `도달일_500℃_전국_평년2026` = 2016~2025 평균 도달일

### 8.4 동적 갱신 범위

| 시트 | 올해 변경 시 |
|------|------------|
| 권역별 / 지점별 | 올해·작년·평년 3행 + 차이 2행 자동 갱신 |
| 기간별(월별) | 올해·작년·평년 블록 자동 갱신 |
| 년도별 | 슬라이딩 윈도우 — 범위 밖 행은 빈칸 처리 |
| 적산온도 도달 | 올해·작년·평년·차이 모두 자동 갱신 |

---

## 9. 확장 가이드

### 9.1 새 작물 추가

`crops_gdd.csv`에 행 1개를 추가한다. 코드 변경 불필요.

```csv
jujube,대추,Jujube,10,35,"500,1000,1500,2000",농진청
```

### 9.2 새 기온 지표 추가

1. `temp_multi.py`의 `_build_region_ref_sheet`와 `_build_station_ref_sheet`의 `metrics` 리스트에 `("지표라벨", "parse_rows키", "집계방법")` 튜플 추가
2. 표시 시트 함수(`_sheet_region_summary` 등)의 `metrics` 리스트에 동일 항목 추가
3. `_daily_average_rows`의 `keys` 리스트에 새 키 추가

### 9.3 새 엑셀 시트 추가

1. `_sheet_새시트(wb, ...)` 함수 작성 (기존 시트 함수 패턴 참고)
2. 필요 시 REF 시트 + VLOOKUP 블록 헬퍼 추가
3. `generate_report()`에서 적절한 순서에 호출 추가

### 9.4 solar → temp 코드 재사용 패턴

solar_multi.py에서 새 기능을 구현한 뒤 temp_multi.py에 동일 기능을 추가할 때:

1. `_daily_average_rows`, `_monthly_metric_by_year` 등 집계 유틸은 양쪽에 동일 코드로 존재한다 (향후 `farmos-commons`로 통합 예정)
2. VLOOKUP 블록 헬퍼(`_write_*_vlookup_block`)는 지표명·REF 시트명만 다르고 구조가 동일하다
3. 설정 시트의 행 번호(`B4`=올해, `B5`=작년, `B6`=평년기간, `B8`=시작일, `B9`=종료일)는 양쪽에서 **하드코딩으로 참조**하므로 반드시 일치해야 한다

---

## 📋 변경 이력

| 날짜 | 버전 | 변경 내용 |
|------|------|----------|
| 2026-09-17 | 1.0 | 초판 작성 |
