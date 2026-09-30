# 프로그램 구조 및 개발자 가이드

> CropWater의 파일 구조, 함수 레퍼런스, 데이터 흐름, 확장 방법을 다룹니다.

---

## 목차

1. [전체 구조 개요](#1-전체-구조-개요)
2. [데이터 흐름](#2-데이터-흐름)
3. [fao56_core.py — 공통 모듈](#3-fao56_corepy--공통-모듈)
4. [cropwater_station.py — 단일 지점](#4-cropwater_stationpy--단일-지점)
5. [cropwater_multi.py — 다지점 비교](#5-cropwater_multipy--다지점-비교)
6. [crops_library.csv — 작물 파라미터 DB](#6-crops_librarycsv--작물-파라미터-db)
7. [확장 가이드](#7-확장-가이드)
8. [설계 원칙](#8-설계-원칙)

---

## 1. 전체 구조 개요

```
cropwater/
├── fao56_core.py          ← 공통 라이브러리 (모든 파일이 import)
│     ├── FAO-56 물리식    ← 순수 함수 (입출력만, 부작용 없음)
│     ├── KMA API 조회     ← requests 기반
│     └── 데이터 유틸      ← 작물 라이브러리·지점 캐시·파일 저장
│
├── cropwater_station.py   ← 단일 지점 심층 분석
│     └── build_workbook() → 6시트 Excel
│
├── cropwater_multi.py     ← 다지점 스크리닝
│     ├── compute_station()        → ETo·Epan
│     ├── compute_water_balance()  → Dr/Ks/DP
│     ├── aggregate_monthly()      → 월별 집계
│     └── build_calendar_sheet()  → 히트맵 + 차트
│
└── (02-Cycle) 단기예보 ETo·ETc 예측 — 7장
      ├── cropwater_fcst.py  ← CLI: calib(Rs 계수) / verify(H2 검증 엑셀)
      ├── fcst_archive.py    ← 과거 단기예보 CSV 파싱 → 발표별 일 입력
      ├── rs_model.py        ← Rs 추정(식50 + 강수유무), rs_coef.csv 2계층
      ├── obs_daily.py       ← 01-Cycle 워크북 → 관측 ETo·Kc
      └── fcst_report.py     ← 검증 엑셀(10시트, 라이브 수식)
```

**의존 관계** — 실행 파일은 서로 독립적이며 `fao56_core`를 공유합니다. 02-Cycle 모듈은 01-Cycle 출력 워크북(관측 기준값)을 입력으로 씁니다.

```
cropwater_station ──┐
                    ├──→ fao56_core ──→ requests, openpyxl
cropwater_multi   ──┘        ▲
                             │
cropwater_fcst ──→ fcst_archive, rs_model, obs_daily, fcst_report ──→ pandas, numpy, openpyxl
```

---

## 2. 데이터 흐름

### ◆ cropwater_station.py

```
CLI 또는 대화형 입력
        │
        ▼
fetch_station_table()     ← 기상청 API허브: 위도·고도·풍속계높이
fetch_asos()              ← 공공데이터포털: ASOS 일자료
        │
        ▼ (일별 list[dict])
build_workbook()
  ├─→ [설정]      파라미터 표시
  ├─→ [원데이터]  ASOS 원자료
  ├─→ [계산과정]  라이브 수식 (svp·eto_pm·kc_of_date)
  ├─→ [결과요약]  집계 수식
  ├─→ [물수지]    라이브 수식 (Dr 연쇄 참조)
  └─→ [계산근거]  정적 텍스트
        │
        ▼
safe_save() → Excel 파일
```

### ◆ cropwater_multi.py

```
CLI 입력
        │
        ▼
fetch_station_table()     ← 기상청 API허브
        │
        ▼ (지점 반복)
fetch_asos()
  → compute_station()
  → compute_water_balance()
  → aggregate_monthly()
        │
        ▼
[지점비교] [일별_전지점] [월별_물수지] [관수필요_달력] [설명]
        │
safe_save() → Excel 파일
```

---

## 3. fao56_core.py — 공통 모듈

### ◆ FAO-56 물리식 함수

| 함수 | 입력 | 출력 | FAO-56 |
| :--- | :--- | :--- | :--- |
| `svp(T)` | 기온 T (°C) | 포화증기압 (kPa) | 식(11) |
| `slope_svp(T)` | 기온 T (°C) | Δ (kPa/°C) | 식(13) |
| `extra_radiation(lat, J)` | 위도(°), 연중일수 | Ra (MJ/m²/day) | 식(21~25) |
| `daylight_hours(lat, J)` | 위도(°), 연중일수 | N (hr) | 식(34) |
| `wind_2m(uz, zw)` | 측정풍속, 측정높이 | u2 (m/s) | 식(47) |
| `kp_class_a(u2, RH, fetch)` | 풍속·습도·풍상거리 | Kp (무차원) | Table 30 |
| `eto_penman_monteith(Tmax, Tmin, Rs, u2, ea, elev, lat, J, Pa)` | 기상값 | ETo (mm/day) | 식(6) |
| `kc_climate_adjust(kc_tab, u2, rhmin, h)` | 표값 Kc·기상값 | Kc_adj | 식(62·65) |
| `stage_of_date(day_ordinal, start_ordinal, L_ini, L_dev, L_mid, L_late)` | ordinal, 생육일수 | 생육단계명 | Table 11 |
| `kc_of_date(day_ordinal, start_ordinal, L_ini, ..., kc_ini, kc_mid, kc_end)` | ordinal, 생육일수, Kc값 | Kc | 그림 25 |

**사용 예시:**
```python
from fao56_core import eto_penman_monteith, kc_of_date, wind_2m
import datetime as dt

eto = eto_penman_monteith(
    Tmax=28.5, Tmin=18.2, Rs=22.0,
    u2=wind_2m(2.0, zw=10),
    ea=1.85, elev=75.82, lat=37.90, J=196, Pa=1003.5  # Pa: hPa 단위
)
# → ETo ≈ 5.23 mm/day

kc = kc_of_date(
    day_ordinal=dt.date(2026, 7, 15).toordinal(),
    start_ordinal=dt.date(2026, 4, 1).toordinal(),
    L_ini=20, L_dev=70, L_mid=90, L_late=30,
    kc_ini=0.50, kc_mid=1.20, kc_end=0.95
)
# → Kc = 1.200 (중기)
```

### ◆ API 조회 함수

| 함수 | 설명 | 반환값 |
| :--- | :--- | :--- |
| `fetch_asos(key, stn, start, end)` | 공공데이터포털 ASOS 일자료 | `list[dict]` |
| `fetch_station_table(hubkey)` | 기상청 API허브 지점일람표 | `(info_dict, name_dict)` |
| `load_station_backup()` | 로컬 캐시 읽기 | `(info_dict, name_dict)` |
| `save_station_backup(info, names)` | 로컬 캐시 저장 | `None` |
| `load_apikeys(path)` | apikey.txt 파싱 | `dict` |

`fetch_asos` 반환값 — 1일 1 dict 구조:

```python
{
    "tm":       "2026-07-01",  # 일자
    "maxTa":    "28.5",        # 최고기온 (°C)
    "minTa":    "18.2",        # 최저기온 (°C)
    "avgRhm":   "75.0",        # 평균상대습도 (%)
    "avgWs":    "1.8",         # 평균풍속 (m/s)
    "avgPv":    "20.5",        # 평균증기압 (hPa)
    "avgTd":    "18.5",        # 평균이슬점온도 (°C)
    "avgPa":    "1003.5",      # 평균기압 (hPa)
    "sumGsr":   "22.0",        # 일사량 합계 (MJ/m²)
    "sumSsHr":  "8.5",         # 일조시간 (hr)
    "sumLrgEv": "6.2",         # 대형증발량 (mm)
    "sumRn":    "0.0",         # 강수량 합계 (mm)
}
```

### ◆ 작물 라이브러리 함수

| 함수 | 설명 |
| :--- | :--- |
| `load_crop_library(path)` | `crops_library.csv` → `dict[crop_id → dict]` |
| `crops_sorted(lib)` | 카테고리·순서 기준 정렬 리스트 반환 |
| `save_crop_override(crop_dict, path)` | 사용자 Kc 오버라이드 저장 |

### ◆ 유틸리티 함수

| 함수 | 설명 |
| :--- | :--- |
| `num(v)` | 문자열·None → float 안전 변환 |
| `safe_save(wb, path)` | 파일명 충돌 시 자동 번호 부여 저장 |

---

## 4. cropwater_station.py — 단일 지점

### ◆ build_workbook(rows, p, out_path)

ASOS 일자료(`rows`)와 파라미터 dict(`p`)를 받아 6시트 Excel을 생성합니다.

```python
p = {
    'lat': 37.90,               # 위도
    'elev': 75.82,              # 해발고도 (m)
    'anem': 10.0,               # 풍속계 높이 (m)
    'fetch': 100.0,             # 증발접시 풍상거리 (m)
    'stn': '101',               # 지점 번호
    'start': '20260401',        # 시작일
    'end': '20260831',          # 종료일
    'meta_source': 'API',       # 지점정보 출처
    'crop': crop_dict,          # load_crop_library() 반환값
    'bud_date': date(2026,4,1), # 생육 시작일
    'is_short_cycle': False,    # 단기작물 여부
    'L_total': 210,             # 전체 생육기간 (일)
    'mulch': 1.0,               # 멀칭 보정계수
    'u2_mid': None,             # 중기 평균 u2 (None → 전체 평균)
    'rh_mid': None,             # 중기 평균 최저 RH
    'u2_end': None,             # 후기 평균 u2
    'rh_end': None,             # 후기 평균 최저 RH
}
```

### ◆ 엑셀 수식 구조

**[계산과정] 시트 주요 열:**

| 열 | 내용 | 비고 |
| :---: | :--- | :--- |
| A | 일자 | 원데이터 참조 |
| U | ETo_PM (mm/day) | FAO-56 PM 수식 전체 |
| V | 생육단계 | 중첩 IF |
| W | 일별 Kc | 생육단계별 선형보간 |
| AA | ETc_PM | `=U2*W2` |

**[물수지] 시트 수식 체인:**

```
G2 = 설정!DR0          ← 초기 고갈량 (첫 행)
G3 = L2                ← 전일 Dr,i (이후 행)

H  = MAX(F-K-G, 0)     [식88] DP = P − ETc_adj − Dr,i-1 ≥ 0 (FAO-56 원식, 2026-09-29)
I  = MAX(G-F, 0)       강수 후 고갈량 (Ks 판단용)
J  = IF(I<=RAW, 1, (TAW-I)/(TAW-RAW))   [식84] Ks
K  = J*E               ETc_adj = Ks × ETc
L  = MIN(MAX(G-F+K+H, 0), TAW)   [식85·86] Dr,i = Dr,i-1 − P + ETc_adj + DP
M  = IF(L>=RAW,"●","") 관수필요 판정
N  = IF(M="●",L,0)     필요 순관수량 In
O  = IF(N>0,N/Ea,0)    필요 총관수량 Ig
```

> Python이 값이 아닌 **수식 문자열**을 셀에 씁니다. 설정 시트의 파라미터를 바꾸면 물수지 시트 전체가 자동 재계산됩니다.

---

## 5. cropwater_multi.py — 다지점 비교

### ◆ compute_station(rows, lat, elev, anem, fetch)

ASOS 원자료 → ETo_PM·ETo_pan 계산.

반환: `(out_list, miss_si, miss_ev)`
- `out_list` 각 요소: `{date, PM, pan, Epan, rain}`

### ◆ compute_water_balance(recs, taw, raw)

FAO-56 식(85) 일별 Dr 추적. 기준작물(Kc=1) · 무관수 가정.
`recs`에 `{DP, Peff, Dr, Ks, irr, In}` 필드를 추가하여 반환합니다.

```python
Dr = 0.0
for rec in recs:
    P        = rec["rain"]
    ETo      = rec["PM"]
    Dr_after = max(Dr - P, 0)                       # Ks 판단용: 비가 먼저 들어간 뒤의 고갈량
    Ks       = 1.0 if Dr_after <= raw \
               else (taw - Dr_after) / (taw - raw)
    ETc_adj  = Ks * ETo
    DP       = max(P - ETc_adj - Dr, 0)             # [식88] FAO-56 원식 (2026-09-29)
    Dr       = min(max(Dr - P + ETc_adj + DP, 0), taw)   # [식85·86]
    rec["Dr"] = Dr
```

- 식(88)은 2026-09-29에 FAO-56 원식으로 바꿨습니다. 이전 간이식 `max(P − Dr, 0)`은 당일 ETc를 빼지 않아, 큰 비가 온 날 DP를 ETc만큼 크게 잡고 그날 끝 고갈량을 0이 아니라 ETc로 남겼습니다.
- Ks는 그날 비가 들어간 뒤의 고갈량으로 판단합니다. FAO-56 예시는 전날 끝 고갈량 Dr,i-1을 씁니다(비 온 날에만 차이). 판단 시점의 확정은 G4 점검 항목입니다.

### ◆ aggregate_monthly(recs)

일별 데이터 → 월별 집계.

반환: `{월번호: {P, Peff, DP, ETo, irr_days, stress_days, In}}`

### ◆ build_calendar_sheet(wb, allrows, stns_list, stn_names, taw, raw)

날짜(행) × 지점(열) Dr 히트맵 + Dr 선형 차트 생성.

- 색상 함수 `_dr_color(dr, raw, taw)`: `D4EDDA`(안전) → `FFF3CD`(주의) → `C82020`(심각) 그라데이션
- RAW 열은 차트 기준선 참조용으로 자동 숨김 처리

---

## 6. crops_library.csv — 작물 파라미터 DB

### ◆ 스키마

| 컬럼 | 타입 | 설명 | 출처 |
| :--- | :--- | :--- | :--- |
| `crop_id` | str | 고유 ID (영문 소문자·언더스코어) | — |
| `name_ko` / `name_en` | str | 한국어·영문 작물명 | — |
| `category` | str | 과수·소채류·곡류 등 | — |
| `kc_ini` / `kc_mid` / `kc_end` | float | Kc 초기·중기·후기 | FAO-56 Table 12 |
| `kc_ini_2~4` | float | 시나리오별 Kc (has_scenarios=1) | FAO-56 Table 12 |
| `L_ini` / `L_dev` / `L_mid` / `L_late` | int | 생육단계별 기간 (일) | FAO-56 Table 11 |
| `h_m` | float | 최대 작물 키 (m) | FAO-56 Table 12 |
| `zr` | float | 근권심도 (m) | FAO-56 Table 22 |
| `p` | float | 고갈계수 | FAO-56 Table 22 |
| `note` | str | 비고·출처 | — |

### ◇ 시나리오 (has_scenarios)

재배 방식(무멀칭/멀칭, 초생재배/청경)에 따라 복수의 Kc 세트를 제공하는 작물.
`has_scenarios=1`이면 `cropwater_station.py` 실행 시 대화형으로 선택합니다.

---

## 7. 확장 가이드

### ◆ 새 작물 추가

`crops_library.csv`에 행 하나를 추가하는 것으로 충분합니다. 값은 FAO-56 Table 11·12·22에서 확인합니다.

```csv
my_crop,내_작물명,My Crop,기타,99,0,0.40,1.10,0.85,,,,,25,35,40,20,0.5,0.50,개인 실측값
```

### ◆ 토성 기본값 변경

실행 시 `--fc`, `--wp` 파라미터로 직접 지정하거나, `cropwater_multi.py` 상단 `add_argument` 기본값을 수정합니다.

| 토성 | FC | WP | AWC (mm/m) |
| :--- | :---: | :---: | :---: |
| 사질토 | 0.10 | 0.04 | 60 |
| 사양토 | 0.14 | 0.06 | 80 |
| 양토 | 0.22 | 0.10 | 120 |
| 식양토 | 0.30 | 0.15 | 150 |
| 식토 | 0.36 | 0.20 | 160 |

### ◇ 실측 관수량 연동 (다음 버전 예정)

현재 물수지는 무관수(자연강우만) 가정입니다. FAO-56 식(85)의 `I` 항에 실측 관수량을 입력하면 Dr이 실측 기반으로 계산됩니다. 데이터 소스는 T21 센서 로그 또는 FarmOS 플랫폼 예정.

```python
# 예정 인터페이스
Dr_end = min(Dr_after - I_net + ETc_adj, taw)
# I_net: 해당 일 실측 순관수량 (mm)
```

### ◇ 단기예보 ETo·ETc 예측 (02-Cycle 진행 중)

기상청 단기예보로 ETo·ETc를 예측하고, 물수지 전망으로 관수 필요 예상일을 알립니다.
- 이론·검증 설계: [THEORY.md 9장](THEORY.md#9-예보-기반-etoetc--기상청-단기예보)
- 게이트 기록: [VALIDATION.md](VALIDATION.md)
- 가설검증 범위: ASOS 101 춘천 · 사과

**서비스 요구사항 (2026-09-29 확정)**

| 갱신 | 실행 시각 | 사용 발표 | 대상 기간 |
| :--- | :--- | :--- | :--- |
| 아침 | 02:10 이후 | 02시 | 오늘 ~ D+3 |
| 저녁 | 17:10 이후 | 17시 | 내일 ~ D+4 |

출력은 엑셀 보고서입니다. 02-Cycle에서는 가설검증용 엑셀을 중심으로 합니다.

**입력 대응 (정정)** — TMN/TMX/REH/WSD만으로는 PM을 계산할 수 없습니다. 단기예보에는 **일사량(Rs)이 없으므로** 추정합니다. 하늘상태(SKY)가 있으면 S4(식50형 + 강수유무 + 구름 비율, 선행일별 계수), 없으면 S3(식50 + 강수유무)입니다. 기압은 고도 기반(식7)으로 대체합니다.

**모듈 (G1~G3 구현, 2026-09-29 SKY 재검증까지)**

| 모듈 | 역할 | 단계 | 상태 |
| :--- | :--- | :---: | :---: |
| `fcst_archive.py` | 과거 단기예보 CSV → 요소별 표(발표, 선행시간, 대상시각, 값, 코드 여부) → 서비스 발표(02·17시)별 일 입력. 입력 형식 세 가지를 첫 줄로 자동 판별(`csv_format`): 포털 CSV(요소별, 여러 달 파일 포함), OpenAPI 응답 CSV(한 파일에 모든 요소, `read_openapi_csv`), 요소별 KST CSV(`read_element_csv`). 쓰지 않는 요소 파일은 건너뛰고 기록. 선택 요소 SKY·POP은 낮 시간 일사 비중으로 가중한 일 값(`sun_weights`). `fcst_daily.py` 계획을 합침 | 1 | 구현 |
| `rs_model.py` + `rs_coef.csv` + `rs_sky_coef.csv` | Rs 추정(S1 식50, S3 식50 + 강수유무, S4 + 하늘상태 구름 비율), 최소제곱 적합, 계수 2계층(stn=0 FAO 기본값 + 지점 행). S4는 지점 × 선행일 행. `rs_coef.csv`는 2025년 관측 S3 계수(2026년 검증용), `rs_coef_2026.csv`는 2026년 관측 S3 계수(2025년 독립 검증용) | 2·3 | 구현 |
| `obs_daily.py` | 01-Cycle 출력 워크북 → 관측 일자료, 관측 ETo(01-Cycle 규칙), Kc(설정 시트 값) | 2 | 구현 |
| `cropwater_fcst.py` | CLI(`calib`, `calib-sky`, `verify`). 예보 ETo·ETc(주 방법 S4/S3, 비교 S1), S4 월 단위 교차검증 계수(`s4_cv`) 또는 다른 해 계수 고정(`--s4-coef`, `load_s4_fixed`), 기준선, 선행시간별 지표·H2 판정, 방법 비교, 월별 지표, 판정 불확실성(블록 부트스트랩), 입력 진단, 오차 분해, 보정 탐색 | 3 | 구현 |
| `fcst_report.py` | H2 검증 엑셀 (라이브 수식). 일별비교 열 배치는 `daily_layout()`이 정하고 다른 시트는 열 키로 참조. 하늘상태가 있으면 SKY계수 시트·S4 열·방법 비교 표 추가 | 3 | 구현 |
| (물수지 전망) | 관측 Dr + 예보 ETc − 예보 유효강수 → 관수 필요 예상일, 서비스용 엑셀 보고서 | 4 | 예정 |
| `kma_fcst.py` + `kma_grid.py` | 운영용 API 수집(02:10·17:10, 페이징, 원자료 보관, 실패 시 직전 발표 사용), 위경도 → 격자 변환 | 5 | 예정 |

**실행**

```bash
# 1) Rs 계수 보정 — 검증 연도와 다른 해의 01-Cycle 워크북으로 (rs_coef.csv에 지점 행 기록)
python cropwater_fcst.py calib --obs output/eto101_apple_20250101_20251231.xlsx --stn 101

# 2) H2 검증 엑셀 — 과거 단기예보 CSV 폴더 + 검증 연도 01-Cycle 워크북
python cropwater_fcst.py verify --fcst data/fcst_101_73134 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101
#    → output/fcst_verify(101)_<격자>_<첫 대상일>_<끝 대상일>.xlsx

# (선택) 다른 격자와 비교 — 같은 관측·계수·Kc로 계산해 '격자비교' 시트 추가
python cropwater_fcst.py verify --fcst data/fcst_101_73134 --compare data/fcst_101_73135 --obs ... --stn 101

# 3) 운영용 S4 계수 — 하늘상태를 포함한 한 생육기의 예보·관측으로 선행일별 적합 (rs_sky_coef.csv)
python cropwater_fcst.py calib-sky --fcst data/fcst_101_73134 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101

# 4) 다른 해 독립 검증 (VALIDATION #13) — S4는 다른 해 운영 계수 고정, S3 비교 계수도 다른 해 관측으로
python cropwater_fcst.py calib --obs output/eto101_apple_20260101_20260928.xlsx --stn 101 --coef rs_coef_2026.csv
python cropwater_fcst.py verify --fcst data/fcst_101_2025 --obs output/eto101_apple_20250101_20251231.xlsx --stn 101 \
       --grid 73_134 --coef rs_coef_2026.csv --s4-coef rs_sky_coef.csv
```

- `--fcst`: 과거 예보 CSV 폴더(또는 파일 목록, 여러 폴더 가능). 파일마다 첫 줄로 형식을 판별합니다.
  - 포털 CSV(요소별 파일): 요소는 파일명 키워드(최고기온·최저기온·1시간기온·습도·풍속·강수량·하늘상태·강수확률)로 판별합니다. 키워드가 없으면 값 분포로 판별합니다(TMX/TMN은 발표시각별 대상일 수 패턴, SKY는 코드 1·3·4, POP은 10% 단위).
  - OpenAPI 응답 CSV(첫 줄 `baseDate,baseTime,category,…`): 한 파일에 모든 요소가 있어 category로 나눕니다. 한 파일이든 월별 파일이든 같습니다.
  - 요소별 KST CSV(첫 줄 `발표일,발표시각,예보일,예보시각,값`): 요소마다 1파일. 파일명 키워드가 없으면 값으로 요소를 판별합니다(아래 '요소 판별'). 격자 정보가 없으므로 `--grid nx_ny`로 지정합니다.
  - 형식을 한 폴더에 섞어도 됩니다. 같은 발표·시각이 겹치면 나중에 읽은 파일(파일명 순) 값을 씁니다.
  - 필수 요소 파일이 없으면 실행 시 경고합니다(`check_archive`의 `missing_elements`). 쓰지 않는 요소(UUU·VVV·VEC·PTY·SNO·WAV) 파일은 건너뛰고 `ignored_files`에 남깁니다.
- `--s4-coef`: 다른 해의 운영 S4 계수 파일. 주면 월 단위 교차검증 대신 이 계수를 선행일별로 그대로 씁니다(독립 연도 검증). SKY계수 시트의 키는 '고정|선행일'이고, 보정 탐색의 ETo 비율은 중첩 교차검증 없이 구합니다(계수가 검증 연도 자료와 무관하므로).
- `--coef`로 S3 비교 계수도 검증 연도와 다른 해의 관측으로 맞춘 파일을 줍니다. 2026년 검증은 2025년 관측 계수(rs_coef.csv), 2025년 검증은 2026년 관측 계수(rs_coef_2026.csv)
- 하늘상태(SKY)가 들어 있으면 주 방법이 S4가 되고, 하늘상태가 없는 서비스 발표 행은 채점에서 뺍니다(사유 '하늘상태 없음').
- `--obs`: 관측 워크북은 첫 발표 7일 전부터 포함해야 합니다(7일평균 기준선).
- `--compare`: 다른 격자의 과거 예보. 선행시간별 성능·입력 편향과 같은 발표·대상일의 직접 차이를 비교합니다(`grid_comparison()`).
- `--coef`: 상대경로가 현재 폴더에 없으면 스크립트 폴더의 같은 이름 파일(저장소의 `rs_coef.csv`)을 씁니다. 지점의 S3 계수가 없으면 FAO-56 기본값(S1)으로 계산하고 경고를 출력합니다.
- 지점의 격자는 `cropwater_fcst.STATIONS`(아래 대표 10개 지점 표)와 비교해, 다르면 요약 시트에 경고를 남깁니다.
- 예보 원자료(CSV)는 용량이 커서 저장소에 올리지 않습니다.

**처리 흐름 (verify)**

```
과거 단기예보 CSV ──→ fcst_archive.load_archive()      형식 자동 판별(포털 / OpenAPI). 요소별 (발표, lead, 대상시각, 값, 코드).
                          │                              결측값(±900) 행 삭제·개수 집계
                          │  check_archive()             G1 점검: 형식, 격자, 요소별 기간·발표 누락·결측값 수, TMX/TMN 행 패턴, 코드값,
                          │                              같은 요소 파일의 기간 겹침 경고, 일부만 있는 발표(같은 시각 발표의 보통 행 수보다 적음)
                          ▼
                    service_table()                      아침 02시 → D+0~D+3, 저녁 17시 → D+1~D+4
                          │                              6요소가 모두 있는 발표만 (빠진 발표는 skipped_runs에 기록)
                          │  daily_inputs()              대상시각마다 '서비스 발표 이전의 가장 최근 발표' 값
                          ▼
01-Cycle 워크북 ──→ obs_daily (관측 ETo·Kc) ──→ cropwater_fcst.forecast_table()
rs_coef.csv ─────→ rs_model.load_coef()   ──┘   예보 ETo(S3·S1), 관측, 지속성, 7일평균, ETc
                          │  s4_cv()                     (SKY가 있으면) S4 계수: 월 단위 교차검증, 선행일별 → ETo_S4
                          │                              (--s4-coef면 다른 해 계수 고정. 행별 묶음 열 s4_fold = 대상월 또는 '고정')
                          │                              주 방법 열: Rs_main·ETo_main·ETc_main (S4, 없으면 S3)
                          │  split_verifiable()          예보 입력(하늘상태 포함)·대상일 관측·기준선이 모두 있는 행만 (뺀 행은 사유와 함께 보관)
                          │
                          ├─ lead_metrics() / cum3() / h2_verdict()     선행시간별 지표·3일 누적(3일 모두 있을 때만)·판정
                          ├─ method_comparison()                        같은 대상일 S4·S3·S1 비교
                          ├─ month_metrics() / month_verdict()          대상일 월별 지표·입력 편향, 달마다 기준 적용(참고)
                          ├─ bootstrap_h2()                             판정 불확실성: 7일 블록 부트스트랩(시드 고정)
                          ├─ input_diagnostics()                        입력 편향·강수 적중
                          ├─ error_attribution() / bias_correction_cv() 오차 분해·보정 탐색(월 단위 교차검증. S4는 학습 행도 검증 달을 뺀 계수로 다시 계산)
                          ├─ grid_comparison()  (--compare)             다른 격자와 비교
                          ▼
                    fcst_report.build_verify_workbook()  → 검증 엑셀
```

**검증 엑셀 (fcst_report.py)** — 시트별 해석은 [RESULTS_GUIDE.md](RESULTS_GUIDE.md)에 있습니다.
- 시트: 요약 / 일별비교 / 3일누적 / 월별 / 입력진단 / 오차분해 / (격자비교) / Rs계수 / (SKY계수) / 관측 / 설정 / 방법 / 차트자료
- 값으로 넣는 것: 예보 일 입력(집계값, 하늘상태 구름 비율 포함), ASOS 관측 일자료, S4 계수(SKY계수 시트, Python 적합값), 오차분해·격자비교 시트(Python 계산)
- 수식으로 계산하는 것: Ra·Rs(S4는 SKY계수 시트를 '묶음|선행일' 키로 찾아 계산)·PM ETo·Kc·ETc·기준선·오차·지표·판정·방법 비교·월별 지표
- 노란 칸(설정·Rs계수·SKY계수·합격 기준·월별 선행일)을 바꾸면 다시 계산됩니다. 월별 시트의 기준은 요약 시트 합격 기준 셀을 따라갑니다.
- 일별비교 열 배치는 `daily_layout(sky)`가 정합니다. 하늘상태가 없으면 이전 배치(A~BB)와 같고, 있으면 입력·Rs·ETo·오차 묶음에 S4 열이 들어갑니다. 다른 시트는 열 문자가 아니라 열 키(`DB`)로 참조합니다.
- LibreOffice 재계산 결과가 Python 계산과 10⁻¹³ 이내로 일치합니다(G3·G3 재검증 점검).
- 요약 시트의 긴 해석 문장(④~)은 여러 행으로 나눠 씁니다. 병합 칸의 줄바꿈·행 높이는 엑셀과 LibreOffice가 다르게 그려 차트와 겹치기 때문입니다.

**격자와 읍면동**
- API는 격자(nx, ny)로 요청합니다.
- 기상자료개방포털은 읍면동을 고르면, 공식 격자표에 따라 그 읍면동이 속한 격자의 예보를 줍니다.
- 대표 ASOS 10개 지점의 격자와 포털 선택 읍면동은 아래와 같습니다. 격자는 활용가이드 LCC 변환값이고, 읍면동은 같은 격자 안에서 관측소와 가장 가까운 대표점입니다.

| ASOS | 지점 | 격자 (nx, ny) | 포털 선택 읍면동 (관측소와 거리) | 같은 격자 읍면동 수 |
| :---: | :--- | :---: | :--- | :---: |
| 101 | 춘천 | 73, 134 | 춘천시 신사우동 (0.5 km) | 10 |
| 216 | 태백 | 95, 119 | 태백시 황지동 (0.5 km) | 4 |
| 119 | 수원 | 60, 120 | 수원시권선구 평동 (1.2 km) | 8 |
| 131 | 청주 | 68, 107 | 청주시흥덕구 복대1동 (0.9 km) | 5 |
| 129 | 서산 | 52, 109 | **없음 → API(nx, ny)로 수집** | 0 |
| 146 | 전주 | 63, 89 | 전주시덕진구 덕진동 (0.5 km) | 14 |
| 156 | 광주 | 59, 75 | 북구 운암2동 (0.7 km) | 13 |
| 136 | 안동 | 90, 106 | 안동시 송하동 (1.0 km) | 6 |
| 192 | 진주 | 79, 75 | 진주시 평거동 (0.0 km) | 1 |
| 189 | 서귀포 | 53, 32 | 서귀포시 중앙동 (0.2 km) | 4 |

**자료 처리 규칙 (활용가이드 2026-06-23판, 변환값은 기본안)**
- **결측:** +900 이상 / −900 이하 값(예: −999.9)
  - 읽을 때 그 행을 빼고 요소별 개수를 셉니다(`check_archive()`의 coverage).
  - 2026년 4~9월(73_134) 확인: 결측값은 7/14 11시~7/15 08시, 8/29 11시~8/30 08시 발표에 몰려 있습니다(발표 전체가 비어 있음).
  - **검증 규칙:** 필수 6요소 중 하나라도 없는 서비스 발표는 통째로 뺍니다(다른 발표로 대신하지 않음). 주 방법이 S4면 하늘상태가 없는 행, 그리고 대상일 관측이 없는 행도 뺍니다. 뺀 발표·행은 요약 시트 '제외' 행에 구간과 사유로 표시합니다.
  - 선택 요소(SKY·POP)는 완전 발표 판정에 쓰지 않습니다. 서비스 발표 자체에 SKY가 없으면 이전 발표로 채우지 않고 비워 둡니다.
  - **운영 규칙(G5, 예정):** 발표를 받지 못하면 직전 발표로 대신하고, 하늘상태가 없으면 S3로 계산합니다. 검증 규칙과 다릅니다.
- **요소 판별:** 파일명 키워드가 우선입니다. 없으면 값 분포로 판별합니다(업로드 과정에서 한글 파일명이 '_'로 바뀌는 경우).
  - 강수: 최솟값 0이고 중앙값 0 (장마철에는 0 비율이 70%대까지 내려가 '0 비율 > 80%' 규칙은 기온으로 잘못 판별함)
  - 하늘상태(SKY): 값이 코드 1·3·4 (드물게 잘못된 0 포함)
  - 강수확률(POP): 0 포함, 0이 아닌 값의 90% 이상이 10% 단위 — 습도(5% 단위, 0 없음)보다 먼저 검사. OpenAPI 자료에는 글피 칸에 66·64 같은 값이 드물게 있어 '모두 10% 단위' 규칙을 바꿨습니다(2026-09-30).
  - 풍향(VEC): 최댓값이 100 초과
  - 요소별 KST CSV만: 예보시각이 06시뿐 → TMN, 15시뿐 → TMX. '강수없음'·'mm' 문자열 → PCP. 음수가 섞인 소수 → 바람성분(UUU·VVV, 값으로는 둘을 구분 못 하며 쓰지 않음). 코드 0~4이고 대부분 0 → PTY
  - 인코딩: UTF-8(BOM 포함)이 아니면 CP949(한글 엑셀 저장본)로 읽습니다.
  - 같은 요소 파일의 기간이 겹치면 경고합니다(판별 오류 신호).
- **하늘상태 일 값 (S4 입력):**
  - SKY 코드 1 맑음, 3 구름많음, 4 흐림. 그 밖의 값(2026-08-26 11시 발표의 0 등)은 결측
  - 구름많음 비율 = Σ w·[SKY=3] / Σ w, 흐림 비율 = Σ w·[SKY=4] / Σ w. w = 그 시각 태양고도의 사인값(밤 0)
  - 태양시 = KST 시각 + (지점 경도 − 135)/15 + 균시차(FAO-56 식32). 지점 경도는 `cropwater_fcst.STATION_LON`
  - 마지막 날(3시간 간격)도 같은 가중을 씁니다(06·09·12·15·18시 값이 주로 반영)
  - POP: 같은 가중의 낮 평균과 하루 최대(0~1). Rs 추정에는 쓰지 않음
- **강수 문자열 (OpenAPI):** 값은 1시간 강수(mm/h)
  - "강수없음" → 0
  - "1mm 미만" → 0.5 (`PCP_LT1_MM`)
  - "1.0mm" ~ "29.0mm" → 그 숫자 (단기 구간은 정수 mm)
  - "30.0~50.0mm" → 40
  - "50.0mm 이상" → 50
  - 숫자로만 온 값은 그대로: 17~23시 발표의 글피(D+3) 1시간 칸은 소수(0.1~4.5 등, "0"은 강수없음), 연장기간은 코드 0~3
  - **포털 자료와의 차이:** 포털 과거자료의 1시간 강수는 정수 mm뿐이고, 1과 2의 빈도비(0.79)가 API의 '1.0mm'·'2.0mm' 빈도비(0.70)와 비슷합니다. 포털에서는 '1mm 미만'이 0으로 기록된 것으로 보입니다. 2025년 4~6월 자료에서 0.5와 0의 차이는 강수유무 22/484행, 운영 S4 계수로 계산한 RMSE 0.03 mm/일 이내입니다(D+1 0.01~0.02, 0.5가 조금 나음). 같은 발표를 두 자료로 대조해 확정합니다([VALIDATION.md](VALIDATION.md) #16).
- **마지막 날(연장기간) 코드값:**
  - 풍속 WSD: 1 → 같은 발표의 직전 정량일 평균 풍속(최대 3.9 m/s), 2 → 6.5, 3 → 11 m/s
  - 강수 PCP: 1 → 1.5, 2 → 9, 3 → 20 mm/h
  - 가이드의 WSD 코드 1 설명 "4 m/s 이상의 약한 바람"은 "미만"의 오기로 보입니다. 실제 자료로 확인합니다.
  - 과거자료 확인 결과: 2026년 포털 자료에서는 WSD 코드가 1과 2만 나왔습니다(73_135 4~6월, 73_134 4~9월). PCP 코드는 4~6월 0·1·2, 4~9월 0·1·2·3입니다.
  - 2025년 OpenAPI 자료(73_134, 4/1~6/4 발표)에는 WSD 코드 3이 6칸 있습니다. PCP 코드는 0·1·2입니다.
  - 마지막 날은 00시(1시간 간격) + 03~21시(3시간 간격) 8개 시각입니다. 03시부터 코드값입니다.
- **발표 시점에 지난 시각 채움:** 대상 시각마다 **서비스 발표 시각 이전(포함)의 가장 최근 발표 값**을 씁니다.
  - 포털 과거자료는 발표 6시간 뒤부터 있으므로, 아침 D0의 00~07시(8시간)는 전날 17·20·23시 발표 값입니다.
  - 운영 API와 OpenAPI로 받은 과거자료는 발표 1시간 뒤부터 있으므로 00~02시(전날 23시 발표)만 채웁니다. 03~07시는 02시 발표 자체의 값입니다.
  - 이 차이는 아침 D+0에만 영향을 줍니다. H2 판정(D+1~D+3)에는 영향이 없습니다.
  - G0의 '21시간 평균' 안을 G1에서 이 규칙으로 바꿨습니다([THEORY.md 9장](THEORY.md#9-예보-기반-etoetc--기상청-단기예보)).
- **과거 예보 CSV 형식 (기상자료개방포털):**
  - 머리행: `format: day(UTC),hour(UTC),forecast,value,day(KST),hour(KST) location:nx_ny Start : YYYYMMDD`
  - UTC 날짜마다 `Start : YYYYMMDD` 행이 있습니다. 여러 달을 한 파일로 받으면 이 행마다 연월을 바꿔 읽습니다(2026-09-29 수정: 전에는 첫 달로만 읽어 두 번째 달부터 날짜가 틀렸음. 한 달씩 받은 파일에는 영향 없음).
  - 발표시각(KST) = Start 행의 연월 + UTC 일·시 + 9시간
  - 1시간 요소: 대상시각 = 발표시각 + forecast(h)
  - 일 요소(TMX·TMN): forecast +6, +7, … = 대상일 순번
    - TMX: 02·05·08·11시 발표는 오늘부터, 14~23시 발표는 내일부터
    - TMN: 02시 발표만 오늘부터
- **과거 예보 CSV 형식 (OpenAPI 응답)** — 단기예보 조회서비스(`getVilageFcst`) 응답 항목을 행으로 모은 파일
  - 머리행: `baseDate,baseTime,category,fcstDate,fcstTime,fcstValue,nx,ny` (시각은 앞의 0이 빠진 `200`, `0` 등)
  - 발표·예보 일시는 KST 그대로입니다. lead = 예보시각 − 발표시각(h). 격자는 nx_ny
  - 한 발표에 14개 요소(category)가 모두 있습니다. 읽는 요소는 필수 6요소 + SKY·POP이고, 나머지(UUU·VVV·VEC·PTY·SNO·WAV)는 읽지 않습니다. WAV(파고)는 육지 격자에서 모두 −999입니다.
  - 발표는 하루 8회(02~23시, 3시간 간격)가 모두 들어 있습니다. 검증 채점에는 02·17시만 쓰고, 나머지는 지나간 시각 채움과 누락 점검에 씁니다.
  - 발표별 행 구조(2025년 4월 확인): 02시 발표 lead +1~+94(D+3은 03시부터 3시간 간격·코드), 17시 발표 +1~+103(D+4가 3시간 간격). 발표당 798~1,052행
  - TMX·TMN: 대상일 = fcstDate(TMN 06시, TMX 15시 칸). 발표시각별 대상일은 포털 규칙과 같습니다.
  - 빈 줄(`,,,,,,,`)·반복된 머리행은 건너뛰고, 같은 발표·요소·예보시각이 두 번 있으면 뒤의 값을 씁니다.
  - 발표 중간이 잘린 경우(일부만 있는 발표)는 `check_archive()`의 `short_issues`로 확인합니다. 예) 2025-05-05 02시 발표는 앞 188줄이 없어 그날 03~18시 값이 없음
- **과거 예보 CSV 형식 (요소별 KST)** — 2025년 자료(2026-09-30 수령)
  - 머리행: `발표일,발표시각,예보일,예보시각,값` (시각은 `0200` 형식). 요소마다 1파일, 격자 정보 없음
  - 일시는 KST, 값 표기는 OpenAPI 응답과 같습니다(강수 문자열, lead +1부터, 연장기간 코드). 먼저 받은 OpenAPI 파일과 겹치는 4/1~6/4 구간의 값이 100% 같았습니다.
  - TMX는 예보시각 15시, TMN은 06시 칸만 있습니다(대상일 = 예보일).
  - 2025년 4/1 02시 ~ 9/30 23시 발표 1,464회(183일 × 8회), 발표당 행 수 모두 정상(일부만 있는 발표 없음), 결측값 없음

---

## 8. 설계 원칙

**라이브 수식 우선**
`cropwater_station.py`는 Python이 값이 아닌 수식 문자열을 셀에 씁니다. 설정값 변경 시 전체 시트가 자동 재계산되며, 셀 단위로 계산 과정을 추적할 수 있습니다.
02-Cycle 검증 엑셀(`fcst_report.py`)도 같은 원칙입니다. 입력을 바꿔 PM을 여러 번 다시 풀어야 하는 오차분해 시트만 값으로 넣고, 시트에 그 사실을 적습니다.

**무관수 가정 명시**
물수지는 자연강우만 반영합니다. 관수가 있었다면 Dr이 실제보다 높게 산출됩니다. 이 한계는 결과요약·물수지 시트 헤더에 명시되어 있습니다.

**캐시 우선**
`stations_backup.csv`는 기상청 API허브 호출 결과를 로컬 캐시합니다. API 연결 실패 시 자동으로 캐시를 사용하고, 성공 시 갱신합니다.

**이중 ETo 크로스체크**
PM과 증발접시(pan) 방식을 동시에 계산합니다. pan/PM(= ETo_pan ÷ ETo_PM)은 두 방법이 일치하면 1.0입니다.
2026년 9개 지점은 모두 0.77~0.85로 비슷해서, 이는 방법 간 계통적 차이로 봅니다.
- 이 비율은 **다른 지점과 동떨어진 지점을 찾는 상대 비교**에 씁니다.
- 일사 자료의 절대 품질은 Rs ≤ Rso 검사로 확인합니다 ([THEORY.md 3장](THEORY.md#3-기준증발산량-eto--fao-56-penman-monteith)).
- 고도화 구현 때 지표를 **역산 계수(ETo_PM ÷ Epan, 10일 이상 합계)**로 바꿀 예정입니다. FAO-56 Class A 계수표(Table 5) 및 Allen-Pruitt Kp와 직접 비교할 수 있기 때문입니다([VALIDATION.md](VALIDATION.md) 후속 조치).

**가설검증 게이트**
고도화 사이클의 각 단계는 THEORY.md를 기준으로 논리·실증 점검을 통과한 뒤 다음 단계로 넘어갑니다. 코드와 docs는 같은 단계에서 함께 갱신합니다. 기록은 [VALIDATION.md](VALIDATION.md)에 남깁니다.
