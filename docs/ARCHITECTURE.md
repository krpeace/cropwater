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
      ├── cropwater_fcst.py  ← CLI: calib·calib-sky(Rs 계수) / verify(H2 검증 엑셀) / errtable·wbverify·service(G4)
      ├── fcst_archive.py    ← 과거 단기예보 CSV 파싱 → 발표별 일 입력(기대 강수 포함), 운영 서비스 표(직전 발표 대체)
      ├── rs_model.py        ← Rs 추정(S3 식50 + 강수유무, S4 + 강수확률·하늘상태), 계수 파일 2계층
      ├── obs_daily.py       ← 01-Cycle 워크북 → 관측 ETo·Kc
      ├── fcst_report.py     ← H2 검증 엑셀(라이브 수식)
      ├── fcst_wb.py         ← (G4) 관측·예보 물수지, 관수 필요 예상일·범위, 오차표, 서비스 전망
      ├── fcst_wb_report.py  ← (G4) 예보 물수지 검증 엑셀 / 서비스(관수 전망) 엑셀
      ├── cropwater_ops.py   ← (G5) 운영 CLI: run(02:10·17:10 수집 → 관수 전망 엑셀) / probe-asos / report(H3) / grid
      ├── kma_fcst.py        ← (G5) 단기예보·ASOS API 조회, 완결성 점검, 원자료 보관, 수집 로그
      ├── kma_grid.py        ← (G5) 위경도 ↔ 단기예보 격자(LCC)
      ├── ops_report.py      ← (G5) H3 보고서 엑셀
      └── ops_replay.py      ← (G5) 모의 API·가짜 시계로 운영 경로 재생(오프라인 검증)
```

**의존 관계** — 실행 파일은 서로 독립적이며 `fao56_core`를 공유합니다. 02-Cycle 모듈은 01-Cycle 출력 워크북(관측 기준값)을 입력으로 씁니다.

```
cropwater_station ──┐
                    ├──→ fao56_core ──→ requests, openpyxl
cropwater_multi   ──┘        ▲
                             │
cropwater_fcst ──→ fcst_archive, rs_model, obs_daily, fcst_report, fcst_wb, fcst_wb_report ──→ pandas, numpy, openpyxl
                                                                  (fcst_wb → fao56_core.wb_step)
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
| `wb_step(dr_prev, P, etc, taw, raw, irr_net=0)` | 전날 끝 Dr, 강수, ETc, TAW, RAW, 순관수량 | (Ks, ETc_adj, DP, Dr) | 식84·85·86·88 (Ks는 Dr,i-1) |

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
| `load_irrigation_log(path)` | 관수 기록 CSV(머리행 날짜·관수량_mm[·메모], UTF-8/CP949) → `{date: 공급 관수량 mm}`. 같은 날은 합침, 음수·잘못된 날짜는 오류 |

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
    'irrig': {},                # 관수 기록 {date: 공급 mm} (--irrig) → 물수지 I열
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

H  = MAX(F+I*Ea-K-G, 0)   [식88] DP = P + I·Ea − ETc_adj − Dr,i-1 ≥ 0 (FAO-56 원식, 2026-09-29)
I  = 관수량 (값, 노란 칸)   공급량 mm. --irrig 관수 기록으로 채우거나 직접 입력(없으면 0 = 무관수)
J  = IF(G<=RAW, 1, (TAW-G)/(TAW-RAW))   [식84] Ks — 전날 끝 고갈량 Dr,i-1로 판단(FAO-56 원식, #15, 2026-09-30)
K  = J*E               ETc_adj = Ks × ETc
L  = MIN(MAX(G-F-I*Ea+K+H, 0), TAW)   [식85·86] Dr,i = Dr,i-1 − P − I·Ea + ETc_adj + DP
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
`recs`에 `{DP, Peff, Dr, Ks, irr, In}` 필드를 추가하여 반환합니다. 한 걸음은 `fao56_core.wb_step`입니다.

```python
def wb_step(dr_prev, P, etc, taw, raw, irr_net=0.0):          # fao56_core
    ks      = 1.0 if dr_prev <= raw else max((taw - dr_prev) / (taw - raw), 0.0)   # [식84] 전날 끝 Dr
    etc_adj = ks * etc
    dp      = max(P + irr_net - etc_adj - dr_prev, 0.0)        # [식88] FAO-56 원식
    dr      = min(max(dr_prev - P - irr_net + etc_adj + dp, 0.0), taw)   # [식85·86]
    return ks, etc_adj, dp, dr
```

- 식(88)은 2026-09-29에 FAO-56 원식으로 바꿨습니다. 이전 간이식 `max(P − Dr, 0)`은 당일 ETc를 빼지 않아, 큰 비가 온 날 DP를 ETc만큼 크게 잡고 그날 끝 고갈량을 0이 아니라 ETc로 남겼습니다.
- Ks는 전날 끝 고갈량 Dr,i-1로 판단합니다(FAO-56 원식, #15, 2026-09-30). 이전에는 그날 비가 먼저 들어간 뒤의 고갈량을 썼습니다(비 온 날에만 차이, 2026년 춘천·사과 관수 필요일 43 → 41일).

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

### ◆ 관수 기록 연동 (2026-09-30 구현)

FAO-56 식(85)의 `I` 항에 농가가 준 관수량을 넣습니다. 비어 있으면 지금까지처럼 무관수입니다.

- **입력:** 관수 기록 CSV — 머리행 `날짜,관수량_mm[,메모]`(영문 `date,amount_mm`도 됨). 날짜는 `YYYY-MM-DD`·`YYYYMMDD`·`YYYY.MM.DD`·`YYYY/MM/DD`. 관수량은 **공급량(mm)**, 10a(1,000 m²)당 1톤 = 1 mm. 같은 날이 여러 줄이면 합칩니다(`fao56_core.load_irrigation_log`).
- **물수지:** 순관수량 I = 공급량 × 관수효율 Ea. `wb_step(dr_prev, P, etc, taw, raw, irr_net)`
- **01-Cycle 워크북:** `cropwater_station.py --irrig 관수기록.csv` → 물수지 시트 I열(노란 칸). 엑셀에서 직접 고쳐도 다시 계산됩니다.
- **예보 물수지(G4):** 관측 물수지(출발 고갈량)에 반영됩니다. `cropwater_fcst.py service --irrig …`, 서비스 엑셀의 '관측 물수지' 시트 노란 칸.
- 데이터 소스는 나중에 T21 센서 로그 또는 FarmOS 플랫폼의 관수 제어 이력으로 바꿀 수 있습니다(같은 CSV 형식으로 내보내면 됨).

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

**입력 대응 (정정)** — TMN/TMX/REH/WSD만으로는 PM을 계산할 수 없습니다. 단기예보에는 **일사량(Rs)이 없으므로** 추정합니다. 하늘상태(SKY)와 강수확률(POP)이 있으면 S4(식50형 + 강수확률 하루 최대 + 구름 비율, 선행일별 계수), 없으면 S3(식50 + 강수유무)입니다. S4의 강수 입력은 #17(2026-09-30)에서 강수유무 → 강수확률 하루 최대로 바꿨습니다. 기압은 고도 기반(식7)으로 대체합니다.

**모듈 (G1~G4 구현, 2026-09-30)**

| 모듈 | 역할 | 단계 | 상태 |
| :--- | :--- | :---: | :---: |
| `fcst_archive.py` | 과거 단기예보 CSV → 요소별 표(발표, 선행시간, 대상시각, 값, 코드 여부) → 서비스 발표(02·17시)별 일 입력. 입력 형식 세 가지를 첫 줄로 자동 판별(`csv_format`): 포털 CSV(요소별, 여러 달 파일 포함), OpenAPI 응답 CSV(한 파일에 모든 요소, `read_openapi_csv`), 요소별 KST CSV(`read_element_csv`). 쓰지 않는 요소 파일은 건너뛰고 기록. 선택 요소 SKY·POP은 낮 시간 일사 비중으로 가중한 일 값(`sun_weights`). `fcst_daily.py` 계획을 합침 | 1 | 구현 |
| `rs_model.py` + `rs_coef.csv` + `rs_sky_coef.csv` | Rs 추정(S1 식50, S3 식50 + 강수유무, S4 + 하늘상태 구름 비율), 최소제곱 적합, 계수 2계층(stn=0 FAO 기본값 + 지점 행). S4는 지점 × 선행일 행이고 `rain_input` 열(`pop_max`)로 강수 입력을 적어 둠(이 열이 없는 이전 형식 = 강수유무 계수는 읽지 않음). `rs_coef.csv`는 2025년 관측 S3 계수(2026년 검증용), `rs_coef_2026.csv`는 2026년 관측 S3 계수(2025년 독립 검증용) | 2·3 | 구현 |
| `obs_daily.py` | 01-Cycle 출력 워크북 → 관측 일자료, 관측 ETo(01-Cycle 규칙), Kc(설정 시트 값) | 2 | 구현 |
| `cropwater_fcst.py` | CLI(`calib`, `calib-sky`, `verify`). 예보 ETo·ETc(주 방법 S4/S3, 비교 S1), S4 월 단위 교차검증 계수(`s4_cv`) 또는 다른 해 계수 고정(`--s4-coef`, `load_s4_fixed`), 기준선, 선행시간별 지표·H2 판정, 방법 비교, 월별 지표, 판정 불확실성(블록 부트스트랩), 입력 진단, 오차 분해, 보정 탐색 | 3 | 구현 |
| `fcst_report.py` | H2 검증 엑셀 (라이브 수식). 일별비교 열 배치는 `daily_layout()`이 정하고 다른 시트는 열 키로 참조. 하늘상태가 있으면 SKY계수 시트·S4 열·방법 비교 표 추가 | 3 | 구현 |
| `fcst_wb.py` | (G4) 관측 물수지(`observed_wb`: 관수 기록·관수 규칙 시나리오·결측일은 그날 아침 D+0 예보로 채움), 발표별 예보 물수지(`forecast_runs`: 아침 출발 = 관측 Dr(D−1), 저녁 = + 오늘 아침 D+0 예보, 경로 중심·빠르면·늦으면 + 비교 경로·참값), 지표(`wb_lead_metrics`, `need_contingency`, `first_need_eval`, `threshold_sensitivity`, `rain_verification`), 오차표(`eto_error_rows`·`save_error_table`·`pooled_errors`·`err_lookup`), 편향 보정 자리(`apply_bias`, 기본 없음 — #10), 한 발표의 전망(`service_outlook`) | 4 | 구현 |
| `fcst_wb_report.py` + `fcst_error_table.csv` | (G4) 예보 물수지 검증 엑셀(7시트, 라이브 수식) / 서비스 엑셀(관수 전망 8시트: 날짜 선택·과습 표시 포함). 오차표는 지점 × 해 × 발표 × 선행일 × 월 행(일·3일 누적) — 서비스는 여러 해를 표본 수로 가중해 합치고, 검증은 검증 연도를 뺀 다른 해 행을 씀. 지점 행이 없으면 stn 0(기본값) 행, 그것도 없으면 전 지점을 합침 | 4 | 구현 |
| `kma_fcst.py` + `kma_grid.py` | 운영용 API 조회(단기예보 페이지 나눔, ASOS 일자료), 응답 코드 분류, 완결성 점검(발표시각별 행 수), 원자료 gzip 보관, 완결 발표만 OpenAPI 형식 월별 CSV로 쌓기, 관측 캐시, 수집 로그. 위경도 → 격자 변환 | 5 | 구현 |
| `cropwater_ops.py` + `ops_report.py` + `ops_replay.py` | 운영 CLI(아래 ◇ 운영 수집기). 슬롯마다 재시도·마감·지난 발표 채움·관측 조회 → 운영 결측 규칙(#12)으로 관수 전망 엑셀 → 슬롯 로그. H3 보고서, 모의 재생 | 5 | 구현 |

**실행**

```bash
# 1) Rs 계수 보정 — 검증 연도와 다른 해의 01-Cycle 워크북으로 (rs_coef.csv에 지점 행 기록)
python cropwater_fcst.py calib --obs output/eto101_apple_20250101_20251231.xlsx --stn 101

# 2) H2 검증 엑셀 — 과거 단기예보 CSV 폴더 + 검증 연도 01-Cycle 워크북
#    (2026년 자료는 OpenAPI로 받은 요소별 KST CSV — 파일에 격자 정보가 없어 --grid로 지정. #14부터 운영·검증 모두 OpenAPI 자료 기준)
python cropwater_fcst.py verify --fcst data/fcst_101_2026 --grid 73_134 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101
#    → output/fcst_verify(101)_<격자>_<첫 대상일>_<끝 대상일>.xlsx

# (선택) 다른 격자와 비교 — 같은 관측·계수·Kc로 계산해 '격자비교' 시트 추가
python cropwater_fcst.py verify --fcst data/fcst_101_73134 --compare data/fcst_101_73135 --obs ... --stn 101

# 3) 운영용 S4 계수 — 하늘상태를 포함한 한 생육기의 예보·관측으로 선행일별 적합 (rs_sky_coef.csv)
python cropwater_fcst.py calib-sky --fcst data/fcst_101_2026 --grid 73_134 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101

# 4) 다른 해 독립 검증 (VALIDATION #13) — S4는 다른 해 운영 계수 고정, S3 비교 계수도 다른 해 관측으로
python cropwater_fcst.py calib --obs output/eto101_apple_20260101_20260928.xlsx --stn 101 --coef rs_coef_2026.csv
python cropwater_fcst.py verify --fcst data/fcst_101_2025 --obs output/eto101_apple_20250101_20251231.xlsx --stn 101 \
       --grid 73_134 --coef rs_coef_2026.csv --s4-coef rs_sky_coef.csv

# 5) (G4) 예보 ETo 오차표 — 해마다 한 번. 검증과 같은 계수 조건으로 (fcst_error_table.csv에 지점·해 행을 바꿔 넣음)
python cropwater_fcst.py errtable --fcst data/fcst_101_2026 --grid 73_134 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101
python cropwater_fcst.py errtable --fcst data/fcst_101_2025 --obs output/eto101_apple_20250101_20251231.xlsx --stn 101 \
       --grid 73_134 --coef rs_coef_2026.csv --s4-coef rs_sky_coef.csv

# 6) (G4) 예보 물수지 검증 엑셀 — 범위의 오차는 검증 연도를 뺀 다른 해 오차표로 채점
python cropwater_fcst.py wbverify --fcst data/fcst_101_2026 --grid 73_134 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101
#    → output/fcst_wbverify(101)_<격자>_<첫 채점 대상일>_<끝 채점 대상일>.xlsx   (--auto-irrigate: 관수 규칙 시나리오, 파일명 _irrig)

# 7) (G4) 서비스 엑셀 — 한 발표의 관수 전망. S4는 운영 계수(rs_sky_coef.csv, 기본값), 오차표는 여러 해 합침
python cropwater_fcst.py service --fcst data/fcst_101_2026 --grid 73_134 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101 \
       --run "2026-05-15 02" --irrig 관수기록.csv
#    → output/fcst_service(101)_20260515_02.xlsx   (--run이 없으면 자료의 가장 최근 서비스 발표)
```

- `--fcst`: 과거 예보 CSV 폴더(또는 파일 목록, 여러 폴더 가능). 파일마다 첫 줄로 형식을 판별합니다.
  - 포털 CSV(요소별 파일): 요소는 파일명 키워드(최고기온·최저기온·1시간기온·습도·풍속·강수량·하늘상태·강수확률)로 판별합니다. 키워드가 없으면 값 분포로 판별합니다(TMX/TMN은 발표시각별 대상일 수 패턴, SKY는 코드 1·3·4, POP은 10% 단위).
  - OpenAPI 응답 CSV(첫 줄 `baseDate,baseTime,category,…`): 한 파일에 모든 요소가 있어 category로 나눕니다. 한 파일이든 월별 파일이든 같습니다.
  - 요소별 KST CSV(첫 줄 `발표일,발표시각,예보일,예보시각,값`): 요소마다 1파일. 파일명 키워드가 없으면 값으로 요소를 판별합니다(아래 '요소 판별'). 격자 정보가 없으므로 `--grid nx_ny`로 지정합니다.
  - 형식을 한 폴더에 섞어도 됩니다. 같은 발표·시각이 겹치면 나중에 읽은 파일(파일명 순) 값을 씁니다.
  - 필수 요소 파일이 없으면 실행 시 경고합니다(`check_archive`의 `missing_elements`). 쓰지 않는 요소(UUU·VVV·VEC·PTY·SNO·WAV) 파일은 건너뛰고 `ignored_files`에 남깁니다.
- `--s4-coef`: 다른 해의 운영 S4 계수 파일. 주면 월 단위 교차검증 대신 이 계수를 선행일별로 그대로 씁니다(독립 연도 검증). SKY계수 시트의 키는 '고정|선행일'이고, 보정 탐색의 ETo 비율은 중첩 교차검증 없이 구합니다(계수가 검증 연도 자료와 무관하므로).
- `--coef`로 S3 비교 계수도 검증 연도와 다른 해의 관측으로 맞춘 파일을 줍니다. 2026년 검증은 2025년 관측 계수(rs_coef.csv), 2025년 검증은 2026년 관측 계수(rs_coef_2026.csv)
- 하늘상태(SKY)와 강수확률(POP)이 모두 들어 있으면 주 방법이 S4가 되고, 둘 중 하나라도 없는 서비스 발표 행은 채점에서 뺍니다(사유 '하늘상태·강수확률 없음').
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

**처리 흐름 (G4: errtable / wbverify / service)**

```
과거·운영 단기예보 + 01-Cycle 워크북 ──→ fcst_wb.prepare()          forecast_table과 같은 예보표(모든 서비스 발표·대상일, ETc_main, 강수·기대 강수·강수확률)
                                        │  apply_bias()              편향 보정 자리 — 기본 없음(#10)
                                        │
   errtable ──→ eto_error_rows() → save_error_table()               예보 ETo 오차표: 발표 × 선행일 × 월(0 = 전체)의 RMSE·MBE·관측 평균, 3일 누적
                                        │
관수 기록 CSV ─→ observed_wb()                                     관측 물수지(식84·85·86·88, Ks = Dr,i-1). 결측일은 fill_from_forecast()(그날 아침 D+0 예보)
                                        │                            (wbverify --auto-irrigate: 전날 끝 Dr ≥ RAW면 Dr만큼 관수)
                                        ▼
   wbverify ──→ forecast_runs()   발표마다 출발 = 관측 Dr(D−1 끝) [저녁은 + 오늘 아침 D+0 예보로 하루]
                                  경로: 중심(기대 강수) · 빠르면(비 없음, ETc×(1+r)) · 늦으면(예보 강수 전부, ETc×(1−r))
                                        · 예보 강수 그대로 · 비 무시 · 관측 강수 · 기준선 · 참값(관측)
                                  r = err_lookup(pooled_errors(오차표, 검증 연도 제외))
                 ├─ wb_lead_metrics() / wb_month_metrics()   예상 Dr 오차(RMSE·MBE·개선율)
                 ├─ need_contingency()                       대상일별 관수 필요 판정 적중
                 ├─ first_need_eval() / first_need_summary() 관수 필요 예상일·판정·범위 적중 (주 지표 3일·참고 포함 4일)
                 ├─ threshold_sensitivity() / rain_verification()
                 ▼
                 fcst_wb_report.build_wbverify_workbook()   → 검증 엑셀
   service ───→ service_outlook() + recent_bias()           한 발표의 전망(오차표는 여러 해 합침, S4 운영 계수)
                 ▼
                 fcst_wb_report.build_service_workbook()    → 서비스 엑셀
```

**예보 물수지 검증 엑셀 (fcst_wb_report.py)**
- 시트: 요약 / 발표별 / 발표요약 / 관측물수지 / 오차표 / 설정 / 방법
- 값으로 넣는 것: 예보 입력(예보 ETc = Kc × 예보 ETo, 예보 강수·기대 강수·강수확률), 관측 ETc·강수, 상대 오차 r, 제외 사유, 요약 E·F표(예보 강수 검증·판정 기준 민감도, Python 계산)
- 수식으로 계산하는 것: 관측 물수지(관수 규칙 시나리오 스위치 포함), 발표별 8개 경로 고갈량(한 칸 수식 `MIN(MAX(Dr − P + Ks·ETc, 0), TAW)`), 관수 필요 판정, 발표요약의 예상일·판정·범위 적중, 요약 A~D표
- 설정 시트의 토양 값·판정 기준 고갈량(기본 = RAW)·평가 대상일 수(3 또는 4)·관수 규칙 스위치를 바꾸면 다시 계산됩니다.
- 예상일(처음 필요 순번)은 MINIFS 대신 순번 1~4의 COUNTIFS 중첩 IF로 씁니다(구형 엑셀·LibreOffice 호환).
- LibreOffice 재계산 결과가 Python 계산과 5 × 10⁻¹³ 이내로 일치합니다(G4 점검).

**서비스 엑셀 (fcst_wb_report.py)** — 한 발표의 관수 전망
- 시트: 관수 전망 / 예보 물수지 / 관측 물수지 / 날짜별 / 오차표 / 편향 점검 / 설정 / 방법
- '관수 전망': 조회일(노란 칸), 전날 끝 Dr·상태(과습 경고·토양 수분 충분 포함), 3일 누적 ETc ± 오차(주 지표), 예보 강수·기대 강수, 관수 필요 예상일(중심·빠르면·늦으면), 권장 관수량(순·공급), 날짜별 표(마지막 날 회색 '참고')
- **날짜 선택 (2026-10-02):** '관수 전망'!B5(조회일, 기본 = 발표일)를 바꾸면 그 날짜의 '지금 토양 상태'와 '앞으로 3일'이 바뀝니다.
  - '날짜별' 시트(`_sheet_timeline`)가 관측 물수지(첫날~관측 마지막 날)와 예보 물수지(관측 전·그날·대상일)를 날짜순으로 이은 표입니다. 모든 칸이 두 시트를 참조하는 수식이라 노란 칸(관수 기록)을 고치면 함께 바뀝니다.
  - 날짜별!S1~S11이 찾기 칸입니다: 조회일, 발표일, 주 지표 첫날(아침 0·저녁 +1), 전날·조회일·대상일 1~4의 행(`MATCH`, 없으면 빈 칸), 과습 판정 행, 주 지표 출발 행. '관수 전망'은 `INDEX(날짜별 열, 찾기 칸)`으로 값을 가져오고, 찾기 칸이 비면 '자료 없음'을 씁니다.
  - 조회일 칸은 목록 유효성 검사(날짜별 A열)이고 오류 알림을 껐습니다. 목록에 없는 날짜를 직접 넣어도 되고, 그때는 '자료 없음'이 나옵니다.
  - '앞으로 3일' 머리는 3일이 모두 예보면 '예보 전망 적용', 모두 관측 물수지면 '관측 자료 적용', 걸치면 '관측 + 예보 자료 적용'으로 바뀝니다.
  - 지난 날짜는 관측 물수지 결과(실제 비·관수 기록)이고 ± 오차·범위가 없습니다. 3일 창이 발표일을 넘으면 넘는 날은 이 발표의 예보입니다. LET·XLOOKUP·MINIFS는 쓰지 않습니다(구형 엑셀·LibreOffice 호환).
- **과습 표시 (2026-10-02):** 상태 칸에 과습 경고(끝 Dr 0)·토양 수분 충분(끝 Dr < 설정의 토양 수분 충분 기준, 기본 RAW ÷ 2)를 관수 쪽 상태보다 먼저 판정합니다. 관측 물수지 행에만, 생육기(Kc > 0)에만 적용합니다. 전날이 '관측 전'(예보로 먼저 진행한 날)이면 관측 마지막 날로 판정합니다. Python 쪽 같은 규칙은 `fcst_wb.soil_status`·`wet_info`입니다([THEORY.md 9장 ◆ 과습 표시](THEORY.md#9-예보-기반-etoetc--기상청-단기예보)).
- '관측 물수지' 노란 칸(관수 기록)을 고치면 어제 끝 Dr → 예보 물수지 → 전망이 다시 계산됩니다. 예보 ETo·강수·r은 값입니다(예보 ETo의 PM 수식은 H2 검증 엑셀에 있음).
- '편향 점검'(#10): 오차표의 월·발표·선행일별 MBE와, 예보표에 있으면 이 발표 이전 최근 30일의 예보 − 관측

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
  - **운영 규칙(G5, #12):** 서비스 발표를 마감까지 받지 못하면 24시간 안의 직전 발표로 대신하고(`fcst_archive.ops_service_table`), 하늘상태·강수확률이 없으면 S3로 계산합니다(`forecast_table(s3_fallback=True)`). 전날 ASOS 관측이 아직 없으면 전날을 예보 하루로 범위를 담아 진행합니다(`fcst_wb.service_outlook(obs_last=…)`). 검증 규칙과 다릅니다(아래 ◇ 운영 수집기).
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
  - POP: 하루 최대(0~1)가 S4의 강수 입력입니다(#17). 같은 가중의 낮 평균은 참고 값
- **강수 문자열 (OpenAPI):** 값은 1시간 강수(mm/h)
  - "강수없음" → 0
  - "1mm 미만" → 0.5 (`PCP_LT1_MM`)
  - "1.0mm" ~ "29.0mm" → 그 숫자 (단기 구간은 정수 mm)
  - "30.0~50.0mm" → 40
  - "50.0mm 이상" → 50
  - 숫자로만 온 값은 그대로: 17~23시 발표의 글피(D+3) 1시간 칸은 소수(0.1~4.5 등, "0"은 강수없음), 연장기간은 코드 0~3
  - **포털 자료와의 차이 (#16에서 확인):** 같은 발표(2026-05-01~05, 40회)를 두 자료로 대조했습니다. 포털은 1 mm 미만을 0으로, 1 mm 이상을 정수로 반올림해 기록합니다('1mm 미만' → 0, 4.8 → 5). 다른 요소는 100% 같습니다([VALIDATION.md](VALIDATION.md) #16).
    - 그래서 포털 자료로 맞춘 계수를 OpenAPI 입력에 쓰면 강수유무 판정이 달라질 수 있습니다. S4의 강수 입력을 강수확률로 바꾼 이유 중 하나입니다(#17). 강수확률은 두 자료가 같습니다.
    - **2026년 4~9월 전체 대조(#14):** 8/26 11시 ~ 8/27 08시 발표 8회에서 포털의 1시간 기온·강수확률·하늘상태가 OpenAPI와 달랐습니다(포털 하늘상태에 코드표에 없는 0이 있던 구간, 포털 쪽 문제로 봄). 나머지는 강수 표기 외 같습니다. 운영 계수·오차표·검증은 모두 OpenAPI 자료로 만듭니다.
    - OpenAPI 자료의 '1mm 미만'은 0.5 mm로 둡니다. 2025년 자료에서 0과 비교해 Rs 오차가 같거나 작았고, 물수지(G4)의 강수량으로도 더 자연스럽습니다.
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


### ◇ 운영 수집기 (G5, 2026-09-30 구현)

예보를 하루 두 번 받아 관수 전망 엑셀을 자동으로 만듭니다. 근거는 [THEORY.md 9장 ◆ 운영 수집 — H3](THEORY.md#9-예보-기반-etoetc--기상청-단기예보), 판정은 [VALIDATION.md](VALIDATION.md) G5입니다.

**실행** — 스케줄러가 부릅니다(아래 등록 방법).

```bash
# 준비: apikey.txt(DATA_GO_KR), ops_config.json(ops_config.example.json을 복사해 고침)
python cropwater_ops.py run --config ops_config.json                          # 02:10·17:10 — 지금 시각의 슬롯
python cropwater_ops.py run --config ops_config.json --slot "2026-10-01 02"   # 특정 슬롯 다시(마감 뒤면 지연 수집)
python cropwater_ops.py run --config ops_config.json --no-wait                # 한 번만 시도(설치 확인용)
python cropwater_ops.py probe-asos --config ops_config.json                   # 전날 ASOS 일자료 조회 가능 시각 점검(매시)
python cropwater_ops.py report --config ops_config.json --since 2026-10-01    # H3 보고서 엑셀
python cropwater_ops.py grid 37.90262 127.73570                               # 농장 좌표 → 격자 73_134
python ops_replay.py --fcst data/fcst_101_2026 --obs output/eto101_apple_20260101_20260928.xlsx \
       --start 2026-06-01 --end 2026-06-30 --data data/replay --faults faults.json   # 오프라인 모의 재생
```

- `run`은 슬롯 결과가 '실패'면 종료 코드 2를 돌려줍니다(스케줄러 기록에서 보임).

**설정 (ops_config.json)**

| 키 | 뜻 | 기본 |
| :--- | :--- | :--- |
| `stn` | ASOS 지점(관측·계수·오차표) | 101 |
| `grid` | 예보 격자 `nx_ny`. 없으면 `lat`·`lon`(농장 좌표)으로 계산, 그것도 없으면 대표 지점 표 | 지점 표 |
| `lat`, `lon` | 농장 좌표(격자 계산용) | — |
| `obs` | 01-Cycle 관측 워크북(설정 시트의 Kc·토양, 원데이터의 관측 이력) | 필수 |
| `data` / `out` | 운영 자료 폴더 / 관수 전망 엑셀 폴더 | `data/ops` / `output/service` |
| `apikey` | 인증키 파일 | `apikey.txt` |
| `coef` / `s4_coef` / `err` | S3 계수 / S4 운영 계수 / 오차표 | `rs_coef.csv` / `rs_sky_coef.csv` / `fcst_error_table.csv` |
| `irrig` | 관수 기록 CSV | 없음 |
| `history` | 함께 읽을 과거 예보 폴더(운영 첫날부터 편향 점검 30일을 채우고 싶을 때) | 없음 |
| `window_days` | 계산에 읽는 최근 예보 기간(일) | 40 |

**한 슬롯의 처리 (`run`)**

```
02:10 (또는 17:10) ─→ Collector.collect_service()
   ├─ 서비스 발표(02·17시) 요청: +0·+5·+10·+15·+20·+30분 (마감 = 발표 + 40분)
   │    fetch_vilage() → 원자료 gzip → check_issue()(발표시각별 행 수) → 필수 6요소가 차면 store_issue()
   │    응답 00 정상 / 03 자료 없음 → 다시 / 서버 오류·시간 초과 → 다시 / 10~33·HTTP 401·403 → 멈춤(인증·요청 오류)
   ├─ 첫 시도 뒤 catch_up(): 지난 24시간 안에 아직 받지 않은 발표를 한 번씩(직전 발표 = 백업, 02시 발표의 00~02시 채움)
   ├─ update_obs(): ASOS 전날(D−1) 조회(asos_d1 기록) + 워크북 뒤 최근 30일의 빠진 날 → 관측 캐시
   ▼
build_service()
   ops_prepare()  최근 window_days일의 받은 발표 → ops_service_table()(서비스 발표가 없으면 직전 발표: src_lead·err_name)
                  → forecast_table(S4 운영 계수 고정, src_lead로 계수 선택, s3_fallback)      ← 관측 = 워크북 + 캐시
   fcst_wb.service_from()  관측 물수지(관수 기록) → 관측이 아직 없는 최근 날은 예보 하루로 범위 포함(obs_last)
                           → service_outlook() → build_service_workbook()(수집 정보 표시)
   ▼
slot_log.csv   결과: 정시 성공 / 백업: S3 / 백업: 직전 발표 (S3) / 지연 수집 / 다시 만듦 / 실패
```

**운영 결측 규칙 구현 (#12)**

| 상황 | 구현 | 서비스 엑셀 표시 |
| :--- | :--- | :--- |
| 서비스 발표를 마감까지 못 받음 | `fcst_archive.ops_service_table`: 24시간 안의 가장 최근 발표(필수 6요소). `daily_inputs(arch, 대체 발표, 서비스 대상일)`. `src_lead`(대체 발표 날짜 기준 선행일) → S4 계수, `err_name`·`err_lead`(02~14시 발표 → 아침, 17~23시 → 저녁) → ± 오차 r. 대체 발표가 담지 못한 마지막 날(참고)은 뺌 | 머리 ※ 줄, 수집 정보 '예보 발표' |
| 하늘상태·강수확률 없음 | 저장할 때 두 요소를 빼고, `forecast_table(s3_fallback=True)`가 그 행을 S3로(`rs_method`) | '일사 추정' |
| 전날 ASOS 관측 없음 | '관측 물수지'는 전날을 예보로 채워 보이고(관수 기록 칸 유지), 전망은 관측 마지막 날 끝 Dr에서 출발해 전날을 예보 하루로 진행(`service_outlook(obs_last=…)`: 세 경로 범위, 관측 강수·관수 기록 반영) | '예보 물수지'의 '관측 전' 행, '어제 관측' |
| 24시간 안에 쓸 발표가 없음 | 엑셀을 만들지 않고 '실패' | — |

**자료 폴더 (`data/ops`, 저장소에 올리지 않음)**

```
raw/fcst/73_134/2026/10/20261001_0200__20261001T021003_p1.json.gz   응답 본문 그대로(시도마다, 받은 시각 포함)
raw/asos/101/2026/09/20260930_20260930__20261001T021012.json.gz
fcst/73_134/vilage_73_134_202610.csv   완결 발표만, OpenAPI 형식(fcst_archive가 그대로 읽음). 하늘상태·강수확률이 덜 찼으면 두 요소는 뺌
fcst/73_134/issues.csv                 받은 발표 목록(발표, 받은 시각, 8요소 완결, 행 수)
obs/asos_101_daily.csv                  ASOS 일자료(날짜마다 가장 최근 조회 값). 워크북에 없는 날짜만 계산에 씀
log/collect_log.csv                     시도마다: 시각, 슬롯, 종류(fcst·asos_d1·asos_fill·asos_probe), 대상, HTTP·응답 코드, 행 수, 상태, 완결 여부
log/slot_log.csv                        슬롯마다: 시작·끝, 받은 시각, 정시 여부, 결과, 계산한 발표, Rs 방법, 전날 관측, 엑셀, 오류
```

**스케줄러 등록**

Windows(명령 프롬프트, 저장소 폴더가 `C:\cropwater`일 때):

```bat
schtasks /Create /TN cropwater_morning /SC DAILY /ST 02:10 /TR "cmd /c cd /d C:\cropwater && python cropwater_ops.py run --config ops_config.json >> data\ops\run.log 2>&1"
schtasks /Create /TN cropwater_evening /SC DAILY /ST 17:10 /TR "cmd /c cd /d C:\cropwater && python cropwater_ops.py run --config ops_config.json >> data\ops\run.log 2>&1"
schtasks /Create /TN cropwater_probe /SC HOURLY /ST 00:40 /TR "cmd /c cd /d C:\cropwater && python cropwater_ops.py probe-asos --config ops_config.json >> data\ops\probe.log 2>&1"
```

- 작업 스케줄러의 작업 속성에서 '작업을 실행하기 위해 절전 모드 해제'와 '예약된 시작 시간을 놓친 경우 가능한 대로 빨리 작업 시작'을 켭니다. 컴퓨터가 꺼져 있었던 슬롯은 '실행 안 됨'(실패)으로 셉니다.
- 관측 점검(`cropwater_probe`)은 조회 가능 시각을 알 때까지 1~2주만 돌리고 지웁니다(`schtasks /Delete /TN cropwater_probe`).

Linux(cron):

```cron
CRON_TZ=Asia/Seoul
10 2 * * *  cd /opt/cropwater && python3 cropwater_ops.py run --config ops_config.json >> data/ops/run.log 2>&1
10 17 * * * cd /opt/cropwater && python3 cropwater_ops.py run --config ops_config.json >> data/ops/run.log 2>&1
40 * * * *  cd /opt/cropwater && python3 cropwater_ops.py probe-asos --config ops_config.json >> data/ops/probe.log 2>&1
```

- 프로그램은 컴퓨터 시간대와 관계없이 UTC + 9시간으로 KST를 계산합니다(한국은 서머타임 없음).
- **국내 인터넷에서 실행합니다.** 공공데이터포털은 해외 접속이 불안정할 수 있습니다(클라우드 작업 공간에서는 접속이 차단됨, 2026-09-30 확인).

**모의 재생 (`ops_replay.py`)** — 과거 예보(OpenAPI CSV)와 01-Cycle 워크북으로 API를 흉내 내고, 가짜 시계로 슬롯을 차례로 처리합니다(네트워크 없이 수집기 논리 검정).
- 장애 주입 JSON: `delay_min`(제공 지연, 분), `missing`(끝내 제공 안 함), `truncate`(잘린 응답 횟수), `drop_opt`(하늘상태·강수확률 없이 제공), `fatal`(인증 오류 시각 범위), `p_http500`·`p_timeout`(요청마다 확률), `seed`
- `--obs-until`: 워크북을 그 날짜까지만 쓰고 뒤는 모의 ASOS로 받아 관측 경로도 검정. `--asos-hour`: 전날 자료가 제공되는 시각
- `--mode light`(기본)는 수집과 출처 판단만 해서 빠릅니다. `full`은 슬롯마다 관수 전망 엑셀을 만듭니다.

---

## 8. 설계 원칙

**라이브 수식 우선**
`cropwater_station.py`는 Python이 값이 아닌 수식 문자열을 셀에 씁니다. 설정값 변경 시 전체 시트가 자동 재계산되며, 셀 단위로 계산 과정을 추적할 수 있습니다.
02-Cycle 검증 엑셀(`fcst_report.py`)도 같은 원칙입니다. 입력을 바꿔 PM을 여러 번 다시 풀어야 하는 오차분해 시트만 값으로 넣고, 시트에 그 사실을 적습니다.

**관수 기록이 없으면 무관수**
관수 기록(`--irrig` 또는 물수지 시트 노란 칸)이 없으면 자연강우만 반영합니다. 실제로 관수했다면 Dr이 실제보다 높게 나옵니다. 결과요약·물수지 시트 머리에 적어 둡니다. 예보 물수지의 예상 Dr도 '예보 기간에 관수하지 않으면'의 값입니다.

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
