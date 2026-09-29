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
└── cropwater_multi.py     ← 다지점 스크리닝
      ├── compute_station()        → ETo·Epan
      ├── compute_water_balance()  → Dr/Ks/DP
      ├── aggregate_monthly()      → 월별 집계
      └── build_calendar_sheet()  → 히트맵 + 차트
```

**의존 관계** — 두 실행 파일은 서로 독립적이며 `fao56_core`만 공유합니다.

```
cropwater_station ──┐
                    ├──→ fao56_core ──→ requests, openpyxl
cropwater_multi   ──┘
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

H  = MAX(F-G, 0)       [식88] DP
I  = MAX(G-F, 0)       강수 후 고갈량
J  = IF(I<=RAW, 1, (TAW-I)/(TAW-RAW))   [식84] Ks
K  = J*E               ETc_adj = Ks × ETc
L  = MIN(I+K, TAW)     [식85] Dr,i
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
    DP       = max(P - Dr, 0)
    Dr_after = max(Dr - P, 0)
    Ks       = 1.0 if Dr_after <= raw \
               else (taw - Dr_after) / (taw - raw)
    Dr       = min(Dr_after + Ks * ETo, taw)
    rec["Dr"] = Dr
```

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

기상청 단기예보(getVilageFcst)로 D+3~D+4 ETo·ETc를 예측하고, 물수지 전망으로 관수 필요 예상일을 알립니다.
이론·검증 설계는 [THEORY.md 9장](THEORY.md#9-예보-기반-etoetc--기상청-단기예보)을, 게이트 기록은 [VALIDATION.md](VALIDATION.md)를 봅니다.

**입력 대응 (정정)** — TMN/TMX/REH/WSD만으로는 PM을 계산할 수 없습니다. 단기예보에는 **일사량(Rs)이 없으므로** 식(50) + 강수유무 보정으로 추정합니다. 기압은 고도 기반(식7)으로 대체합니다.

**예정 모듈 (단계별 구현 후 이 절을 갱신)**

| 모듈 | 역할 | 단계 |
| :--- | :--- | :---: |
| `kma_fcst.py` | 발표시각 선택, totalCount 기준 페이징, 오류코드 처리, 원자료 JSON 보관, 실패 시 직전 발표분 사용(이중화) | 1 |
| `kma_grid.py` | 위경도 → 격자(nx, ny) 변환 (활용가이드 LCC 식, 공식 격자표와 대조 확인) | 1 |
| `rs_model.py` + `rs_coef.csv` | Rs 추정. 계수 2계층: FAO-56 기본값 + 지점·계절 보정값 | 2 |
| `fcst_daily.py` | 시간 → 일 집계 (TMX/TMN 우선, ea 시간평균, u2, 강수합·강수유무, 선행시간·완결성 표시) | 3 |
| `cropwater_fcst.py` | 예측 실행(CLI): ETo·ETc·물수지 전망 → 엑셀 + JSON(FarmOS·Fcore 연동용) | 3~4 |

**자료 처리 규칙 (활용가이드 2026-06-23판)**
- 결측: +900 이상 / −900 이하 값
- 강수 문자열: "강수없음", "1mm 미만", "30.0~50.0mm", "50.0mm 이상"
- 연장기간(마지막 날) 코드값: PCP 1~3, SNO 1~2, WSD 1~3
- WSD 코드 1의 가이드 설명 "4 m/s 이상의 약한 바람"은 "미만"의 오기로 보입니다. 실제 응답으로 확인합니다.
- 격자: 춘천 ASOS (73,134) ≠ 신북 사과원 (73,135), 태백 ASOS (95,119) ≠ 태백 실증포장 (96,118)
  - 서비스는 농장 좌표 격자를 씁니다.
  - 검증은 ASOS 격자를 씁니다.

---

## 8. 설계 원칙

**라이브 수식 우선**
`cropwater_station.py`는 Python이 값이 아닌 수식 문자열을 셀에 씁니다. 설정값 변경 시 전체 시트가 자동 재계산되며, 셀 단위로 계산 과정을 추적할 수 있습니다.

**무관수 가정 명시**
물수지는 자연강우만 반영합니다. 관수가 있었다면 Dr이 실제보다 높게 산출됩니다. 이 한계는 결과요약·물수지 시트 헤더에 명시되어 있습니다.

**캐시 우선**
`stations_backup.csv`는 기상청 API허브 호출 결과를 로컬 캐시합니다. API 연결 실패 시 자동으로 캐시를 사용하고, 성공 시 갱신합니다.

**이중 ETo 크로스체크**
PM과 증발접시(pan) 방식을 동시에 계산합니다. pan/PM(= ETo_pan ÷ ETo_PM)은 두 방법이 일치하면 1.0입니다.
2026년 9개 지점은 모두 0.77~0.85로 비슷해서, 이는 방법 간 계통적 차이로 봅니다.
- 이 비율은 **다른 지점과 동떨어진 지점을 찾는 상대 비교**에 씁니다.
- 일사 자료의 절대 품질은 Rs ≤ Rso 검사로 확인합니다 ([THEORY.md 3장](THEORY.md#3-기준증발산량-eto--fao-56-penman-monteith)).

**가설검증 게이트**
고도화 사이클의 각 단계는 THEORY.md를 기준으로 논리·실증 점검을 통과한 뒤 다음 단계로 넘어갑니다. 코드와 docs는 같은 단계에서 함께 갱신합니다. 기록은 [VALIDATION.md](VALIDATION.md)에 남깁니다.
