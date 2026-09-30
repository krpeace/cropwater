# 🌾 CropWater — FAO-56 기반 노지 작물 물수지 분석 도구

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![R&D](https://img.shields.io/badge/국가R%26D-RS--2025--02305238-orange)](https://www.nrf.re.kr/)
[![FAO-56](https://img.shields.io/badge/기준-FAO--56%20Penman--Monteith-yellowgreen)](https://www.fao.org/4/x0490e/x0490e00.htm)

> 기상청 ASOS 및 FAO-56 기반 일별 증발산량·근권 물수지 자동 분석 도구

---

## 📌 목차

1. [프로젝트 소개](#-프로젝트-소개)
2. [주요 기능](#-주요-기능)
3. [설치](#-설치)
4. [인증키 설정](#-인증키-설정)
5. [사용법](#-사용법)
6. [엑셀 출력 구성](#-엑셀-출력-구성)
7. [지원 작물 목록](#-지원-작물-목록)
8. [디렉터리 구조](#-디렉터리-구조)
9. [문서](#-문서)
10. [연구 과제 및 기여](#-연구-과제-및-기여)

---

## 🌱 프로젝트 소개

국가 R&D 과제(RS-2025-02305238, 주관: 주식회사 파모스)로 개발된 **물리 모델 기반 작물 물수지 분석 도구**입니다. 기상청 ASOS 공공 데이터와 FAO-56 표준 모델을 결합해 노지 작물의 일별 수분 부족량과 최적 관수 시점을 산출합니다.

**검증 사례**
- 🍎 노지 사과 — 춘천 (ASOS 101), 2026년 4~8월
- 🥬 여름 배추 — 태백 (ASOS 216), 2026년 7~8월

---

## ✨ 주요 기능

| 기능 | 설명 |
| :--- | :--- |
| **기준증발산량 (ETo)** | FAO-56 Penman-Monteith 계산식 적용, 일사 결측 시 Ångström 보정 |
| **작물증발산량 (ETc)** | 생육단계별 Kc 보간, 국지 기상 보정 및 멀칭 보정 |
| **일별 근권 물수지** | TAW·RAW·Dr·Ks·심층침투(DP) 추적 및 스트레스 보정 증발산량(ETc_adj) 산정. 관수 기록(날짜·관수량 CSV 또는 엑셀 노란 칸)을 넣으면 Dr에 반영 |
| **다지점 일괄 분석** | 30개 이상 지점 동시 계산 및 관수 필요 달력(히트맵) 시각화 |
| **지점 메타 자동화** | 기상청 API 연동, 위도·고도·풍속계 높이 자동 수집 및 로컬 캐싱 |
| **51개 작물 라이브러리** | FAO-56 표준 파라미터(Kc·Zr·p) 내장 (`crops_library.csv`) |
| **단기예보 ETo·ETc 예측** (02-Cycle, H2 검증 완료) | 기상청 단기예보(아침 02시 → 오늘~D+3, 저녁 17시 → 내일~D+4)로 ETo·ETc 예측, 일사는 기온교차·강수확률·하늘상태로 추정(S4), 과거 예보로 선행시간별·월별 검증 |
| **관수 필요 예상일** (02-Cycle G4) | 어제 끝 관측 Dr + 예보 ETc − 기대 강수(강수량 × 강수확률) → 며칠 뒤 Dr ≥ RAW가 되는 날과 범위(빠르면·늦으면), 3일 누적 ETc ± 오차, 권장 관수량. 서비스(관수 전망) 엑셀 |

---

## 🔧 설치

```bash
git clone https://github.com/your-org/cropwater.git
cd cropwater

# 가상환경 생성 및 활성화
python -m venv .venv
.venv\Scripts\activate      # Windows
# source .venv/bin/activate  # macOS/Linux

pip install -r requirements.txt
```

---

## 🔑 인증키 설정

두 포털에서 별도로 인증키를 발급받아야 합니다.

| 키 | 포털 | 용도 |
| :--- | :--- | :--- |
| `DATA_GO_KR` | [공공데이터포털](https://data.go.kr) | ASOS 일자료 조회 |
| `KMA_HUB` | [기상청 API허브](https://apihub.kma.go.kr) | 지점일람표(위도·고도) 조회 |

```bash
cp apikey.txt.example apikey.txt   # 복사 후 실제 키 입력
```

> ⚠️ `apikey.txt`는 `.gitignore`로 관리됩니다. 절대 GitHub에 올리지 마세요.

---

## 🚀 사용법

### cropwater_station.py — 단일 지점

```bash
python cropwater_station.py --stn 101 --crop apple --start 20260401 --end 20260831 --bud 20260401
python cropwater_station.py --stn 216 --crop kimchi_cabbage --start 20260701 --end 20260831 --bud 20260706
```

| 파라미터 | 설명 | 예시 |
| :--- | :--- | :--- |
| `--stn` | ASOS 지점 번호 | `101` |
| `--crop` | 작물 ID | `apple` |
| `--start` / `--end` | 조회 기간 YYYYMMDD (미래 날짜는 전날로 자동 조정) | `20260401` |
| `--bud` | 생육 시작일 (발아·정식일) | `20260401` |
| `--irrig` | 관수 기록 CSV (머리행 `날짜,관수량_mm[,메모]`, 공급량 mm = 10a당 톤) → 물수지 시트 I열 | `관수기록.csv` |
| `--out` | 출력 파일명 (생략 시 자동 생성) | `result.xlsx` |

> 파라미터를 생략하면 대화형 프롬프트가 순서대로 묻습니다.

---

### cropwater_multi.py — 다지점 비교

```bash
python cropwater_multi.py --stns 101,119,216 --start 20260401 --end 20260831
python cropwater_multi.py --stns 101,119,131,146,156,136,216 --start 20260401 --end 20260831 --zr 1.0 --pdep 0.50
```

| 파라미터 | 설명 | 기본값 |
| :--- | :--- | :--- |
| `--stns` | 지점 번호 (쉼표 구분, 공백 없이) | 9개 기본 지점 |
| `--start` / `--end` | 조회 기간 | 당해 1월 1일 / 어제 |
| `--zr` | 근권심도 (m) | `0.50` |
| `--fc` / `--wp` | 포장용수량 / 위조점 (m³/m³) | `0.22` / `0.10` (양토) |
| `--pdep` | 고갈계수 p | `0.40` |

---

### cropwater_fcst.py — 단기예보 ETo·ETc 예측 검증 (02-Cycle)

```bash
# Rs 계수 보정: 검증 연도와 다른 해의 cropwater_station.py 출력 워크북 → rs_coef.csv
python cropwater_fcst.py calib --obs output/eto101_apple_20250101_20251231.xlsx --stn 101

# H2 검증: 과거 단기예보 CSV 폴더(포털 요소별 CSV 또는 OpenAPI 응답 CSV) + 검증 연도 워크북 → 검증 엑셀
#   하늘상태(SKY) CSV가 함께 있으면 주 방법이 S4(하늘상태 반영)가 됨
python cropwater_fcst.py verify --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101

# 운영용 S4 계수: 하늘상태 포함 과거 예보 + 관측 → rs_sky_coef.csv (선행일별)
python cropwater_fcst.py calib-sky --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101

# 다른 해 독립 검증: S4는 다른 해 운영 계수 고정, S3 비교 계수도 다른 해 관측으로 (VALIDATION #13)
python cropwater_fcst.py calib --obs output/eto101_apple_20260101_20260928.xlsx --stn 101 --coef rs_coef_2026.csv
python cropwater_fcst.py verify --fcst data/fcst_101_2025 --obs output/eto101_apple_20250101_20251231.xlsx --stn 101 \
       --grid 73_134 --coef rs_coef_2026.csv --s4-coef rs_sky_coef.csv

# (G4) 예보 ETo 오차표(해마다) → fcst_error_table.csv / 예보 물수지 검증 엑셀 / 서비스(관수 전망) 엑셀
python cropwater_fcst.py errtable --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101
python cropwater_fcst.py wbverify --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101
python cropwater_fcst.py service  --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101 \
       --run "2026-05-15 02" --irrig 관수기록.csv
```

| 파라미터 | 설명 |
| :--- | :--- |
| `--fcst` | 과거 단기예보 CSV 폴더 또는 파일들 (필수 TMX·TMN·TMP·REH·WSD·PCP, 선택 SKY·POP. 여러 달을 한 파일로 받아도 됨). 형식은 첫 줄로 자동 판별: 기상자료개방포털 요소별 CSV, OpenAPI(단기예보 조회서비스) 응답을 모은 CSV(`baseDate,baseTime,category,fcstDate,fcstTime,fcstValue,nx,ny`, 한 파일에 모든 요소), 요소별 KST CSV(`발표일,발표시각,예보일,예보시각,값`). 하루 8회 발표를 모두 받음. 필수 요소가 빠지면 경고, 쓰지 않는 요소(바람성분 등) 파일은 건너뜀 |
| `--obs` | `cropwater_station.py` 출력 워크북 (관측 기준값·Kc 설정). 첫 발표 7일 전부터 포함 |
| `--stn` | ASOS 지점 번호 (rs_coef.csv 행 선택) |
| `--coef` / `--out` | 계수 파일(기본 `rs_coef.csv`, 현재 폴더에 없으면 스크립트 폴더) / 출력 파일명 (기본 `output/fcst_verify(지점)_격자_시작_끝.xlsx`) |
| `--compare` | (선택) 비교할 다른 격자의 과거 예보 → '격자비교' 시트 |
| `--s4-coef` | (선택) 다른 해의 운영 S4 계수 파일(`rs_sky_coef.csv`). 주면 교차검증 대신 그대로 적용 — 독립 연도 검증 |
| `--grid` | (선택) 파일에 격자 정보가 없을 때(요소별 KST CSV) 격자 `nx_ny`, 예: `73_134` |
| `--err` | (errtable·wbverify·service) 예보 ETo 오차표, 기본 `fcst_error_table.csv`. wbverify는 검증 연도를 뺀 다른 해 행으로 범위를 채점 |
| `--irrig` | (wbverify·service) 관수 기록 CSV → 관측 물수지(출발 고갈량)에 반영 |
| `--auto-irrigate` | (wbverify) 관수 규칙 시나리오: 전날 끝 Dr ≥ RAW면 그 Dr만큼 관수 |
| `--run` | (service) 발표시각 `YYYY-MM-DD HH`(02 또는 17). 없으면 자료의 가장 최근 서비스 발표. S4는 운영 계수(`--s4-coef` 기본 `rs_sky_coef.csv`) |

> 가설·합격 기준·게이트 판정은 [docs/VALIDATION.md](docs/VALIDATION.md), 이론은 [THEORY.md 9장](docs/THEORY.md)에 있습니다.
>
> ASOS 101 춘천·사과 검증 결과: 두 해 모두 H2 기준 충족 (주 방법 S4 = 기온교차 + 강수확률 하루 최대 + 하늘상태)
> - 2026년(월 단위 교차검증), 대상일 4/2~9/19: D+1 RMSE 아침 0.94·저녁 0.82 mm/일, 지속성 대비 45~54% 개선
> - 2025년(다른 해, 2026년 계수 고정), 대상일 4/1~10/4: D+1 RMSE 아침 0.94·저녁 0.84 mm/일, 최소 개선율 34%·42% — 아침 발표는 여유가 작음
> - 비교 S3(기온교차 + 강수유무): 2026년 D+1 0.99·0.88, 2025년 1.01·0.92(아침 미달). 한여름(6~8월) 한두 달이 약함
>
> 예보 물수지(G4, 두 해, 사과·양토 RAW 60 mm): 3일 예상 고갈량 오차 7~15 mm로 예보 없이 보는 것보다 22~42% 작음. 오차의 대부분은 강수 예보에서 옴(강수가 완벽하면 0.8~2.0 mm).
> 관수 필요 예상일은 사건의 절반가량이 같은 날이고, 범위(빠르면~늦으면)가 실제 날짜를 모두 담음 — 실제 날짜가 '빠르면'보다 앞선 적 없음

---

**주요 ASOS 지점 번호**

| 번호 | 지점 | 번호 | 지점 | 번호 | 지점 |
| :---: | :--- | :---: | :--- | :---: | :--- |
| 101 | 춘천 | 119 | 수원 | 131 | 청주 |
| 133 | 대전 | 136 | 안동 | 143 | 대구 |
| 146 | 전주 | 156 | 광주 | 192 | 진주 |
| 216 | 태백 | 184 | 제주 | 189 | 서귀포 |

---

## 📊 엑셀 출력 구성

### cropwater_station.py (6시트)

| 시트 | 내용 |
| :--- | :--- |
| **설정** | 지점·작물·토양 파라미터 (노란 셀: 직접 편집 가능, 변경 시 전체 자동 재계산) |
| **원데이터** | ASOS 일별 원자료 |
| **계산과정** | FAO-56 전 과정 라이브 수식 |
| **결과요약** | ETo·ETc 기간 합계 + 물수지 요약 (ΣPeff, 유효강수율, 관수 필요 일수 등) |
| **물수지** | 일별 Dr·Ks·DP·ETc_adj·관수필요(●) 라이브 수식. I열(노란 칸) = 관수량 입력 |
| **계산근거** | FAO-56 식번호·출처·적용 설명 |

### cropwater_multi.py (5시트)

| 시트 | 내용 |
| :--- | :--- |
| **지점비교** | 지점별 ETo 평균 + 물수지 요약 컬럼 (녹색=충분·적색=부족) |
| **일별_전지점** | 전 지점 일별 Peff·Dr·Ks·관수필요 |
| **월별_물수지** | 지점 × 월 교차 집계 |
| **관수필요_달력** | Dr 히트맵 (흰→연녹→연노→진적) + Dr 선형 차트 |
| **설명** | 계산 방법·파라미터 |

### cropwater_fcst.py verify (11시트 + 하늘상태가 있으면 SKY계수, --compare면 격자비교)

| 시트 | 내용 |
| :--- | :--- |
| **요약** | H2 합격 기준(노란 셀)·판정, 선행시간별 ETo·ETc 지표, Rs 방법 비교(S4일 때), 3일 누적, 주요 발견, 차트 |
| **일별비교** | 발표 × 대상일 예보 입력·Rs·PM ETo·기준선·ETc·입력 오차 (라이브 수식) |
| **3일누적** | 발표별 첫 3일 합 오차 |
| **월별** | 월별 성능, 달마다 기준 적용(참고), 월별 입력 편향 (라이브 수식, 선행일 선택) |
| **입력진단** | 입력 편향·강수 적중 |
| **오차분해** | 입력 교체 오차 분해, 보정 탐색, 판정 불확실성(블록 부트스트랩) (Python 계산값) |
| **격자비교** (선택) | 두 격자 예보의 성능·입력 편향·직접 차이 (Python 계산값) |
| **Rs계수 / SKY계수 / 관측 / 설정** | S3 계수·H1 재검증 / S4 교차검증·운영 계수(하늘상태가 있을 때) / ASOS 관측 ETo·Kc / 지점·예보·Kc 설정 |
| **방법 / 차트자료** | 정의·규칙·한계 / 차트 원본 |

### cropwater_fcst.py wbverify (7시트) — 예보 물수지 검증 (G4)

| 시트 | 내용 |
| :--- | :--- |
| **요약** | 예상 고갈량 오차(중심·비교 경로·기준선), 관수 필요 판정 적중, 관수 필요 예상일·범위 적중, 월별, 예보 강수 검증, 판정 기준 민감도, 해석 |
| **발표별 / 발표요약** | 발표 × 대상일 8개 경로 고갈량(라이브 수식) / 발표마다 예상일(중심·빠르면·늦으면·참값)·판정 |
| **관측물수지 / 오차표 / 설정 / 방법** | 관측 물수지(관수 규칙 스위치) / 범위의 상대 오차 r / 토양·판정 기준·평가 기간 / 방법 |

### cropwater_fcst.py service (7시트) — 관수 전망 (G4)

| 시트 | 내용 |
| :--- | :--- |
| **관수 전망** | 어제 끝 Dr·상태, 3일 누적 ETc ± 오차(주 지표), 예보·기대 강수, 관수 필요 예상일(빠르면·늦으면), 권장 관수량, 날짜별 표(마지막 날 '참고') |
| **예보 물수지 / 관측 물수지** | 세 경로 고갈량(수식) / 어제까지 물수지 — 노란 칸에 관수량을 적으면 전망이 다시 계산됨 |
| **오차표 / 편향 점검 / 설정 / 방법** | 예보 ETo 오차 / 월·선행일별 편향(보정은 안 함) / 토양 값 / 방법 |

📖 자세한 해석 방법 → [docs/RESULTS_GUIDE.md](docs/RESULTS_GUIDE.md)

---

## 🌿 지원 작물 목록

`--crop` 파라미터에 `crop_id`를 입력합니다.

| 분류 | crop_id |
| :--- | :--- |
| 과수 | `apple` `pear` `cherry` `citrus` `grape` `almond` `olive` |
| 베리 | `strawberry` `blueberry` |
| 가지과 | `tomato` `pepper` `eggplant` |
| 박과 | `cucumber` `watermelon` `melon` `pumpkin` |
| 엽채류 | `cabbage` `kimchi_cabbage` `broccoli` `lettuce` `spinach` `onion` |
| 근채류 | `potato` `sweet_potato` `radish` `carrot` `sugar_beet` |
| 두과 | `soybean` `pea` `chickpea` `green_bean` |
| 곡류 | `rice` `barley` `wheat` `maize` `sorghum` |
| 기타 | `alfalfa` `turfgrass_cool` `turfgrass_warm` `sunflower` `rapeseed` … |

전체 파라미터는 `crops_library.csv` 참조.

---

## 📁 디렉터리 구조

```
cropwater/
├── README.md
├── LICENSE
├── requirements.txt
├── apikey.txt.example          ← 인증키 템플릿 (.gitignore에 apikey.txt 등록 필수)
├── .gitignore
│
├── fao56_core.py               ← 공통 모듈 (FAO-56 물리식·API 조회·작물 라이브러리)
├── cropwater_station.py        ← 단일 지점 분석 → 6시트 Excel
├── cropwater_multi.py          ← 다지점 비교 → 히트맵 + 차트
│
├── cropwater_fcst.py           ← (02-Cycle) 단기예보 ETo·ETc 예측 CLI: calib / calib-sky / verify / errtable / wbverify / service
├── fcst_archive.py             ← 과거 단기예보 CSV 파싱(포털·OpenAPI 형식) → 발표별 일 입력(기대 강수 포함)
├── rs_model.py                 ← 일사량(Rs) 추정: S3 식(50) + 강수유무, S4 + 강수확률·하늘상태 구름 비율, 계수 적합
├── obs_daily.py                ← cropwater_station 워크북 → 관측 ETo·Kc
├── fcst_report.py              ← H2 검증 엑셀 (11~13시트, 라이브 수식)
├── fcst_wb.py                  ← (G4) 관측·예보 물수지, 관수 필요 예상일·범위, 오차표, 서비스 전망
├── fcst_wb_report.py           ← (G4) 예보 물수지 검증 엑셀 / 서비스(관수 전망) 엑셀
│
├── crops_library.csv           ← 51개 작물 Kc·Zr·p (FAO-56 Table 12·22)
├── stations_backup.csv         ← ASOS 지점 메타 캐시
├── rs_coef.csv                 ← Rs 추정 계수 S3 (FAO 기본값 + 지점 보정값)
├── rs_coef_2026.csv            ← S3 계수 (2026년 관측, 2025년 독립 검증용)
├── rs_sky_coef.csv             ← Rs 추정 계수 S4 (지점 × 선행일, calib-sky)
├── fcst_error_table.csv        ← 예보 ETo 오차표 (지점 × 해 × 발표 × 선행일 × 월, errtable)
│
├── output/                     ← 분석 결과물 저장 폴더 (.gitignore로 xlsx 제외)
│   └── .gitkeep
│
├── tests/
│   ├── test_fao56.py           ← FAO-56 핵심 계산·물수지 한 걸음·관수 기록 단위 테스트 (pytest)
│   ├── test_fcst.py            ← 02-Cycle 예보 모듈 단위 테스트
│   └── test_wb.py              ← 02-Cycle G4 예보 물수지 단위 테스트
│
└── docs/
    ├── THEORY.md               ← 관수·토양학 이론 (농대생 입문용)
    ├── ARCHITECTURE.md         ← 코드 구조·함수 레퍼런스 (개발자용)
    ├── RESULTS_GUIDE.md        ← 엑셀 결과 해석 방법
    └── VALIDATION.md           ← 사이클별 가설·게이트 점검 기록
```

---

## 📚 문서

| 문서 | 대상 | 내용 |
| :--- | :--- | :--- |
| [THEORY.md](docs/THEORY.md) | 농대생·입문자 | ETo·ETc·TAW·RAW·물수지 이론, 예보 기반 ETo(9장) |
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | 개발자·연구자 | 함수 레퍼런스·데이터 흐름·확장 방법 |
| [RESULTS_GUIDE.md](docs/RESULTS_GUIDE.md) | 현장 사용자 | 엑셀 시트별 해석·주의사항 |
| [VALIDATION.md](docs/VALIDATION.md) | 연구자·개발자 | 고도화 사이클별 가설·합격기준·게이트 점검 기록 |

---

## 🔬 연구 과제 및 기여

| 항목 | 내용 |
| :--- | :--- |
| 과제명 | 노지 작물 수분스트레스 진단·정밀자동관수 패키지기술 산업화 |
| 과제번호 | RS-2025-02305238 |
| 주관기관 | 주식회사 파모스 |
| 지원기관 | 농림축산식품부 |

**참고 문헌**
- Allen, R.G. et al. (1998). *Crop Evapotranspiration — FAO Irrigation and Drainage Paper No. 56.* FAO, Rome.
- KMA ASOS API: https://data.go.kr
- 기상청 API허브: https://apihub.kma.go.kr

새 작물 추가는 `crops_library.csv`에 행 하나를 추가하는 것으로 충분합니다. PR과 Issue 모두 환영합니다.

---

## 📄 라이선스

MIT License © 2026 주식회사 파모스 (FarmOS Corp.)
