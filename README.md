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
| **일별 근권 물수지** | TAW·RAW·Dr·Ks·심층침투(DP) 추적 및 스트레스 보정 증발산량(ETc_adj) 산정 |
| **다지점 일괄 분석** | 30개 이상 지점 동시 계산 및 관수 필요 달력(히트맵) 시각화 |
| **지점 메타 자동화** | 기상청 API 연동, 위도·고도·풍속계 높이 자동 수집 및 로컬 캐싱 |
| **51개 작물 라이브러리** | FAO-56 표준 파라미터(Kc·Zr·p) 내장 (`crops_library.csv`) |

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
| **물수지** | 일별 Dr·Ks·DP·ETc_adj·관수필요(●) 라이브 수식 |
| **계산근거** | FAO-56 식번호·출처·적용 설명 |

### cropwater_multi.py (5시트)

| 시트 | 내용 |
| :--- | :--- |
| **지점비교** | 지점별 ETo 평균 + 물수지 요약 컬럼 (녹색=충분·적색=부족) |
| **일별_전지점** | 전 지점 일별 Peff·Dr·Ks·관수필요 |
| **월별_물수지** | 지점 × 월 교차 집계 |
| **관수필요_달력** | Dr 히트맵 (흰→연녹→연노→진적) + Dr 선형 차트 |
| **설명** | 계산 방법·파라미터 |

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
├── crops_library.csv           ← 51개 작물 Kc·Zr·p (FAO-56 Table 12·22)
├── stations_backup.csv         ← ASOS 지점 메타 캐시
│
├── output/                     ← 분석 결과물 저장 폴더 (.gitignore로 xlsx 제외)
│   └── .gitkeep
│
├── tests/
│   └── test_fao56.py           ← FAO-56 핵심 계산 단위 테스트 (pytest)
│
└── docs/
    ├── THEORY.md               ← 관수·토양학 이론 (농대생 입문용)
    ├── ARCHITECTURE.md         ← 코드 구조·함수 레퍼런스 (개발자용)
    └── RESULTS_GUIDE.md        ← 엑셀 결과 해석 방법
```

---

## 📚 문서

| 문서 | 대상 | 내용 |
| :--- | :--- | :--- |
| [THEORY.md](docs/THEORY.md) | 농대생·입문자 | ETo·ETc·TAW·RAW·물수지 이론 |
| [ARCHITECTURE.md](docs/ARCHITECTURE.md) | 개발자·연구자 | 함수 레퍼런스·데이터 흐름·확장 방법 |
| [RESULTS_GUIDE.md](docs/RESULTS_GUIDE.md) | 현장 사용자 | 엑셀 시트별 해석·주의사항 |

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
