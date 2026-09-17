"""
temp_core.py — 온도·적산온도(GDD) 계산 엔진

farmos-cropweather 프로젝트의 온도 분석 핵심 모듈.
solar_core.py에서 수집한 ASOS 데이터(avg_ta, max_ta, min_ta)를
입력으로 받아 적산온도 계산, 도달일자 산출 등을 수행한다.

데이터 수집:
  solar_core.fetch_all_stations() → parse_rows() 가 반환하는
  dict 리스트를 그대로 사용. 각 dict에는 다음 키가 포함:
    year, month, day, date, avg_ta, max_ta, min_ta, ...

계산 방법 (WEATHER_THEORY.md 참조):
  방법 1 (기본): GDD = max((Tmax + Tmin)/2 - Tbase, 0)
  방법 2:        GDD = max(avgTa - Tbase, 0)
  방법 3:        방법 1 + Tupper 반영 (AquaCrop 방식)

참고문헌:
  - Paredes et al. (2025), Agric. Water Manag., 319, 109755
  - McMaster & Wilhelm (1997), Agric. For. Meteorol., 87(4)
  - FAO56rev (Pereira et al., 2025)
"""

import csv
import math
from datetime import date, timedelta
from collections import defaultdict
from pathlib import Path

# ---------------------------------------------------------------------------
# 상수
# ---------------------------------------------------------------------------

DEFAULT_T_BASE = 10.0
DEFAULT_T_UPPER = 35.0
DEFAULT_METHOD = "maxmin"       # "maxmin" | "avg" | "aquacrop"
DEFAULT_MILESTONES = [500, 1000, 1500, 2000, 2500, 3000, 3500, 4000]

CROPS_GDD_CSV = Path(__file__).parent / "crops_gdd.csv"


# ═══════════════════════════════════════════════════════════════════════════
# 1. 작물 GDD 라이브러리
# ═══════════════════════════════════════════════════════════════════════════

def load_crops_gdd(csv_path=None):
    """crops_gdd.csv 로드 → {crop_id: {...}} 딕셔너리 반환.

    Returns:
        dict[str, dict]: 키=crop_id, 값={
            "crop_id", "crop_name_ko", "crop_name_en",
            "t_base" (float), "t_upper" (float|None),
            "milestones" (list[int]), "source" (str)
        }

    Examples:
        >>> lib = load_crops_gdd()
        >>> lib["apple"]["t_base"]
        4.0
        >>> lib["apple"]["milestones"]
        [500, 1000, 1500, 2000, 2500, 3000, 3500, 4000]
    """
    path = Path(csv_path) if csv_path else CROPS_GDD_CSV
    if not path.exists():
        raise FileNotFoundError(f"작물 GDD 라이브러리를 찾을 수 없습니다: {path}")

    crops = {}
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            crop_id = row["crop_id"].strip()
            t_upper_raw = row.get("t_upper", "").strip()
            milestones_raw = row.get("gdd_milestones", "").strip()

            crops[crop_id] = {
                "crop_id": crop_id,
                "crop_name_ko": row["crop_name_ko"].strip(),
                "crop_name_en": row["crop_name_en"].strip(),
                "t_base": float(row["t_base"]),
                "t_upper": float(t_upper_raw) if t_upper_raw else None,
                "milestones": [int(x) for x in milestones_raw.split(",")]
                              if milestones_raw else DEFAULT_MILESTONES,
                "source": row.get("source", "").strip(),
            }
    return crops


def list_crops(csv_path=None):
    """작물 목록을 (crop_id, 한글명, Tbase) 튜플 리스트로 반환."""
    lib = load_crops_gdd(csv_path)
    return [(v["crop_id"], v["crop_name_ko"], v["t_base"])
            for v in sorted(lib.values(), key=lambda x: x["crop_id"])]


# ═══════════════════════════════════════════════════════════════════════════
# 2. 일별 GDD 계산
# ═══════════════════════════════════════════════════════════════════════════

def calc_gdd(avg_ta, max_ta, min_ta, t_base=DEFAULT_T_BASE,
             t_upper=DEFAULT_T_UPPER, method=DEFAULT_METHOD):
    """단일 일의 GDD를 계산한다.

    Args:
        avg_ta: 일평균기온 (℃) — ASOS avgTa
        max_ta: 일최고기온 (℃) — ASOS maxTa
        min_ta: 일최저기온 (℃) — ASOS minTa
        t_base: 기준온도 (℃)
        t_upper: 상한온도 (℃, method="aquacrop" 시 사용)
        method: "maxmin" | "avg" | "aquacrop"

    Returns:
        float: GDD 값 (≥ 0), 입력이 None이면 None

    Raises:
        ValueError: 알 수 없는 method
    """
    if method == "maxmin":
        if max_ta is None or min_ta is None:
            return None
        return max((max_ta + min_ta) / 2.0 - t_base, 0.0)

    elif method == "avg":
        if avg_ta is None:
            return None
        return max(avg_ta - t_base, 0.0)

    elif method == "aquacrop":
        if avg_ta is None or max_ta is None or min_ta is None:
            return None
        if avg_ta < t_base:
            return 0.0
        elif avg_ta > t_upper:
            return t_upper - t_base
        else:
            return max((max_ta + min_ta) / 2.0 - t_base, 0.0)

    else:
        raise ValueError(f"알 수 없는 GDD 계산 방법: {method!r}  "
                         f"('maxmin', 'avg', 'aquacrop' 중 선택)")


def add_gdd_to_rows(rows, t_base=DEFAULT_T_BASE, t_upper=DEFAULT_T_UPPER,
                    method=DEFAULT_METHOD):
    """파싱된 행 리스트에 "gdd" 키를 추가한다 (in-place).

    solar_core.parse_rows() 가 반환하는 dict 리스트를 입력으로 받는다.

    Args:
        rows: [{"year", "month", "day", "avg_ta", "max_ta", "min_ta", ...}, ...]
        t_base, t_upper, method: calc_gdd 파라미터

    Returns:
        rows (동일 리스트, gdd 키 추가됨)
    """
    for r in rows:
        r["gdd"] = calc_gdd(
            r.get("avg_ta"), r.get("max_ta"), r.get("min_ta"),
            t_base, t_upper, method,
        )
    return rows


# ═══════════════════════════════════════════════════════════════════════════
# 3. 온도·GDD 집계
# ═══════════════════════════════════════════════════════════════════════════

def daily_average_rows(rows):
    """다중 지점 일별 데이터 → 일별 지점평균 행으로 변환.

    기상청 전국 통계 방식과 동일.
    solar_multi.py의 _daily_average_rows()와 같은 로직이나,
    온도·GDD 키에 특화된 버전.
    """
    temp_keys = ["avg_ta", "max_ta", "min_ta", "gdd"]

    daily = defaultdict(lambda: defaultdict(list))
    for r in rows:
        day_key = (r["year"], r["month"], r["day"])
        for key in temp_keys:
            v = r.get(key)
            if v is not None:
                daily[day_key][key].append(v)

    result = []
    for (year, month, day), metrics in sorted(daily.items()):
        row = {"year": year, "month": month, "day": day}
        for key, vals in metrics.items():
            row[key] = sum(vals) / len(vals)
        result.append(row)

    return result


def monthly_stats(rows, metric_key, agg_type):
    """일별 데이터 → {year: {month: value}} 월별 집계.

    Args:
        rows: 일별 행 리스트 (단일 지점 또는 daily_average_rows 결과)
        metric_key: "avg_ta" | "max_ta" | "min_ta" | "gdd"
        agg_type: "avg" (월 내 일평균) | "cumsum" (월 내 합계)

    Returns:
        dict[int, dict[int, float]]: {year: {month: value}}
    """
    buckets = defaultdict(lambda: defaultdict(list))
    for r in rows:
        v = r.get(metric_key)
        if v is not None:
            buckets[r["year"]][r["month"]].append(v)

    result = {}
    for year, months in buckets.items():
        result[year] = {}
        for month, vals in months.items():
            if agg_type == "avg":
                result[year][month] = sum(vals) / len(vals)
            else:  # cumsum
                result[year][month] = sum(vals)
    return result


def period_stats(rows, metric_key, agg_type):
    """일별 데이터 → {year: value} 기간 전체 집계.

    Args:
        rows: 일별 행 리스트
        metric_key: "avg_ta" | "max_ta" | "min_ta" | "gdd"
        agg_type: "avg" (기간 내 일평균) | "cumsum" (기간 내 합계)

    Returns:
        dict[int, float]: {year: value}
    """
    buckets = defaultdict(list)
    for r in rows:
        v = r.get(metric_key)
        if v is not None:
            buckets[r["year"]].append(v)

    result = {}
    for year, vals in buckets.items():
        if agg_type == "avg":
            result[year] = sum(vals) / len(vals)
        else:
            result[year] = sum(vals)
    return result


# ═══════════════════════════════════════════════════════════════════════════
# 4. 평년 산출
# ═══════════════════════════════════════════════════════════════════════════

def calc_normal_value(year_vals, base_year, normal_years=10):
    """연도별 값 딕셔너리에서 평년 평균 계산.

    Args:
        year_vals: {year: value}
        base_year: 기준 연도 (올해)
        normal_years: 평년 산정 기간 (기본 10)

    Returns:
        float | None
    """
    n_start = base_year - normal_years
    n_end = base_year - 1
    vals = [year_vals[y] for y in range(n_start, n_end + 1) if y in year_vals]
    return sum(vals) / len(vals) if vals else None


def calc_normal_monthly(monthly_by_year, base_year, normal_years=10):
    """연도별 월별 값에서 평년 월별 평균 계산.

    Args:
        monthly_by_year: {year: {month: value}}
        base_year: 기준 연도
        normal_years: 평년 산정 기간

    Returns:
        dict[int, float|None]: {month: value}
    """
    n_start = base_year - normal_years
    n_end = base_year - 1
    normal_vals = {}
    for month in range(1, 13):
        yr_vals = []
        for yr in range(n_start, n_end + 1):
            v = monthly_by_year.get(yr, {}).get(month)
            if v is not None:
                yr_vals.append(v)
        normal_vals[month] = (sum(yr_vals) / len(yr_vals)) if yr_vals else None
    return normal_vals


# ═══════════════════════════════════════════════════════════════════════════
# 5. 적산온도 도달 일자
# ═══════════════════════════════════════════════════════════════════════════

def gdd_reaching_dates(rows, t_base=DEFAULT_T_BASE, t_upper=DEFAULT_T_UPPER,
                       method=DEFAULT_METHOD, milestones=None,
                       start_month=1, start_day=1):
    """연도별 적산온도 도달 일자 산출.

    1월 1일(또는 지정 시작일)부터 일별 GDD를 누적하여,
    각 임계값을 처음 초과하는 날짜를 기록한다.

    Args:
        rows: 일별 행 리스트 (단일 지점 또는 daily_average_rows 결과)
              gdd 키가 없으면 자동 계산
        t_base, t_upper, method: GDD 계산 파라미터
        milestones: 임계값 리스트 (기본: [500, 1000, ..., 4000])
        start_month, start_day: 적산 시작 월-일

    Returns:
        dict[int, dict[int, str|None]]:
            {year: {milestone: "M-D" or None}}
            None = 해당 연도에 미도달
    """
    if milestones is None:
        milestones = DEFAULT_MILESTONES
    milestones = sorted(milestones)

    # gdd 키가 없으면 추가
    if rows and "gdd" not in rows[0]:
        add_gdd_to_rows(rows, t_base, t_upper, method)

    # 연도별 일별 데이터 정렬
    by_year = defaultdict(list)
    for r in rows:
        by_year[r["year"]].append(r)
    for year in by_year:
        by_year[year].sort(key=lambda x: (x["month"], x["day"]))

    results = {}
    for year, year_rows in sorted(by_year.items()):
        cumsum = 0.0
        reached = {}
        remaining = list(milestones)

        for r in year_rows:
            # 적산 시작일 이전은 건너뜀
            if (r["month"], r["day"]) < (start_month, start_day):
                continue

            gdd = r.get("gdd")
            if gdd is None:
                continue

            cumsum += gdd

            # 아직 미도달인 임계값 중 초과한 것 기록
            while remaining and cumsum >= remaining[0]:
                threshold = remaining.pop(0)
                reached[threshold] = f"{r['month']}-{r['day']}"

        # 미도달 임계값은 None
        for m in remaining:
            reached[m] = None

        results[year] = reached

    return results


def gdd_reaching_dates_normal(yearly_dates, base_year, normal_years=10):
    """평년 적산온도 도달 일자 산출.

    과거 N년간 각 연도의 도달일자를 DOY(Day of Year)로 변환하여
    평균 DOY를 구하고, 다시 "M-D" 형식으로 변환한다.

    Args:
        yearly_dates: gdd_reaching_dates()의 반환값
        base_year: 기준 연도
        normal_years: 평년 산정 기간

    Returns:
        dict[int, str|None]: {milestone: "M-D" or None}
    """
    n_start = base_year - normal_years
    n_end = base_year - 1

    # 모든 임계값 수집
    all_milestones = set()
    for year_data in yearly_dates.values():
        all_milestones.update(year_data.keys())

    normal = {}
    for ms in sorted(all_milestones):
        doys = []
        for yr in range(n_start, n_end + 1):
            if yr not in yearly_dates:
                continue
            date_str = yearly_dates[yr].get(ms)
            if date_str is None:
                continue
            # "M-D" → DOY
            month, day = map(int, date_str.split("-"))
            try:
                d = date(yr, month, day)
                doy = d.timetuple().tm_yday
                doys.append(doy)
            except ValueError:
                continue

        if doys:
            avg_doy = round(sum(doys) / len(doys))
            # DOY → "M-D" (기준연도 사용, 윤년 고려)
            try:
                ref_date = date(base_year, 1, 1) + timedelta(days=avg_doy - 1)
                normal[ms] = f"{ref_date.month}-{ref_date.day}"
            except (ValueError, OverflowError):
                normal[ms] = None
        else:
            normal[ms] = None

    return normal


# ═══════════════════════════════════════════════════════════════════════════
# 6. 동기간 비교 유틸
# ═══════════════════════════════════════════════════════════════════════════

def filter_same_period(rows, start_md, end_md):
    """MM-DD 기반 동기간 필터.

    Args:
        rows: 일별 행 리스트
        start_md: (month, day) 시작
        end_md: (month, day) 종료

    Returns:
        동기간에 해당하는 행만 포함한 리스트
    """
    filtered = []
    for r in rows:
        md = (r["month"], r["day"])
        if start_md <= md <= end_md:
            filtered.append(r)
    return filtered


# ═══════════════════════════════════════════════════════════════════════════
# 7. 편의 함수 — 분석 파이프라인
# ═══════════════════════════════════════════════════════════════════════════

def analyze_temperature(parsed_rows, t_base=DEFAULT_T_BASE,
                        t_upper=DEFAULT_T_UPPER, method=DEFAULT_METHOD,
                        ref_year=None, normal_years=10,
                        milestones=None):
    """단일 지점 또는 권역의 종합 온도 분석.

    solar_core.parse_rows()의 반환값을 입력으로 받아
    GDD 계산 → 월별·기간 집계 → 도달일자 → 평년까지 한번에 수행한다.

    Args:
        parsed_rows: [{"year", "month", "day", "avg_ta", "max_ta", "min_ta", ...}]
        t_base, t_upper, method: GDD 계산 파라미터
        ref_year: 기준 연도 (기본: 데이터의 최대 연도)
        normal_years: 평년 기간
        milestones: 도달일자 임계값 리스트

    Returns:
        dict: {
            "params": {t_base, t_upper, method, ref_year, normal_years},
            "monthly_avg_ta":  {year: {month: float}},
            "monthly_max_ta":  {year: {month: float}},
            "monthly_min_ta":  {year: {month: float}},
            "monthly_gdd":     {year: {month: float}},
            "period_avg_ta":   {year: float},
            "period_gdd":      {year: float},
            "reaching_dates":  {year: {milestone: "M-D"|None}},
            "normal_reaching": {milestone: "M-D"|None},
        }
    """
    if ref_year is None:
        ref_year = max(r["year"] for r in parsed_rows)

    # GDD 계산
    add_gdd_to_rows(parsed_rows, t_base, t_upper, method)

    # 월별 집계
    m_avg = monthly_stats(parsed_rows, "avg_ta", "avg")
    m_max = monthly_stats(parsed_rows, "max_ta", "avg")
    m_min = monthly_stats(parsed_rows, "min_ta", "avg")
    m_gdd = monthly_stats(parsed_rows, "gdd", "cumsum")

    # 기간 전체 집계
    p_avg = period_stats(parsed_rows, "avg_ta", "avg")
    p_gdd = period_stats(parsed_rows, "gdd", "cumsum")

    # 도달 일자
    reaching = gdd_reaching_dates(
        parsed_rows, t_base, t_upper, method, milestones,
    )
    normal_reaching = gdd_reaching_dates_normal(
        reaching, ref_year, normal_years,
    )

    return {
        "params": {
            "t_base": t_base,
            "t_upper": t_upper,
            "method": method,
            "ref_year": ref_year,
            "normal_years": normal_years,
        },
        "monthly_avg_ta": m_avg,
        "monthly_max_ta": m_max,
        "monthly_min_ta": m_min,
        "monthly_gdd": m_gdd,
        "period_avg_ta": p_avg,
        "period_gdd": p_gdd,
        "reaching_dates": reaching,
        "normal_reaching": normal_reaching,
    }


# ═══════════════════════════════════════════════════════════════════════════
# 8. 자체 테스트 (python temp_core.py 로 실행)
# ═══════════════════════════════════════════════════════════════════════════

def _self_test():
    """기본 동작 검증용 자체 테스트."""
    print("=" * 60)
    print("temp_core.py 자체 테스트")
    print("=" * 60)

    # --- 1. GDD 계산 ---
    print("\n[1] calc_gdd 단위 테스트")

    # 방법 1 (maxmin): (20+8)/2 - 10 = 4.0
    assert calc_gdd(13.5, 20, 8, 10, method="maxmin") == 4.0
    print("  ✓ maxmin: (20+8)/2 - 10 = 4.0")

    # 방법 1: 기준온도 미만 → 0
    assert calc_gdd(5.0, 8, 2, 10, method="maxmin") == 0.0
    print("  ✓ maxmin: (8+2)/2=5 < 10 → 0.0")

    # 방법 2 (avg): 15 - 10 = 5.0
    assert calc_gdd(15, 20, 10, 10, method="avg") == 5.0
    print("  ✓ avg: 15 - 10 = 5.0")

    # 방법 2: 기준온도 미만 → 0
    assert calc_gdd(8, 12, 4, 10, method="avg") == 0.0
    print("  ✓ avg: 8 < 10 → 0.0")

    # 방법 3 (aquacrop): avg > Tupper → Tupper - Tbase
    assert calc_gdd(36, 40, 32, 10, 35, method="aquacrop") == 25.0
    print("  ✓ aquacrop: avg=36 > Tupper=35 → 35-10 = 25.0")

    # 방법 3: 정상 범위
    assert calc_gdd(20, 25, 15, 10, 35, method="aquacrop") == 10.0
    print("  ✓ aquacrop: 정상 범위 → (25+15)/2 - 10 = 10.0")

    # None 처리
    assert calc_gdd(None, 20, 10, 10, method="maxmin") == 5.0   # avg 불필요
    assert calc_gdd(15, None, 10, 10, method="maxmin") is None   # max 필요
    assert calc_gdd(None, 20, 10, 10, method="avg") is None      # avg 필요
    print("  ✓ None 입력 → 방법별 필수 필드 없으면 None")

    # --- 2. 작물 라이브러리 ---
    print("\n[2] crops_gdd.csv 로드 테스트")
    try:
        lib = load_crops_gdd()
        print(f"  ✓ {len(lib)}개 작물 로드 완료")
        apple = lib.get("apple")
        if apple:
            assert apple["t_base"] == 4.0
            assert apple["t_upper"] == 36.0
            assert 500 in apple["milestones"]
            print(f"  ✓ apple: Tbase={apple['t_base']}, "
                  f"Tupper={apple['t_upper']}, "
                  f"milestones={apple['milestones']}")
        grape = lib.get("grape")
        if grape:
            assert grape["t_base"] == 10.0
            print(f"  ✓ grape: Tbase={grape['t_base']}")
    except FileNotFoundError:
        print("  ⚠ crops_gdd.csv 파일 없음 — 건너뜀")

    # --- 3. 도달 일자 계산 ---
    print("\n[3] gdd_reaching_dates 테스트")

    # 합성 데이터: 4월 1일부터 30일간, 매일 avg=20, max=25, min=15
    test_rows = []
    for d in range(1, 31):
        test_rows.append({
            "year": 2024, "month": 4, "day": d,
            "avg_ta": 20.0, "max_ta": 25.0, "min_ta": 15.0,
        })
    # Tbase=10 → 일 GDD = (25+15)/2 - 10 = 10
    # 10일 → 100, 50일 → 500 ... 하지만 30일만 있으므로 300까지
    reaching = gdd_reaching_dates(
        test_rows, t_base=10, milestones=[100, 200, 300, 500],
    )
    assert reaching[2024][100] == "4-10"   # 10일 × 10 = 100
    assert reaching[2024][200] == "4-20"
    assert reaching[2024][300] == "4-30"
    assert reaching[2024][500] is None     # 미도달
    print(f"  ✓ 100℃ → {reaching[2024][100]}, "
          f"200℃ → {reaching[2024][200]}, "
          f"300℃ → {reaching[2024][300]}, "
          f"500℃ → {reaching[2024][500]}")

    # --- 4. 평년 도달일자 ---
    print("\n[4] gdd_reaching_dates_normal 테스트")

    # 3개 연도 합성 데이터
    multi_year_rows = []
    for yr in [2022, 2023, 2024]:
        for d in range(1, 181):  # 1/1 ~ 6/29 (180일)
            m = 1
            accum = d
            for days_in_m in [31, 28, 31, 30, 31, 30]:
                if accum <= days_in_m:
                    break
                accum -= days_in_m
                m += 1
            day = accum
            # 약간의 연도별 변동 추가
            base_temp = 10 + 10 * math.sin((d - 80) / 365 * 2 * math.pi)
            multi_year_rows.append({
                "year": yr, "month": m, "day": day,
                "avg_ta": base_temp + (yr - 2023) * 0.5,
                "max_ta": base_temp + 5 + (yr - 2023) * 0.5,
                "min_ta": base_temp - 5 + (yr - 2023) * 0.5,
            })

    reaching_multi = gdd_reaching_dates(
        multi_year_rows, t_base=10, milestones=[500, 1000],
    )
    normal = gdd_reaching_dates_normal(reaching_multi, base_year=2024, normal_years=3)
    print(f"  ✓ 평년 도달일자 (2022~2024 평균):")
    for ms, d in sorted(normal.items()):
        print(f"    {ms}℃ → {d}")

    # --- 5. 집계 함수 ---
    print("\n[5] 월별·기간 집계 테스트")

    add_gdd_to_rows(test_rows, t_base=10)
    m_gdd = monthly_stats(test_rows, "gdd", "cumsum")
    p_gdd = period_stats(test_rows, "gdd", "cumsum")
    assert m_gdd[2024][4] == 300.0   # 30일 × 10
    assert p_gdd[2024] == 300.0
    print(f"  ✓ 2024-04 GDD 합: {m_gdd[2024][4]}")
    print(f"  ✓ 2024 기간 GDD 합: {p_gdd[2024]}")

    # --- 완료 ---
    print("\n" + "=" * 60)
    print("모든 테스트 통과 ✓")
    print("=" * 60)


if __name__ == "__main__":
    _self_test()
