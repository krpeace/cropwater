"""
test_solar.py — solar_core 단위 테스트 (pytest)

API 호출 없이 순수 계산 함수만 검증한다.
네트워크·인증키가 없는 CI 환경에서도 실행 가능.

실행:
  pytest tests/test_solar.py -v
"""

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from solar_core import (
    safe_float,
    get_date_chunks,
    is_finalized_year,
    _cache_path,
    parse_rows,
    _avg,
    _sum,
    aggregate_monthly,
    aggregate_yearly,
    calc_normal,
    MAX_QUERY_YEARS,
    CACHE_DIR,
)


# ═══════════════════════════════════════════════════════════════════════════
# 유틸 함수
# ═══════════════════════════════════════════════════════════════════════════

class TestSafeFloat:
    def test_normal_string(self):
        assert safe_float("3.14") == 3.14

    def test_integer_string(self):
        assert safe_float("42") == 42.0

    def test_already_float(self):
        assert safe_float(1.5) == 1.5

    def test_empty_string(self):
        assert safe_float("") is None

    def test_whitespace_string(self):
        assert safe_float("  ") is None

    def test_none_input(self):
        assert safe_float(None) is None

    def test_non_numeric_string(self):
        assert safe_float("abc") is None

    def test_custom_default(self):
        assert safe_float("", default=0.0) == 0.0
        assert safe_float(None, default=-1.0) == -1.0

    def test_negative(self):
        assert safe_float("-5.2") == -5.2


# ═══════════════════════════════════════════════════════════════════════════
# 날짜 청크 분할
# ═══════════════════════════════════════════════════════════════════════════

class TestGetDateChunks:
    def test_single_chunk(self):
        # 1년 → MAX_QUERY_YEARS=2이면 청크 1개
        chunks = get_date_chunks(date(2024, 1, 1), date(2024, 12, 31), max_years=2)
        assert len(chunks) == 1
        assert chunks[0] == (date(2024, 1, 1), date(2024, 12, 31))

    def test_two_chunks(self):
        # 3년 범위, max_years=2 → 2개 청크
        chunks = get_date_chunks(date(2022, 1, 1), date(2024, 6, 30), max_years=2)
        assert len(chunks) == 2
        assert chunks[0] == (date(2022, 1, 1), date(2023, 12, 31))
        assert chunks[1] == (date(2024, 1, 1), date(2024, 6, 30))

    def test_five_chunks(self):
        # 10년 범위, max_years=2 → 5개 청크
        chunks = get_date_chunks(date(2016, 1, 1), date(2025, 12, 31), max_years=2)
        assert len(chunks) == 5

    def test_end_date_respected(self):
        # 마지막 청크의 end가 end_date를 넘지 않음
        chunks = get_date_chunks(date(2016, 1, 1), date(2026, 9, 15), max_years=2)
        assert chunks[-1][1] == date(2026, 9, 15)

    def test_same_day(self):
        chunks = get_date_chunks(date(2026, 1, 1), date(2026, 1, 1), max_years=2)
        assert len(chunks) == 1
        assert chunks[0] == (date(2026, 1, 1), date(2026, 1, 1))

    def test_default_max_years(self):
        # MAX_QUERY_YEARS=2 기본값 동작 확인
        chunks = get_date_chunks(date(2016, 1, 1), date(2026, 9, 15))
        # 2016~2017, 2018~2019, 2020~2021, 2022~2023, 2024~2025, 2026~
        assert len(chunks) == 6


# ═══════════════════════════════════════════════════════════════════════════
# 확정 연도 판별
# ═══════════════════════════════════════════════════════════════════════════

class TestIsFinalizedYear:
    def test_past_year(self):
        assert is_finalized_year(2020) is True

    def test_far_past(self):
        assert is_finalized_year(2016) is True

    def test_current_year(self):
        assert is_finalized_year(date.today().year) is False

    def test_future_year(self):
        assert is_finalized_year(date.today().year + 1) is False


# ═══════════════════════════════════════════════════════════════════════════
# 캐시 경로
# ═══════════════════════════════════════════════════════════════════════════

class TestCachePath:
    def test_path_format(self):
        p = _cache_path(108, date(2016, 1, 1), date(2017, 12, 31))
        assert p == CACHE_DIR / "108_2016_2017.csv"

    def test_single_year(self):
        p = _cache_path(184, date(2026, 1, 1), date(2026, 9, 15))
        assert p == CACHE_DIR / "184_2026_2026.csv"


# ═══════════════════════════════════════════════════════════════════════════
# 원시 데이터 파싱
# ═══════════════════════════════════════════════════════════════════════════

class TestParseRows:
    def _make_raw(self, tm, sum_gsr="15.2", sum_ss_hr="8.5",
                  ss_dur="10.0", avg_ta="22.3", max_ta="28.0",
                  min_ta="18.0", sum_rn="0.0", stn_id="108"):
        return {
            "tm": tm, "stnId": stn_id, "stnNm": "서울",
            "sumGsr": sum_gsr, "sumSsHr": sum_ss_hr, "ssDur": ss_dur,
            "hr1MaxIcsr": "2.1", "avgTa": avg_ta,
            "maxTa": max_ta, "minTa": min_ta, "sumRn": sum_rn,
        }

    def test_basic_parsing(self):
        rows = parse_rows([self._make_raw("2026-07-15")])
        assert len(rows) == 1
        r = rows[0]
        assert r["year"] == 2026
        assert r["month"] == 7
        assert r["day"] == 15
        assert r["sum_gsr"] == 15.2
        assert r["sum_ss_hr"] == 8.5
        assert r["avg_ta"] == 22.3
        assert r["max_ta"] == 28.0
        assert r["min_ta"] == 18.0

    def test_date_formats(self):
        # YYYY-MM-DD 형식
        rows1 = parse_rows([self._make_raw("2026-01-01")])
        assert rows1[0]["date"] == date(2026, 1, 1)

        # YYYYMMDD 형식
        rows2 = parse_rows([self._make_raw("20260101")])
        assert rows2[0]["date"] == date(2026, 1, 1)

    def test_ss_rate_calculation(self):
        # 일조율 = sum_ss_hr / ss_dur × 100
        rows = parse_rows([self._make_raw("2026-07-01",
                                          sum_ss_hr="8.0", ss_dur="10.0")])
        assert rows[0]["ss_rate"] == 80.0

    def test_ss_rate_zero_dur(self):
        # ss_dur=0 → 일조율 None (0 나누기 방지)
        rows = parse_rows([self._make_raw("2026-07-01",
                                          sum_ss_hr="5.0", ss_dur="0")])
        assert rows[0]["ss_rate"] is None

    def test_missing_values(self):
        rows = parse_rows([self._make_raw("2026-07-01",
                                          sum_gsr="", avg_ta="")])
        assert rows[0]["sum_gsr"] is None
        assert rows[0]["avg_ta"] is None

    def test_invalid_date_skipped(self):
        rows = parse_rows([self._make_raw("invalid-date")])
        assert rows == []

    def test_empty_tm_skipped(self):
        rows = parse_rows([self._make_raw("")])
        assert rows == []

    def test_multiple_rows(self):
        raw = [self._make_raw(f"2026-07-{d:02d}") for d in range(1, 6)]
        rows = parse_rows(raw)
        assert len(rows) == 5
        assert [r["day"] for r in rows] == [1, 2, 3, 4, 5]

    def test_stn_id_parsing(self):
        rows = parse_rows([self._make_raw("2026-07-01", stn_id="184")])
        assert rows[0]["stn_id"] == 184


# ═══════════════════════════════════════════════════════════════════════════
# 집계 헬퍼 (_avg, _sum)
# ═══════════════════════════════════════════════════════════════════════════

class TestAggHelpers:
    def test_avg_normal(self):
        assert _avg([1.0, 2.0, 3.0]) == 2.0

    def test_avg_with_none(self):
        assert _avg([1.0, None, 3.0]) == 2.0

    def test_avg_all_none(self):
        assert _avg([None, None]) is None

    def test_avg_empty(self):
        assert _avg([]) is None

    def test_sum_normal(self):
        assert _sum([10.0, 20.0, 30.0]) == 60.0

    def test_sum_with_none(self):
        assert _sum([10.0, None, 30.0]) == 40.0

    def test_sum_all_none(self):
        assert _sum([None, None]) is None

    def test_sum_empty(self):
        assert _sum([]) is None


# ═══════════════════════════════════════════════════════════════════════════
# 월별 집계
# ═══════════════════════════════════════════════════════════════════════════

def _make_parsed(year, month, day, sum_gsr, sum_ss_hr, ss_dur):
    """테스트용 parsed row 생성 헬퍼."""
    ss_rate = round(sum_ss_hr / ss_dur * 100, 1) if ss_dur and ss_dur > 0 else None
    return {
        "date": date(year, month, day),
        "year": year, "month": month, "day": day,
        "stn_id": 108, "stn_nm": "서울",
        "sum_gsr": sum_gsr, "sum_ss_hr": sum_ss_hr,
        "ss_dur": ss_dur, "ss_rate": ss_rate,
        "hr1_max_icsr": None, "avg_ta": 20.0,
        "max_ta": 25.0, "min_ta": 15.0, "sum_rn": 0.0,
        "_raw": {},
    }


class TestAggregateMonthly:
    def test_single_month(self):
        rows = [
            _make_parsed(2026, 7, 1, sum_gsr=20.0, sum_ss_hr=9.0, ss_dur=10.0),
            _make_parsed(2026, 7, 2, sum_gsr=10.0, sum_ss_hr=5.0, ss_dur=10.0),
            _make_parsed(2026, 7, 3, sum_gsr=15.0, sum_ss_hr=7.5, ss_dur=10.0),
        ]
        result = aggregate_monthly(rows)
        assert len(result) == 1
        m = result[0]
        assert m["year"] == 2026
        assert m["month"] == 7
        assert m["days_count"] == 3
        assert m["sum_gsr"] == 45.0            # 합계
        assert abs(m["avg_gsr"] - 15.0) < 0.01  # 평균

    def test_two_months(self):
        rows = [
            _make_parsed(2026, 6, 30, 18.0, 8.0, 10.0),
            _make_parsed(2026, 7,  1, 20.0, 9.0, 10.0),
        ]
        result = aggregate_monthly(rows)
        assert len(result) == 2
        assert result[0]["month"] == 6
        assert result[1]["month"] == 7

    def test_sorted_output(self):
        rows = [
            _make_parsed(2026, 7, 1, 10.0, 5.0, 10.0),
            _make_parsed(2026, 5, 1, 10.0, 5.0, 10.0),
            _make_parsed(2026, 6, 1, 10.0, 5.0, 10.0),
        ]
        result = aggregate_monthly(rows)
        assert [r["month"] for r in result] == [5, 6, 7]

    def test_none_values_excluded(self):
        rows = [
            _make_parsed(2026, 7, 1, None, None, None),
            _make_parsed(2026, 7, 2, 20.0, 8.0, 10.0),
        ]
        result = aggregate_monthly(rows)
        assert result[0]["sum_gsr"] == 20.0  # None 제외하고 합산
        assert result[0]["days_count"] == 2  # 행 수는 2개


# ═══════════════════════════════════════════════════════════════════════════
# 연도별 집계
# ═══════════════════════════════════════════════════════════════════════════

class TestAggregateYearly:
    def test_single_year(self):
        rows = [
            _make_parsed(2025, 1, d, sum_gsr=float(d*2),
                         sum_ss_hr=float(d), ss_dur=10.0)
            for d in range(1, 4)
        ]
        result = aggregate_yearly(rows)
        assert len(result) == 1
        y = result[0]
        assert y["year"] == 2025
        assert y["days_count"] == 3
        assert y["sum_gsr"] == 2.0 + 4.0 + 6.0  # 합계

    def test_two_years_sorted(self):
        rows = [
            _make_parsed(2025, 6, 1, 20.0, 8.0, 10.0),
            _make_parsed(2024, 6, 1, 15.0, 6.0, 10.0),
        ]
        result = aggregate_yearly(rows)
        assert result[0]["year"] == 2024
        assert result[1]["year"] == 2025


# ═══════════════════════════════════════════════════════════════════════════
# 평년 계산
# ═══════════════════════════════════════════════════════════════════════════

class TestCalcNormal:
    def _make_multi_year_rows(self, years):
        """여러 연도에 걸쳐 1월 데이터만 있는 합성 데이터."""
        rows = []
        for year in years:
            for day in range(1, 4):
                rows.append(_make_parsed(
                    year, 1, day,
                    sum_gsr=float(year - 2000),   # 연도별로 다른 값
                    sum_ss_hr=5.0,
                    ss_dur=10.0,
                ))
        return rows

    def test_normal_calculation(self):
        # 2016~2025 데이터, ref_year=2026, normal_years=10
        rows = self._make_multi_year_rows(range(2016, 2026))
        result = calc_normal(rows, ref_year=2026, normal_years=10)

        # 1월 평균 sum_gsr = 평년(2016~2025)의 각 연도 1월합 평균
        # 각 연도 1월합: 3일 × (year-2000) = 3×16=48, 3×17=51, ..., 3×25=75
        # 연도별 월합 목록: [48, 51, 54, 57, 60, 63, 66, 69, 72, 75]
        # 평균: sum([48..75, step3]) / 10 = 615 / 10 = 61.5
        jan_sum_gsr = result[1]["sum_gsr"]
        assert abs(jan_sum_gsr - 61.5) < 0.01

    def test_normal_period_recorded(self):
        rows = self._make_multi_year_rows(range(2016, 2026))
        result = calc_normal(rows, ref_year=2026, normal_years=10)
        assert result[1]["normal_start"] == 2016
        assert result[1]["normal_end"] == 2025

    def test_all_months_present(self):
        rows = self._make_multi_year_rows([2024, 2025])
        result = calc_normal(rows, ref_year=2026, normal_years=2)
        assert set(result.keys()) == set(range(1, 13))

    def test_no_data_month_returns_none(self):
        # 1월 데이터만 있음 → 2~12월은 None
        rows = self._make_multi_year_rows([2024, 2025])
        result = calc_normal(rows, ref_year=2026, normal_years=2)
        assert result[1]["sum_gsr"] is not None   # 1월 ✓
        assert result[7]["sum_gsr"] is None        # 7월 데이터 없음 → None

    def test_out_of_range_years_excluded(self):
        # ref_year=2026, normal_years=2 → 평년=2024~2025
        # 2023 데이터는 평년 범위 밖 → 포함되지 않아야 함
        rows_normal = self._make_multi_year_rows([2024, 2025])
        rows_extra = self._make_multi_year_rows([2023])
        result_with = calc_normal(rows_normal + rows_extra, ref_year=2026, normal_years=2)
        result_without = calc_normal(rows_normal, ref_year=2026, normal_years=2)
        # 2023이 포함되면 평균이 달라져야 하지만, 범위 밖이므로 동일해야 함
        assert abs(result_with[1]["sum_gsr"] - result_without[1]["sum_gsr"]) < 0.001
