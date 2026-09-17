"""
test_temp.py — temp_core GDD 계산 단위 테스트 (pytest)

실행:
  pytest tests/test_temp.py -v
"""

import sys
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from temp_core import calc_gdd, load_crops_gdd, gdd_reaching_dates


class TestCalcGDD:
    """calc_gdd 단위 테스트"""

    def test_maxmin_normal(self):
        # (20+8)/2 - 10 = 4.0
        assert calc_gdd(13.5, 20, 8, 10, method="maxmin") == 4.0

    def test_maxmin_below_base(self):
        # (8+2)/2 = 5 < 10 → 0
        assert calc_gdd(5.0, 8, 2, 10, method="maxmin") == 0.0

    def test_avg_normal(self):
        assert calc_gdd(15, 20, 10, 10, method="avg") == 5.0

    def test_avg_below_base(self):
        assert calc_gdd(8, 12, 4, 10, method="avg") == 0.0

    def test_aquacrop_above_upper(self):
        # avg=36 > Tupper=35 → 35-10 = 25
        assert calc_gdd(36, 40, 32, 10, 35, method="aquacrop") == 25.0

    def test_aquacrop_normal_range(self):
        assert calc_gdd(20, 25, 15, 10, 35, method="aquacrop") == 10.0

    def test_none_handling_maxmin(self):
        assert calc_gdd(None, 20, 10, 10, method="maxmin") == 5.0  # avg 불필요
        assert calc_gdd(15, None, 10, 10, method="maxmin") is None

    def test_none_handling_avg(self):
        assert calc_gdd(None, 20, 10, 10, method="avg") is None


class TestCropsGDD:
    """crops_gdd.csv 로드 테스트"""

    def test_load_crops(self):
        lib = load_crops_gdd()
        assert len(lib) >= 40

    def test_apple_params(self):
        lib = load_crops_gdd()
        assert lib["apple"]["t_base"] == 4.0
        assert lib["apple"]["t_upper"] == 36.0
        assert 500 in lib["apple"]["milestones"]

    def test_grape_params(self):
        lib = load_crops_gdd()
        assert lib["grape"]["t_base"] == 10.0


class TestReachingDates:
    """gdd_reaching_dates 테스트"""

    def test_synthetic_data(self):
        # 30일간 매일 GDD=10 → 100℃=10일, 200℃=20일, 300℃=30일, 500℃=미도달
        rows = [
            {"year": 2024, "month": 4, "day": d,
             "avg_ta": 20.0, "max_ta": 25.0, "min_ta": 15.0}
            for d in range(1, 31)
        ]
        reaching = gdd_reaching_dates(rows, t_base=10, milestones=[100, 200, 300, 500])

        assert reaching[2024][100] == "4-10"
        assert reaching[2024][200] == "4-20"
        assert reaching[2024][300] == "4-30"
        assert reaching[2024][500] is None
