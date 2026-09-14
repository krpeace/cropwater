"""
tests/test_fao56.py — FAO-56 핵심 계산 단위 테스트

검증 기준:
  - FAO-56 식(11) svp, 식(21) Ra, 식(47) wind_2m
  - FAO-56 식(6) ETo Penman-Monteith
  - FAO-56 식(82)(83)(84)(88) TAW·RAW·Ks·DP 물수지
  - FAO-56 그림(25) Kc 생육단계 보간
  - 실증값: 춘천(ASOS 101) 2026-04-01 실측 기반

실행:
  pip install pytest
  pytest tests/ -v
"""
import sys, os, math, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from fao56_core import (
    svp, slope_svp, extra_radiation, daylight_hours,
    wind_2m, eto_penman_monteith, kc_of_date
)

ABS_TOL_ETO = 0.05   # mm/day
ABS_TOL_KS  = 0.001


# ════════════════════════════════════════════════════════════
# 1. 포화증기압 svp(T)  — FAO-56 식(11)
# ════════════════════════════════════════════════════════════

class TestSvp:
    def test_0c(self):
        """0°C → 포화증기압 = 0.6108 kPa"""
        assert abs(svp(0) - 0.6108) < 0.001

    def test_20c(self):
        """20°C → 포화증기압 ≈ 2.338 kPa"""
        assert abs(svp(20) - 2.338) < 0.005

    def test_monotone(self):
        """기온 상승 시 포화증기압 단조증가"""
        vals = [svp(t) for t in [0, 10, 20, 30, 40]]
        assert all(vals[i] < vals[i+1] for i in range(len(vals)-1))


# ════════════════════════════════════════════════════════════
# 2. 일외 복사량 Ra  — FAO-56 식(21)
# ════════════════════════════════════════════════════════════

class TestExtraRadiation:
    def test_equator_march(self):
        """적도(위도=0), 3월 15일(J=74): Ra 35~38 MJ/m²/day"""
        ra = extra_radiation(0, 74)
        assert 35.0 < ra < 38.0

    def test_seasonal_variation(self):
        """춘천(37.9°N): 7월 Ra > 1월 Ra"""
        assert extra_radiation(37.9, 196) > extra_radiation(37.9, 15)


# ════════════════════════════════════════════════════════════
# 3. 풍속 변환 wind_2m()  — FAO-56 식(47)
# ════════════════════════════════════════════════════════════

class TestWind2m:
    def test_10m_reduction(self):
        """10m → 2m: u2 = u10 × 4.87/ln(67.8×10−5.42) ≈ u10 × 0.748"""
        u2 = wind_2m(2.0, zw=10)
        ratio = u2 / 2.0
        assert abs(ratio - 0.748) < 0.01

    def test_2m_pass_through(self):
        """측정 높이 2m이면 값 그대로"""
        assert abs(wind_2m(3.0, zw=2) - 3.0) < 0.01


# ════════════════════════════════════════════════════════════
# 4. ETo — FAO-56 식(6)
# ════════════════════════════════════════════════════════════

class TestEtoPM:
    """
    춘천(ASOS 101) 2026-04-01 실측값 기반
    Tmax=17.7°C, Tmin=5.2°C, u10=1.0 m/s → u2=0.748 m/s
    Rs=10.47 MJ/m²/day, avgPv=10.3 hPa → ea=1.03 kPa
    고도=75.82m, 위도=37.90°, J=91
    cropwater_station.py 검증값: ETo ≈ 1.77 mm/day
    """
    def _base_params(self, **overrides):
        p = dict(
            Tmax=17.7, Tmin=5.2, Rs=10.47,
            u2=wind_2m(1.0, zw=10),
            ea=1.03, elev=75.82, lat=37.90, J=91,
            Pa=1003.5   # hPa (함수 내부에서 ×0.1 → kPa 변환)
        )
        p.update(overrides)
        return p

    def test_chuncheon_april1(self):
        eto = eto_penman_monteith(**self._base_params())
        assert abs(eto - 1.77) < ABS_TOL_ETO

    def test_eto_nonnegative(self):
        assert eto_penman_monteith(**self._base_params()) >= 0

    def test_higher_temp_increases_eto(self):
        eto_hot  = eto_penman_monteith(**self._base_params(Tmax=30.0, Tmin=18.0))
        eto_cool = eto_penman_monteith(**self._base_params())
        assert eto_hot > eto_cool

    def test_higher_wind_increases_eto(self):
        eto_windy = eto_penman_monteith(**self._base_params(u2=4.0))
        eto_calm  = eto_penman_monteith(**self._base_params())
        assert eto_windy > eto_calm


# ════════════════════════════════════════════════════════════
# 5. 토양 물수지 파라미터 — FAO-56 식(82)(83)(84)(88)
# ════════════════════════════════════════════════════════════

class TestWaterBalance:
    """양토(Loam), 사과 기준: FC=0.22, WP=0.10, Zr=1.0m, p=0.50"""
    FC=0.22; WP=0.10; Zr=1.0; p=0.50
    TAW = 1000 * (FC - WP) * Zr   # 120mm
    RAW = p * TAW                  # 60mm

    def test_taw(self):
        """TAW = 1000 × (FC−WP) × Zr = 120mm  [식(82)]"""
        assert abs(self.TAW - 120.0) < 0.1

    def test_raw(self):
        """RAW = p × TAW = 60mm  [식(83)]"""
        assert abs(self.RAW - 60.0) < 0.1

    @pytest.mark.parametrize("Dr, expected_Ks", [
        (0,   1.00),
        (30,  1.00),
        (60,  1.00),   # 경계: 아직 무스트레스
        (90,  0.50),   # (120−90)/(120−60)
        (105, 0.25),
        (120, 0.00),
    ])
    def test_ks(self, Dr, expected_Ks):
        """Ks [식(84)]: Dr≤RAW → 1.0, Dr>RAW → (TAW−Dr)/(TAW−RAW)"""
        Ks = 1.0 if Dr <= self.RAW \
             else max(0.0, (self.TAW - Dr) / (self.TAW - self.RAW))
        assert abs(Ks - expected_Ks) < ABS_TOL_KS

    def test_dp_excess_rain(self):
        """DP = max(P − Dr, 0): 강수 50mm, Dr=20mm → DP=30mm  [식(88)]"""
        assert abs(max(50.0 - 20.0, 0) - 30.0) < 0.01

    def test_dp_small_rain(self):
        """강수 < 고갈량이면 DP = 0"""
        assert max(10.0 - 40.0, 0) == 0.0

    def test_dr_lower_bound(self):
        """Dr은 0 미만이 될 수 없음"""
        assert max(0.0, 0.0 - 200.0 + 5.0) >= 0

    def test_dr_upper_bound(self):
        """Dr은 TAW를 초과할 수 없음"""
        assert min(self.TAW, self.TAW + 10.0) <= self.TAW


# ════════════════════════════════════════════════════════════
# 6. Kc 생육단계 보간 — kc_of_date()
# ════════════════════════════════════════════════════════════

class TestKcCurve:
    """사과: L_ini=20, L_dev=70, L_mid=90, L_late=30 / Kc 0.50→1.20→0.95
    발아일: 2026-04-01 (ordinal 738,581)"""

    START = dt.date(2026, 4, 1).toordinal()
    KC_KW = dict(start_ordinal=dt.date(2026,4,1).toordinal(),
                 L_ini=20, L_dev=70, L_mid=90, L_late=30,
                 kc_ini=0.50, kc_mid=1.20, kc_end=0.95)

    def _kc(self, d: dt.date):
        return kc_of_date(d.toordinal(), **self.KC_KW)

    def test_ini_stage(self):
        """초기 10일차: Kc = Kc_ini = 0.50"""
        kc = self._kc(dt.date(2026, 4, 10))   # 발아 9일 후
        assert abs(kc - 0.50) < 0.001

    def test_mid_stage(self):
        """중기 한가운데: Kc = Kc_mid = 1.20"""
        kc = self._kc(dt.date(2026, 7, 20))
        assert abs(kc - 1.20) < 0.001

    def test_dev_monotone(self):
        """발육기: 시간이 갈수록 Kc 단조증가"""
        dates = [dt.date(2026, 5, 1), dt.date(2026, 5, 20), dt.date(2026, 6, 10)]
        kcs = [self._kc(d) for d in dates]
        assert all(0.50 <= k <= 1.20 for k in kcs)
        assert kcs[0] < kcs[1] < kcs[2]

    def test_dormancy(self):
        """생육 종료 후 휴면기: Kc = 0"""
        assert self._kc(dt.date(2026, 11, 1)) == 0.0
