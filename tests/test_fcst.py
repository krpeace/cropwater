# -*- coding: utf-8 -*-
"""02-Cycle 단기예보 모듈 단위 테스트 (pytest)
  fcst_archive: 포털 CSV 파싱, 발표·선행시간 → 대상시각, TMX/TMN 대상일, 지나간 시각 채움, 연장기간 코드 변환
  rs_model    : Rs 추정·계수 적합·계수 파일 2계층
  obs_daily   : Kc (01-Cycle 설정과 같은 식)
  cropwater_fcst / fcst_report : 지표·판정, 엑셀 PM 수식 = 파이썬 PM
"""
import math, os, sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from fao56_core import eto_penman_monteith, extra_radiation, kc_climate_adjust, slope_svp, svp
from fcst_archive import (Archive, _daily_targets, _hourly_targets, daily_inputs, detect_element,
                          read_portal_csv, service_table)
from rs_model import fit, load_coef, ra_rso, rs_s1, rs_s3, save_coef
from obs_daily import kc_params, kc_series
from cropwater_fcst import h2_verdict, lead_metrics, pressure_from_elev
from fcst_report import pm

TS = pd.Timestamp


# ── fcst_archive ──────────────────────────────────────────────────────
def test_read_portal_csv_issue_time_and_location(tmp_path):
    p = tmp_path / "춘천시신사우동_1시간기온_20260401_20260430.csv"
    p.write_text(" format: day(UTC),hour(UTC),forecast,value,day(KST),hour(KST) location:73_134 Start : 20260401\n"
                 "1,1700,+6,11.0,2,0200\n1,1700,+7,12.0,2,0200\n"
                 "Start : 20260402\n2,0200,+6,17.0,2,1100\n", encoding="utf-8")
    loc, df = read_portal_csv(p)
    assert loc == "73_134"
    assert list(df["issue"]) == [TS("2026-04-02 02:00"), TS("2026-04-02 02:00"), TS("2026-04-02 11:00")]
    assert detect_element(str(p), df) == "TMP"          # 파일명 키워드
    h = _hourly_targets(df)
    assert list(h["target"])[:2] == [TS("2026-04-02 08:00"), TS("2026-04-02 09:00")]   # 발표 + lead


def test_daily_targets_follow_guide_table():
    rows = []
    for hour in (2, 5, 8, 11, 14, 17, 20, 23):
        n_tmx = 3 if hour == 14 else 4
        for k in range(n_tmx):
            rows.append((TS(f"2026-04-02 {hour:02d}:00"), 6 + k, 20.0 + k))
    df = pd.DataFrame(rows, columns=["issue", "forecast", "value"])
    t = _daily_targets(df, "TMX")
    first = t.groupby("issue")["target_date"].min()
    assert first[TS("2026-04-02 11:00")] == TS("2026-04-02")    # 11시 발표까지 오늘 최고기온 포함
    assert first[TS("2026-04-02 14:00")] == TS("2026-04-03")
    assert first[TS("2026-04-02 17:00")] == TS("2026-04-03")
    t = _daily_targets(df, "TMN")
    first = t.groupby("issue")["target_date"].min()
    assert first[TS("2026-04-02 02:00")] == TS("2026-04-02")    # 오늘 최저기온은 02시 발표에만
    assert first[TS("2026-04-02 05:00")] == TS("2026-04-03")


def test_detect_daily_element_by_row_pattern():
    rows = []
    for d in range(3):
        for hour in (2, 5, 8, 11, 14, 17, 20, 23):
            n = 4 if hour in (2, 17, 20, 23) else 3             # TMN 패턴
            rows += [(TS(f"2026-04-0{d + 1} {hour:02d}:00"), 6 + k, 5.0) for k in range(n)]
    df = pd.DataFrame(rows, columns=["issue", "forecast", "value"])
    assert detect_element("x.csv", df) == "TMN"


def _hourly(issue, leads, values):
    return _hourly_targets(pd.DataFrame({"issue": TS(issue), "forecast": leads, "value": values}))


def _archive():
    """전날 17시 발표(00~07시 채움용) + 02시 서비스 발표(08시~D+3, 마지막 날 3시간·코드)"""
    a = Archive()
    run = "2026-04-02 02:00"
    lead_q = list(range(6, 71))                         # 08시 ~ D+3 00시 (1시간)
    lead_c = list(range(73, 92, 3))                     # D+3 03~21시 (3시간, 코드)
    leads = lead_q + lead_c
    prev = ("2026-04-01 17:00", list(range(7, 15)))     # 04-02 00~07시
    a.hourly["TMP"] = pd.concat([_hourly(prev[0], prev[1], [10.0] * 8), _hourly(run, leads, [20.0] * len(leads))])
    a.hourly["REH"] = pd.concat([_hourly(prev[0], prev[1], [80.0] * 8), _hourly(run, leads, [50.0] * len(leads))])
    wsd = [2.0] * len(lead_q) + [2.0] * len(lead_c)     # 코드 2 = 6.5 m/s
    a.hourly["WSD"] = pd.concat([_hourly(prev[0], prev[1], [1.0] * 8), _hourly(run, leads, wsd)])
    pcp = [0.0] * len(leads)
    pcp[leads.index(10)] = 1.5                          # 04-02 12시 1.5 mm
    pcp[leads.index(76)] = 1.0                          # D+3 06시 코드 1 → 1.5 mm/h × 3 h
    a.hourly["PCP"] = pd.concat([_hourly(prev[0], prev[1], [0.0] * 8), _hourly(run, leads, pcp)])
    d = pd.DataFrame({"issue": TS(run), "forecast": [6, 7, 8, 9], "value": [18.0, 19.0, 20.0, 21.0]})
    a.daily["TMX"] = _daily_targets(d, "TMX")
    a.daily["TMN"] = _daily_targets(d.assign(value=[5.0, 6.0, 7.0, 8.0]), "TMN")
    return a, TS(run)


def test_daily_inputs_fill_and_codes():
    a, run = _archive()
    out = daily_inputs(a, run, [run.normalize() + pd.Timedelta(days=k) for k in range(4)]).set_index("lead_day")
    d0 = out.loc[0]
    assert d0.hours == 24 and d0.filled == 8           # 00~07시는 전날 17시 발표로 채움
    ea_expected = (8 * svp(10.0) * 0.80 + 16 * svp(20.0) * 0.50) / 24
    assert d0.ea == pytest.approx(ea_expected)
    assert d0.u10 == pytest.approx((8 * 1.0 + 16 * 2.0) / 24)
    assert d0.rain == pytest.approx(1.5) and d0.rain_flag == 1
    assert (d0.Tmax, d0.Tmin) == (18.0, 5.0)
    d3 = out.loc[3]
    assert d3.ext and d3.hours == 8                    # 00시(1시간) + 03~21시(3시간)
    assert d3.u10 == pytest.approx((2.0 + 7 * 6.5) / 8)
    assert d3.rain == pytest.approx(4.5) and d3.rain_flag == 1


def test_service_table_runs():
    a, run = _archive()
    st = service_table(a, runs={"아침": (2, range(0, 4))})
    assert list(st.lead_day) == [0, 1, 2, 3] and set(st.run) == {run}


# ── rs_model ──────────────────────────────────────────────────────────
def test_rs_s3_clip_and_fit_recovers_coefficients(tmp_path):
    lat, elev = 37.9, 75.8
    dates = pd.date_range("2025-04-01", "2025-09-30")
    rng = np.random.default_rng(0)
    tmin = rng.uniform(5, 20, len(dates)); dT = rng.uniform(4, 14, len(dates))
    flag = (np.arange(len(dates)) % 3 == 0).astype(int)
    ra, rso = ra_rso(lat, elev, dates.dayofyear)
    rs = rs_s3(tmin + dT, tmin, flag, ra, rso, -0.1, 0.19, -0.11)
    obs = pd.DataFrame({"date": dates, "Tmax": tmin + dT, "Tmin": tmin, "Rs": rs, "rain": flag * 5.0})
    cf = fit(obs, lat, elev)
    assert (cf["a"], cf["b"], cf["c"]) == pytest.approx((-0.1, 0.19, -0.11), abs=1e-9)
    assert cf["n"] == len(dates) and cf["rmse_rs"] < 1e-9
    # 제한: 0.05Ra ~ Rso
    assert rs_s1(40, 0, ra[:1], rso[:1], 0.30)[0] == pytest.approx(rso[0])
    assert rs_s3(10, 10, 1, ra[:1], rso[:1], -0.1, 0.19, -0.11)[0] == pytest.approx(0.05 * ra[0])
    # 계수 파일 2계층
    path = tmp_path / "rs_coef.csv"
    save_coef("101", cf, path, note="test")
    assert load_coef("101", path)["a"] == pytest.approx(-0.1, abs=1e-4)
    d = load_coef("999", path)
    assert d["method"] == "S1" and d["krs"] == 0.16 and d["a"] is None


# ── Kc (01-Cycle 설정 시트와 같은 값) ─────────────────────────────────
def test_kc_params_and_series():
    st = {"시나리오 번호 (1~4)": 3, "생육 시작일 (발아기/정식일)": pd.Timestamp("2026-04-01"),
          "L_ini (초기, 일)": 20, "L_dev (발육, 일)": 70, "L_mid (중기, 일)": 90, "L_late (후기, 일)": 30,
          "생육중기 초목 수고 h (m)": 3.2, "중기 평균 u2 (m/s)": 0.998, "중기 평균 RHmin (%)": 55.33,
          "후기 평균 u2 (m/s)": 1.5, "후기 평균 RHmin (%)": 55, "멀칭 Kc 보정계수": 1}
    kp = kc_params(st)
    assert kp["kc_mid"] == pytest.approx(kc_climate_adjust(1.20, 0.998, 55.33, 3.2))
    kc = kc_series(pd.to_datetime(["2026-03-31", "2026-04-01", "2026-04-21", "2026-06-30"]), kp)
    assert kc[0] == 0.0 and kc[1] == 0.5
    assert kc[2] == pytest.approx(0.5 + (kp["kc_mid"] - 0.5) * 1 / 70)
    assert kc[3] == pytest.approx(kp["kc_mid"])


# ── 지표·판정 ─────────────────────────────────────────────────────────
def test_lead_metrics_and_verdict():
    o = np.array([3.0, 4.0, 5.0, 4.0])
    rows = []
    for rn, ks in (("아침", (0, 1, 2, 3)), ("저녁", (1, 2, 3, 4))):
        for k in ks:
            for i, v in enumerate(o):
                rows.append(dict(run_name=rn, lead_day=k, ETo_obs=v, ETo_S3=v + 0.5, ETo_pers=v + (1.0 if i % 2 else -1.0),
                                 ETo_7d=v + 0.8))
    met = lead_metrics(pd.DataFrame(rows))
    r = met.iloc[1]
    assert r.RMSE == pytest.approx(0.5) and r.MBE == pytest.approx(0.5)
    assert r.RMSE_pers == pytest.approx(1.0) and r.skill_pers == pytest.approx(0.5)
    v = h2_verdict(met)
    assert v["아침"]["pass_"] and v["저녁"]["pass_"]
    assert not h2_verdict(met, rmse_d1_max=0.4)["아침"]["pass_"]
    assert not h2_verdict(met, skill_min=0.6)["저녁"]["pass_"]


# ── 엑셀 PM 수식 = fao56_core PM ──────────────────────────────────────
def _eval_excel(expr):
    py = expr.lstrip("=").replace("^", "**").replace("SQRT", "math.sqrt").replace("MIN", "min")
    return eval(py, {"math": math, "min": min})


@pytest.mark.parametrize("tmax,tmin,ea,u2,rs,J", [(25.0, 12.0, 1.2, 1.5, 18.0, 150), (8.0, -2.0, 0.6, 0.8, 9.0, 95)])
def test_excel_pm_formula_matches_python(tmax, tmin, ea, u2, rs, J):
    lat, elev = 37.90262, 75.82
    ra = extra_radiation(lat, J); rso = (0.75 + 2e-5 * elev) * ra
    t = (tmax + tmin) / 2; es = (svp(tmax) + svp(tmin)) / 2; d = slope_svp(t)
    g = 0.000665 * pressure_from_elev(elev)
    f = pm(*(repr(x) for x in (tmax, tmin, ea, u2, rs, rso, t, es, d, g)))
    assert _eval_excel(f) == pytest.approx(eto_penman_monteith(tmax, tmin, rs, u2, ea, elev, lat, J), rel=1e-12)


# ── 격자 비교 ─────────────────────────────────────────────────────────
def test_grid_comparison_direct_difference():
    from cropwater_fcst import grid_comparison
    rng = np.random.default_rng(1)
    rows = []
    for rn, h, ks in (("아침", 2, (0, 1, 2, 3)), ("저녁", 17, (1, 2, 3, 4))):
        for d in range(6):
            run = TS("2026-05-01") + pd.Timedelta(days=d, hours=h)
            for k in ks:
                o = 4 + rng.normal()
                rows.append(dict(run=run, run_name=rn, lead_day=k, target=run.normalize() + pd.Timedelta(days=k),
                                 Tmax=25 + rng.normal(), Tmin=12.0, Tmax_obs=26.0, Tmin_obs=12.0, ea=1.2, ea_obs=1.2,
                                 u10=1.5, u10_obs=1.8, Rs_S3=18.0, Rs_obs=19.0, rain=0.0, rain_flag=0, flag_obs=0,
                                 ETo_S3=o - 0.2, ETo_obs=o, ETo_pers=o + 1.0, ETo_7d=o + 0.8))
    a = pd.DataFrame(rows)
    b = a.assign(Tmax=a.Tmax - 1.0, ETo_S3=a.ETo_S3 - 0.1)
    gc = grid_comparison(a, b)
    d = gc["diff"].set_index("item")["mean"]
    assert d["최고기온 (℃)"] == pytest.approx(1.0) and d["예보 ETo (mm/일)"] == pytest.approx(0.1)
    assert d["풍속 u10 (m/s)"] == pytest.approx(0.0) and gc["rain_agree"] == 1.0 and gc["n"] == len(a)
    m = gc["metrics"]
    ra = m[(m.grid == "A") & (m.run_name == "아침") & (m.lead_day == 1)].RMSE.iloc[0]
    rb = m[(m.grid == "B") & (m.run_name == "아침") & (m.lead_day == 1)].RMSE.iloc[0]
    assert ra == pytest.approx(0.2) and rb == pytest.approx(0.3)


# ── 결측·자료 공백 ────────────────────────────────────────────────────
def test_missing_values_dropped_and_counted(tmp_path):
    p = tmp_path / "x.csv"
    p.write_text(" format: day(UTC),hour(UTC),forecast,value,day(KST),hour(KST) location:73_134 Start : 20260701\n"
                 "1,1700,+6,21.0,2,0200\n1,1700,+7,-999.900024,2,0200\n1,1700,+8,901.0,2,0200\n", encoding="utf-8")
    loc, df = read_portal_csv(p)
    assert len(df) == 1 and df.attrs["n_missing"] == 2


def test_detect_rainy_month_precipitation():
    # 장마철: 비 오는 시각이 30%여도(0 비율 70%) 강수로 판별해야 함 (기온·풍속과 구분)
    v = [0.0] * 70 + [1.0, 2.0, 5.0, 12.0, 30.0, 3.0] * 5
    df = pd.DataFrame({"issue": [TS("2026-07-01 02:00") + pd.Timedelta(hours=3 * (i // 20)) for i in range(len(v))],
                       "forecast": [6 + i % 20 for i in range(len(v))], "value": v})
    assert detect_element("x.csv", df) == "PCP"


def test_service_table_skips_incomplete_runs():
    a, run = _archive()
    # 같은 날 17시 발표가 TMP에만 있고 나머지 요소에는 없음 → 저녁 발표는 제외되어야 함
    ev = TS("2026-04-02 17:00")
    a.hourly["TMP"] = pd.concat([a.hourly["TMP"], _hourly(str(ev), [7, 8], [15.0, 14.0])])
    st = service_table(a, runs={"아침": (2, range(0, 4)), "저녁": (17, range(1, 5))})
    assert set(st.run) == {run}
    skipped = {(n, t) for n, t, _ in a.skipped_runs}
    assert ("저녁", ev) in skipped
    miss = dict(((n, t), m) for n, t, m in a.skipped_runs)[("저녁", ev)]
    assert "TMX" in miss and "TMP" not in miss


def test_split_verifiable_and_strict_cum3():
    from cropwater_fcst import cum3, split_verifiable
    rows = []
    for k in range(4):   # 아침 발표 1회, D+2부터 관측 없음
        rows.append(dict(run_name="아침", run=TS("2026-09-27 02:00"), lead_day=k, Tmax=25.0, Tmin=15.0, ea=1.5,
                         u10=1.0, rain=0.0, ETo_S3=3.0, ETo_obs=(3.2 if k < 2 else np.nan), ETo_pers=3.1, ETo_7d=3.0))
    df, dropped = split_verifiable(pd.DataFrame(rows))
    assert len(df) == 2 and set(dropped.drop_reason) == {"대상일 관측 없음"}
    c, s = cum3(df)
    assert len(c) == 0          # D+0~D+2 중 D+2가 없으므로 3일 누적에서 제외


# ── 월별 지표·판정 불확실성·Rs 계수 재보정 ────────────────────────────
def _month_df(err_apr=0.5, err_may=-1.0):
    rows = []
    for rn, h, ks in (("아침", 2, (0, 1, 2, 3)), ("저녁", 17, (1, 2, 3, 4))):
        for day in pd.date_range("2026-04-01", "2026-05-31"):
            run = day + pd.Timedelta(hours=h)
            for k in ks:
                t = day + pd.Timedelta(days=k)
                o = 3.0 + (t.day % 3)
                e = err_apr if t.month == 4 else err_may
                rows.append(dict(run=run, run_date=day, run_name=rn, lead_day=k, target=t, ETo_obs=o, ETo_S3=o + e,
                                 ETo_pers=o + (1.0 if t.day % 2 else -1.0), ETo_7d=o + 0.8,
                                 Tmax=24.5, Tmax_obs=25.0, Tmin=12.3, Tmin_obs=12.0, ea=1.2, ea_obs=1.1, u10=1.5, u10_obs=2.0,
                                 Rs_S3=17.0, Rs_obs=19.0, rain_flag=int(t.day % 4 == 0), flag_obs=0))
    df = pd.DataFrame(rows)
    return df[df.target.dt.month.isin([4, 5])].reset_index(drop=True)


def test_month_metrics_and_verdict():
    from cropwater_fcst import month_metrics, month_verdict
    mm = month_metrics(_month_df())
    a = mm[(mm.run_name == "아침") & (mm.lead_day == 1) & (mm.month == 4)].iloc[0]
    assert a.RMSE == pytest.approx(0.5) and a.MBE == pytest.approx(0.5) and a.skill_pers == pytest.approx(0.5)
    assert a.dTmax == pytest.approx(-0.5) and a.ddT == pytest.approx(-0.8) and a.dRs == pytest.approx(-2.0)
    assert a.false_alarm == int((a.rain_fcst * a.n).round()) and a.miss == 0
    m = mm[(mm.run_name == "저녁") & (mm.lead_day == 2) & (mm.month == 5)].iloc[0]
    assert m.RMSE == pytest.approx(1.0) and m.skill_pers == pytest.approx(0.0)
    assert list(mm.run_name.unique()) == ["아침", "저녁"] and set(mm.lead_day) == {1, 2, 3}
    v = month_verdict(mm).set_index(["run_name", "month"])
    assert v.loc[("아침", 4), "pass_"] and not v.loc[("아침", 5), "pass_"]
    assert v.loc[("저녁", 5), "rmse_d1"] == pytest.approx(1.0)


def test_bootstrap_h2_bounds_and_reproducible():
    from cropwater_fcst import bootstrap_h2
    good = bootstrap_h2(_month_df(0.5, 0.5), n_boot=200)
    assert good["pass_all"] == 1.0 and good["n_days"] == 60      # 5/31 발표는 대상일이 6월이라 빠짐
    assert good["runs"]["아침"]["rmse_d1"] == pytest.approx((0.5, 0.5, 0.5))
    assert good["runs"]["저녁"]["min_skill"] == pytest.approx((0.5, 0.5, 0.5))
    bad = bootstrap_h2(_month_df(1.2, 1.2), n_boot=200)
    assert bad["runs"]["아침"]["pass_frac"] == 0.0 and bad["pass_all"] == 0.0
    mix = _month_df(0.5, -1.1)          # 4월 좋음, 5월 기준 초과 → 뽑힌 날에 따라 판정이 갈림
    a, b = bootstrap_h2(mix, n_boot=300), bootstrap_h2(mix, n_boot=300)
    assert a == b                          # 시드 고정 → 재현
    assert 0.0 < a["runs"]["아침"]["pass_frac"] < 1.0
    lo, mid, hi = a["runs"]["아침"]["rmse_d1"]
    assert lo < mid < hi


def test_fit_rs_forecast_recovers_coefficients():
    from cropwater_fcst import fit_rs_forecast
    rng = np.random.default_rng(3)
    n = 120
    tx = 20 + 8 * rng.random(n); tn = tx - (4 + 10 * rng.random(n)); flag = (rng.random(n) < 0.3).astype(int)
    ra = 30 + 10 * rng.random(n)
    t = pd.DataFrame(dict(Tmax=tx, Tmin=tn, rain_flag=flag, Ra=ra,
                          Rs_obs=(-0.05 + 0.17 * np.sqrt(tx - tn) - 0.12 * flag) * ra))
    a, b, c = fit_rs_forecast(t)
    assert (a, b, c) == pytest.approx((-0.05, 0.17, -0.12), abs=1e-9)


def test_wrap_text_display_width():
    from fcst_report import _disp_width, _wrap_text
    assert _disp_width("가a") == 3
    s = "④ 판정 불확실성: " + " ".join(["아침 0.88~1.08"] * 20)
    parts = _wrap_text(s, 60)
    assert all(_disp_width(p) <= 60 for p in parts) and len(parts) > 1
    assert " ".join(p.strip() for p in parts) == s
