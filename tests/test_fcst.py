# -*- coding: utf-8 -*-
"""02-Cycle 단기예보 모듈 단위 테스트 (pytest)
  fcst_archive: 포털 CSV·OpenAPI 응답 CSV 파싱, 발표·선행시간 → 대상시각, TMX/TMN 대상일, 지나간 시각 채움, 연장기간 코드 변환
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
                rows.append(dict(run_name=rn, lead_day=k, ETo_obs=v, ETo_S3=v + 0.5, ETo_main=v + 0.5,
                                 ETo_pers=v + (1.0 if i % 2 else -1.0),
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
                                 u10=1.5, u10_obs=1.8, Rs_S3=18.0, Rs_main=18.0, Rs_obs=19.0, rain=0.0, rain_flag=0,
                                 flag_obs=0, ETo_S3=o - 0.2, ETo_main=o - 0.2, ETo_obs=o, ETo_pers=o + 1.0, ETo_7d=o + 0.8))
    a = pd.DataFrame(rows)
    b = a.assign(Tmax=a.Tmax - 1.0, ETo_S3=a.ETo_S3 - 0.1, ETo_main=a.ETo_main - 0.1)
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
                rows.append(dict(run=run, run_date=day, run_name=rn, lead_day=k, target=t, ETo_obs=o, ETo_S3=o + e, ETo_main=o + e,
                                 ETo_pers=o + (1.0 if t.day % 2 else -1.0), ETo_7d=o + 0.8,
                                 Tmax=24.5, Tmax_obs=25.0, Tmin=12.3, Tmin_obs=12.0, ea=1.2, ea_obs=1.1, u10=1.5, u10_obs=2.0,
                                 Rs_S3=17.0, Rs_main=17.0, Rs_obs=19.0, rain_flag=int(t.day % 4 == 0), flag_obs=0))
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


# ── 하늘상태(SKY)·강수확률(POP): 여러 달 파일, 판별, 낮 시간 가중, S4 ──────────────
def test_read_multi_month_file_updates_month_at_start_lines(tmp_path):
    p = tmp_path / "sky.csv"
    p.write_text(" format: day(UTC),hour(UTC),forecast,value,day(KST),hour(KST)  location:73_134 Start : 20260430\n"
                 "30,2300,+6,1.0,31,0800\nStart : 20260501\n1,0200,+6,3.0,1,1100\n1,1700,+6,4.0,2,0200\n", encoding="utf-8")
    _, df = read_portal_csv(p)
    assert list(df["issue"]) == [TS("2026-05-01 08:00"), TS("2026-05-01 11:00"), TS("2026-05-02 02:00")]


def test_detect_sky_and_pop_and_filter_invalid_sky(tmp_path):
    iss = [TS("2026-06-01 02:00") + pd.Timedelta(hours=3 * (i // 60)) for i in range(600)]
    rng = np.random.default_rng(2)
    sky = pd.DataFrame({"issue": iss, "forecast": [6 + i % 60 for i in range(600)],
                        "value": rng.choice([1.0, 3.0, 4.0], 600)})
    assert detect_element("x.csv", sky) == "SKY"
    pop = sky.assign(value=rng.choice([0.0, 0.0, 10.0, 20.0, 30.0, 60.0, 80.0], 600))
    assert detect_element("x.csv", pop) == "POP"               # 10% 단위·0 포함 → 습도(5% 단위)가 아님
    reh = sky.assign(value=rng.choice([55.0, 60.0, 65.0, 70.0, 85.0, 95.0], 600))
    assert detect_element("x.csv", reh) == "REH"
    p = tmp_path / "하늘상태.csv"
    p.write_text(" format: day(UTC),hour(UTC),forecast,value,day(KST),hour(KST)  location:73_134 Start : 20260601\n"
                 "1,0200,+6,3.0,1,1100\n1,0200,+7,0.0,1,1100\n1,0200,+8,4.0,1,1100\n", encoding="utf-8")
    a = Archive(); e = a.add_file(str(p))
    assert e == "SKY" and len(a.hourly["SKY"]) == 2 and a.n_missing["SKY"] == 1   # 코드표에 없는 0은 결측


def test_sun_weights_daytime_shape():
    from fcst_archive import sun_weights
    t = pd.date_range("2026-06-21 00:00", periods=24, freq="h")
    w = sun_weights(t, 37.9, 127.74)
    assert w[0] == 0 and w[23] == 0 and w[12] > w[9] > w[7] > 0          # 태양시 정오 ≈ 12:30
    assert abs(w[12] - w[13]) < 0.02
    w0 = sun_weights(t)                                                  # 위도 없으면 06~18시 같은 가중치
    assert w0.sum() == 13 and w0[5] == 0 and w0[6] == 1


def test_daily_inputs_sky_fractions_and_optional_completeness():
    a, run = _archive()
    lead = a.hourly["TMP"][a.hourly["TMP"].issue == run].lead.tolist()
    t = [run + pd.Timedelta(hours=h) for h in lead]
    val = [4.0 if tt.normalize() == run.normalize() + pd.Timedelta(days=1) else 1.0 for tt in t]   # D+1 종일 흐림
    a.hourly["SKY"] = _hourly(str(run), lead, val)
    a.hourly["POP"] = _hourly(str(run), lead, [30.0] * len(lead))
    out = daily_inputs(a, run, [run.normalize() + pd.Timedelta(days=k) for k in range(4)], 37.9, 127.74).set_index("lead_day")
    assert out.loc[1, "sky_overcast"] == pytest.approx(1.0) and out.loc[1, "sky_cloudy"] == 0
    assert out.loc[2, "sky_overcast"] == 0 and out.loc[2, "pop"] == pytest.approx(0.30)
    # 선택 요소가 없는 발표는 필수 6요소만으로 완전 발표로 본다
    assert run in a.complete_issues


def test_rs_s4_fit_recovers_and_per_row_coefficients():
    from rs_model import fit_s4, rs_s4
    rng = np.random.default_rng(5)
    n = 200
    tx = 20 + 8 * rng.random(n); tn = tx - (4 + 10 * rng.random(n)); fl = (rng.random(n) < 0.3).astype(float)
    cl, ov = rng.random(n) * 0.5, rng.random(n) * 0.5
    ra = 30 + 10 * rng.random(n)
    co = np.array([0.25, 0.11, -0.05, -0.09, -0.23])
    rs = (co[0] + co[1] * np.sqrt(tx - tn) + co[2] * fl + co[3] * cl + co[4] * ov) * ra
    assert fit_s4(tx, tn, fl, cl, ov, ra, rs) == pytest.approx(co, abs=1e-9)
    rso = ra * 0.76
    one = rs_s4(tx, tn, fl, cl, ov, ra, rso, co)
    per_row = rs_s4(tx, tn, fl, cl, ov, ra, rso, np.tile(co, (n, 1)))
    assert np.allclose(one, per_row) and (one <= rso + 1e-12).all()


def test_s4_cv_excludes_own_month():
    from cropwater_fcst import s4_cv
    rng = np.random.default_rng(7)
    rows = []
    for day in pd.date_range("2026-04-01", "2026-06-30"):
        for k in (1, 2):
            tx = 22 + 6 * rng.random(); tn = tx - 8 - 4 * rng.random()
            cl, ov = 0.5 * rng.random(), 0.5 * rng.random()
            rows.append(dict(target=day, lead_day=k, Tmax=tx, Tmin=tn, rain_flag=float(rng.random() < 0.2),
                             pop_max=float(rng.choice([0.0, 0.3, 0.6])), sky_cloudy=cl, sky_overcast=ov, Ra=38.0, Rso=29.0,
                             Rs_obs=(0.25 + 0.11 * np.sqrt(tx - tn) - 0.1 * ov) * 38.0 + rng.normal()))
    df = pd.DataFrame(rows)
    co, tab = s4_cv(df)
    assert set(tab.fold) == {"2026-04", "2026-05", "2026-06"} and set(tab.lead_day) == {1, 2}
    df2 = df.copy()
    df2.loc[df2.target.dt.month == 5, "Rs_obs"] += 5.0          # 5월 관측을 바꿔도
    co2, _ = s4_cv(df2)
    may = (df.target.dt.month == 5).values
    assert np.allclose(co[may], co2[may])                       # 5월 행의 계수는 그대로(5월 자료를 쓰지 않음)
    assert not np.allclose(co[~may], co2[~may])                 # 다른 달 계수는 5월 자료를 씀


def test_daily_layout_backward_compatible_letters():
    from fcst_report import daily_layout
    from openpyxl.utils import get_column_letter as CL
    col = {k: CL(i) for i, (k, _, _) in enumerate(daily_layout(False), 1)}
    assert (col["flag"], col["rs3"], col["e3"], col["eo"], col["x3"], col["kc"], col["etc"], col["drs"], col["judge"]) == \
           ("O", "R", "Y", "AA", "AD", "AH", "AI", "BA", "BB")          # 하늘상태가 없으면 이전 배치와 같음
    keys = [k for k, _, _ in daily_layout(True)]
    assert {"cld", "ovc", "fold", "krow", "rs4", "e4", "x4", "drs3"} <= set(keys)


def test_sky_coef_file_roundtrip(tmp_path):
    from rs_model import load_sky_coef, save_sky_coef
    p = tmp_path / "rs_sky_coef.csv"
    tab = pd.DataFrame([dict(lead_day=k, a=0.2 + k / 100, b=0.1, c=-0.05 * k, d=-0.1, e=-0.2, n=100,
                             fit_start="2026-04-02", fit_end="2026-09-19", rmse_rs=4.5) for k in range(5)])
    save_sky_coef("101", tab, p, note="시험")
    save_sky_coef("216", tab.assign(a=0.3), p)
    c = load_sky_coef("101", p)
    assert sorted(c) == [0, 1, 2, 3, 4] and c[2] == pytest.approx((0.22, 0.1, -0.1, -0.1, -0.2))
    assert load_sky_coef("216", p)[0][0] == pytest.approx(0.3) and load_sky_coef("999", p) == {}
    # 이전 형식(강수 입력 열 없음 = 강수유무 계수)은 강수확률 기준 S4에 쓰지 않음
    old = tmp_path / "old.csv"
    old.write_text("stn,lead_day,a,b,c,d,e\n101,1,0.25,0.11,-0.04,-0.09,-0.23\n", encoding="utf-8")
    with pytest.raises(ValueError):
        load_sky_coef("101", old)
    from cropwater_fcst import load_s4_fixed
    with pytest.raises(ValueError):
        load_s4_fixed("101", str(old))
    assert list(load_s4_fixed("101", str(p)).lead_day) == [0, 1, 2, 3, 4]


# ── OpenAPI 응답 CSV (단기예보 조회서비스, 2025년 자료) ─────────────────────
API_HDR = "baseDate,baseTime,category,fcstDate,fcstTime,fcstValue,nx,ny\n"


def _api(issue, rows):
    """rows: [(요소, 예보시각, 값)] → 응답 CSV 줄들 (시각은 앞의 0을 뺀 API 표기: 200, 0 …)"""
    i = TS(issue)
    return "".join(f"{i:%Y%m%d},{i.hour * 100},{c},{TS(t):%Y%m%d},{TS(t).hour * 100},{v},73,134\n" for c, t, v in rows)


def _api_file(tmp_path):
    prev = _api("2025-03-31 23:00", [(c, f"2025-04-01 {h:02d}:00", v) for h in range(6)
                                     for c, v in (("TMP", 10), ("PCP", "강수없음"))])
    run = _api("2025-04-01 02:00", [("TMP", "2025-04-01 03:00", 20), ("PCP", "2025-04-01 03:00", "1mm 미만"),
                                    ("TMP", "2025-04-01 04:00", 20), ("PCP", "2025-04-01 04:00", "2.0mm"),
                                    ("TMP", "2025-04-01 05:00", 20), ("PCP", "2025-04-01 05:00", "강수없음"),
                                    ("TMN", "2025-04-01 06:00", 5), ("TMX", "2025-04-01 15:00", 18),
                                    ("WAV", "2025-04-01 03:00", -999)])
    again = _api("2025-04-01 02:00", [("TMP", "2025-04-01 05:00", 21)])       # 같은 칸이 다시 오면 뒤의 값
    p = tmp_path / "단기예보_20250401.csv"
    p.write_text("\ufeff" + API_HDR + prev + ",,,,,,,\n,,,,,,,\n" + API_HDR + run + again, encoding="utf-8")
    return p


def test_read_openapi_csv_strings_blank_rows_duplicates(tmp_path):
    from fcst_archive import PCP_LT1_MM, is_openapi_csv, pcp_mm, read_openapi_csv
    p = _api_file(tmp_path)
    assert is_openapi_csv(p)
    q = tmp_path / "portal.csv"
    q.write_text(" format: day(UTC),hour(UTC),forecast,value,day(KST),hour(KST) location:73_134 Start : 20260401\n", encoding="utf-8")
    assert not is_openapi_csv(q)
    locs, t = read_openapi_csv(p)
    assert locs == {"73_134"} and set(t) == {"TMP", "PCP", "TMX", "TMN"}      # 쓰지 않는 요소(WAV)는 읽지 않음
    tmp = t["TMP"][t["TMP"].issue == TS("2025-04-01 02:00")]
    assert list(tmp.forecast) == [1, 2, 3] and list(tmp.value) == [20, 20, 21]   # lead = 예보시각 − 발표시각
    pcp = t["PCP"][t["PCP"].issue == TS("2025-04-01 02:00")]
    assert list(pcp.value) == [PCP_LT1_MM, 2.0, 0.0]
    assert list(t["TMX"].target_date) == [TS("2025-04-01")] and list(t["TMN"].value) == [5]
    assert (pcp_mm("30.0~50.0mm"), pcp_mm("50.0mm 이상"), pcp_mm("0.3"), pcp_mm("1.0mm 미만")) == (40.0, 50.0, 0.3, PCP_LT1_MM)
    assert math.isnan(pcp_mm("?"))


def test_openapi_archive_fills_only_hours_before_first_lead(tmp_path):
    from fcst_archive import PCP_LT1_MM, check_archive
    a = Archive()
    assert set(a.add_file(str(_api_file(tmp_path)))) == {"TMP", "PCP", "TMX", "TMN"}
    run = TS("2025-04-01 02:00")
    d = daily_inputs(a, run, [TS("2025-04-01")]).iloc[0]
    assert d.hours == 6 and d.filled == 3            # 00~02시만 전날 23시 발표, 03시부터는 02시 발표 자체의 값
    assert (d.Tmax, d.Tmin) == (18, 5) and d.rain == pytest.approx(PCP_LT1_MM + 2.0) and d.rain_flag == 1
    rep = check_archive(a)
    assert rep["formats"] == ["openapi"] and rep["location"] == ["73_134"]


def test_check_archive_reports_partial_issue(tmp_path):
    from fcst_archive import check_archive
    txt = API_HDR
    for day, n in (("2025-04-01", 5), ("2025-04-02", 5), ("2025-04-03", 3)):      # 4/3 발표는 앞부분이 잘림
        txt += _api(f"{day} 02:00", [("TMP", TS(f"{day} 08:00") - pd.Timedelta(hours=h), 15) for h in range(n)][::-1])
    p = tmp_path / "api.csv"
    p.write_text(txt, encoding="utf-8")
    a = Archive(); a.add_file(str(p))
    assert check_archive(a)["short_issues"] == [("2025-04-03 02:00:00", "TMP", 3, 5)]


# ── 요소별 KST CSV (발표일,발표시각,예보일,예보시각,값) · 다른 해 고정 계수 ──────────────
KST_HDR = "발표일,발표시각,예보일,예보시각,값\n"


def _kst(rows):
    """rows: [(발표시각, 예보시각, 값)] → 요소별 KST CSV 줄들"""
    return "".join(f"{TS(i):%Y%m%d},{TS(i):%H%M},{TS(t):%Y%m%d},{TS(t):%H%M},{v}\n" for i, t, v in rows)


def _kst_hourly(values, hours=range(3, 24), days=("2025-06-01", "2025-06-02")):
    rows = []
    for d in days:
        for k, h in enumerate(hours):
            rows.append((f"{d} 02:00", f"{d} {h:02d}:00", values[k % len(values)]))
    return rows


def test_element_csv_format_detection_and_values(tmp_path):
    from fcst_archive import csv_format, read_element_csv
    pop = [0, 0, 30, 60, 20, 0, 30, 0, 20, 60, 0, 30, 20, 0, 60, 30, 0, 20, 30, 66, 0]      # 66 같은 값은 드묾
    cases = {"SKY": [1, 3, 4, 4, 1], "POP": pop, "REH": [55, 60, 75, 85, 90, 95],
             "TMP": [12, 15, 18, 21, 23, -1], "WSD": [0.8, 1.3, 2.6, 3.1, 1.7],
             "PCP": ["강수없음", "강수없음", "1mm 미만", "2.0mm", "강수없음"], "UUU/VVV": [-0.2, 0.5, -1.3, 2.1, -0.7]}
    for elem, vals in cases.items():
        p = tmp_path / f"x_{len(elem)}.csv"
        p.write_text("﻿" + KST_HDR + _kst(_kst_hourly(vals)), encoding="utf-8")
        assert csv_format(p) == "element"
        e, df = read_element_csv(p)
        assert e == elem, (elem, e)
        if elem == "PCP":
            assert sorted(set(df.value)) == [0.0, 0.5, 2.0]
    # 하루 1칸 요소: 06시 → TMN, 15시 → TMX (대상일 = 예보일)
    p = tmp_path / "d.csv"
    p.write_text(KST_HDR + _kst([("2025-06-01 02:00", "2025-06-01 06:00", 14.0), ("2025-06-01 02:00", "2025-06-02 06:00", 15.0)]),
                 encoding="utf-8")
    e, df = read_element_csv(p)
    assert e == "TMN" and list(df.target_date) == [TS("2025-06-01"), TS("2025-06-02")]
    # 한글 엑셀 저장본(CP949)도 읽음
    q = tmp_path / "cp949.csv"
    q.write_bytes((KST_HDR + _kst(_kst_hourly(cases["PCP"]))).encode("cp949"))
    assert csv_format(q) == "element" and read_element_csv(q)[0] == "PCP"


def test_archive_element_files_skip_unused_and_report_missing(tmp_path):
    from fcst_archive import check_archive
    d = tmp_path / "f"; d.mkdir()
    (d / "a.csv").write_text(KST_HDR + _kst(_kst_hourly([12, 15, 18, 21, 23])), encoding="utf-8")        # TMP
    (d / "b.csv").write_text(KST_HDR + _kst(_kst_hourly([-0.2, 0.5, -1.3, 2.1])), encoding="utf-8")      # UUU/VVV
    from fcst_archive import load_archive
    a = load_archive([str(d)])
    rep = check_archive(a)
    assert "TMP" in a.hourly and rep["formats"] == ["element"] and rep["ignored_files"] == [("b.csv", "UUU/VVV")]
    assert "WSD" in rep["missing_elements"] and "TMP" not in rep["missing_elements"]


def test_forecast_table_fixed_s4_coefficients():
    from cropwater_fcst import S4_FIXED, bias_correction_cv, forecast_table, s4_is_fixed
    from rs_model import rs_s4
    rng = np.random.default_rng(3)
    rows = []
    for day in pd.date_range("2025-04-02", "2025-05-31"):
        for k in (1, 2):
            tx = 20 + 6 * rng.random()
            rows.append(dict(run_name="아침", run=day - pd.Timedelta(days=k) + pd.Timedelta(hours=2), target=day, lead_day=k,
                             Tmax=tx, Tmin=tx - 9, ea=1.2, u10=1.5, rain=0.0, rain_flag=0, sky_cloudy=0.3 * rng.random(),
                             sky_overcast=0.3 * rng.random(), pop=0.1, pop_max=0.2))
    st = pd.DataFrame(rows)
    dates = pd.date_range("2025-03-20", "2025-06-05")
    obs = pd.DataFrame(dict(date=dates, ETo_obs=3 + rng.random(len(dates)), Tmax=24.0, Tmin=13.0, ea_obs=1.1, u10=1.4,
                            Rs=18.0, rain=0.0, pa=np.nan))
    fixed = pd.DataFrame(dict(lead_day=[1, 2], a=[0.25, 0.30], b=[0.11, 0.10], c=[-0.04, -0.12], d=[-0.09, -0.13],
                              e=[-0.23, -0.17], n=[320, 320], fit_start="2026-04-02", fit_end="2026-09-18", note=""))
    kp = kc_params({"시나리오 번호 (1~4)": 3, "생육 시작일 (발아기/정식일)": pd.Timestamp("2025-04-01"),
                    "L_ini (초기, 일)": 20, "L_dev (발육, 일)": 70, "L_mid (중기, 일)": 90, "L_late (후기, 일)": 30,
                    "생육중기 초목 수고 h (m)": 3.2, "중기 평균 u2 (m/s)": 1.0, "중기 평균 RHmin (%)": 55,
                    "후기 평균 u2 (m/s)": 1.5, "후기 평균 RHmin (%)": 55, "멀칭 Kc 보정계수": 1})
    df = forecast_table(st, obs, 37.9, 76.5, dict(a=None), kp, fixed)
    assert s4_is_fixed(df) and set(df.s4_fold) == {S4_FIXED}
    k1 = df[df.lead_day == 1]
    exp = rs_s4(k1.Tmax, k1.Tmin, k1.pop_max, k1.sky_cloudy, k1.sky_overcast, k1.Ra, k1.Rso,        # S4 강수 입력 = 강수확률 하루 최대
                fixed.iloc[0][["a", "b", "c", "d", "e"]].to_numpy(float))
    assert np.allclose(k1.Rs_S4, exp)
    assert list(df.attrs["s4_table"].fold.unique()) == [S4_FIXED]
    # 보정 탐색: 고정 계수면 중첩 교차검증이 필요 없어 대상월 2개로도 ETo 비율 보정을 구함
    m, _, folds, rows = bias_correction_cv(df, 37.9, 76.5, dict(a=None))
    assert folds == ["2025-04", "2025-05"] and rows["B3"].notna().all()
