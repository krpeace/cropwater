# -*- coding: utf-8 -*-
"""02-Cycle 4단계(G4) 예보 물수지 단위 테스트 (pytest)
  fcst_archive : 기대 강수(시각별 강수량 × 강수확률)
  fcst_wb      : 관측 물수지(관수 기록·관수 규칙·결측 채움), 발표별 예보 물수지(아침·저녁 출발, 세 경로, 참값),
                 관수 필요 예상일·판정·범위, 오차표(합치기·다른 해·월 칸 대체), 편향 보정 자리, 서비스 전망
  fcst_wb_report: 한 칸 물수지 수식 = fao56_core.wb_step
"""
import datetime as dt
import os
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from fao56_core import wb_step
from fcst_archive import daily_inputs
import fcst_wb as W

TS = pd.Timestamp
SOIL = dict(fc=0.22, wp=0.10, zr=1.0, p=0.5, taw=120.0, raw=60.0, dr0=0.0, ea=0.95)
KP1 = dict(bud=dt.date(2026, 1, 1), L=(1, 1, 500, 1), kc_ini=1.0, kc_mid=1.0, kc_end=1.0)   # Kc = 1 고정


def _obs(days, eto, rain):
    return pd.DataFrame({"date": pd.date_range(days[0], periods=len(eto)), "ETo_obs": eto, "rain": rain})


# ── 기대 강수 ─────────────────────────────────────────────────────────
def test_rain_exp_is_pcp_times_pop():
    from test_fcst import _archive, _hourly
    a, run = _archive()
    lead = a.hourly["PCP"][a.hourly["PCP"].issue == run].lead.tolist()
    a.hourly["POP"] = _hourly(str(run), lead, [60.0] * len(lead))
    out = daily_inputs(a, run, [run.normalize() + pd.Timedelta(days=k) for k in range(4)], 37.9, 127.74).set_index("lead_day")
    assert out.loc[0, "rain"] == pytest.approx(1.5) and out.loc[0, "rain_exp"] == pytest.approx(0.9)    # 1.5 mm × 60%
    assert out.loc[3, "rain"] == pytest.approx(4.5) and out.loc[3, "rain_exp"] == pytest.approx(2.7)    # 코드 1 → 4.5 mm × 60%
    assert out.loc[1, "rain_exp"] == 0.0


# ── 관측 물수지 ───────────────────────────────────────────────────────
def test_observed_wb_matches_wb_step_with_irrigation_log():
    eto, rain = [5.0] * 10, [0, 0, 0, 0, 30, 0, 0, 0, 0, 0]
    obs = _obs(["2026-05-01"], eto, rain)
    irr = {dt.date(2026, 5, 8): 20 / 0.95}                          # 공급 21.05 mm → 순 20 mm
    w = W.observed_wb(obs, KP1, SOIL, irrig=irr)
    dr = 0.0
    for r, p, d in zip(w.itertuples(), rain, obs.date):
        inet = 20.0 if d == TS("2026-05-08") else 0.0
        ks, ea_, dp, dr = wb_step(dr, p, 5.0, 120.0, 60.0, inet)
        assert r.Dr == pytest.approx(dr) and r.Ks == pytest.approx(ks)
    assert w.loc[w.date == TS("2026-05-08"), "I_net"].iloc[0] == pytest.approx(20.0)
    assert not w.filled.any() and not w.tainted.any()


def test_observed_wb_auto_irrigation_rule():
    soil = dict(SOIL, taw=40.0, raw=20.0)
    w = W.observed_wb(_obs(["2026-05-01"], [5.0] * 6, [0.0] * 6), KP1, soil, auto_irrigate=True)
    assert list(w.Dr.round(6)[:4]) == [5.0, 10.0, 15.0, 20.0]       # 4일째 끝 Dr = RAW
    assert w.I_net.iloc[4] == pytest.approx(20.0)                   # 5일째 Dr(20)만큼 관수
    assert w.Dr.iloc[4] == pytest.approx(5.0) and w.I_gross.iloc[4] == pytest.approx(20 / 0.95)


def test_observed_wb_fill_and_taint():
    obs = _obs(["2026-05-01"], [5.0, 5.0, np.nan, 5.0, 5.0, 5.0], [0, 0, 0, 0, 50, 0])
    obs = obs[obs.date != TS("2026-05-04")]                        # 5/4 행 자체가 없음
    fill = {TS("2026-05-03"): (4.0, 0.0), TS("2026-05-04"): (3.0, 2.0)}
    w = W.observed_wb(obs, KP1, SOIL, fill=fill).set_index("date")
    assert w.loc["2026-05-03", "ETo"] == 4.0 and w.loc["2026-05-03", "ETo_src"] == "예보"
    assert w.loc["2026-05-04", "P"] == 2.0 and w.loc["2026-05-04", "filled"]      # 행이 없는 날은 강수도 예보로
    assert w.loc["2026-05-02", "tainted"] == False and w.loc["2026-05-04", "tainted"]
    assert w.loc["2026-05-05", "Dr"] == 0.0 and not w.loc["2026-05-05", "tainted"]   # 큰 비로 Dr 0 → 영향 끝


# ── 발표별 예보 물수지 ────────────────────────────────────────────────
def _ft(D, etc=4.0):
    rows = []
    for rn, h, leads in (("아침", 2, range(0, 4)), ("저녁", 17, range(1, 5))):
        run = TS(D) + pd.Timedelta(hours=h)
        for k in leads:
            rain = 10.0 if k == 1 else 0.0
            rows.append(dict(run_name=rn, run=run, lead_day=k, target=TS(D) + pd.Timedelta(days=k), Kc=1.0,
                             ETc_main=etc, ETo_main=etc, rain=rain, rain_exp=0.6 * rain, pop_max=0.6 if k == 1 else 0.1,
                             ETc_pers=5.0, ETo_obs=5.0))
    return pd.DataFrame(rows)


def test_forecast_runs_morning_and_evening_paths():
    D = "2026-06-10"
    obs = _obs(["2026-06-01"], [5.0] * 15, [0.0] * 15)
    owb = W.observed_wb(obs, KP1, SOIL)
    ft = _ft(D)
    runs = W.forecast_runs(ft, owb, SOIL)
    dr0 = float(owb.set_index("date").loc[TS(D) - pd.Timedelta(days=1), "Dr"])     # 9일 × 5 = 45
    assert dr0 == pytest.approx(45.0)
    m = runs[runs.run_name == "아침"].sort_values("order")
    exp = W.path(dr0, [4.0] * 4, [0.0, 6.0, 0.0, 0.0], SOIL)                         # 중심: 기대 강수
    assert list(m.Dr_center) == pytest.approx(exp)
    assert list(m.Dr_true) == pytest.approx(W.path(dr0, [5.0] * 4, [0.0] * 4, SOIL))
    assert list(m.Dr_none) == pytest.approx(W.path(dr0, [4.0] * 4, [0.0] * 4, SOIL))
    rel = W.ERR_DEFAULT_REL
    assert list(m.Dr_early) == pytest.approx(W.path(dr0, [4.0 * (1 + rel)] * 4, [0.0] * 4, SOIL))
    assert list(m.Dr_late) == pytest.approx(W.path(dr0, [4.0 * (1 - rel)] * 4, [0.0, 10.0, 0.0, 0.0], SOIL))
    e = runs[runs.run_name == "저녁"].sort_values("order")
    start = wb_step(dr0, 0.0, 4.0, 120.0, 60.0)[3]                                  # 그날 D는 아침 D+0 예보로
    assert e.dr_start.iloc[0] == pytest.approx(start)
    assert list(e.Dr_center) == pytest.approx(W.path(dr0, [4.0] * 5, [0.0, 6.0, 0.0, 0.0, 0.0], SOIL)[1:])
    assert e.dr_start_true.iloc[0] == pytest.approx(wb_step(dr0, 0.0, 5.0, 120.0, 60.0)[3])
    assert runs.ok.all() and list(m.main) == [True, True, True, False]


def test_forecast_runs_excludes_rows_after_filled_target():
    D = "2026-06-10"
    eto = [5.0] * 15
    eto[10] = np.nan                                                                 # 6/11 관측 없음
    obs = _obs(["2026-06-01"], eto, [0.0] * 15)
    ft = _ft(D)
    owb = W.observed_wb(obs, KP1, SOIL, fill=W.fill_from_forecast(ft))
    runs = W.forecast_runs(ft, owb, SOIL)
    m = runs[runs.run_name == "아침"].sort_values("order")
    assert list(m.ok) == [True, False, False, False]
    assert set(m.reason[~m.ok]) == {"예보 기간에 관측 결측일(예보로 채움)"}


def test_first_need_and_categories():
    assert W.first_need([10, 70, 80], 60) == 2 and W.first_need([10, 20], 60) is None
    assert W.first_need([70, 80], 60, orders=[2, 3]) == 2
    assert W.need_category(2, 2) == "같은 날" and W.need_category(1, 2) == "하루 빠름"
    assert W.need_category(3, 1) == "이틀 이상 늦음" and W.need_category(None, 2) == "실제만 필요(놓침)"
    assert W.need_category(2, None) == "예보만 필요(헛경보)" and W.need_category(None, None) == "둘 다 기간 안 필요 없음"


def test_first_need_eval_range_coverage():
    base = dict(run_name="아침", ok=True, dr_start=50.0, dr_start_true=50.0, lead_day=0)
    rows = []
    # 발표 1: 중심 없음, 빠르면 2, 참값 2 → 놓침이지만 범위 안
    for k, (c, e, l, t) in enumerate([(52, 55, 50, 55), (55, 62, 51, 61), (58, 70, 52, 66)], 1):
        rows.append(dict(base, run=TS("2026-06-01 02:00"), order=k, Dr_center=c, Dr_early=e, Dr_late=l, Dr_true=t))
    # 발표 2: 참값이 빠르면보다 앞
    for k, (c, e, l, t) in enumerate([(52, 55, 50, 61), (55, 62, 51, 66), (58, 70, 52, 70)], 1):
        rows.append(dict(base, run=TS("2026-06-02 02:00"), order=k, Dr_center=c, Dr_early=e, Dr_late=l, Dr_true=t))
    fe = W.first_need_eval(pd.DataFrame(rows), 60.0)
    a, b = fe.iloc[0], fe.iloc[1]
    assert (a.fcst, a.early, a.late, a.true, a.category, a.covered) == (None, 2, None, 2, "실제만 필요(놓침)", True)
    assert b.true == 1 and b.early == 2 and b.true_before_early and not b.covered
    s = W.first_need_summary(fe).iloc[0]
    assert s.n_range_event == 2 and s.coverage == pytest.approx(0.5) and s.true_before_early == 1


# ── 오차표 ────────────────────────────────────────────────────────────
def _ft_err(year, bias):
    rows = []
    for d in pd.date_range(f"{year}-05-01", f"{year}-06-30"):
        for rn, h, leads in (("아침", 2, range(0, 4)), ("저녁", 17, range(1, 5))):
            for k in leads:
                obs = 4.0 + (d.day % 3)
                rows.append(dict(run_name=rn, run=d + pd.Timedelta(hours=h), lead_day=k, target=d + pd.Timedelta(days=k),
                                 ETo_main=obs + bias + (0.5 if k % 2 else -0.5), ETo_obs=obs))
    return pd.DataFrame(rows)


def test_error_table_pool_exclude_and_lookup(tmp_path):
    rows = pd.concat([W.eto_error_rows(_ft_err(2025, 0.2), "101", 2025), W.eto_error_rows(_ft_err(2026, -0.2), "101", 2026)])
    p = tmp_path / "err.csv"
    W.save_error_table(rows, p)
    W.save_error_table(W.eto_error_rows(_ft_err(2026, -0.2), "101", 2026), p)       # 같은 지점·해는 바꿔 넣음
    back = W.load_error_rows(str(p))
    assert len(back) == len(rows)
    pool = W.pooled_errors(back, "101")
    one = pool[(pool.kind == "day") & (pool.run_name == "아침") & (pool.lead_day == 1) & (pool.month == 0)].iloc[0]
    assert one.mbe == pytest.approx(0.5, abs=1e-9) and one.years == "2025,2026"     # +0.2+0.5, −0.2+0.5 → 평균 0.5
    other = W.pooled_errors(back, "101", exclude_year=2026)
    o1 = other[(other.kind == "day") & (other.run_name == "아침") & (other.lead_day == 1) & (other.month == 0)].iloc[0]
    assert o1.mbe == pytest.approx(0.7) and o1.years == "2025"
    rmse, rel, n, src = W.err_lookup(pool, "day", "아침", 1, 5)
    assert src == "월별" and n >= 15 and rel == pytest.approx(rmse / pool[(pool.kind == "day") & (pool.run_name == "아침")
                                                                           & (pool.lead_day == 1) & (pool.month == 5)].obs_mean.iloc[0])
    assert W.err_lookup(pool, "day", "아침", 1, 9)[3] == "전체 월"                   # 9월 칸 없음 → 전체 월
    assert W.err_lookup(None, "day", "아침", 1, 5)[1] == W.ERR_DEFAULT_REL
    assert W.pooled_errors(back, "999").stn_src.iloc[0].startswith("전 지점")        # 지점 행이 없으면 다른 지점 오차
    c3 = pool[(pool.kind == "cum3") & (pool.run_name == "저녁")]
    assert set(c3.lead_day) == {1}


def test_apply_bias_hook():
    ft = _ft("2026-06-10")
    assert W.apply_bias(ft) is ft                                                   # #10: 기본은 보정 없음
    t = pd.DataFrame(dict(run_name=["아침"], lead_day=[0], factor=[1.1]))
    b = W.apply_bias(ft, t)
    m = (b.run_name == "아침") & (b.lead_day == 0)
    assert np.allclose(b.loc[m, "ETc_main"], 4.4) and np.allclose(b.loc[~m, "ETc_main"], 4.0)


# ── 서비스 전망 ───────────────────────────────────────────────────────
def test_service_outlook_need_and_advice():
    D = "2026-06-10"
    obs = _obs(["2026-05-25"], [5.0] * 20, [0.0] * 20)                              # 6/9 끝 Dr ≥ RAW (Ks < 1로 80보다 조금 작음)
    owb = W.observed_wb(obs, KP1, SOIL)
    ol = W.service_outlook(_ft(D), owb, SOIL, None, "아침", TS(D) + pd.Timedelta(hours=2))
    start = float(owb.set_index("date").loc[TS("2026-06-09"), "Dr"])
    assert start > 60 and ol["start"] == pytest.approx(start) and ol["need"]["center"] == 0
    assert ol["advice"][0] == pytest.approx(start) and ol["advice"][1] == pytest.approx(start / 0.95)
    obs2 = _obs(["2026-05-31"], [5.0] * 20, [0.0] * 20)                             # 6/9 끝 Dr 50 → 54, 52(비 6), 56, 60
    ol2 = W.service_outlook(_ft(D), W.observed_wb(obs2, KP1, SOIL), SOIL, None, "저녁", TS(D) + pd.Timedelta(hours=17))
    d = ol2["days"]
    k = ol2["need"]["center"]
    assert k == 3 and d.Dr_center.iloc[2] == pytest.approx(60.0) and ol2["need"]["early"] < 3
    assert ol2["cum3"][0] == pytest.approx(12.0) and len(ol2["pre"]) == 1


# ── 엑셀 한 칸 물수지 수식 = wb_step ─────────────────────────────────────
def _eval(expr, cells):
    import re
    py = expr.lstrip("=")
    for k, v in cells.items():
        py = re.sub(rf"\b{k}\b", repr(v), py)
    return eval(py.replace("^", "**"), {"MIN": min, "MAX": max, "IF": lambda c, a, b: a if c else b})


@pytest.mark.parametrize("prev,p,etc", [(0.0, 0.0, 5.0), (55.0, 3.0, 4.0), (70.0, 30.0, 6.0), (90.0, 0.0, 5.0), (10.0, 40.0, 3.0)])
def test_excel_wb_formula_matches_wb_step(prev, p, etc):
    from fcst_wb_report import _wb_formula
    f = _wb_formula("PREV", "RAIN", "ETC", {"RAW": "60", "TAW": "120"})
    assert _eval(f, {"PREV": prev, "RAIN": p, "ETC": etc}) == pytest.approx(wb_step(prev, p, etc, 120.0, 60.0)[3])
