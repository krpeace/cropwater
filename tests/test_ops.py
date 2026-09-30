# -*- coding: utf-8 -*-
"""02-Cycle 5단계(G5) 운영 수집기 단위 테스트 (pytest)
  kma_grid     : 위경도 → 격자 (대표 10개 지점 = 활용가이드 변환값)
  kma_fcst     : 응답 해석(JSON·XML 인증 오류), 응답 코드 분류, 페이지 나눔, 완결성 점검, 저장 → fcst_archive로 다시 읽기
  fcst_archive : 운영 서비스 표 — 서비스 발표가 없으면 직전 발표(대체 발표 기준 선행일·오차 종류)
  cropwater_fcst.forecast_table : S3 대체(하늘상태 없음), 대체 발표 선행일로 S4 계수 선택
  cropwater_ops : 슬롯 시각, 재시도·마감·백업·인증 오류, H3 집계
  fcst_wb      : 대체 발표의 오차 칸, 아침 관측 지연(전날 관측을 예보로 채운 출발)
"""
import datetime as dt
import json
import os
import sys

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import kma_fcst as K
from fcst_archive import load_archive, ops_service_table, daily_inputs, service_table, error_kind

TS = pd.Timestamp


# ── 격자 ─────────────────────────────────────────────────────────────
def test_grid_reference_stations():
    from kma_grid import latlon_to_grid, grid_to_latlon
    from fao56_core import load_station_backup
    from cropwater_fcst import STATIONS
    info, _ = load_station_backup(os.path.join(os.path.dirname(__file__), "..", "stations_backup.csv"))
    for stn, (_, grid) in STATIONS.items():
        lat, lon = info[int(stn)][:2]
        assert "%d_%d" % latlon_to_grid(lat, lon) == grid
    lat, lon = grid_to_latlon(73, 134)
    assert latlon_to_grid(lat, lon) == (73, 134)


# ── 응답 해석 ─────────────────────────────────────────────────────────
def _resp(code, items=None, total=None):
    body = {"response": {"header": {"resultCode": code, "resultMsg": "M"}}}
    if items is not None:
        body["response"]["body"] = {"items": {"item": items}, "totalCount": len(items) if total is None else total}
    return json.dumps(body).encode()


def test_parse_and_classify():
    p = K.parse_response(_resp("00", [{"a": 1}]))
    assert p["code"] == "00" and p["items"] == [{"a": 1}] and p["total"] == 1
    x = K.parse_response(b"<OpenAPI_ServiceResponse><cmmMsgHeader><returnAuthMsg>SERVICE_KEY_IS_NOT_REGISTERED_ERROR"
                         b"</returnAuthMsg><returnReasonCode>30</returnReasonCode></cmmMsgHeader></OpenAPI_ServiceResponse>")
    assert x["code"] == "30" and "SERVICE_KEY" in x["msg"]
    assert K.classify(200, "00") == "ok" and K.classify(200, "03") == "nodata"
    assert K.classify(200, "30") == "fatal" and K.classify(401, "") == "fatal"
    assert K.classify(500, "") == "retry" and K.classify(200, "99") == "retry" and K.classify(200, "22") == "retry"


def test_service_key_encoding():
    assert K.service_key("ab+c/d==") == "ab%2Bc%2Fd%3D%3D"       # 디코딩 키 → 인코딩
    assert K.service_key("ab%2Bc") == "ab%2Bc"                   # 인코딩 키는 그대로


def test_fetch_pages_follows_total():
    items = [{"i": i} for i in range(2500)]
    calls = []

    def http(url, timeout):
        page = int(dict(x.split("=") for x in url.split("?")[1].split("&"))["pageNo"])
        calls.append(page)
        return 200, _resp("00", items[(page - 1) * 1000: page * 1000], total=2500)
    r = K.fetch_pages(lambda p: f"http://x/?pageNo={p}&numOfRows=1000", http)
    assert r["status"] == "ok" and len(r["items"]) == 2500 and calls == [1, 2, 3] and len(r["bodies"]) == 3

    def boom(url, timeout):
        raise TimeoutError("t")
    assert K.fetch_pages(lambda p: "http://x/", boom)["status"] == "retry"


# ── 완결성·저장 ──────────────────────────────────────────────────────
def _issue_items(issue, n_hourly=None, sky="1", nx=73, ny=134, drop=()):
    """발표 1개의 API 항목(8요소, 정상 행 수). 1시간 요소는 lead 1부터 n개, TMX·TMN은 발표시각별 대상일 수"""
    issue = TS(issue)
    h = issue.hour
    n = n_hourly or K.HOURLY_N[h]
    vals = dict(TMP="20", REH="60", WSD="2.1", PCP="강수없음", SKY=sky, POP="20")
    out = []
    for cat, v in vals.items():
        for k in range(1, n + 1):
            t = issue + pd.Timedelta(hours=k)
            out.append(dict(baseDate=f"{issue:%Y%m%d}", baseTime=f"{issue:%H}00", category=cat, fcstDate=f"{t:%Y%m%d}",
                            fcstTime=f"{t:%H}00", fcstValue=v, nx=nx, ny=ny))
    first_x = 0 if h <= 11 else 1
    first_n = 0 if h == 2 else 1
    for cat, first, cnt, hh, v in (("TMX", first_x, K.TMX_N[h], "1500", "25"), ("TMN", first_n, K.TMN_N[h], "0600", "12")):
        for k in range(cnt):
            d = issue.normalize() + pd.Timedelta(days=first + k)
            out.append(dict(baseDate=f"{issue:%Y%m%d}", baseTime=f"{issue:%H}00", category=cat, fcstDate=f"{d:%Y%m%d}",
                            fcstTime=hh, fcstValue=v, nx=nx, ny=ny))
    return [x for x in out if x["category"] not in drop]


def test_check_issue_counts():
    iss = TS("2026-06-10 02:00")
    c = K.check_issue(K.items_frame(_issue_items(iss)), iss)
    assert c["complete8"] and c["counts"]["TMP"] == 78 and c["counts"]["TMX"] == 4
    it = _issue_items(iss)
    it.pop(3)                                                         # TMP 한 칸 빠짐
    c = K.check_issue(K.items_frame(it), iss)
    assert not c["required_ok"] and "TMP 77/78" in c["missing"]
    it = _issue_items(iss)
    it[[i for i, x in enumerate(it) if x["category"] == "SKY"][0]]["fcstValue"] = "0"   # 코드표 밖 값 → 빠진 칸
    c = K.check_issue(K.items_frame(it), iss)
    assert c["required_ok"] and not c["optional_ok"]
    it = _issue_items(iss)
    it[0]["fcstValue"] = "-999"                                       # 결측값
    assert not K.check_issue(K.items_frame(it), iss)["required_ok"]
    c = K.check_issue(K.items_frame(_issue_items(TS("2026-06-10 17:00"))), TS("2026-06-10 17:00"))
    assert c["complete8"] and c["expected"]["TMP"] == 87


def test_store_roundtrip_and_optional_drop(tmp_path):
    st = K.Store(str(tmp_path))
    now = TS("2026-06-10 02:11")
    for iss, drop in ((TS("2026-06-09 23:00"), ()), (TS("2026-06-10 02:00"), ("SKY",))):
        df = K.items_frame(_issue_items(iss, drop=drop))
        chk = K.check_issue(df, iss)
        st.store_issue("73_134", iss, df, chk, now)
    assert st.has_issue("73_134", TS("2026-06-10 02:00")) and not st.issue_complete8("73_134", TS("2026-06-10 02:00"))
    a = load_archive(st.fcst_files("73_134"))
    assert "73_134" in a.location
    assert TS("2026-06-10 02:00") in set(a.complete_issues)
    assert TS("2026-06-10 02:00") not in a.issue_set("POP")          # 하늘상태가 덜 찬 발표는 강수확률도 뺌(→ S3)
    assert TS("2026-06-09 23:00") in a.issue_set("SKY")
    paths = st.save_raw("fcst", "73_134", "20260610_0200", [(1, b'{"x":1}')], now)
    import gzip
    assert gzip.open(paths[0]).read() == b'{"x":1}'                  # 원자료는 받은 그대로


# ── 운영 서비스 표 ───────────────────────────────────────────────────
def _arch_from_items(tmp_path, issues):
    st = K.Store(str(tmp_path))
    for iss in issues:
        df = K.items_frame(_issue_items(iss))
        st.store_issue("73_134", iss, df, K.check_issue(df, iss), TS(iss))
    return load_archive(st.fcst_files("73_134"))


def test_ops_service_table_uses_previous_issue(tmp_path):
    a = _arch_from_items(tmp_path, [TS("2026-06-09 23:00"), TS("2026-06-10 02:00"), TS("2026-06-10 14:00")])
    st = ops_service_table(a, start=TS("2026-06-10"), end=TS("2026-06-10 17:00"))
    m = st[st.run == TS("2026-06-10 02:00")]
    assert not m.backup.any() and list(m.src_lead) == list(m.lead_day) == [0, 1, 2, 3]
    ref = service_table(a)
    ref = ref[ref.run == TS("2026-06-10 02:00")]
    assert np.allclose(m.ea.values, ref.ea.values) and np.allclose(m.u10.values, ref.u10.values)
    e = st[st.run == TS("2026-06-10 17:00")]                          # 17시 없음 → 14시 발표로(D+4는 없음)
    assert e.backup.all() and set(e.src_run) == {TS("2026-06-10 14:00")}
    assert list(e.lead_day) == [1, 2, 3, 4] and list(e.src_lead) == [1, 2, 3, 4]
    assert e.set_index("lead_day").loc[4, "Tmax"] is None or pd.isna(e.set_index("lead_day").loc[4, "Tmax"])
    assert set(e.err_name) == {"아침"} and error_kind(TS("2026-06-09 23:00")) == "저녁"
    a2 = _arch_from_items(tmp_path / "b", [TS("2026-06-09 23:00")])
    m2 = ops_service_table(a2, start=TS("2026-06-10"), end=TS("2026-06-10 02:00"))
    assert m2.backup.all() and list(m2.src_lead) == [1, 2, 3, 4] and set(m2.err_name) == {"저녁"}
    ref2 = daily_inputs(a2, TS("2026-06-09 23:00"), [TS("2026-06-10") + pd.Timedelta(days=k) for k in range(4)])
    assert np.allclose(m2.ea.values, ref2.ea.values)
    a3 = _arch_from_items(tmp_path / "c", [TS("2026-06-08 23:00")])     # 27시간 전 → 대신할 발표 없음
    ops_service_table(a3, start=TS("2026-06-10"), end=TS("2026-06-10 02:00"))
    assert ("아침", TS("2026-06-10 02:00")) in a3.failed_slots


def test_forecast_table_s3_fallback_and_src_lead():
    from cropwater_fcst import forecast_table, load_s4_fixed
    from rs_model import load_coef
    base = os.path.join(os.path.dirname(__file__), "..")
    s4 = load_s4_fixed("101", os.path.join(base, "rs_sky_coef.csv"))
    coef = load_coef("101", os.path.join(base, "rs_coef.csv"))
    run = TS("2026-06-10 02:00")
    st = pd.DataFrame(dict(run_name="아침", run=run, target=[TS("2026-06-11"), TS("2026-06-12"), TS("2026-06-11")],
                           lead_day=[1, 2, 1], src_lead=[1, 2, 2], Tmax=28.0, Tmin=18.0, ea=1.8, u10=2.0, rain=0.0, rain_flag=0,
                           sky_cloudy=[0.2, np.nan, 0.2], sky_overcast=[0.1, np.nan, 0.1], pop=0.2, pop_max=[0.3, np.nan, 0.3],
                           rain_exp=0.0))
    obs = pd.DataFrame(dict(date=pd.date_range("2026-06-01", periods=15), ETo_obs=4.0, Tmax=27.0, Tmin=17.0, ea_obs=1.7,
                            u10=1.5, Rs=20.0, rain=0.0, pa=1000.0))
    kp = dict(bud=dt.date(2026, 4, 1), L=(20, 70, 90, 30), kc_ini=0.5, kc_mid=1.1, kc_end=0.9)
    ft = forecast_table(st, obs, 37.9, 75.8, coef, kp, s4, s3_fallback=True)
    assert list(ft.rs_method) == ["S4", "S3", "S4"]
    assert ft.ETo_main.iloc[1] == pytest.approx(ft.ETo_S3.iloc[1])
    c = s4.set_index("lead_day")
    assert ft.s4_a.iloc[2] == pytest.approx(c.loc[2, "a"]) and ft.s4_a.iloc[0] == pytest.approx(c.loc[1, "a"])
    assert ft.ETo_S4.iloc[2] != pytest.approx(ft.ETo_S4.iloc[0])      # 같은 입력이라도 대체 발표 선행일의 계수
    ft0 = forecast_table(st.drop(columns="src_lead"), obs, 37.9, 75.8, coef, kp, s4)
    assert "rs_method" not in ft0 and np.isnan(ft0.ETo_main.iloc[1])  # 검증 경로는 그대로(S4 없으면 비움)


# ── 슬롯·수집 ─────────────────────────────────────────────────────────
def test_current_slot():
    from cropwater_ops import current_slot, parse_slot
    assert current_slot(TS("2026-10-01 02:10")) == ("아침", TS("2026-10-01 02:00"))
    assert current_slot(TS("2026-10-01 02:09:30")) == ("아침", TS("2026-10-01 02:00"))    # 1분 일찍 불려도
    assert current_slot(TS("2026-10-01 17:12")) == ("저녁", TS("2026-10-01 17:00"))
    assert current_slot(TS("2026-10-01 01:00")) == ("저녁", TS("2026-09-30 17:00"))
    assert parse_slot("2026-10-01 17") == ("저녁", TS("2026-10-01 17:00"))


def _collector(tmp_path, items, faults=None, t0="2026-06-10 02:10"):
    from cropwater_ops import Collector
    from ops_replay import MockApi, SimClock
    clock = SimClock(TS(t0))
    api = MockApi(items, {}, clock, faults)
    cfg = dict(stn="101", grid="73_134", data=str(tmp_path), obs=None)
    return Collector(cfg, K.Store(str(tmp_path)), api, clock, key="K"), api, clock


def _items(issues):
    return {TS(i): _issue_items(i) for i in issues}


ISS = ["2026-06-09 17:00", "2026-06-09 20:00", "2026-06-09 23:00", "2026-06-10 02:00"]


def test_collect_delay_retry_on_time(tmp_path):
    c, api, clock = _collector(tmp_path, _items(ISS), {"delay_min": {"2026-06-10 02:00": 12}})
    r = c.collect_service("아침", TS("2026-06-10 02:00"))
    assert r["complete8"] and r["attempts"] == 4                       # +0, +5, +10 자료 없음 → +15 받음
    assert TS("2026-06-10 02:25") <= r["received_at"] <= TS("2026-06-10 02:40")
    assert c.store.has_issue("73_134", TS("2026-06-09 23:00"))         # 첫 시도 뒤 직전 발표를 먼저 확보
    log = c.store.read_log("collect_log.csv")
    assert (log.status == "nodata").sum() >= 3


def test_collect_missing_goes_to_backup(tmp_path):
    c, api, clock = _collector(tmp_path, _items(ISS), {"missing": ["2026-06-10 02:00"]})
    r = c.collect_service("아침", TS("2026-06-10 02:00"))
    assert not r["complete8"] and r["attempts"] == 6 and clock.now() <= TS("2026-06-10 02:41")
    from cropwater_ops import light_service
    a = load_archive(c.store.fcst_files("73_134"))
    st = ops_service_table(a, start=TS("2026-06-10"), end=TS("2026-06-10 02:00"))
    assert set(st.src_run) == {TS("2026-06-09 23:00")} and st.backup.all()


def test_collect_truncated_then_complete(tmp_path):
    c, api, clock = _collector(tmp_path, _items(ISS), {"truncate": {"2026-06-10 02:00": 1}})
    r = c.collect_service("아침", TS("2026-06-10 02:00"))
    assert r["complete8"] and r["attempts"] == 2
    log = c.store.read_log("collect_log.csv")
    assert "incomplete" in set(log.status)


def test_collect_fatal_stops(tmp_path):
    c, api, clock = _collector(tmp_path, _items(ISS), {"fatal": [["2026-06-10 00:00", "2026-06-10 03:00"]]})
    r = c.collect_service("아침", TS("2026-06-10 02:00"))
    assert not r["complete8"] and r["attempts"] == 1 and "30" in c.fatal
    assert api.calls == 1                                              # 인증 오류면 다시 요청하지 않음


def test_collect_optional_missing_is_partial(tmp_path):
    c, api, clock = _collector(tmp_path, _items(ISS), {"drop_opt": ["2026-06-10 02:00"]})
    r = c.collect_service("아침", TS("2026-06-10 02:00"))
    assert r["required_ok"] and not r["complete8"] and r["attempts"] == 6     # 마감까지 8요소를 기다린 뒤 S3
    assert c.store.has_issue("73_134", TS("2026-06-10 02:00"))


def test_h3_summary_counts_missing_slots():
    from cropwater_ops import h3_summary
    sl = pd.DataFrame([
        dict(slot_run="2026-06-10 02:00", run_name="아침", started_at="2026-06-10 02:10:00", on_time="True", result="정시 성공", excel="a"),
        dict(slot_run="2026-06-10 17:00", run_name="저녁", started_at="2026-06-10 17:10:00", on_time="False",
             result="백업: 직전 발표", excel="b"),
        dict(slot_run="2026-06-10 17:00", run_name="저녁", started_at="2026-06-10 18:30:00", on_time="False", result="지연 수집", excel="b"),
        dict(slot_run="2026-06-11 17:00", run_name="저녁", started_at="2026-06-11 17:10:00", on_time="False", result="실패", excel=np.nan),
        dict(slot_run="2026-06-11 17:00", run_name="저녁", started_at="2026-06-11 19:00:00", on_time=np.nan, result="다시 만듦", excel="c"),
    ])
    s, tab, _ = h3_summary(sl, pd.DataFrame())
    assert s["n_expected"] == 4 and s["n_on_time"] == 1                # 6/11 02시는 실행 기록 없음 → 실패
    assert s["n_backup_auto"] == 1 and s["n_not_on_time"] == 3          # 6/11 17시: 처음 실행이 실패 → 뒤의 재실행 엑셀을 섞지 않음
    assert tab.set_index("slot_run").loc[TS("2026-06-10 17:00"), "final_result"] == "지연 수집"
    assert tab.set_index("slot_run").loc[TS("2026-06-11 02:00"), "first_result"] == "실행 안 됨"
    assert s["verdict"].startswith("미달")


# ── 서비스 전망: 대체 발표 오차 칸, 관측 지연 ─────────────────────────────
def test_service_outlook_backup_uses_source_error():
    import fcst_wb as W
    from test_wb import KP1, SOIL, _ft, _obs
    D = "2026-06-10"
    ft = _ft(D)
    m = ft.run_name == "아침"
    ft["err_name"], ft["err_lead"], ft["backup"] = "아침", ft["lead_day"], False
    ft.loc[m, "err_name"], ft.loc[m, "err_lead"], ft.loc[m, "backup"] = "저녁", ft.loc[m, "lead_day"] + 1, True
    err = pd.DataFrame([dict(kind="day", run_name=rn, lead_day=k, month=0, n=100, rmse=0.1 * (k + 1) + (0.05 if rn == "저녁" else 0),
                             mbe=0.0, obs_mean=4.0) for rn in ("아침", "저녁") for k in range(5)] +
                       [dict(kind="cum3", run_name="저녁", lead_day=1, month=0, n=100, rmse=1.2, mbe=0.0, obs_mean=12.0),
                        dict(kind="cum3", run_name="아침", lead_day=0, month=0, n=100, rmse=2.4, mbe=0.0, obs_mean=12.0)])
    owb = W.observed_wb(_obs(["2026-06-01"], [5.0] * 12, [0.0] * 12), KP1, SOIL)
    ol = W.service_outlook(ft, owb, SOIL, err, "아침", TS(D) + pd.Timedelta(hours=2))
    assert list(ol["days"].rel.round(4)) == [round((0.1 * (k + 2) + 0.05) / 4, 4) for k in range(4)]   # 저녁 D+1~D+4 칸
    assert ol["cum3"][3] == pytest.approx(0.1)                          # 3일 누적도 대체 발표(저녁 D+1) 칸


def test_forecast_runs_morning_obs_lag():
    import fcst_wb as W
    from fao56_core import wb_step
    from test_wb import KP1, SOIL, _ft, _obs
    ft = pd.concat([_ft("2026-06-09"), _ft("2026-06-10")], ignore_index=True)
    owb = W.observed_wb(_obs(["2026-06-01"], [5.0] * 15, [0.0] * 15), KP1, SOIL)
    r0 = W.forecast_runs(ft, owb, SOIL)
    r1 = W.forecast_runs(ft, owb, SOIL, morning_obs_lag=True)
    k = (r1.run_name == "아침") & (r1.run == TS("2026-06-10 02:00"))
    o = owb.set_index("date")
    lag0 = wb_step(float(o.loc[TS("2026-06-08"), "Dr"]), 0.0, 4.0, 120.0, 60.0)[3]      # 6/9는 그날 아침 D+0 예보(ETc 4)로
    assert r1[k].dr_start.iloc[0] == pytest.approx(lag0) and r1[k].dr_start_true.iloc[0] == pytest.approx(45.0)
    assert list(r1[k].Dr_true) == pytest.approx(list(r0[k].Dr_true))                     # 참값은 그대로
    assert list(r1[k].Dr_center) == pytest.approx(W.path(lag0, [4.0] * 4, [0.0, 6.0, 0.0, 0.0], SOIL))
    e = (r1.run_name == "저녁")
    assert list(r1[e].Dr_center) == pytest.approx(list(r0[e].Dr_center))                 # 저녁 발표는 영향 없음


def test_service_outlook_obs_lag_matches_range_path():
    """전날 관측이 없을 때 서비스 전망(전날을 예보 하루로 진행, 범위 포함) = 검증 경로 forecast_runs('range')"""
    import fcst_wb as W
    from fao56_core import wb_step
    from test_wb import KP1, SOIL, _ft, _obs
    ft = pd.concat([_ft("2026-06-09"), _ft("2026-06-10")], ignore_index=True)
    ft.loc[(ft.run == TS("2026-06-09 02:00")) & (ft.lead_day == 0), ["rain", "rain_exp"]] = [8.0, 4.8]   # 6/9 비 예보
    full = _obs(["2026-06-01"], [5.0] * 15, [0.0] * 15)
    owb_full = W.observed_wb(full, KP1, SOIL)
    runs = W.forecast_runs(ft, owb_full, SOIL, morning_obs_lag="range")
    g = runs[runs.run == TS("2026-06-10 02:00")].sort_values("order")
    lagged = full[full.date < TS("2026-06-09")]                                         # 6/9 관측 아직 없음
    owb = W.observed_wb(lagged, KP1, SOIL, fill=W.fill_from_forecast(ft), end=TS("2026-06-09"))
    assert W.last_observed(owb, TS("2026-06-10")) == TS("2026-06-08")
    ol = W.service_outlook(ft, owb, SOIL, None, "아침", TS("2026-06-10 02:00"), obs_last=TS("2026-06-08"))
    assert [p["kind"] for p in ol["pre"]] == ["lag"] and ol["lag_days"] == [TS("2026-06-09")]
    for k in ("center", "early", "late"):
        assert list(ol["days"][f"Dr_{k}"]) == pytest.approx(list(g[f"Dr_{k}"]))
    dr8 = float(owb.set_index("date").loc[TS("2026-06-08"), "Dr"])
    assert ol["start"] == pytest.approx(wb_step(dr8, 4.8, 4.0, 120.0, 60.0)[3])         # 중심 = 기대 강수
    assert ol["pre"][0]["Dr_early"] == pytest.approx(wb_step(dr8, 0.0, 4.0 * 1.25, 120.0, 60.0)[3])   # 빠르면 = 비 없음
    # 전날 관수 기록(관측 물수지 노란 칸)과 관측 강수(ASOS 행은 있고 일사만 빠짐)는 세 경로 모두에 들어감
    part = full[full.date <= TS("2026-06-09")].copy()
    part.loc[part.date == TS("2026-06-09"), ["ETo_obs", "rain"]] = [np.nan, 3.0]
    owb2 = W.observed_wb(part, KP1, SOIL, irrig={dt.date(2026, 6, 9): 10 / 0.95}, fill=W.fill_from_forecast(ft))
    ol2 = W.service_outlook(ft, owb2, SOIL, None, "아침", TS("2026-06-10 02:00"), obs_last=W.last_observed(owb2, TS("2026-06-10")))
    p = ol2["pre"][0]
    assert p["rain_known"] and p["rain"] == 3.0 and p["irr"] == pytest.approx(10.0)
    assert p["Dr_early"] == pytest.approx(wb_step(dr8, 3.0, 5.0, 120.0, 60.0, 10.0)[3])


def test_late_rerun_is_late_collection(tmp_path):
    from cropwater_ops import slot_result
    c, api, clock = _collector(tmp_path, _items(ISS), {"delay_min": {"2026-06-10 02:00": 45}})
    r = c.collect_service("아침", TS("2026-06-10 02:00"))
    assert not r["required_ok"] and r["attempts"] == 6                   # 02:55에야 제공 → 마감(02:40)까지 못 받음
    clock.t = TS("2026-06-10 03:10")                                    # 다시 실행
    r2 = c.collect_service("아침", TS("2026-06-10 02:00"))
    assert r2["late"] and r2["attempts"] == 1 and r2["complete8"]
    assert slot_result(r2, dict(ok=True, backup=False, rs_method="S4")) == "지연 수집"
    assert slot_result(r, dict(ok=True, backup=True, rs_method="S4")) == "백업: 직전 발표"
    r3 = c.collect_service("아침", TS("2026-06-10 02:00"))                # 또 실행 → 이미 받음
    assert r3["rerun"] and slot_result(r3, dict(ok=True, backup=False, rs_method="S4")) == "다시 만듦"
