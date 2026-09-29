#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
obs_daily.py — ASOS 관측 일자료 로더 (02-Cycle 검증용 기준값)

01-Cycle cropwater_station.py 출력 워크북의 '원데이터' 시트를 읽어 표준 열 이름으로 돌려준다.
  date, Tmax, Tmin, RHmean, RHmin, u10(평균풍속), pv(hPa), td, pa(hPa), Rs, ss, epan, rain
관측 기준 ETo(ETo_obs)는 01-Cycle과 같은 규칙(ea=증기압 우선, 기압=관측값)으로 계산한다.
"""
import datetime as dt

import numpy as np
import openpyxl
import pandas as pd

from fao56_core import eto_penman_monteith, svp, wind_2m

_COLS = ["date", "Tmax", "Tmin", "Tavg", "RHmean", "RHmin", "u10", "pv", "td", "pa", "Rs", "ss", "epan", "rain"]


def load_station_workbook(path):
    """01-Cycle 출력 워크북 → (DataFrame, meta: lat, elev, anem, settings{설정 시트 항목: 값})"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = [r for r in wb["원데이터"].iter_rows(min_row=2, values_only=True) if r and r[0] is not None]
    df = pd.DataFrame([r[:14] for r in rows], columns=_COLS)
    df["date"] = pd.to_datetime(df["date"])
    for c in _COLS[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["rain"] = df["rain"].fillna(0.0)          # ASOS 일강수는 무강수일에 비어 있음
    settings = {}
    for r in wb["설정"].iter_rows(min_row=1, max_row=80, values_only=True):
        if r and isinstance(r[0], str) and len(r) > 1 and r[1] is not None:
            settings[r[0].strip()] = r[1]
    meta = {"settings": settings}
    for key, label in (("lat", "위도 (°)"), ("elev", "고도 (m)"), ("anem", "풍속계 높이 (m)")):
        if label in settings:
            meta[key] = float(settings[label])
    return df, meta


def kc_params(settings, crop_scenarios=None):
    """01-Cycle 설정 시트 값 → Kc 계산 인자 (시나리오 표값 + 현지기상 보정 + 멀칭).
       예보 ETc는 01-Cycle 워크북과 같은 Kc를 써야 관측 ETc와 비교할 수 있다."""
    from fao56_core import KC_SCENARIOS, kc_climate_adjust
    g = lambda k, d=None: settings.get(k, d)
    scn = int(g("시나리오 번호 (1~4)", 3) or 3)
    table = (crop_scenarios or {}).get(scn) or KC_SCENARIOS.get(scn, KC_SCENARIOS[3])
    kc_ini, kc_mid, kc_end = table[:3]
    h = float(g("생육중기 초목 수고 h (m)", 3.2))
    mulch = float(g("멀칭 Kc 보정계수", 1.0) or 1.0)
    u2m, rhm = float(g("중기 평균 u2 (m/s)", 1.5)), float(g("중기 평균 RHmin (%)", 55))
    u2e, rhe = float(g("후기 평균 u2 (m/s)", 1.5)), float(g("후기 평균 RHmin (%)", 55))
    bud = g("생육 시작일 (발아기/정식일)")
    bud = pd.Timestamp(bud).date() if bud is not None else None
    return dict(scenario=scn, bud=bud,
                L=(int(g("L_ini (초기, 일)", 20)), int(g("L_dev (발육, 일)", 70)),
                   int(g("L_mid (중기, 일)", 90)), int(g("L_late (후기, 일)", 30))),
                kc_ini=kc_ini * mulch,
                kc_mid=kc_climate_adjust(kc_mid, u2m, rhm, h) * mulch,
                kc_end=kc_climate_adjust(kc_end, u2e, rhe, h, is_end=True) * mulch,
                table=(kc_ini, kc_mid, kc_end), h=h, mulch=mulch, u2_mid=u2m, rh_mid=rhm, u2_end=u2e, rh_end=rhe)


def kc_series(dates, kp):
    """날짜 목록 → 일별 Kc (fao56_core.kc_of_date, 01-Cycle 계산과정 시트와 같은 식)"""
    from fao56_core import kc_of_date
    if kp.get("bud") is None:
        return np.full(len(dates), np.nan)
    b = kp["bud"].toordinal()
    return np.array([kc_of_date(pd.Timestamp(d).date().toordinal(), b, *kp["L"],
                                kp["kc_ini"], kp["kc_mid"], kp["kc_end"]) for d in dates])


def add_eto_obs(df, lat, elev, anem=10.0):
    """관측 기준 ETo (01-Cycle 동일 규칙: ea = 증기압 → 이슬점 → 평균습도, 기압 = 관측값)"""
    out, eas = [], []
    for r in df.itertuples():
        ea = np.nan
        if not (pd.isna(r.Tmax) or pd.isna(r.Tmin)):
            es = (svp(r.Tmax) + svp(r.Tmin)) / 2
            ea = r.pv * 0.1 if not pd.isna(r.pv) else (svp(r.td) if not pd.isna(r.td) else r.RHmean / 100 * es)
        eas.append(ea)
        if any(pd.isna(v) for v in (r.Tmax, r.Tmin, r.u10, r.Rs, ea)):
            out.append(np.nan); continue
        J = r.date.timetuple().tm_yday
        out.append(eto_penman_monteith(r.Tmax, r.Tmin, r.Rs, wind_2m(r.u10, anem), ea, elev, lat, J,
                                       None if pd.isna(r.pa) else r.pa))
    df = df.copy()
    df["ETo_obs"] = out
    df["ea_obs"] = eas                                   # ETo_obs에 실제로 쓴 ea (증기압 → 이슬점 → 평균습도)
    df["rain_flag_obs"] = (df["rain"] >= 1.0).astype(int)
    return df
