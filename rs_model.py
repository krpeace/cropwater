#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rs_model.py — 일사량(Rs) 추정 (02-Cycle 2단계)

단기예보에는 일사·일조가 없으므로 기온교차와 강수유무로 Rs를 추정한다 (THEORY.md 9장).
  S1  FAO-56 식(50)      : Rs = kRs·√(Tmax−Tmin)·Ra,  kRs = 0.16(내륙)/0.19(해안)
  S3  식(50) + 강수유무   : Rs/Ra = a + b·√(Tmax−Tmin) + c·(일강수 ≥ 1 mm)
  추정값은 [0.05·Ra, Rso] 범위로 제한.

계수 2계층 (rs_coef.csv): stn=0 행은 FAO-56 기본값, 지점 행은 관측 Rs로 맞춘 보정값.
지점 보정값은 검증 기간과 겹치지 않는 해의 자료로 정한다(예: 2026년 검증 → 2025년 자료).

S4 (G3 재검증, 2026-09-29) — 하늘상태(SKY) 예보로 구름 정보를 더함
  Rs/Ra = a + b·√(Tmax−Tmin) + c·강수유무 + d·구름많음 비율 + e·흐림 비율     (선행일 k마다 계수)
  구름 비율은 낮 시간 일사 비중으로 가중한 값(fcst_archive). 관측 운량이 없으므로 계수는 '예보 입력 → 관측 Rs'로
  맞춘다(예보 오차까지 흡수). 계수 파일: rs_sky_coef.csv (stn, lead_day, a~e).
"""
import csv, datetime as dt, math, os

import numpy as np
import pandas as pd

from fao56_core import extra_radiation

RS_COEF_DEFAULT = "rs_coef.csv"
RS_SKY_COEF_DEFAULT = "rs_sky_coef.csv"
_COLS = ["stn", "method", "krs", "a", "b", "c", "fit_start", "fit_end", "n", "rmse_rs", "note"]
_SKY_COLS = ["stn", "lead_day", "a", "b", "c", "d", "e", "fit_start", "fit_end", "n", "rmse_rs", "note"]
S4_NAMES = ("a", "b", "c", "d", "e")


def ra_rso(lat, elev, doy):
    ra = np.array([extra_radiation(lat, int(j)) for j in np.atleast_1d(doy)])
    return ra, (0.75 + 2e-5 * elev) * ra


def clip_rs(rs, ra, rso):
    return np.minimum(np.maximum(rs, 0.05 * ra), rso)


def rs_s1(tmax, tmin, ra, rso, krs=0.16):
    dT = np.maximum(np.asarray(tmax, float) - np.asarray(tmin, float), 0.0)
    return clip_rs(krs * np.sqrt(dT) * ra, ra, rso)


def rs_s3(tmax, tmin, rain_flag, ra, rso, a, b, c):
    dT = np.maximum(np.asarray(tmax, float) - np.asarray(tmin, float), 0.0)
    return clip_rs((a + b * np.sqrt(dT) + c * np.asarray(rain_flag, float)) * ra, ra, rso)


def s4_design(tmax, tmin, rain_flag, cloudy, overcast):
    """S4 설명변수 [1, √(Tmax−Tmin), 강수유무, 구름많음 비율, 흐림 비율]"""
    dT = np.maximum(np.asarray(tmax, float) - np.asarray(tmin, float), 0.0)
    return np.column_stack([np.ones(len(dT)), np.sqrt(dT), np.asarray(rain_flag, float),
                            np.asarray(cloudy, float), np.asarray(overcast, float)])


def rs_s4(tmax, tmin, rain_flag, cloudy, overcast, ra, rso, coef):
    """S4 Rs. coef: (a, b, c, d, e) 한 벌 또는 행마다 한 벌(n×5 배열)"""
    X = s4_design(tmax, tmin, rain_flag, cloudy, overcast)
    co = np.asarray(coef, float)
    kt = X @ co if co.ndim == 1 else (X * co).sum(axis=1)
    ra = np.asarray(ra, float)
    return clip_rs(kt * ra, ra, np.asarray(rso, float))


def fit_s4(tmax, tmin, rain_flag, cloudy, overcast, ra, rs_obs):
    """S4 계수 최소제곱 적합(Rs/Ra에 대해). 반환: (a, b, c, d, e) 배열"""
    X = s4_design(tmax, tmin, rain_flag, cloudy, overcast)
    y = np.asarray(rs_obs, float) / np.asarray(ra, float)
    ok = np.isfinite(X).all(axis=1) & np.isfinite(y)
    return np.linalg.lstsq(X[ok], y[ok], rcond=None)[0]


def save_sky_coef(stn, table, path=RS_SKY_COEF_DEFAULT, note=""):
    """선행일별 S4 계수표(열 lead_day, a~e, n, fit_start, fit_end, rmse_rs)를 rs_sky_coef.csv에 기록(같은 지점 행은 교체)"""
    rows = []
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            rows = [r for r in csv.DictReader(f) if str(r.get("stn")) != str(stn)]
    for r in table.to_dict("records"):
        rows.append(dict(stn=stn, lead_day=int(r["lead_day"]), **{k: round(float(r[k]), 4) for k in S4_NAMES},
                         fit_start=r.get("fit_start", ""), fit_end=r.get("fit_end", ""), n=int(r.get("n", 0)),
                         rmse_rs=round(float(r.get("rmse_rs", float("nan"))), 3), note=note))
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=_SKY_COLS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in _SKY_COLS})


def load_sky_coef(stn, path=RS_SKY_COEF_DEFAULT):
    """지점의 선행일별 S4 계수 {lead_day: (a, b, c, d, e)}. 없으면 빈 dict"""
    if not os.path.exists(path):
        return {}
    with open(path, encoding="utf-8-sig") as f:
        return {int(r["lead_day"]): tuple(float(r[k]) for k in S4_NAMES)
                for r in csv.DictReader(f) if str(r.get("stn")) == str(stn)}


def fit(obs, lat, elev, months=range(4, 10)):
    """관측 일자료(DataFrame: date, Tmax, Tmin, Rs, rain)로 kRs(S1형)와 a·b·c(S3)를 최소제곱 적합"""
    d = obs.dropna(subset=["Tmax", "Tmin", "Rs"]).copy()
    d = d[d["date"].dt.month.isin(list(months)) & (d["Tmax"] > d["Tmin"])]
    ra, rso = ra_rso(lat, elev, d["date"].dt.dayofyear)
    x = np.sqrt(d["Tmax"] - d["Tmin"]).values
    flag = (d["rain"].fillna(0) >= 1.0).astype(float).values
    y = d["Rs"].values
    krs = float((x * ra * y).sum() / ((x * ra) ** 2).sum())
    X = np.column_stack([np.ones(len(d)), x, flag])
    a, b, c = np.linalg.lstsq(X, y / ra, rcond=None)[0]
    est = rs_s3(d["Tmax"], d["Tmin"], flag, ra, rso, a, b, c)
    return dict(krs=krs, a=float(a), b=float(b), c=float(c), n=int(len(d)),
                fit_start=str(d["date"].min().date()), fit_end=str(d["date"].max().date()),
                rmse_rs=float(np.sqrt(np.mean((est - y) ** 2))))


def load_coef(stn, path=RS_COEF_DEFAULT):
    """지점 계수(없으면 FAO-56 기본값 행). 반환: dict(method, krs, a, b, c, source)"""
    rows = []
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    pick = [r for r in rows if str(r.get("stn")) == str(stn)] or [r for r in rows if str(r.get("stn")) == "0"]
    if not pick:
        return dict(method="S1", krs=0.16, a=None, b=None, c=None, source="FAO-56 식(50) 내륙 기본값")
    r = pick[-1]
    num = lambda k: float(r[k]) if r.get(k) not in (None, "") else None
    return dict(method=r.get("method") or ("S3" if num("a") is not None else "S1"), krs=num("krs") or 0.16,
                a=num("a"), b=num("b"), c=num("c"),
                source=f"{os.path.basename(path)} stn={r.get('stn')} ({r.get('fit_start', '')}~{r.get('fit_end', '')})",
                stn=r.get("stn"), fit_start=r.get("fit_start", ""), fit_end=r.get("fit_end", ""),
                n=r.get("n", ""), rmse_rs=num("rmse_rs"), note=r.get("note", ""))


def save_coef(stn, coef, path=RS_COEF_DEFAULT, note=""):
    """지점 계수를 rs_coef.csv에 기록(같은 지점 행은 교체). 파일이 없으면 FAO 기본값 행과 함께 생성."""
    rows = []
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            rows = [r for r in csv.DictReader(f) if str(r.get("stn")) != str(stn)]
    if not any(str(r.get("stn")) == "0" for r in rows):
        rows.insert(0, dict(stn=0, method="S1", krs=0.16, a="", b="", c="", fit_start="", fit_end="", n="",
                            rmse_rs="", note="FAO-56 식(50) 내륙 기본값 (해안 0.19)"))
    rows.append(dict(stn=stn, method="S3", krs=round(coef["krs"], 4), a=round(coef["a"], 4), b=round(coef["b"], 4),
                     c=round(coef["c"], 4), fit_start=coef.get("fit_start", ""), fit_end=coef.get("fit_end", ""),
                     n=coef.get("n", ""), rmse_rs=round(coef.get("rmse_rs", float("nan")), 3), note=note))
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=_COLS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in _COLS})
