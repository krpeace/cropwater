#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cropwater_fcst.py — 기상청 단기예보 기반 ETo·ETc 예측과 검증 (02-Cycle 3단계, H2)

[하는 일]
  1) 과거 단기예보(포털 CSV)를 서비스 발표(아침 02시 → 오늘~D+3, 저녁 17시 → 내일~D+4)별 일 입력으로 집계
     (fcst_archive.py)
  2) 예보 ETo = FAO-56 PM [식6]. 예보에 없는 Rs는 식(50)+강수유무 보정(rs_model.py, rs_coef.csv)
  3) 예보 ETc = Kc × 예보 ETo  (Kc는 01-Cycle 워크북 설정 시트와 같은 값)
  4) ASOS 관측 ETo(01-Cycle 규칙)와 비교: 선행시간별 RMSE·MBE·R², 지속성·7일평균 기준선, 3일 누적
  5) 검증 엑셀(라이브 수식) 출력

[실행]
  # Rs 계수 보정 (검증 연도와 다른 해의 관측으로)
  python cropwater_fcst.py calib --obs output/eto101_apple_20250101_20251231.xlsx --stn 101
  # H2 검증 엑셀
  python cropwater_fcst.py verify --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101

[입력]
  --fcst : 기상자료개방포털 '단기예보(격자)' CSV 폴더(또는 파일들). 필수 요소 TMX·TMN·TMP·REH·WSD·PCP
  --obs  : 01-Cycle cropwater_station.py 출력 워크북(원데이터·설정 시트). 검증 기간 + 7일 전부터 포함
  rs_coef.csv : 지점별 Rs 계수(없으면 FAO-56 kRs 0.16)

근거: docs/THEORY.md 9장, 설계: docs/ARCHITECTURE.md 7장, 판정 기록: docs/VALIDATION.md
"""
import argparse, math, os

import numpy as np
import pandas as pd

from fao56_core import eto_penman_monteith, wind_2m
from fcst_archive import RAIN_FLAG_MM, SERVICE_RUNS, check_archive, load_archive, service_table
from obs_daily import add_eto_obs, kc_params, kc_series, load_station_workbook
from rs_model import fit, load_coef, ra_rso, rs_s1, rs_s3, save_coef

FCST_ANEM = 10.0            # 단기예보 풍속(WSD) 기준 높이 (m)
KRS_FAO = 0.16              # FAO-56 식(50) 내륙 기본값
H2_SKILL_MIN = 0.30         # H2: 지속성 대비 RMSE 개선율 기준 (VALIDATION.md, 2026-09-29 확정)
H2_RMSE_D1_MAX = 1.0        # H2: D+1 RMSE 상한 (mm/일)
H2_LEADS = (1, 2, 3)        # H2 판정 선행일
# 대표 ASOS 지점과 단기예보 격자(nx_ny) — ARCHITECTURE.md 7장 표
STATIONS = {"101": ("춘천", "73_134"), "216": ("태백", "95_119"), "119": ("수원", "60_120"),
            "131": ("청주", "68_107"), "129": ("서산", "52_109"), "146": ("전주", "63_89"),
            "156": ("광주", "59_75"), "136": ("안동", "90_106"), "192": ("진주", "79_75"),
            "189": ("서귀포", "53_32")}


def pressure_from_elev(elev):
    """FAO-56 식(7) 표준대기압 (kPa)"""
    return 101.3 * ((293 - 0.0065 * elev) / 293) ** 5.26


def eto_series(tmax, tmin, ea, u10, rs, dates, lat, elev, anem=FCST_ANEM, pa_hpa=None):
    """행별 PM ETo. pa_hpa=None이면 고도로 기압 추정(예보), 값이 있으면 관측 기압"""
    out = []
    pa = [None] * len(dates) if pa_hpa is None else list(pa_hpa)
    for tx, tn, e, u, r, d, p in zip(tmax, tmin, ea, u10, rs, dates, pa):
        if any(pd.isna(v) for v in (tx, tn, e, u, r)):
            out.append(np.nan); continue
        out.append(eto_penman_monteith(tx, tn, r, wind_2m(u, anem), e, elev, lat, pd.Timestamp(d).dayofyear,
                                       None if (p is None or pd.isna(p)) else p))
    return np.array(out)


def rs_estimate(tmax, tmin, flag, dates, lat, elev, coef, method="S3"):
    """Rs 추정 → (Rs, Ra, Rso). S3 = 식(50)+강수유무(계수 a·b·c, 없으면 S1로 대체),
       S2 = 식(50) 지점 kRs, S1 = 식(50) FAO kRs 0.16"""
    ra, rso = ra_rso(lat, elev, pd.DatetimeIndex(dates).dayofyear)
    if method == "S3" and coef.get("a") is not None:
        return rs_s3(tmax, tmin, flag, ra, rso, coef["a"], coef["b"], coef["c"]), ra, rso
    krs = coef.get("krs", KRS_FAO) if method == "S2" else KRS_FAO
    return rs_s1(tmax, tmin, ra, rso, krs), ra, rso


# ── 예보 ETo·ETc와 기준선 ────────────────────────────────────────────────
def forecast_table(st, obs, lat, elev, coef, kp):
    """서비스 표(fcst_archive.service_table) + 관측 → 예보 ETo(S3·S1), 관측·기준선, Kc·ETc"""
    df = st.copy()
    df["target"] = pd.to_datetime(df["target"])
    df["run_date"] = df["run"].dt.normalize()
    rs3, ra, rso = rs_estimate(df.Tmax, df.Tmin, df.rain_flag, df.target, lat, elev, coef, "S3")
    rs1, _, _ = rs_estimate(df.Tmax, df.Tmin, df.rain_flag, df.target, lat, elev, coef, "S1")
    df["Ra"], df["Rso"], df["Rs_S3"], df["Rs_S1"] = ra, rso, rs3, rs1
    df["u2"] = df["u10"].map(lambda u: wind_2m(u, FCST_ANEM))
    df["ETo_S3"] = eto_series(df.Tmax, df.Tmin, df.ea, df.u10, df.Rs_S3, df.target, lat, elev)
    df["ETo_S1"] = eto_series(df.Tmax, df.Tmin, df.ea, df.u10, df.Rs_S1, df.target, lat, elev)
    o = obs.set_index("date")
    look = lambda col, dates: o[col].reindex(pd.DatetimeIndex(dates)).values
    for c_obs, c in (("ETo_obs", "ETo_obs"), ("Tmax", "Tmax_obs"), ("Tmin", "Tmin_obs"), ("ea_obs", "ea_obs"),
                     ("u10", "u10_obs"), ("Rs", "Rs_obs"), ("rain", "rain_obs"), ("pa", "pa_obs")):
        df[c] = look(c_obs, df.target)
    df["flag_obs"] = (df["rain_obs"] >= RAIN_FLAG_MM).astype(int)
    # 기준선: 발표일 전날(가장 최근의 완결된 관측일) 값, 최근 7일 평균
    df["ETo_pers"] = look("ETo_obs", df.run_date - pd.Timedelta(days=1))
    roll7 = o["ETo_obs"].rolling(7, min_periods=7).mean()
    df["ETo_7d"] = roll7.reindex(pd.DatetimeIndex(df.run_date - pd.Timedelta(days=1))).values
    df["Kc"] = kc_series(df.target, kp)
    for c in ("S3", "S1"):
        df[f"ETc_{c}"] = df["Kc"] * df[f"ETo_{c}"]
    df["ETc_obs"] = df["Kc"] * df["ETo_obs"]
    df["ETc_pers"] = df["Kc"] * df["ETo_pers"]
    df["ETc_7d"] = df["Kc"] * df["ETo_7d"]
    return df


def _stats(f, o):
    m = ~(pd.isna(f) | pd.isna(o))
    f, o = np.asarray(f, float)[m], np.asarray(o, float)[m]
    e = f - o
    r = np.corrcoef(f, o)[0, 1] if len(f) > 2 and f.std() > 0 and o.std() > 0 else np.nan
    return dict(n=int(m.sum()), obs_mean=o.mean(), fcst_mean=f.mean(), MBE=e.mean(), MAE=np.abs(e).mean(),
                RMSE=math.sqrt((e ** 2).mean()), R2=r ** 2)


def lead_metrics(df, fcol="ETo_S3", ocol="ETo_obs", pcol="ETo_pers", mcol="ETo_7d"):
    """발표(아침·저녁) × 선행일별 지표와 기준선 대비 개선율"""
    rows = []
    for (rn, k), g in df.groupby(["run_name", "lead_day"], sort=False):
        s = _stats(g[fcol], g[ocol])
        sp, sm = _stats(g[pcol], g[ocol]), _stats(g[mcol], g[ocol])
        s.update(run_name=rn, lead_day=int(k), RMSE_pers=sp["RMSE"], RMSE_7d=sm["RMSE"],
                 skill_pers=1 - s["RMSE"] / sp["RMSE"], skill_7d=1 - s["RMSE"] / sm["RMSE"])
        rows.append(s)
    cols = ["run_name", "lead_day", "n", "obs_mean", "fcst_mean", "MBE", "MAE", "RMSE", "R2",
            "RMSE_pers", "skill_pers", "RMSE_7d", "skill_7d"]
    order = {n: i for i, n in enumerate(SERVICE_RUNS)}
    out = pd.DataFrame(rows)[cols]
    return out.sort_values(["run_name", "lead_day"], key=lambda s: s.map(order) if s.name == "run_name" else s,
                           ignore_index=True)


def cum3(df, fcol="ETo_S3", ocol="ETo_obs", pcol="ETo_pers"):
    """발표별 처음 3개 대상일(아침 D0~D+2, 저녁 D+1~D+3) 누적"""
    rows = []
    for (rn, run), g in df.sort_values("lead_day").groupby(["run_name", "run"], sort=False):
        g3 = g.head(3)
        if len(g3) < 3 or g3[[fcol, ocol, pcol]].isna().any().any():
            continue
        rows.append(dict(run_name=rn, run=run, leads=f"D+{g3.lead_day.min()}~D+{g3.lead_day.max()}",
                         fcst=g3[fcol].sum(), obs=g3[ocol].sum(), pers=g3[pcol].sum()))
    c = pd.DataFrame(rows)
    res = []
    for rn, g in c.groupby("run_name", sort=False):
        s, sp = _stats(g.fcst, g.obs), _stats(g.pers, g.obs)
        res.append(dict(run_name=rn, leads=g.leads.iloc[0], n=s["n"], obs_mean=s["obs_mean"], MBE=s["MBE"],
                        RMSE=s["RMSE"], rel_RMSE=s["RMSE"] / s["obs_mean"], RMSE_pers=sp["RMSE"],
                        skill_pers=1 - s["RMSE"] / sp["RMSE"]))
    return c, pd.DataFrame(res)


def h2_verdict(met, skill_min=H2_SKILL_MIN, rmse_d1_max=H2_RMSE_D1_MAX, leads=H2_LEADS):
    """H2 판정: 발표마다 D+1~D+3 개선율 ≥ skill_min, D+1 RMSE ≤ rmse_d1_max"""
    out = {}
    for rn, g in met.groupby("run_name", sort=False):
        g = g.set_index("lead_day")
        ok_skill = all(g.loc[k, "skill_pers"] >= skill_min for k in leads if k in g.index)
        ok_d1 = g.loc[1, "RMSE"] <= rmse_d1_max if 1 in g.index else False
        out[rn] = dict(skill_ok=bool(ok_skill), d1_ok=bool(ok_d1), pass_=bool(ok_skill and ok_d1),
                       min_skill=float(min(g.loc[k, "skill_pers"] for k in leads if k in g.index)),
                       rmse_d1=float(g.loc[1, "RMSE"]) if 1 in g.index else float("nan"))
    return out


# ── 입력 진단과 오차 분해 ────────────────────────────────────────────────
def input_diagnostics(df):
    """발표 × 선행일별 입력 편향(예보−관측)·RMSE, 강수유무 적중"""
    rows = []
    for (rn, k), g in df.groupby(["run_name", "lead_day"], sort=False):
        r = dict(run_name=rn, lead_day=int(k), n=len(g))
        for name, f, o in (("Tmax", g.Tmax, g.Tmax_obs), ("Tmin", g.Tmin, g.Tmin_obs),
                           ("dT", g.Tmax - g.Tmin, g.Tmax_obs - g.Tmin_obs), ("ea", g.ea, g.ea_obs),
                           ("u10", g.u10, g.u10_obs), ("Rs", g.Rs_S3, g.Rs_obs)):
            s = _stats(f, o)
            r[f"{name}_MBE"], r[f"{name}_RMSE"] = s["MBE"], s["RMSE"]
        hit = int(((g.rain_flag == 1) & (g.flag_obs == 1)).sum())
        miss = int(((g.rain_flag == 0) & (g.flag_obs == 1)).sum())
        fa = int(((g.rain_flag == 1) & (g.flag_obs == 0)).sum())
        cn = int(((g.rain_flag == 0) & (g.flag_obs == 0)).sum())
        r.update(hit=hit, miss=miss, false_alarm=fa, correct_neg=cn,
                 POD=hit / (hit + miss) if hit + miss else np.nan,
                 FAR=fa / (hit + fa) if hit + fa else np.nan,
                 CSI=hit / (hit + miss + fa) if hit + miss + fa else np.nan)
        rows.append(r)
    return pd.DataFrame(rows)


def error_attribution(df, lat, elev, coef):
    """예보 입력을 하나씩 관측값으로 바꿨을 때의 ETo RMSE (S3). 줄어든 만큼이 그 입력 예보오차의 몫.
       '관측 입력 전부'는 구조오차(Rs 추정)만 남은 상태 = H1 조건(기압만 고도 추정)."""
    variants = {
        "예보 입력 그대로": dict(),
        "기온 → 관측": dict(Tmax="Tmax_obs", Tmin="Tmin_obs"),
        "습도(ea) → 관측": dict(ea="ea_obs"),
        "풍속 → 관측": dict(u10="u10_obs"),
        "강수유무 → 관측": dict(rain_flag="flag_obs"),
        "관측 입력 전부 (Rs만 추정)": dict(Tmax="Tmax_obs", Tmin="Tmin_obs", ea="ea_obs", u10="u10_obs",
                                     rain_flag="flag_obs"),
        "참고: Rs만 관측": dict(Rs="Rs_obs"),
    }
    res = {}
    for name, rep in variants.items():
        col = lambda c: df[rep.get(c, c)]
        if "Rs" in rep:
            rs = df["Rs_obs"].values
        else:
            rs, _, _ = rs_estimate(col("Tmax"), col("Tmin"), col("rain_flag"), df.target, lat, elev, coef, "S3")
        res[name] = eto_series(col("Tmax"), col("Tmin"), col("ea"), col("u10"), rs, df.target, lat, elev)
    rows = []
    for (rn, k), idx in df.groupby(["run_name", "lead_day"], sort=False).groups.items():
        r = dict(run_name=rn, lead_day=int(k))
        for name, v in res.items():
            e = v[df.index.get_indexer(idx)] - df.loc[idx, "ETo_obs"].values
            r[name] = math.sqrt(np.nanmean(e ** 2))
        rows.append(r)
    return pd.DataFrame(rows), list(variants)


def bias_correction_cv(df, lat, elev, coef):
    """탐색: 지점 편향 보정의 효과를 월 단위 교차검증으로 평가(보정값은 다른 달 자료로 추정).
       B1 기온(Tmax·Tmin) 가산 보정, B2 기온 + 풍속 가산 보정, B3 ETo 비율 보정. 발표 × 선행일별로 추정.
       반환: (지표표 long: method·발표·선행일, 전체기간 입력 편향 표, 교차검증 묶음 목록)"""
    d = df.copy()
    per = d.target.dt.to_period("M")
    days_last = d.loc[per == per.max(), "target"].dt.normalize().nunique()
    # 마지막 달 대상일이 10일 미만이면(예: 7/1~7/4) 앞 달과 한 묶음
    d["fold"] = per.where(~((per == per.max()) & (days_last < 10)), per.max() - 1).astype(str)
    if d.fold.nunique() < 2:
        raise ValueError("편향 보정 교차검증에는 대상일이 2개 달 이상 필요합니다")
    cols = {"B1": np.full(len(d), np.nan), "B2": np.full(len(d), np.nan), "B3": np.full(len(d), np.nan)}
    pos = {ix: i for i, ix in enumerate(d.index)}
    for f in sorted(d.fold.unique()):
        tr, te = d[d.fold != f], d[d.fold == f]
        for (rn, k), g in te.groupby(["run_name", "lead_day"]):
            t = tr[(tr.run_name == rn) & (tr.lead_day == k)]
            dtx, dtn = (t.Tmax - t.Tmax_obs).mean(), (t.Tmin - t.Tmin_obs).mean()
            du = (t.u10 - t.u10_obs).mean()
            ratio = t.ETo_obs.sum() / t.ETo_S3.sum()
            tx, tn = g.Tmax - dtx, g.Tmin - dtn
            rs, _, _ = rs_estimate(tx, tn, g.rain_flag, g.target, lat, elev, coef, "S3")
            idx = [pos[i] for i in g.index]
            cols["B1"][idx] = eto_series(tx, tn, g.ea, g.u10, rs, g.target, lat, elev)
            cols["B2"][idx] = eto_series(tx, tn, g.ea, (g.u10 - du).clip(lower=0.1), rs, g.target, lat, elev)
            cols["B3"][idx] = g.ETo_S3 * ratio
    names = {"ETo_S3": "보정 없음(원자료)", "B1": "기온 보정", "B2": "기온+풍속 보정", "B3": "ETo 비율 보정"}
    out = []
    for c, name in names.items():
        if c != "ETo_S3":
            d[c] = cols[c]
        m = lead_metrics(d, c)
        m.insert(0, "method", name)
        _, c3 = cum3(d, c)
        m = m.merge(c3[["run_name", "RMSE"]].rename(columns={"RMSE": "RMSE_3d"}), on="run_name", how="left")
        out.append(m)
    bias = (df.assign(dTmax=df.Tmax - df.Tmax_obs, dTmin=df.Tmin - df.Tmin_obs, du10=df.u10 - df.u10_obs)
              .groupby(["run_name", "lead_day"], sort=False)[["dTmax", "dTmin", "du10"]].mean().reset_index())
    return pd.concat(out, ignore_index=True), bias, sorted(d.fold.unique())


# ── H1 (관측 입력 + Rs 추정) — G2 기록용 ──────────────────────────────────
def h1_table(obs, lat, elev, coef, start, end, anem=10.0):
    """관측 입력(기압 포함)에 Rs만 추정 → 관측 ETo와 비교 (S1 FAO kRs, S2 지점 kRs, S3 강수유무 보정)"""
    d = obs[(obs.date >= pd.Timestamp(start)) & (obs.date <= pd.Timestamp(end))].dropna(subset=["ETo_obs"]).copy()
    out = {}
    for m in ("S1", "S2", "S3"):
        rs, _, _ = rs_estimate(d.Tmax, d.Tmin, d.rain_flag_obs, d.date, lat, elev, coef, m)
        est = eto_series(d.Tmax, d.Tmin, d.ea_obs, d.u10, rs, d.date, lat, elev, anem=anem, pa_hpa=d.pa)
        s = _stats(est, d.ETo_obs)
        s["SUMERR"] = (np.nansum(est) - d.ETo_obs.sum()) / d.ETo_obs.sum()
        s["Rs_RMSE"] = _stats(rs, d.Rs)["RMSE"]
        out[m] = (s, est)
    return d, out


# ── 실행 ────────────────────────────────────────────────────────────────
def load_obs(path):
    obs, meta = load_station_workbook(path)
    obs = add_eto_obs(obs, meta["lat"], meta["elev"], meta.get("anem", 10.0))
    return obs, meta


def run_verify(fcst_paths, obs_path, stn, coef_path="rs_coef.csv", analysis=True):
    """검증 전체 계산. analysis=True면 오차분해·편향보정 탐색까지"""
    arch = load_archive(fcst_paths)
    rep = check_archive(arch)
    st = service_table(arch)
    obs, meta = load_obs(obs_path)
    coef = load_coef(stn, coef_path)
    kp = kc_params(meta["settings"])
    df = forecast_table(st, obs, meta["lat"], meta["elev"], coef, kp)
    name, grid = STATIONS.get(str(stn), ("", ""))
    res = dict(stn=str(stn), stn_name=name, stn_grid=grid, arch=arch, check=rep, table=df, obs=obs, meta=meta,
               coef=coef, kp=kp,
               h1_start=pd.Timestamp(kp["bud"]) if kp.get("bud") else df.target.min(), h1_end=obs.date.max())
    if analysis:
        res["attr"], res["attr_names"] = error_attribution(df, meta["lat"], meta["elev"], coef)
        res["bc"], res["bc_bias"], res["bc_folds"] = bias_correction_cv(df, meta["lat"], meta["elev"], coef)
        res["findings"] = findings(res)
    return res


def findings(res):
    """요약 시트의 해석 문장(④~). 수식으로 연결할 수 없는 Python 분석 결과와 자료 조건 경고"""
    out, df = [], res["table"]
    att, names = res["attr"], res["attr_names"]
    # 번호(④~)는 마지막에 순서대로 붙임
    d1 = att[att.lead_day == 1]
    base = d1[names[0]].mean()
    cand = {n: base - d1[n].mean() for n in names[1:5]}
    top = max(cand, key=cand.get)
    struct = d1[names[5]].mean()
    out.append(f"오차 분해(D+1, 두 발표 평균 RMSE {base:.2f}): '{top}' 교체 시 {cand[top]:.2f} mm/일 감소로 가장 큼. "
               f"예보가 완벽해도 남는 구조오차(Rs 추정)는 {struct:.2f} mm/일 (오차분해 ①)")
    bc = res["bc"]
    raw = bc[(bc.method == "보정 없음(원자료)") & (bc.lead_day == 1)].RMSE.mean()
    b2 = bc[(bc.method == "기온+풍속 보정") & (bc.lead_day == 1)].RMSE.mean()
    out.append(f"탐색: 월 단위 교차검증한 기온+풍속 편향 보정은 D+1 RMSE를 {raw:.2f} → {b2:.2f} mm/일"
               f"({(1 - b2 / raw):.0%} 감소), 편향을 거의 0으로 줄임 (오차분해 ②). 채택 여부는 다음 게이트에서 결정")
    loc = res["check"]["location"]
    if res.get("stn_grid") and loc != [res["stn_grid"]]:
        out.append(f"주의: 예보 격자 {', '.join(loc)}가 ASOS {res['stn']} 격자({res['stn_grid']})와 다릅니다. "
                   f"대표성 오차(특히 기온 편향)가 달라질 수 있어 {res['stn_grid']} 자료로 재확인이 필요합니다")
    months = sorted(df.target.dt.month.unique())
    if not set(range(4, 10)) <= set(months):
        out.append(f"기간: 대상일 {df.target.min():%Y-%m-%d}~{df.target.max():%Y-%m-%d}만 포함한 "
                   f"중간 결과입니다. 생육기(4~9월) 전체 판정에는 나머지 달의 예보 자료가 필요합니다")
    return [f"{'④⑤⑥⑦⑧⑨'[i]} {t}" for i, t in enumerate(out)]


def main(argv=None):
    ap = argparse.ArgumentParser(description="단기예보 기반 ETo·ETc 예측 검증 (02-Cycle)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("calib", help="관측 워크북으로 Rs 계수 보정 → rs_coef.csv")
    c.add_argument("--obs", required=True); c.add_argument("--stn", required=True)
    c.add_argument("--coef", default="rs_coef.csv"); c.add_argument("--months", default="4-9")
    v = sub.add_parser("verify", help="과거 단기예보로 H2 검증 엑셀 작성")
    v.add_argument("--fcst", nargs="+", required=True); v.add_argument("--obs", required=True)
    v.add_argument("--stn", required=True); v.add_argument("--coef", default="rs_coef.csv")
    v.add_argument("--out", default=None)
    a = ap.parse_args(argv)

    if a.cmd == "calib":
        m0, m1 = (int(x) for x in a.months.split("-"))
        obs, meta = load_obs(a.obs)
        cf = fit(obs, meta["lat"], meta["elev"], months=range(m0, m1 + 1))
        save_coef(a.stn, cf, a.coef, note=f"{os.path.basename(a.obs)} {m0}~{m1}월 관측 Rs로 보정")
        print(f"[계수] stn={a.stn} {cf}")
        print(f"[저장] {a.coef}")
        return

    res = run_verify(a.fcst, a.obs, a.stn, a.coef)
    df = res["table"]
    met = lead_metrics(df)
    print(f"[예보] 격자 {res['check']['location']}, 발표 {res['check'].get('issues')}회, "
          f"누락 {len(res['check'].get('missing_issues', []))}회")
    print(f"[Rs 계수] {res['coef']['source']}")
    print(met.round(3).to_string(index=False))
    _, h1 = h1_table(res["obs"], res["meta"]["lat"], res["meta"]["elev"], res["coef"], res["h1_start"], res["h1_end"],
                     res["meta"].get("anem", 10.0))
    for m, (s, _) in h1.items():
        print(f"[H1 {m}] {res['h1_start']:%Y-%m-%d}~{res['h1_end']:%Y-%m-%d} RMSE {s['RMSE']:.3f}, 합계오차 {s['SUMERR']:+.1%}")
    for rn, v in h2_verdict(met).items():
        print(f"[H2] {rn}: {'통과' if v['pass_'] else '미달'} (최소 개선율 {v['min_skill']:.0%}, D+1 RMSE {v['rmse_d1']:.2f})")
    out = a.out or f"output/fcst_verify({a.stn})_{df.target.min():%Y%m%d}_{df.target.max():%Y%m%d}.xlsx"
    from fcst_report import build_verify_workbook
    saved = build_verify_workbook(res, out)
    print(f"[완료] {saved}")


if __name__ == "__main__":
    main()
