#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fcst_wb.py — 예보 물수지: 어제까지의 관측 물수지 + 예보 ETc − 예보 강수 → 관수 필요 예상일 (02-Cycle 4단계, G4)

[하는 일]
  1) 관측 물수지(observed_wb): 01-Cycle 워크북의 관측 ETo·강수 + Kc + 관수 기록 → 날마다 끝 고갈량 Dr
     - FAO-56 식(84·85·86·88), Ks는 전날 끝 고갈량 Dr,i-1 (fao56_core.wb_step, VALIDATION #15)
     - 관측이 빠진 날은 그날 아침 발표의 D+0 예보로 채우고 표시(운영 규칙). 채운 날에 기대는 상태(tainted)도 표시
  2) 예보 물수지(forecast_runs): 발표마다 출발 고갈량에서 대상일마다 wb_step (관수 없음 = '관수하지 않으면')
     - 아침(02시): 출발 = 관측 Dr(D−1 끝) → D+0~D+3
     - 저녁(17시): 출발 = 관측 Dr(D−1 끝) + 그날(D)의 아침 발표 D+0 예보로 하루 진행 → D+1~D+4
     - 예보 ETc = Kc × 예보 ETo(주 방법, 편향 보정 없음 — VALIDATION #10)
     - 세 경로(#18·#19): 중심 = 기대 강수(시각별 강수량 × 강수확률),
                        빠르면 = 비 없음 + ETc × (1 + r), 늦으면 = 예보 강수 전부 + ETc × (1 − r)
                        r = 월·발표·선행일별 예보 ETo 상대 오차(RMSE ÷ 관측 평균, 오차표 fcst_error_table.csv)
     - 비교 경로: 예보 강수 그대로, 비 무시(0), 관측 강수(완벽한 강수 예보), 기준선(어제 관측 ETc가 이어지고 비 없음)
  3) 검증(참값): 같은 출발(관측 Dr, D−1 끝)에서 관측 ETc·관측 강수로 관수 없이 진행한 고갈량
  4) 서비스 표시(#18): 처음 3개 대상일이 주 지표, 마지막 날(아침 D+3·저녁 D+4)은 참고. 관수 필요 예상일과 범위(빠르면·늦으면)

근거: docs/THEORY.md 6·9장, 설계: docs/ARCHITECTURE.md 7장, 판정 기록: docs/VALIDATION.md (G4)
"""
import math
import os

import numpy as np
import pandas as pd

from fao56_core import wb_step
from obs_daily import kc_series

MAIN_DAYS = 3                   # 서비스 주 지표: 처음 3개 대상일 (#18)
ERR_MIN_N = 15                  # 오차표: 월 칸 표본이 이보다 적으면 전체 월 값을 씀
ERR_DEFAULT_REL = 0.25          # 오차표에 해당 칸이 없을 때의 상대 오차(보수적 기본값)
PATHS = {                       # 예보 물수지 경로 (열 이름 Dr_<키>)
    "center": "중심: 기대 강수(강수량 × 강수확률)",
    "early": "빠르면: 비 없음 + ETc 오차만큼 많게",
    "late": "늦으면: 예보 강수 전부 + ETc 오차만큼 적게",
    "fcst": "비교: 예보 강수 그대로",
    "none": "비교: 비 무시(0)",
    "obs": "비교: 관측 강수(완벽한 강수 예보)",
    "pers": "기준선: 어제 관측 ETc 지속 + 비 없음",
}


# ── 입력 준비 ────────────────────────────────────────────────────────────
def prepare(fcst_paths, obs_path, stn, coef_path="rs_coef.csv", s4_coef_path=None, grid=None, bias=None):
    """과거(또는 운영) 단기예보 + 01-Cycle 워크북 → 예보표(모든 서비스 발표·대상일), 관측, 메타, Kc 인자, 점검 결과.
       예보 ETo·ETc는 cropwater_fcst.forecast_table과 같다(주 방법 S4, 없으면 S3). S4 계수는 s4_coef_path가 있으면 그 계수 고정,
       없으면 월 단위 교차검증(과거 자료 검증용). bias: 편향 보정표(apply_bias, 기본 없음 — #10)"""
    from cropwater_fcst import STATION_LON, coef_file, forecast_table, load_obs, load_s4_fixed
    from fcst_archive import check_archive, load_archive, service_table
    from obs_daily import kc_params
    from rs_model import load_coef
    obs, meta = load_obs(obs_path)
    arch = load_archive(fcst_paths)
    if grid and not arch.location:
        arch.location.add(grid)
    chk = check_archive(arch)
    st = service_table(arch, lat=meta["lat"], lon=STATION_LON.get(str(stn)))
    kp = kc_params(meta["settings"])
    s4_fixed = load_s4_fixed(stn, s4_coef_path) if s4_coef_path else None
    ft = forecast_table(st, obs, meta["lat"], meta["elev"], load_coef(stn, coef_file(coef_path)), kp, s4_fixed)
    ft.attrs = {}
    ft = apply_bias(ft, bias)
    return dict(ft=ft, obs=obs, meta=meta, kp=kp, check=chk, arch=arch, s4_fixed=s4_fixed, stn=str(stn),
                skipped_runs=list(getattr(arch, "skipped_runs", [])))


def apply_bias(ft, table=None):
    """편향 보정 자리(#10 결정: 보정하지 않음). table(run_name, lead_day, factor)을 주면 예보 ETo·ETc에 곱한다.
       기본은 None → 그대로. 보정을 쓰려면 한 해에서 구한 고정값이 다른 해로 옮겨 가지 않았다는 점(VALIDATION #10)을 먼저 확인"""
    if table is None or len(table) == 0:
        return ft
    f = ft.merge(table[["run_name", "lead_day", "factor"]], on=["run_name", "lead_day"], how="left")["factor"].fillna(1.0).values
    out = ft.copy()
    for c in ("ETo_main", "ETc_main"):
        out[c] = out[c] * f
    out["bias_factor"] = f
    return out


def fill_from_forecast(ft, col="ETo_main"):
    """관측이 빠진 날을 채울 값: 그날 아침 발표 D+0 예보 {date: (ETo, 기대 강수)}"""
    m = ft[(ft.run_name == "아침") & (ft.lead_day == 0) & ft[col].notna()]
    rain = m["rain_exp"].fillna(m["rain"]) if "rain_exp" in m else m["rain"]
    return {pd.Timestamp(t): (float(e), float(0.0 if pd.isna(r) else r)) for t, e, r in zip(m.target, m[col], rain)}


# ── 토양·관측 물수지 ─────────────────────────────────────────────────────
def soil_params(settings):
    """01-Cycle 설정 시트 값 → 토양·관수 파라미터 (TAW·RAW는 FAO-56 식82·83)"""
    def g(k, d):
        v = settings.get(k)
        return float(d if v is None or isinstance(v, str) else v)
    fc, wp = g("포장용수량 θFC (m³/m³)", 0.22), g("위조점 θWP (m³/m³)", 0.10)
    zr, p = g("근권심도 Zr (m)", 1.0), g("토양수분고갈계수 p", 0.5)
    taw = 1000 * (fc - wp) * zr
    return dict(fc=fc, wp=wp, zr=zr, p=p, taw=taw, raw=p * taw,
                dr0=g("초기 고갈량 Dr₀ (mm)", 0.0), ea=g("관수효율 Ea", 0.95))


def observed_wb(obs, kp, soil, irrig=None, fill=None, auto_irrigate=False, start=None, end=None):
    """관측 물수지 — 날마다 끝 고갈량.
       obs  : cropwater_fcst.load_obs 결과(date, ETo_obs, rain)
       irrig: {date: 공급 관수량 mm} (관수 기록). 근권에는 × Ea
       fill : {date: (ETo, 강수)} 관측 ETo가 없는 날 대신 쓸 값(운영: 그날 아침 발표 D+0 예보). 강수는 관측 행이 아예 없을 때만 씀.
              채울 값도 없으면 ETc 0
       auto_irrigate: 검증 시나리오 'FAO-56 기본 관수 규칙' — 전날 끝 Dr ≥ RAW면 그날 Dr만큼(순량) 관수해 포장용수량으로
       end  : 마지막 날(관측보다 뒤면 fill로 채움 — 운영에서 관측이 늦게 올 때)
       반환 열: date, Kc, ETo, ETo_src(관측/예보/없음), ETc, P, I_gross, I_net, Dr_prev, Ks, ETc_adj, DP, Dr, need,
               filled(관측이 빠져 채운 날), tainted(채운 날에 기대는 상태: 채운 뒤 Dr이 0이 되기 전까지), obs_row(ASOS 행이 있음 = 강수는 관측)"""
    o = obs.drop_duplicates("date").set_index("date").sort_index()
    idx = pd.date_range(pd.Timestamp(start) if start is not None else o.index.min(),
                        pd.Timestamp(end) if end is not None else o.index.max(), freq="D")
    eto = o["ETo_obs"].reindex(idx)
    rain = o["rain"].reindex(idx)
    have_row = pd.Series(idx.isin(o.index), index=idx)
    src = pd.Series(np.where(eto.notna(), "관측", "없음"), index=idx)
    if fill:
        fe = pd.Series({pd.Timestamp(k): v[0] for k, v in fill.items()}).reindex(idx)
        fr = pd.Series({pd.Timestamp(k): v[1] for k, v in fill.items()}).reindex(idx)
        m = eto.isna() & fe.notna()
        eto[m], src[m] = fe[m], "예보"
        mr = (~have_row) & fr.notna()
        rain[mr] = fr[mr]
    kc = kc_series(idx, kp)
    irrig = {pd.Timestamp(k): v for k, v in (irrig or {}).items()}
    taw, raw, ea = soil["taw"], soil["raw"], soil["ea"]
    rows, dr, tainted = [], soil["dr0"], False
    for d, k, e, p, s, row in zip(idx, kc, eto.values, rain.values, src.values, have_row.values):
        filled = (s != "관측") or not row
        P = 0.0 if pd.isna(p) else float(p)
        etc = 0.0 if (pd.isna(e) or pd.isna(k)) else float(k * e)
        ig = float(irrig.get(d, 0.0))
        inet = ig * ea
        if auto_irrigate and dr >= raw:
            inet += dr
            ig += dr / ea
        ks, etc_adj, dp, dr_new = wb_step(dr, P, etc, taw, raw, inet)
        tainted = filled or (tainted and dr_new > 0)
        rows.append(dict(date=d, Kc=k, ETo=e, ETo_src=s, ETc=etc, P=P, I_gross=ig, I_net=inet, Dr_prev=dr, Ks=ks,
                         ETc_adj=etc_adj, DP=dp, Dr=dr_new, need=dr_new >= raw, filled=filled, tainted=tainted, obs_row=bool(row)))
        dr = dr_new
    return pd.DataFrame(rows)


def path(dr0, etc, rain, soil):
    """출발 고갈량 dr0에서 관수 없이 날마다 wb_step → 끝 고갈량 목록"""
    out, dr = [], dr0
    for e, p in zip(etc, rain):
        dr = wb_step(dr, p, e, soil["taw"], soil["raw"])[3]
        out.append(dr)
    return out


def first_need(drs, raw, orders=None):
    """경로에서 처음 Dr ≥ RAW인 순번(orders가 없으면 1부터 센 위치). 없으면 None"""
    for i, v in zip(orders if orders is not None else range(1, len(drs) + 1), drs):
        if not pd.isna(v) and v >= raw:
            return int(i)
    return None


# ── 예보 ETo 오차표 (서비스의 '± 오차'와 관수 필요 예상일 범위) ─────────────────
ERR_COLS = ["stn", "year", "kind", "run_name", "lead_day", "month", "n", "rmse", "mbe", "obs_mean"]


def eto_error_rows(ft, stn, year, col="ETo_main"):
    """예보표 → 오차표 행: kind='day'(발표 × 선행일 × 월, month=0은 전체), kind='cum3'(처음 3개 대상일 합, 선행일 = 첫 선행일)"""
    from fcst_archive import SERVICE_RUNS
    d = ft[ft[col].notna() & ft.ETo_obs.notna()].copy()
    d["month"] = pd.to_datetime(d.target).dt.month
    rows = []

    def add(kind, rn, k, m, f, o):
        e = np.asarray(f, float) - np.asarray(o, float)
        rows.append(dict(stn=str(stn), year=int(year), kind=kind, run_name=rn, lead_day=int(k), month=int(m), n=len(e),
                         rmse=float(np.sqrt(np.mean(e ** 2))), mbe=float(e.mean()), obs_mean=float(np.mean(o))))
    for (rn, k), g in d.groupby(["run_name", "lead_day"]):
        add("day", rn, k, 0, g[col], g.ETo_obs)
        for m, gm in g.groupby("month"):
            add("day", rn, k, m, gm[col], gm.ETo_obs)
    first = {n: min(ls) for n, (_, ls) in SERVICE_RUNS.items()}
    c3 = []
    for (rn, run), g in d.groupby(["run_name", "run"]):
        k0 = first.get(rn, int(g.lead_day.min()))
        g3 = g[g.lead_day.isin([k0, k0 + 1, k0 + 2])]
        if len(g3) == 3:
            c3.append(dict(run_name=rn, k0=k0, month=int(g3.month.iloc[0]), f=g3[col].sum(), o=g3.ETo_obs.sum()))
    c3 = pd.DataFrame(c3)
    if len(c3):
        for (rn, k0), g in c3.groupby(["run_name", "k0"]):
            add("cum3", rn, k0, 0, g.f, g.o)
            for m, gm in g.groupby("month"):
                add("cum3", rn, k0, m, gm.f, gm.o)
    return pd.DataFrame(rows, columns=ERR_COLS)


def save_error_table(rows, path):
    """오차표 CSV에 (지점, 해) 행을 바꿔 넣는다. 다른 지점·해의 행은 그대로 둔다"""
    old = load_error_rows(path)
    if len(old):
        keys = set(zip(rows.stn.astype(str), rows.year.astype(int)))
        old = old[[(s, y) not in keys for s, y in zip(old.stn.astype(str), old.year.astype(int))]]
    out = pd.concat([old, rows], ignore_index=True) if len(old) else rows
    out = out.sort_values(["stn", "year", "kind", "run_name", "lead_day", "month"])
    out.to_csv(path, index=False, encoding="utf-8-sig", float_format="%.5f")
    return out


def load_error_rows(path):
    if not path or not os.path.exists(path):
        return pd.DataFrame(columns=ERR_COLS)
    t = pd.read_csv(path, encoding="utf-8-sig", dtype={"stn": str})
    return t[ERR_COLS]


def pooled_errors(rows, stn, exclude_year=None):
    """오차표 행 → 지점의 여러 해를 합친 표(칸별 MSE를 표본 수로 가중). 지점 행이 없으면 stn '0'(기본값) 행.
       exclude_year: 검증 연도를 뺀다(다른 해 오차로 채점해 공정하게). 반환 열: kind, run_name, lead_day, month, n, rmse, mbe, obs_mean, years, stn_src"""
    t = rows[rows.stn.astype(str) == str(stn)]
    src = str(stn)
    if not len(t):
        t, src = rows[rows.stn.astype(str) == "0"], "0"
    if not len(t):                                    # 지점·기본값 행이 모두 없으면 있는 지점을 모두 합쳐 씀(다른 지점 오차)
        t, src = rows, "전 지점(" + ",".join(sorted(set(rows.stn.astype(str)))) + ")"
    if exclude_year is not None:
        t = t[t.year.astype(int) != int(exclude_year)]
    if not len(t):
        return pd.DataFrame(columns=["kind", "run_name", "lead_day", "month", "n", "rmse", "mbe", "obs_mean", "years", "stn_src"])
    t = t.assign(sse=t.n * t.rmse ** 2, se=t.n * t.mbe, so=t.n * t.obs_mean)
    g = t.groupby(["kind", "run_name", "lead_day", "month"]).agg(n=("n", "sum"), sse=("sse", "sum"), se=("se", "sum"),
                                                                 so=("so", "sum"), years=("year", lambda s: ",".join(str(x) for x in sorted(set(s)))))
    g = g.reset_index()
    g["rmse"], g["mbe"], g["obs_mean"] = np.sqrt(g.sse / g.n), g.se / g.n, g.so / g.n
    g["stn_src"] = src
    return g[["kind", "run_name", "lead_day", "month", "n", "rmse", "mbe", "obs_mean", "years", "stn_src"]]


def err_lookup(err, kind, run_name, lead_day, month):
    """오차표 칸 → (RMSE, 상대 오차 = RMSE ÷ 관측 평균, 표본 수, '월별'/'전체 월'/'기본값').
       월 칸 표본이 ERR_MIN_N보다 적으면 전체 월 값"""
    if err is None or len(err) == 0:
        return np.nan, ERR_DEFAULT_REL, 0, "기본값"
    t = err[(err.kind == kind) & (err.run_name == run_name) & (err.lead_day == int(lead_day))]
    m = t[t.month == int(month)]
    if len(m) and int(m.n.iloc[0]) >= ERR_MIN_N:
        r, src = m.iloc[0], "월별"
    else:
        a = t[t.month == 0]
        if not len(a):
            return np.nan, ERR_DEFAULT_REL, 0, "기본값"
        r, src = a.iloc[0], "전체 월"
    return float(r.rmse), float(r.rmse / r.obs_mean) if r.obs_mean > 0 else ERR_DEFAULT_REL, int(r.n), src


# ── 발표별 예보 물수지 ───────────────────────────────────────────────────
def _run_inputs(g, pre, o, etc_col):
    """한 발표의 경로 입력 목록(저녁은 그날 D를 앞에 붙임)"""
    tgt = list(g.target)
    raw_rain = g["rain"].astype(float).fillna(0.0)
    exp_rain = (g["rain_exp"].astype(float).fillna(raw_rain) if "rain_exp" in g else raw_rain).fillna(0.0)
    x = dict(etc=[p["etc"] for p in pre] + list(g[etc_col].astype(float)),
             rain=[p["rain"] for p in pre] + list(raw_rain),
             rain_exp=[p["rain_exp"] for p in pre] + list(exp_rain),
             rel=[p["rel"] for p in pre] + list(g["rel"]),
             etc_t=[p["etc_t"] for p in pre] + [float(o.loc[t, "ETc"]) if t in o.index else np.nan for t in tgt],
             rain_t=[p["rain_t"] for p in pre] + [float(o.loc[t, "P"]) if t in o.index else np.nan for t in tgt],
             filled=[p["filled"] for p in pre] + [bool(o.loc[t, "filled"]) if t in o.index else True for t in tgt],
             etc_p=[p["etc_p"] for p in pre] + list(g["ETc_pers"].astype(float)),
             irr=[p.get("irr", 0.0) for p in pre] + [0.0] * len(tgt),
             known=[bool(p.get("rain_known", False)) for p in pre] + [False] * len(tgt))
    return x


def run_paths(dr0, x, soil):
    """출발 고갈량과 입력 목록 → 경로별 끝 고갈량 목록 {PATHS 키 + 'true'}.
       x['irr'](순관수량, 먼저 진행하는 날에만 — 관측 물수지 노란 칸)은 모든 경로의 물 공급에 더한다(식85의 I, 한 칸 식에서 P와 같은 자리)"""
    z = lambda v: [0.0 if pd.isna(a) else a for a in v]
    n = len(x["etc"])
    irr = x.get("irr") or [0.0] * n
    known = x.get("known") or [False] * n             # 관측 강수를 아는 날(ASOS 행은 있고 ETo 입력만 빠짐): '빠르면'도 그 강수
    w = lambda rain: [a + b for a, b in zip(rain, irr)]
    return dict(center=path(dr0, x["etc"], w(x["rain_exp"]), soil),
                early=path(dr0, [e * (1 + r) for e, r in zip(x["etc"], x["rel"])],
                           w([p if k else 0.0 for p, k in zip(x["rain"], known)]), soil),
                late=path(dr0, [e * max(1 - r, 0.0) for e, r in zip(x["etc"], x["rel"])], w(x["rain"]), soil),
                fcst=path(dr0, x["etc"], w(x["rain"]), soil),
                none=path(dr0, x["etc"], w([0.0] * n), soil),
                obs=path(dr0, x["etc"], w(z(x["rain_t"])), soil),
                pers=path(dr0, z(x["etc_p"]), w([0.0] * n), soil),
                true=path(dr0, z(x["etc_t"]), w(z(x["rain_t"])), soil))


def _pre_day(m, o, day, etc_col, rain_obs=None, irr=0.0):
    """관측 대신 그날 아침 발표 D+0 예보로 하루를 먼저 진행할 입력(세 경로 범위 포함). m: 그날 아침 D+0 예보 행, o: 관측 물수지(날짜 색인)
       rain_obs: 그날 관측 강수를 이미 알면(ASOS 행은 있고 ETo 입력만 빠짐) 세 경로 모두 그 값. irr: 그날 순관수량(관측 물수지 노란 칸)"""
    rain0 = 0.0 if pd.isna(m["rain"]) else float(m["rain"])
    rexp = float(m["rain_exp"]) if "rain_exp" in m and pd.notna(m["rain_exp"]) else rain0
    known = rain_obs is not None and not pd.isna(rain_obs)
    if known:
        rain0 = rexp = float(rain_obs)
    has = day in o.index
    return dict(etc=float(m[etc_col]), rain=rain0, rain_exp=rexp, rel=float(m["rel"]),
                etc_t=float(o.loc[day, "ETc"]) if has else np.nan, rain_t=float(o.loc[day, "P"]) if has else np.nan,
                filled=bool(o.loc[day, "filled"]) if has else True,
                etc_p=float(m["ETc_pers"]) if "ETc_pers" in m and pd.notna(m["ETc_pers"]) else np.nan,
                irr=float(irr), rain_known=known, date=pd.Timestamp(day), Kc=float(m["Kc"]), ETo=float(m["ETo_main"]),
                pop_max=m.get("pop_max", np.nan))


def forecast_runs(ft, owb, soil, err=None, etc_col="ETc_main", morning_obs_lag=None):
    """발표 × 대상일 예보 물수지와 참값.
       ft : forecast_table 결과(모든 행). 열 run_name, run, lead_day, target, Kc, ETc_main, rain, rain_exp, pop_max, ETc_pers
       owb: observed_wb 결과(출발 고갈량과 참값의 관측 ETc·강수)
       err: pooled_errors 결과(상대 오차 r). 없으면 r = ERR_DEFAULT_REL
       morning_obs_lag: (G5) 아침 02:10에 전날(D−1) ASOS 일자료가 아직 없다고 볼 때의 출발. 참값은 그대로 관측 Dr(D−1 끝)
            'fill' : 관측 Dr(D−2 끝)에서 D−1을 그날 아침 발표 D+0 예보(ETo, 기대 강수)로 채워 관측처럼 한 걸음 → 세 경로가 같은 출발
                     (G4에서 구현한 observed_wb의 fill과 같은 값)
            'range': D−1을 저녁 발표의 '그날'처럼 예보 하루로 먼저 진행(중심 = 기대 강수, 빠르면 = 비 없음·ETc × (1 + r),
                     늦으면 = 예보 강수 전부·ETc × (1 − r)) — 중심은 'fill'과 같고 범위가 D−1의 불확실성을 담음
            None   : 관측이 있다고 봄(검증 기본)
       반환: 행 = 발표 × 대상일. 경로 Dr_<PATHS 키>, 참값 Dr_true, 출발 고갈량 dr_start(중심 경로의 출발)·dr_start_true,
            유효 여부 ok와 제외 사유 reason"""
    lag = {True: "fill", False: None}.get(morning_obs_lag, morning_obs_lag)
    o = owb.set_index("date")
    lag_fill = fill_from_forecast(ft) if lag == "fill" else {}
    d = ft.copy()
    d["month"] = pd.to_datetime(d.target).dt.month
    d["rel"] = [err_lookup(err, "day", rn, k, m)[1] for rn, k, m in zip(d.run_name, d.lead_day, d.month)]
    mor0 = d[(d.run_name == "아침") & (d.lead_day == 0)].set_index("target")
    out = []
    for (rn, run), g in d.groupby(["run_name", "run"], sort=False):
        g = g.sort_values("lead_day")
        D = pd.Timestamp(run).normalize()
        prev = D - pd.Timedelta(days=1)
        base = prev                           # 출발 관측일(이날 끝 관측 Dr에서 출발)
        reason = ""
        # 출발 고갈량이 관측 결측일(예보로 채움)에 기대더라도 예보·참값이 같은 출발에서 시작하므로 비교는 공정하다 → 빼지 않음.
        # 참값이 관측 입력을 써야 하므로 대상일(저녁은 그날 D 포함)에 채운 날이 있으면 그 행부터 뺀다(아래 ok)
        if prev not in o.index:
            reason = "출발일 관측 물수지 없음"
        elif g[etc_col].isna().any():
            reason = "예보 ETc 없음(주 방법 입력 결측)"
        pre = []                              # 저녁: 그날(D) 하루를 먼저 진행 — 그날 아침 발표 D+0 예보
        if not reason and rn == "저녁":
            if D not in mor0.index or pd.isna(mor0.loc[D, etc_col]):
                reason = "그날 아침 발표 D+0 예보 없음"
            elif D not in o.index:
                reason = "그날 관측 물수지 없음"
            else:
                pre = [_pre_day(mor0.loc[D], o, D, etc_col)]
        if lag == "range" and rn == "아침" and not reason:
            base = prev - pd.Timedelta(days=1)
            if base not in o.index or prev not in mor0.index or pd.isna(mor0.loc[prev, etc_col]):
                reason = "관측 지연 계산: 전전날 관측 물수지 또는 전날 아침 D+0 예보 없음"
            else:
                pre = [_pre_day(mor0.loc[prev], o, prev, etc_col)]
        dr_obs0 = float(o.loc[base, "Dr"]) if base in o.index else np.nan
        dr_f0 = dr_obs0                        # 예보 경로의 출발(관측 지연 'fill'이 아니면 참값과 같음)
        if lag == "fill" and rn == "아침" and not reason:
            p2 = prev - pd.Timedelta(days=1)
            if p2 not in o.index or prev not in lag_fill:
                reason = "관측 지연 계산: 전전날 관측 물수지 또는 전날 아침 D+0 예보 없음"
            else:
                e_, r_ = lag_fill[prev]
                dr_f0 = wb_step(float(o.loc[p2, "Dr"]), r_, float(o.loc[prev, "Kc"]) * e_, soil["taw"], soil["raw"])[3]
        x = _run_inputs(g, pre, o, etc_col)
        n0 = len(pre)
        res = run_paths(dr_f0, x, soil) if not reason else {k: [np.nan] * len(x["etc"]) for k in list(PATHS) + ["true"]}
        if not reason and dr_f0 != dr_obs0:
            res["true"] = run_paths(dr_obs0, x, soil)["true"]
        start = {k: (v[n0 - 1] if n0 else (dr_obs0 if k == "true" else dr_f0)) for k, v in res.items()}
        for i, (_, r) in enumerate(g.iterrows()):
            j = n0 + i
            ok_true = not reason and not any(x["filled"][:j + 1])
            row = dict(run_name=rn, run=r["run"], run_date=D, lead_day=int(r["lead_day"]), target=r["target"],
                       order=i + 1, main=(i + 1) <= MAIN_DAYS, month=int(r["month"]), Kc=r["Kc"],
                       ETc_fcst=x["etc"][j], rel=x["rel"][j], rain_fcst=x["rain"][j], rain_exp=x["rain_exp"][j],
                       pop_max=r.get("pop_max", np.nan), ETc_true=x["etc_t"][j], rain_true=x["rain_t"][j],
                       ETc_pers=x["etc_p"][j], dr_obs_prev=dr_obs0, dr_start=start["center"], dr_start_true=start["true"],
                       pre_etc=pre[0]["etc"] if pre else np.nan, pre_rain=pre[0]["rain"] if pre else np.nan,
                       pre_rain_exp=pre[0]["rain_exp"] if pre else np.nan, pre_rel=pre[0]["rel"] if pre else np.nan,
                       pre_etc_t=pre[0]["etc_t"] if pre else np.nan, pre_rain_t=pre[0]["rain_t"] if pre else np.nan,
                       pre_etc_p=pre[0]["etc_p"] if pre else np.nan,
                       ok=bool(ok_true), reason=reason or ("" if ok_true else "예보 기간에 관측 결측일(예보로 채움)"))
            for k in list(PATHS) + ["true"]:
                row[f"Dr_{k}"] = res[k][j]
            out.append(row)
    return pd.DataFrame(out)


# ── 검증 지표 ────────────────────────────────────────────────────────────
WB_COLS = ("Dr_center", "Dr_fcst", "Dr_none", "Dr_obs", "Dr_pers")


def wb_lead_metrics(runs, cols=WB_COLS):
    """발표 × 순번(1~4)별 예상 Dr 오차: RMSE·MBE, 기준선(Dr_pers) 대비 개선율"""
    d = runs[runs.ok]
    rows = []
    for (rn, k), g in d.groupby(["run_name", "order"], sort=False):
        r = dict(run_name=rn, order=int(k), lead_day=int(g.lead_day.iloc[0]), n=len(g))
        for c in cols:
            e = g[c] - g.Dr_true
            r[f"RMSE_{c}"], r[f"MBE_{c}"] = math.sqrt((e ** 2).mean()), e.mean()
        for c in cols:
            r[f"skill_{c}"] = 1 - r[f"RMSE_{c}"] / r["RMSE_Dr_pers"] if r.get("RMSE_Dr_pers", 0) > 0 else np.nan
        rows.append(r)
    return pd.DataFrame(rows).sort_values(["run_name", "order"], key=lambda s: s.map({"아침": 0, "저녁": 1})
                                          if s.name == "run_name" else s, ignore_index=True)


def wb_month_metrics(runs, col="Dr_center", days=MAIN_DAYS):
    """대상일 월별 예상 Dr 오차(주 지표 기간), 예보·관측 강수 평균"""
    d = runs[runs.ok & (runs.order <= days)]
    rows = []
    for (rn, m), g in d.groupby(["run_name", "month"], sort=False):
        e = g[col] - g.Dr_true
        rows.append(dict(run_name=rn, month=int(m), n=len(g), RMSE=math.sqrt((e ** 2).mean()), MBE=e.mean(),
                         RMSE_obsrain=math.sqrt(((g.Dr_obs - g.Dr_true) ** 2).mean()),
                         rain_fcst=g.rain_fcst.mean(), rain_exp=g.rain_exp.mean(), rain_obs=g.rain_true.mean()))
    return pd.DataFrame(rows)


def need_contingency(runs, col, raw):
    """대상일별 '관수 필요(Dr ≥ RAW)' 판정 적중: 발표 × 순번별 적중·놓침·헛경보·정상, POD·FAR·CSI"""
    d = runs[runs.ok]
    rows = []
    for (rn, k), g in d.groupby(["run_name", "order"], sort=False):
        f, t = g[col] >= raw, g.Dr_true >= raw
        h, m, fa, cn = int((f & t).sum()), int((~f & t).sum()), int((f & ~t).sum()), int((~f & ~t).sum())
        rows.append(dict(run_name=rn, order=int(k), lead_day=int(g.lead_day.iloc[0]), n=len(g), hit=h, miss=m,
                         false_alarm=fa, correct_neg=cn, POD=h / (h + m) if h + m else np.nan,
                         FAR=fa / (h + fa) if h + fa else np.nan, CSI=h / (h + m + fa) if h + m + fa else np.nan,
                         accuracy=(h + cn) / len(g)))
    return pd.DataFrame(rows)


NEED_CATS = ["같은 날", "하루 빠름", "하루 늦음", "이틀 이상 빠름", "이틀 이상 늦음", "예보만 필요(헛경보)",
             "실제만 필요(놓침)", "둘 다 기간 안 필요 없음"]


def need_category(f, t):
    """예보·참값의 관수 필요 순번(None = 기간 안 없음) → 판정"""
    if f is None and t is None:
        return "둘 다 기간 안 필요 없음"
    if f is None:
        return "실제만 필요(놓침)"
    if t is None:
        return "예보만 필요(헛경보)"
    return {0: "같은 날", -1: "하루 빠름", 1: "하루 늦음"}.get(f - t, "이틀 이상 빠름" if f < t else "이틀 이상 늦음")


def first_need_eval(runs, raw, col="Dr_center", days=MAIN_DAYS):
    """발표마다 관수 필요 예상일(처음 Dr ≥ RAW인 대상일)과 범위(빠르면·늦으면) — 출발 고갈량이 RAW 미만인 발표만.
       days: 볼 대상일 수(주 지표 3, 참고 포함 4)"""
    d = runs[runs.ok & (runs.order <= days)]
    rows = []
    for (rn, run), g in d.groupby(["run_name", "run"], sort=False):
        g = g.sort_values("order")
        if len(g) < min(days, 3) or g.dr_start.iloc[0] >= raw or g.dr_start_true.iloc[0] >= raw:
            continue
        od = list(g.order)
        f, t = first_need(g[col], raw, od), first_need(g.Dr_true, raw, od)
        e, l = first_need(g.Dr_early, raw, od), first_need(g.Dr_late, raw, od)
        inf = math.inf
        lo, hi, tt = (e if e is not None else inf), (l if l is not None else inf), (t if t is not None else inf)
        rows.append(dict(run_name=rn, run=run, n_days=len(g), dr_start=g.dr_start.iloc[0], fcst=f, true=t, early=e, late=l,
                         category=need_category(f, t), covered=bool(lo <= tt <= hi),
                         true_before_early=bool(tt < lo), true_after_late=bool(tt > hi),
                         event=(f is not None) or (t is not None) or (e is not None)))
    return pd.DataFrame(rows)


def first_need_summary(fe):
    """first_need_eval 표 → 발표별 판정 개수·비율과 범위 적중"""
    rows = []
    for rn, g in fe.groupby("run_name", sort=False):
        r = dict(run_name=rn, n=len(g))
        for c in NEED_CATS:
            r[c] = int((g.category == c).sum())
        ev = g[(g.fcst.notna()) | (g.true.notna())]
        r["n_event"] = len(ev)
        r["within1"] = float(ev.category.isin(["같은 날", "하루 빠름", "하루 늦음"]).mean()) if len(ev) else np.nan
        r["late_or_miss"] = float(ev.category.isin(["하루 늦음", "이틀 이상 늦음", "실제만 필요(놓침)"]).mean()) if len(ev) else np.nan
        er = g[g.event]
        r["n_range_event"] = len(er)
        r["coverage"] = float(er.covered.mean()) if len(er) else np.nan
        r["true_before_early"] = int(er.true_before_early.sum())
        r["true_after_late"] = int(er.true_after_late.sum())
        r["early_only"] = int((g.early.notna() & g.fcst.isna()).sum())
        r["early_only_true"] = int((g.early.notna() & g.fcst.isna() & g.true.notna()).sum())
        rows.append(r)
    return pd.DataFrame(rows)


def threshold_sensitivity(runs, thresholds, days=MAIN_DAYS, col="Dr_center"):
    """관수 기준 고갈량(RAW 대신 더 낮은 값으로 자주 관수하는 농가)별 관수 필요 예상일 적중 — 사건 수를 늘려 시점 정확도를 본다.
       경로(Ks 포함)는 그대로 두고 판정 기준만 바꾼다. 반환: 기준 × 발표별 사건 수·같은 날·±1일·늦음/놓침·범위 적중"""
    rows = []
    for th in thresholds:
        fe = first_need_eval(runs, th, col=col, days=days)
        for rn, g in list(fe.groupby("run_name", sort=False)) + [("전체", fe)]:
            ev = g[(g.fcst.notna()) | (g.true.notna())]
            er = g[g.event]
            rows.append(dict(threshold=th, run_name=rn, n_runs=len(g), n_event=len(ev),
                             same=int((ev.category == "같은 날").sum()),
                             within1=float(ev.category.isin(["같은 날", "하루 빠름", "하루 늦음"]).mean()) if len(ev) else np.nan,
                             early=int(ev.category.isin(["하루 빠름", "이틀 이상 빠름", "예보만 필요(헛경보)"]).sum()),
                             late_or_miss=int(ev.category.isin(["하루 늦음", "이틀 이상 늦음", "실제만 필요(놓침)"]).sum()),
                             n_range_event=len(er), coverage=float(er.covered.mean()) if len(er) else np.nan,
                             true_before_early=int(er.true_before_early.sum())))
    return pd.DataFrame(rows)


def rain_verification(runs):
    """예보 강수 검증(대상일 일합계): 발표 × 순번별 평균(예보·기대·관측), 예보/관측 비, 비 온 날(≥ 1 mm) 적중·헛경보"""
    d = runs[runs.ok]
    rows = []
    for (rn, k), g in d.groupby(["run_name", "order"], sort=False):
        f, o = g.rain_fcst >= 1, g.rain_true >= 1
        rows.append(dict(run_name=rn, order=int(k), lead_day=int(g.lead_day.iloc[0]), n=len(g), fcst=g.rain_fcst.mean(),
                         exp=g.rain_exp.mean(), obs=g.rain_true.mean(),
                         ratio=g.rain_fcst.sum() / g.rain_true.sum() if g.rain_true.sum() > 0 else np.nan,
                         ratio_exp=g.rain_exp.sum() / g.rain_true.sum() if g.rain_true.sum() > 0 else np.nan,
                         f_days=float(f.mean()), o_days=float(o.mean()),
                         POD=float((f & o).sum() / o.sum()) if o.sum() else np.nan,
                         FAR=float((f & ~o).sum() / f.sum()) if f.sum() else np.nan))
    return pd.DataFrame(rows)


def run_wbverify(fcst_paths, obs_path, stn, coef_path="rs_coef.csv", s4_coef_path=None, grid=None,
                 err_path="fcst_error_table.csv", irrig=None, auto_irrigate=False, thresholds=(30, 40, 50, 60),
                 morning_obs_lag=False):
    """G4 검증 전체: 예보표 → 관측 물수지 → 발표별 예보 물수지(범위는 검증 연도를 뺀 오차표) → 지표.
       morning_obs_lag: (G5) 아침 발표 때 전날 관측이 없다고 보고 출발을 예보로 채운 값으로(forecast_runs)"""
    p = prepare(fcst_paths, obs_path, stn, coef_path, s4_coef_path, grid)
    return wbverify_from(p, stn, err_path, irrig, auto_irrigate, thresholds, morning_obs_lag)


def wbverify_from(p, stn, err_path="fcst_error_table.csv", irrig=None, auto_irrigate=False, thresholds=(30, 40, 50, 60),
                  morning_obs_lag=False):
    """prepare 결과로 G4 검증 계산"""
    from cropwater_fcst import coef_file
    ft = p["ft"]
    year = int(pd.to_datetime(ft.target).min().year)
    soil = soil_params(p["meta"]["settings"])
    rows = load_error_rows(coef_file(err_path)) if err_path else pd.DataFrame(columns=ERR_COLS)
    err = pooled_errors(rows, stn, exclude_year=year)
    owb = observed_wb(p["obs"], p["kp"], soil, irrig=irrig, fill=fill_from_forecast(ft), auto_irrigate=auto_irrigate)
    runs = forecast_runs(ft, owb, soil, err, morning_obs_lag=morning_obs_lag)
    raw = soil["raw"]
    res = dict(p, stn=str(stn), soil=soil, year=year, err=err, err_rows=rows, err_path=err_path, owb=owb, runs=runs,
               auto_irrigate=auto_irrigate, irrig=irrig or {}, thresholds=tuple(thresholds), morning_obs_lag=morning_obs_lag)
    res["lead"] = wb_lead_metrics(runs)
    res["month"] = wb_month_metrics(runs)
    res["cont"] = need_contingency(runs, "Dr_center", raw)
    res["fe3"], res["fe4"] = first_need_eval(runs, raw, days=3), first_need_eval(runs, raw, days=4)
    res["fe3_sum"], res["fe4_sum"] = first_need_summary(res["fe3"]), first_need_summary(res["fe4"])
    res["sens"] = threshold_sensitivity(runs, thresholds)
    res["rain"] = rain_verification(runs)
    res["findings"] = wb_findings(res)
    return res


def wb_findings(res):
    """G4 검증 해석 문장"""
    out = []
    lead, soil = res["lead"], res["soil"]
    m = lead[lead.order <= MAIN_DAYS]
    out.append(f"① 예보 물수지(중심: 기대 강수) 3일 예상 Dr 오차 RMSE {m.RMSE_Dr_center.min():.1f}~{m.RMSE_Dr_center.max():.1f} mm, "
               f"기준선(어제 ETc 지속·비 없음) 대비 {m.skill_Dr_center.min():.0%}~{m.skill_Dr_center.max():.0%} 개선")
    out.append(f"② 강수를 관측값으로 바꾸면(완벽한 강수 예보) 오차가 {m.RMSE_Dr_obs.min():.1f}~{m.RMSE_Dr_obs.max():.1f} mm로 줄어듦 "
               f"→ 예보 물수지 오차의 대부분은 강수 예보에서 옴(ETc 예보 오차는 작음)")
    rv = res["rain"]
    r12 = rv[rv.order <= 2]
    if len(r12) and r12.obs.sum() > 0:
        out.append(f"③ 예보 강수량은 관측보다 많음: 처음 이틀 예보/관측 비 {r12.ratio.min():.2f}~{r12.ratio.max():.2f} "
                   f"(기대 강수 {r12.ratio_exp.min():.2f}~{r12.ratio_exp.max():.2f})")
    mm = res["month"]
    if len(mm):
        w = mm.sort_values("RMSE", ascending=False).iloc[0]
        out.append(f"④ 가장 약한 달: {int(w.month)}월 {w.run_name} RMSE {w.RMSE:.1f} mm (예보 강수 {w.rain_fcst:.1f}·관측 {w.rain_obs:.1f} mm/일)")
    for mark, (rn, s) in zip("⑤⑥", res["fe3_sum"].set_index("run_name").iterrows()):
        out.append(f"{mark} {rn} 관수 필요 예상일(3일, RAW {soil['raw']:.0f} mm): 사건 {int(s.n_event)}회 중 같은 날 {int(s['같은 날'])}회, "
                   f"늦음·놓침 {int(s['하루 늦음'] + s['이틀 이상 늦음'] + s['실제만 필요(놓침)'])}회, "
                   f"범위(빠르면~늦으면) 적중 {s.coverage:.0%} ({int(s.n_range_event)}회), 실제가 '빠르면'보다 앞선 경우 {int(s.true_before_early)}회")
    sens = res.get("sens")
    if sens is not None and len(sens):
        a = sens[sens.run_name == "전체"]
        out.append(f"⑦ 판정 기준을 {', '.join(str(int(t)) for t in a.threshold)} mm로 바꿔도 범위 적중 {a.coverage.min():.0%}~{a.coverage.max():.0%}, "
                   f"실제가 '빠르면'보다 앞선 경우 {int(a.true_before_early.sum())}회")
    return out


# ── 한 발표의 서비스 전망 ─────────────────────────────────────────────────
def day_forecast(ft, day, etc_col="ETc_main"):
    """하루(day)를 먼저 진행할 예보 행: 그날 아침 발표 D+0 → 없으면 그날을 대상으로 한 가장 최근 서비스 발표(그날 02시까지)"""
    day = pd.Timestamp(day)
    t = pd.to_datetime(ft.target)
    m = ft[(ft.run_name == "아침") & (ft.lead_day == 0) & (t == day) & ft[etc_col].notna()]
    if len(m):
        return m.iloc[0]
    m = ft[(t == day) & ft[etc_col].notna() & (pd.to_datetime(ft.run) <= day + pd.Timedelta(hours=2))]
    return m.sort_values("run").iloc[-1] if len(m) else None


def service_outlook(ft, owb, soil, err, run_name, run, etc_col="ETc_main", obs_last=None):
    """한 서비스 발표의 전망(서비스 엑셀·CLI 공용).
       obs_last: 관측이 있는 마지막 날(없으면 D−1). 그 뒤 D−1까지는 관측이 아직 없는 날 → 아침 D+0 예보로 '먼저 진행하는 날'
                 (저녁 발표의 그날 D와 같이 세 경로 범위를 담음 — VALIDATION G5 관측 지연, #12). 그날 관측 강수(ASOS 행)나
                 관수 기록(관측 물수지 노란 칸)이 있으면 그 값을 씀
       반환: dict(days=대상일 표[선행일, 날짜, 주/참고, ETo, Kc, ETc, ETc 오차, 강수 예보·기대 강수·강수확률, 세 경로 Dr, 상태],
                 start = 출발 고갈량(중심, 먼저 진행한 날 뒤), dr_obs_prev = 관측 마지막 날 끝 Dr, prev = 관측 마지막 날,
                 pre = 먼저 진행한 날 목록(관측 지연 날 'lag' + 저녁의 그날 'today'), lag_days,
                 need = {center, early, late: 순번 또는 None}, cum3 = (예보 ETc 3일 합, 오차), advice = (순관수량, 공급 관수량))"""
    d = ft[(ft.run_name == run_name) & (ft.run == pd.Timestamp(run))].sort_values("lead_day").copy()
    if d.empty:
        raise ValueError(f"{run_name} {run} 발표가 예보표에 없습니다")
    if d[etc_col].isna().any():
        raise ValueError(f"{run_name} {run} 발표의 예보 ETc가 비어 있습니다(필수 요소 결측)")
    d["month"] = pd.to_datetime(d.target).dt.month
    # 운영(#12): 직전 발표로 대신한 행은 대체 발표 종류·선행일의 오차(err_name·err_lead, ops_service_table). 검증 표에는 없음 → 서비스 발표 칸
    en = d["err_name"] if "err_name" in d else pd.Series(run_name, index=d.index)
    el = d["err_lead"] if "err_lead" in d else d["lead_day"]
    e = [err_lookup(err, "day", n_, k, m) for n_, k, m in zip(en, el, d.month)]
    d["eto_rmse"], d["rel"], d["err_n"], d["err_src"] = [x[0] for x in e], [x[1] for x in e], [x[2] for x in e], [x[3] for x in e]
    d["etc_err"] = d[etc_col] * d.rel                  # ± 오차 = 예보 ETc × 상대 오차 r (범위 경로와 같은 크기)
    D = pd.Timestamp(run).normalize()
    o = owb.set_index("date")
    prev = pd.Timestamp(obs_last) if obs_last is not None else D - pd.Timedelta(days=1)
    if prev not in o.index:
        raise ValueError(f"관측 물수지에 {prev:%Y-%m-%d}(출발일)가 없습니다 — 01-Cycle 워크북 기간 확인")
    lag_days = list(pd.date_range(prev + pd.Timedelta(days=1), D - pd.Timedelta(days=1)))
    pre = []
    for day, kind in [(x, "lag") for x in lag_days] + ([(D, "today")] if run_name == "저녁" else []):
        m = day_forecast(ft, day, etc_col)
        if m is None:
            raise ValueError(f"{day:%Y-%m-%d}을 진행할 예보가 없습니다(아침 발표 D+0)" if kind == "lag" else
                             f"저녁 발표는 그날({D:%Y-%m-%d}) 아침 발표의 D+0 예보가 필요합니다")
        m = m.copy()
        m["rel"] = err_lookup(err, "day", "아침", 0, day.month)[1]
        has = day in o.index
        rain_obs = float(o.loc[day, "P"]) if (kind == "lag" and has and bool(o.loc[day].get("obs_row", False))) else None
        irr = float(o.loc[day, "I_net"]) if (kind == "lag" and has) else 0.0
        pre.append(dict(_pre_day(m, o, day, etc_col, rain_obs, irr), kind=kind))
    x = _run_inputs(d, pre, o, etc_col)
    dr0 = float(o.loc[prev, "Dr"])
    res = run_paths(dr0, x, soil)
    n0 = len(pre)
    for k in ("center", "early", "late", "fcst"):
        d[f"Dr_{k}"] = res[k][n0:]
    for i, p in enumerate(pre):
        p.update({f"Dr_{k}": res[k][i] for k in ("center", "early", "late")})
    d["order"] = range(1, len(d) + 1)
    d["main"] = d.order <= MAIN_DAYS
    raw = soil["raw"]
    starts = {k: (res[k][n0 - 1] if n0 else dr0) for k in ("center", "early", "late")}
    start = starts["center"]
    need = {k: (0 if starts[k] >= raw else first_need(res[k][n0:], raw)) for k in ("center", "early", "late")}
    main = d[d.main]
    c3 = err_lookup(err, "cum3", run_name, int(main.lead_day.iloc[0]), int(main.month.iloc[0]))
    if "err_name" in main and bool(main.get("backup", pd.Series(False)).any()):
        cb = err_lookup(err, "cum3", str(main.err_name.iloc[0]), int(main.err_lead.iloc[0]), int(main.month.iloc[0]))
        c3 = cb if cb[3] != "기본값" else c3             # 대체 발표 칸이 오차표에 없으면 서비스 발표 칸(THEORY 9장 #12)
    s3 = float(main[etc_col].sum())
    cum3 = (s3, s3 * c3[1], c3[3], c3[1])                # (3일 합, ± 오차 = 합 × 3일 누적 상대 오차, 출처, 상대 오차)
    k_need = need["center"]
    dr_need = start if k_need == 0 else (float(d.Dr_center.iloc[k_need - 1]) if k_need else np.nan)
    advice = (dr_need, dr_need / soil["ea"]) if not pd.isna(dr_need) else (np.nan, np.nan)
    return dict(days=d, start=start, dr_obs_prev=dr0, prev=prev, pre=pre, lag_days=lag_days, need=need, cum3=cum3,
                advice=advice, run_name=run_name, run=pd.Timestamp(run))


def last_observed(owb, before):
    """관측 물수지에서 관측 ETo가 있는 마지막 날(before 전). 없으면 None"""
    m = owb[(owb.ETo_src == "관측") & (pd.to_datetime(owb.date) < pd.Timestamp(before))]
    return pd.Timestamp(m.date.max()) if len(m) else None


def recent_bias(ft, run, days=30, col="ETo_main"):
    """편향 점검(#10): 이 발표 이전의 발표 중 대상일이 최근 days일 안이고 관측이 있는 행 → 발표 × 선행일별 예보 − 관측"""
    run = pd.Timestamp(run)
    t = pd.to_datetime(ft.target)
    d = ft[(ft.run < run) & (t >= run.normalize() - pd.Timedelta(days=days)) & (t < run.normalize())
           & ft[col].notna() & ft.ETo_obs.notna()]
    rows = []
    for (rn, k), g in d.groupby(["run_name", "lead_day"]):
        rows.append(dict(run_name=rn, lead_day=int(k), n=len(g), fcst=g[col].mean(), obs=g.ETo_obs.mean(),
                         mbe=(g[col] - g.ETo_obs).mean(), pct=g[col].sum() / g.ETo_obs.sum() - 1))
    return pd.DataFrame(rows)


def latest_run(ft, etc_col="ETc_main"):
    """예보표에서 입력이 모두 있는 가장 최근 서비스 발표 (구분, 발표시각)"""
    ok = ft.groupby(["run_name", "run"])[etc_col].apply(lambda s: s.notna().all())
    ok = ok[ok]
    if ok.empty:
        raise ValueError("입력이 모두 있는 서비스 발표가 없습니다")
    rn, run = max(ok.index, key=lambda x: x[1])
    return rn, run


def service_prepare(fcst_paths, obs_path, stn, run=None, coef_path="rs_coef.csv", s4_coef_path="rs_sky_coef.csv",
                    grid=None, err_path="fcst_error_table.csv", irrig=None, recent_days=30):
    """서비스 엑셀 준비: 운영 S4 계수(고정)로 예보표 → 관측 물수지(출발일까지, 결측은 예보로 채움, 관수 기록 반영)
       → 한 발표의 전망(service_outlook), 편향 점검(recent_bias). run: '2026-05-15 02' 같은 발표시각(없으면 가장 최근 발표)"""
    p = prepare(fcst_paths, obs_path, stn, coef_path, s4_coef_path, grid)
    return service_from(p, stn, run, err_path, irrig, recent_days)


def service_from(p, stn, run=None, err_path="fcst_error_table.csv", irrig=None, recent_days=30):
    """prepare 결과로 한 발표의 서비스 전망 준비"""
    from cropwater_fcst import STATIONS, coef_file, main_method
    ft = p["ft"]
    if run is None:
        rn, run = latest_run(ft)
    else:
        run = pd.Timestamp(run)
        hit = ft[ft.run == run]
        if hit.empty:
            raise ValueError(f"{run} 발표가 예보표에 없습니다(서비스 발표는 02시·17시)")
        rn = hit.run_name.iloc[0]
    soil = soil_params(p["meta"]["settings"])
    rows = load_error_rows(coef_file(err_path)) if err_path else pd.DataFrame(columns=ERR_COLS)
    err = pooled_errors(rows, stn)
    D = pd.Timestamp(run).normalize()
    owb = observed_wb(p["obs"], p["kp"], soil, irrig=irrig, fill=fill_from_forecast(ft), end=D - pd.Timedelta(days=1))
    # 관측이 아직 없는 최근 날(보통 아침 02:10의 전날)은 관측 물수지에 예보로 채워 보이되(관수 기록 칸), 전망은 그날부터 범위를 담아
    # 예보 하루로 진행한다(VALIDATION G5: 채운 값을 관측처럼 쓰면 범위 적중이 80%·71%로 떨어짐)
    last = last_observed(owb, D)
    obs_last = last if (last is not None and last < D - pd.Timedelta(days=1)) else None
    ol = service_outlook(ft, owb, soil, err, rn, run, obs_last=obs_last)
    name, grid_std = STATIONS.get(str(stn), ("", ""))
    return dict(outlook=ol, owb=owb, soil=soil, err=err, err_path=err_path, meta=p["meta"], stn=str(stn), stn_name=name,
                grid="-".join(p["check"].get("location") or []) or grid_std, main=main_method(ft), irrig=irrig or {}, ft=ft,
                recent=recent_bias(ft, run, recent_days), recent_days=recent_days,
                obs_last=p["obs"].date.max(), check=p["check"])
