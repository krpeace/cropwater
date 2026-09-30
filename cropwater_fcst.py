#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cropwater_fcst.py — 기상청 단기예보 기반 ETo·ETc 예측과 검증 (02-Cycle 3단계, H2)

[하는 일]
  1) 과거 단기예보(포털 CSV 또는 OpenAPI 응답 CSV)를 서비스 발표(아침 02시 → 오늘~D+3, 저녁 17시 → 내일~D+4)별 일 입력으로 집계
     (fcst_archive.py)
  2) 예보 ETo = FAO-56 PM [식6]. 예보에 없는 Rs는 추정한다(rs_model.py)
     - S4(주 방법, 하늘상태 예보가 있을 때): 식(50)형 + 강수유무 + 낮 시간 구름많음·흐림 비율, 선행일별 계수
       검증에서는 월 단위 교차검증 계수(대상월을 뺀 나머지 달로 맞춤), 운영 계수는 rs_sky_coef.csv
       다른 해 독립 검증(--s4-coef)에서는 다른 해의 운영 계수를 그대로 적용(계수 고정)
     - S3(비교, 하늘상태가 없을 때 주 방법): 식(50)+강수유무 보정, 다른 해 관측으로 정한 계수(rs_coef.csv)
  3) 예보 ETc = Kc × 예보 ETo  (Kc는 01-Cycle 워크북 설정 시트와 같은 값)
  4) ASOS 관측 ETo(01-Cycle 규칙)와 비교: 선행시간별 RMSE·MBE·R², 지속성·7일평균 기준선, 3일 누적
  5) 검증 엑셀(라이브 수식) 출력

[실행]
  # Rs 계수 보정 (검증 연도와 다른 해의 관측으로)
  python cropwater_fcst.py calib --obs output/eto101_apple_20250101_20251231.xlsx --stn 101
  # H2 검증 엑셀 (하늘상태·강수확률 CSV를 같은 폴더에 넣으면 S4로 검증)
  python cropwater_fcst.py verify --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101
  # 다른 해 독립 검증 — S4는 다른 해 운영 계수 고정, S3 비교 계수도 다른 해 관측으로 (VALIDATION #13)
  python cropwater_fcst.py calib --obs output/eto101_apple_20260101_20260928.xlsx --stn 101 --coef rs_coef_2026.csv
  python cropwater_fcst.py verify --fcst data/fcst_101_2025 --obs output/eto101_apple_20250101_20251231.xlsx --stn 101 \
         --grid 73_134 --coef rs_coef_2026.csv --s4-coef rs_sky_coef.csv
  # 운영용 S4 계수(선행일별) — 한 생육기 전체의 예보·관측으로 맞춰 rs_sky_coef.csv에 기록
  python cropwater_fcst.py calib-sky --fcst data/fcst_101 --obs output/eto101_apple_20260101_20260928.xlsx --stn 101

[입력]
  --fcst : 과거 단기예보 CSV 폴더(또는 파일들). 필수 요소 TMX·TMN·TMP·REH·WSD·PCP, 선택 요소 SKY(하늘상태)·POP(강수확률)
           - 기상자료개방포털 '단기예보(격자)' CSV: 요소별 파일
           - 단기예보 조회서비스(OpenAPI) 응답 CSV(baseDate,baseTime,category,fcstDate,fcstTime,fcstValue,nx,ny): 한 파일에 모든 요소
           형식은 첫 줄로 자동 판별. 한 폴더에 섞어 넣어도 됨(같은 발표·시각은 나중에 읽은 파일 값)
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
from rs_model import S4_NAMES, S4_RAIN_DEFAULT, fit, fit_s4, load_coef, ra_rso, rs_s1, rs_s3, rs_s4, save_coef, save_sky_coef

FCST_ANEM = 10.0            # 단기예보 풍속(WSD) 기준 높이 (m)
KRS_FAO = 0.16              # FAO-56 식(50) 내륙 기본값
H2_SKILL_MIN = 0.30         # H2: 지속성 대비 RMSE 개선율 기준 (VALIDATION.md, 2026-09-29 확정)
H2_RMSE_D1_MAX = 1.0        # H2: D+1 RMSE 상한 (mm/일)
H2_LEADS = (1, 2, 3)        # H2 판정 선행일
H2_BOOT_N = 2000            # 판정 불확실성: 부트스트랩 반복 수
H2_BOOT_BLOCK = 7           # 블록 길이(일). 날씨가 며칠 이어져 예보 오차끼리 상관이 있으므로 하루 단위로 뽑지 않음
H2_BOOT_SEED = 20260929     # 재현용 난수 시드
# 대표 ASOS 지점과 단기예보 격자(nx_ny) — ARCHITECTURE.md 7장 표
STATIONS = {"101": ("춘천", "73_134"), "216": ("태백", "95_119"), "119": ("수원", "60_120"),
            "131": ("청주", "68_107"), "129": ("서산", "52_109"), "146": ("전주", "63_89"),
            "156": ("광주", "59_75"), "136": ("안동", "90_106"), "192": ("진주", "79_75"),
            "189": ("서귀포", "53_32")}
# 관측소 경도(°E) — 하늘상태를 낮 시간 일사 비중으로 가중할 때 태양시 보정에만 씀(0.1° 차이 = 0.4분)
STATION_LON = {"101": 127.7357, "216": 128.9893, "119": 126.9830, "131": 127.4407, "129": 126.4939,
               "146": 127.1172, "156": 126.8916, "136": 128.7073, "192": 128.0400, "189": 126.5653}
S4_RAIN = S4_RAIN_DEFAULT   # S4 강수 입력: 강수확률 하루 최대(pop_max, 0~1). #17 결정(2026-09-30). 이전(G3 재검증·#13)은 rain_flag
S4_FEATS = ("Tmax", "Tmin", S4_RAIN, "sky_cloudy", "sky_overcast", "Ra", "Rso")
S4_FIXED = "고정"      # 다른 해 운영 계수를 그대로 쓸 때의 계수 묶음 이름(SKY계수 시트 키 '고정|k')


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
def month_folds(target):
    """교차검증 묶음 = 대상일의 달. 마지막 달 대상일이 10일 미만이면(예: 7/1~7/4) 앞 달과 한 묶음"""
    t = pd.to_datetime(pd.Series(target))
    per = t.dt.to_period("M")
    days_last = t[per == per.max()].dt.normalize().nunique()
    return per.where(~((per == per.max()) & (days_last < 10)), per.max() - 1).astype(str)


def s4_cv(df, exclude=()):
    """S4 계수의 월 단위 교차검증(선행일별).
       각 행에는 그 행의 묶음(대상월)과 exclude 묶음을 뺀 나머지 행으로 맞춘 같은 선행일의 계수를 준다.
       반환: (행별 계수 n×5 배열 — 없으면 NaN, 계수표 DataFrame[fold, lead_day, n, a~e])"""
    fold = month_folds(df.target).values
    feat = df[list(S4_FEATS)].notna().all(axis=1).values
    fit_ok = feat & df["Rs_obs"].notna().values
    lead = df["lead_day"].values
    co = np.full((len(df), len(S4_NAMES)), np.nan)
    rows = []
    for f in sorted(set(fold[feat])):
        for k in sorted(set(lead[feat & (fold == f)])):
            tr = fit_ok & (lead == k) & (fold != f) & ~np.isin(fold, list(exclude))
            if tr.sum() < 20:
                continue
            t = df[tr]
            c = fit_s4(t.Tmax, t.Tmin, t[S4_RAIN], t.sky_cloudy, t.sky_overcast, t.Ra, t.Rs_obs)
            co[feat & (fold == f) & (lead == k)] = c
            rows.append(dict(fold=f, lead_day=int(k), n=int(tr.sum()), **dict(zip(S4_NAMES, c))))
    return co, pd.DataFrame(rows, columns=["fold", "lead_day", "n", *S4_NAMES])


def s4_fit_all(df):
    """운영용 S4 계수: 자료 전체로 선행일별 적합 → DataFrame[lead_day, a~e, n, fit_start, fit_end, rmse_rs]"""
    ok = df[list(S4_FEATS)].notna().all(axis=1) & df["Rs_obs"].notna()
    rows = []
    for k, t in df[ok].groupby("lead_day"):
        c = fit_s4(t.Tmax, t.Tmin, t[S4_RAIN], t.sky_cloudy, t.sky_overcast, t.Ra, t.Rs_obs)
        est = rs_s4(t.Tmax, t.Tmin, t[S4_RAIN], t.sky_cloudy, t.sky_overcast, t.Ra, t.Rso, c)
        rows.append(dict(lead_day=int(k), **dict(zip(S4_NAMES, c)), n=len(t),
                         fit_start=str(t.target.min().date()), fit_end=str(t.target.max().date()),
                         rmse_rs=float(np.sqrt(np.mean((est - t.Rs_obs.values) ** 2)))))
    return pd.DataFrame(rows)


def load_s4_fixed(stn, path):
    """다른 해 운영 S4 계수(rs_sky_coef.csv) → DataFrame[lead_day, a~e, n, fit_start, fit_end, note] (지점 행)"""
    t = pd.read_csv(coef_file(path), encoding="utf-8-sig", dtype={"stn": str})
    t = t[t["stn"].str.strip() == str(stn)]
    if t.empty:
        raise ValueError(f"{path}에 지점 {stn}의 S4 계수가 없습니다(calib-sky로 먼저 만드세요)")
    got = set(t["rain_input"].fillna("rain_flag")) if "rain_input" in t else {"rain_flag"}    # 열이 없으면 이전 형식
    if got != {S4_RAIN}:
        raise ValueError(f"{path}: S4 계수의 강수 입력이 {sorted(got)}입니다({S4_RAIN} 필요). calib-sky로 다시 만드세요")
    for c in ("n", "fit_start", "fit_end", "note"):
        if c not in t:
            t[c] = ""
    t = t.assign(lead_day=t.lead_day.astype(int))
    return t[["lead_day", *S4_NAMES, "n", "fit_start", "fit_end", "note"]].sort_values("lead_day").reset_index(drop=True)


def s4_is_fixed(df):
    """S4 계수가 고정(다른 해 운영 계수)인지 — 행별 계수 묶음 열로 판단"""
    return "s4_fold" in df and len(df) > 0 and bool((df["s4_fold"] == S4_FIXED).all())


def forecast_table(st, obs, lat, elev, coef, kp, s4_fixed=None):
    """서비스 표(fcst_archive.service_table) + 관측 → 예보 ETo(S4·S3·S1), 관측·기준선, Kc·ETc.
       주 방법(★): 하늘상태가 있으면 S4, 없으면 S3 → 열 Rs_main·ETo_main·ETc_main.
       S4 계수: s4_fixed(load_s4_fixed 결과)가 있으면 그 계수를 선행일별로 고정 적용(다른 해 독립 검증),
       없으면 월 단위 교차검증. 계수표는 df.attrs["s4_table"], 행별 묶음은 s4_fold 열(대상월 또는 '고정')."""
    df = st.copy()
    df["target"] = pd.to_datetime(df["target"])
    df["run_date"] = df["run"].dt.normalize()
    for c in ("rain_flag", "sky_cloudy", "sky_overcast", "pop", "pop_max", "rain_exp"):
        if c in df:
            df[c] = pd.to_numeric(df[c])
    rs3, ra, rso = rs_estimate(df.Tmax, df.Tmin, df.rain_flag, df.target, lat, elev, coef, "S3")
    rs1, _, _ = rs_estimate(df.Tmax, df.Tmin, df.rain_flag, df.target, lat, elev, coef, "S1")
    df["Ra"], df["Rso"], df["Rs_S3"], df["Rs_S1"] = ra, rso, rs3, rs1
    df["u2"] = df["u10"].map(lambda u: np.nan if pd.isna(u) else wind_2m(u, FCST_ANEM))   # 풍속 없음 → 결측(채점에서 빠짐)
    df["ETo_S3"] = eto_series(df.Tmax, df.Tmin, df.ea, df.u10, df.Rs_S3, df.target, lat, elev)
    df["ETo_S1"] = eto_series(df.Tmax, df.Tmin, df.ea, df.u10, df.Rs_S1, df.target, lat, elev)
    o = obs.set_index("date")
    look = lambda col, dates: o[col].reindex(pd.DatetimeIndex(dates)).values
    for c_obs, c in (("ETo_obs", "ETo_obs"), ("Tmax", "Tmax_obs"), ("Tmin", "Tmin_obs"), ("ea_obs", "ea_obs"),
                     ("u10", "u10_obs"), ("Rs", "Rs_obs"), ("rain", "rain_obs"), ("pa", "pa_obs")):
        df[c] = look(c_obs, df.target)
    df["flag_obs"] = (df["rain_obs"] >= RAIN_FLAG_MM).astype(int)
    methods = ["S3", "S1"]
    if "sky_cloudy" in df and df["sky_cloudy"].notna().any() and S4_RAIN in df and df[S4_RAIN].notna().any():
        if s4_fixed is None:
            co, tab = s4_cv(df)
            df["s4_fold"] = month_folds(df.target).values
        else:
            m = s4_fixed.set_index("lead_day")
            co = np.array([m.loc[k, list(S4_NAMES)].to_numpy(float) if k in m.index else np.full(len(S4_NAMES), np.nan)
                           for k in df.lead_day]).reshape(len(df), len(S4_NAMES))
            tab = s4_fixed.assign(fold=S4_FIXED)[["fold", "lead_day", "n", *S4_NAMES]]
            df["s4_fold"] = S4_FIXED
        for j, n in enumerate(S4_NAMES):
            df[f"s4_{n}"] = co[:, j]
        df["Rs_S4"] = rs_s4(df.Tmax, df.Tmin, df[S4_RAIN], df.sky_cloudy, df.sky_overcast, df.Ra, df.Rso, co)
        df["ETo_S4"] = eto_series(df.Tmax, df.Tmin, df.ea, df.u10, df.Rs_S4, df.target, lat, elev)
        df.attrs["s4_table"] = tab
        methods.insert(0, "S4")
    main = methods[0]
    df["Rs_main"], df["ETo_main"] = df[f"Rs_{main}"], df[f"ETo_{main}"]
    # 기준선: 발표일 전날(가장 최근의 완결된 관측일) 값, 최근 7일 평균
    df["ETo_pers"] = look("ETo_obs", df.run_date - pd.Timedelta(days=1))
    roll7 = o["ETo_obs"].rolling(7, min_periods=7).mean()
    df["ETo_7d"] = roll7.reindex(pd.DatetimeIndex(df.run_date - pd.Timedelta(days=1))).values
    df["Kc"] = kc_series(df.target, kp)
    for c in methods + ["main"]:
        df[f"ETc_{c}"] = df["Kc"] * df[f"ETo_{c}"]
    df["ETc_obs"] = df["Kc"] * df["ETo_obs"]
    df["ETc_pers"] = df["Kc"] * df["ETo_pers"]
    df["ETc_7d"] = df["Kc"] * df["ETo_7d"]
    return df


def main_method(df):
    """주 방법 이름: 하늘상태 S4 열이 있으면 S4, 아니면 S3"""
    return "S4" if "ETo_S4" in df.columns else "S3"


def split_verifiable(df):
    """채점할 수 있는 행과 뺀 행. 예보 입력(주 방법에 필요한 하늘상태 포함)·대상일 관측·기준선 관측이 모두 있어야 한다.
       반환: (검증 표, 뺀 행 + drop_reason)"""
    no_in = df[["Tmax", "Tmin", "ea", "u10", "rain"]].isna().any(axis=1)
    no_sky = (~no_in) & df["ETo_main"].isna() if "ETo_main" in df else pd.Series(False, index=df.index)
    no_obs = df["ETo_obs"].isna()
    no_base = df[["ETo_pers", "ETo_7d"]].isna().any(axis=1)
    reason = np.select([no_in, no_sky, no_obs, no_base],
                       ["예보 입력 결측", "하늘상태·강수확률 없음", "대상일 관측 없음", "기준선 관측 없음"], "")
    d = df.assign(drop_reason=reason)
    keep = d[reason == ""].drop(columns="drop_reason").reset_index(drop=True)
    keep.attrs = dict(df.attrs)
    return keep, d[reason != ""].reset_index(drop=True)


def run_gaps(runs, max_gap_h=15):
    """제외된 서비스 발표 [(구분, 발표시각, 빠진 요소)] → 연속 구간 목록 [(첫 발표, 끝 발표, 횟수, 빠진 요소)]"""
    ts = sorted(runs, key=lambda x: x[1])
    out = []
    for name, t, miss in ts:
        if out and (t - out[-1][1]) <= pd.Timedelta(hours=max_gap_h):
            a, _, n, m = out[-1]
            out[-1] = (a, t, n + 1, sorted(set(m) | set(miss)))
        else:
            out.append((t, t, 1, sorted(miss)))
    return out

def _stats(f, o):
    m = ~(pd.isna(f) | pd.isna(o))
    f, o = np.asarray(f, float)[m], np.asarray(o, float)[m]
    e = f - o
    r = np.corrcoef(f, o)[0, 1] if len(f) > 2 and f.std() > 0 and o.std() > 0 else np.nan
    return dict(n=int(m.sum()), obs_mean=o.mean(), fcst_mean=f.mean(), MBE=e.mean(), MAE=np.abs(e).mean(),
                RMSE=math.sqrt((e ** 2).mean()), R2=r ** 2)


def lead_metrics(df, fcol="ETo_main", ocol="ETo_obs", pcol="ETo_pers", mcol="ETo_7d"):
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


def cum3(df, fcol="ETo_main", ocol="ETo_obs", pcol="ETo_pers"):
    """발표별 처음 3개 대상일(아침 D0~D+2, 저녁 D+1~D+3) 누적. 세 날이 모두 채점 가능할 때만 포함"""
    rows = []
    first = {n: min(ls) for n, (_, ls) in SERVICE_RUNS.items()}
    for (rn, run), g in df.groupby(["run_name", "run"], sort=False):
        k0 = first.get(rn, int(g.lead_day.min()))
        g3 = g[g.lead_day.isin([k0, k0 + 1, k0 + 2])]
        if len(g3) < 3 or g3[[fcol, ocol, pcol]].isna().any().any():
            continue
        rows.append(dict(run_name=rn, run=run, leads=f"D+{k0}~D+{k0 + 2}",
                         fcst=g3[fcol].sum(), obs=g3[ocol].sum(), pers=g3[pcol].sum()))
    c = pd.DataFrame(rows, columns=["run_name", "run", "leads", "fcst", "obs", "pers"])
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


# ── 월별 성능과 판정 불확실성 ────────────────────────────────────────────
def month_metrics(df, leads=H2_LEADS, fcol="ETo_main"):
    """발표 × 선행일 × 대상일 월별 예보 ETo 지표와 입력 편향(예보 − 관측). 엑셀 월별 시트와 같은 정의"""
    rows = []
    d = df[df.lead_day.isin(leads)]
    for (rn, k, m), g in d.groupby(["run_name", "lead_day", d.target.dt.month], sort=False):
        s, sp = _stats(g[fcol], g.ETo_obs), _stats(g.ETo_pers, g.ETo_obs)
        rows.append(dict(run_name=rn, lead_day=int(k), month=int(m), n=s["n"], obs_mean=s["obs_mean"],
                         fcst_mean=s["fcst_mean"], MBE=s["MBE"], RMSE=s["RMSE"], rel_RMSE=s["RMSE"] / s["obs_mean"],
                         RMSE_pers=sp["RMSE"], skill_pers=1 - s["RMSE"] / sp["RMSE"],
                         dTmax=(g.Tmax - g.Tmax_obs).mean(), dTmin=(g.Tmin - g.Tmin_obs).mean(),
                         ddT=((g.Tmax - g.Tmin) - (g.Tmax_obs - g.Tmin_obs)).mean(), dea=(g.ea - g.ea_obs).mean(),
                         du10=(g.u10 - g.u10_obs).mean(), dRs=(g.Rs_main - g.Rs_obs).mean(),
                         rain_fcst=g.rain_flag.mean(), rain_obs=g.flag_obs.mean(),
                         false_alarm=int(((g.rain_flag == 1) & (g.flag_obs == 0)).sum()),
                         miss=int(((g.rain_flag == 0) & (g.flag_obs == 1)).sum())))
    order = {n: i for i, n in enumerate(SERVICE_RUNS)}
    return pd.DataFrame(rows).sort_values(["run_name", "lead_day", "month"], ignore_index=True,
                                          key=lambda s: s.map(order) if s.name == "run_name" else s)


def month_verdict(mm, skill_min=H2_SKILL_MIN, rmse_d1_max=H2_RMSE_D1_MAX):
    """참고: 달마다 H2 기준을 적용한 결과(원인 진단용). H2 판정 자체는 전체 기간으로 한다"""
    rows = []
    for (rn, m), g in mm.groupby(["run_name", "month"], sort=False):
        g = g.set_index("lead_day")
        ms, r1 = float(g.skill_pers.min()), float(g.loc[1, "RMSE"]) if 1 in g.index else float("nan")
        rows.append(dict(run_name=rn, month=int(m), n_d1=int(g.loc[1, "n"]) if 1 in g.index else 0,
                         min_skill=ms, rmse_d1=r1, pass_=bool(ms >= skill_min and r1 <= rmse_d1_max)))
    return pd.DataFrame(rows)


def bootstrap_h2(df, n_boot=H2_BOOT_N, block=H2_BOOT_BLOCK, seed=H2_BOOT_SEED, skill_min=H2_SKILL_MIN,
                 rmse_d1_max=H2_RMSE_D1_MAX, leads=H2_LEADS, fcol="ETo_main"):
    """H2 판정의 표본 불확실성 — 이동 블록 부트스트랩.
       발표일을 block일 묶음으로 복원추출한다. 같은 날의 아침·저녁 발표와 선행일을 함께 뽑아 서로의 상관을 보존.
       반환: runs={구분: dict(rmse_d1, min_skill = (5%, 50%, 95%) 백분위, pass_frac = 두 기준을 모두 충족한 비율)},
             pass_all(모든 발표 동시 충족 비율), n_boot, block, n_days, seed"""
    ks = sorted(set(leads) | {1})
    d = df[df.lead_day.isin(ks)]
    d = d.assign(e2=(d[fcol] - d.ETo_obs) ** 2, p2=(d.ETo_pers - d.ETo_obs) ** 2)
    runs = list(dict.fromkeys(d.run_name))
    days = np.array(sorted(d.run_date.unique()))
    pos = {t: i for i, t in enumerate(days)}
    ri, ki = {r: i for i, r in enumerate(runs)}, {k: i for i, k in enumerate(ks)}
    N = len(days)
    E, P, Cn = (np.zeros((N, len(runs), len(ks))) for _ in range(3))
    for (t, rn, k), g in d.groupby(["run_date", "run_name", "lead_day"]):
        i, j, l = pos[t], ri[rn], ki[k]
        E[i, j, l], P[i, j, l], Cn[i, j, l] = g.e2.sum(), g.p2.sum(), len(g)
    rng = np.random.default_rng(seed)
    nb = -(-N // block)
    starts = rng.integers(0, max(N - block + 1, 1), size=(n_boot, nb))
    idx = np.minimum((starts[:, :, None] + np.arange(block)).reshape(n_boot, -1)[:, :N], N - 1)
    Es, Ps, Cs = E[idx].sum(axis=1), P[idx].sum(axis=1), Cn[idx].sum(axis=1)
    with np.errstate(invalid="ignore", divide="ignore"):
        rmse, skill = np.sqrt(Es / Cs), 1 - np.sqrt(Es / Ps)
    q = lambda a: tuple(float(x) for x in np.nanpercentile(a, [5, 50, 95]))
    out, ok_all = {}, np.ones(n_boot, bool)
    for rn, j in ri.items():
        r1 = rmse[:, j, ki[1]]
        ms = np.nanmin(skill[:, j, [ki[k] for k in leads]], axis=1)
        ok = (r1 <= rmse_d1_max) & (ms >= skill_min)
        ok_all &= ok
        out[rn] = dict(rmse_d1=q(r1), min_skill=q(ms), pass_frac=float(ok.mean()))
    return dict(runs=out, pass_all=float(ok_all.mean()), n_boot=n_boot, block=block, n_days=N, seed=seed)


# ── 입력 진단과 오차 분해 ────────────────────────────────────────────────
def input_diagnostics(df):
    """발표 × 선행일별 입력 편향(예보−관측)·RMSE, 강수유무 적중"""
    rows = []
    for (rn, k), g in df.groupby(["run_name", "lead_day"], sort=False):
        r = dict(run_name=rn, lead_day=int(k), n=len(g))
        for name, f, o in (("Tmax", g.Tmax, g.Tmax_obs), ("Tmin", g.Tmin, g.Tmin_obs),
                           ("dT", g.Tmax - g.Tmin, g.Tmax_obs - g.Tmin_obs), ("ea", g.ea, g.ea_obs),
                           ("u10", g.u10, g.u10_obs), ("Rs", g.Rs_main, g.Rs_obs)):
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


def rain_input(df):
    """주 방법의 강수 입력 열: S4는 강수확률 하루 최대(S4_RAIN), S3는 강수유무"""
    return S4_RAIN if main_method(df) == "S4" else "rain_flag"


def main_rs(df, lat, elev, coef, tmax=None, tmin=None, rain=None, s4coef=None):
    """주 방법의 Rs를 (바꾼) 입력으로 다시 계산. S4는 행별 계수(s4_a~e 열 또는 s4coef), S3는 rs_coef 계수.
       rain: 주 방법의 강수 입력(S4 강수확률 하루 최대, S3 강수유무). 없으면 df의 해당 열"""
    tmax = df.Tmax if tmax is None else tmax
    tmin = df.Tmin if tmin is None else tmin
    rain = df[rain_input(df)] if rain is None else rain
    if main_method(df) == "S4":
        co = df[[f"s4_{n}" for n in S4_NAMES]].values if s4coef is None else s4coef
        return rs_s4(tmax, tmin, rain, df.sky_cloudy, df.sky_overcast, df.Ra, df.Rso, co)
    return rs_estimate(tmax, tmin, rain, df.target, lat, elev, coef, "S3")[0]


def error_attribution(df, lat, elev, coef):
    """예보 입력을 하나씩 관측값으로 바꿨을 때의 ETo RMSE (주 방법). 줄어든 만큼이 그 입력 예보오차의 몫.
       '관측 입력 전부'는 Rs 추정만 남은 상태(S3면 H1 조건과 같음. S4는 하늘상태 관측이 없어 예보 그대로)."""
    all_obs = "관측 입력 전부 (Rs만 추정)" if main_method(df) == "S3" else "관측 입력 전부 (Rs만 추정, 하늘상태는 예보)"
    rc = rain_input(df)       # S4: 강수확률 하루 최대 → '관측'은 비 온 날(≥ 1 mm) 1, 아니면 0 (완벽한 강수확률)
    rain_name = "강수유무 → 관측" if rc == "rain_flag" else "강수확률 → 관측 (비 온 날 1)"
    variants = {
        "예보 입력 그대로": dict(),
        "기온 → 관측": dict(Tmax="Tmax_obs", Tmin="Tmin_obs"),
        "습도(ea) → 관측": dict(ea="ea_obs"),
        "풍속 → 관측": dict(u10="u10_obs"),
        rain_name: {rc: "flag_obs"},
        all_obs: {"Tmax": "Tmax_obs", "Tmin": "Tmin_obs", "ea": "ea_obs", "u10": "u10_obs", rc: "flag_obs"},
        "참고: Rs만 관측": dict(Rs="Rs_obs"),
    }
    res = {}
    for name, rep in variants.items():
        col = lambda c: df[rep.get(c, c)]
        if "Rs" in rep:
            rs = df["Rs_obs"].values
        else:
            rs = main_rs(df, lat, elev, coef, col("Tmax"), col("Tmin"), col(rc))
        res[name] = eto_series(col("Tmax"), col("Tmin"), col("ea"), col("u10"), rs, df.target, lat, elev)
    rows = []
    for (rn, k), idx in df.groupby(["run_name", "lead_day"], sort=False).groups.items():
        r = dict(run_name=rn, lead_day=int(k))
        for name, v in res.items():
            e = v[df.index.get_indexer(idx)] - df.loc[idx, "ETo_obs"].values
            r[name] = math.sqrt(np.nanmean(e ** 2))
        rows.append(r)
    return pd.DataFrame(rows), list(variants)


def fit_rs_forecast(t):
    """Rs 식(S3) 계수를 예보 입력(예보 일교차·예보 강수유무)과 관측 Rs로 최소제곱 적합 → (a, b, c).
       rs_model.fit과 같은 식이지만 입력이 관측이 아니라 예보다(예보 입력의 계통오차까지 흡수)."""
    t = t[(t.Tmax > t.Tmin) & t.Rs_obs.notna()]
    X = np.column_stack([np.ones(len(t)), np.sqrt(t.Tmax - t.Tmin).values, t.rain_flag.values.astype(float)])
    return np.linalg.lstsq(X, (t.Rs_obs / t.Ra).values, rcond=None)[0]


def bias_correction_cv(df, lat, elev, coef):
    """탐색: 지점 보정의 효과를 월 단위 교차검증으로 평가(보정값은 다른 달 자료로 추정). 대상은 주 방법의 예보.
       B1 기온(Tmax·Tmin) 가산 보정, B2 기온 + 풍속 가산 보정, B3 ETo 비율 보정,
       B4 Rs 계수 재보정(예보 입력 → 관측 Rs; 주 방법이 S3일 때만. S4는 이미 예보 입력으로 맞춘 계수).
       S4는 계수 자체가 교차검증 값이므로, B3 비율을 구하는 학습 행의 예보도 검증 달을 뺀 계수로 다시 계산한다(중첩 교차검증).
       반환: (지표표 long: method·발표·선행일, 전체기간 입력 편향 표, 교차검증 묶음 목록,
              행별 보정 ETo 표(열 B1~B3(B4), df와 같은 index))"""
    d = df.copy()
    main = main_method(d)
    d["fold"] = month_folds(d.target).values
    if d.fold.nunique() < 2:
        raise ValueError("편향 보정 교차검증에는 대상일이 2개 달 이상 필요합니다")
    keys = ("B1", "B2", "B3") + (("B4",) if main == "S3" else ())
    cols = {b: np.full(len(d), np.nan) for b in keys}
    pos = {ix: i for i, ix in enumerate(d.index)}
    for f in sorted(d.fold.unique()):
        tr, te = d[d.fold != f], d[d.fold == f]
        base_tr = tr.ETo_main
        if main == "S4" and not s4_is_fixed(d):   # 학습 행의 S4 예보를 검증 달 f까지 뺀 계수로 다시 계산(고정 계수면 그대로)
            co_f, _ = s4_cv(d, exclude=(f,))
            rs_f = rs_s4(tr.Tmax, tr.Tmin, tr[S4_RAIN], tr.sky_cloudy, tr.sky_overcast, tr.Ra, tr.Rso,
                         co_f[[pos[i] for i in tr.index]])
            base_tr = pd.Series(eto_series(tr.Tmax, tr.Tmin, tr.ea, tr.u10, rs_f, tr.target, lat, elev), index=tr.index)
        for (rn, k), g in te.groupby(["run_name", "lead_day"]):
            m = (tr.run_name == rn) & (tr.lead_day == k)
            t = tr[m]
            dtx, dtn = (t.Tmax - t.Tmax_obs).mean(), (t.Tmin - t.Tmin_obs).mean()
            du = (t.u10 - t.u10_obs).mean()
            ok = base_tr[m].notna() & t.ETo_obs.notna()
            ratio = t.ETo_obs[ok].sum() / base_tr[m][ok].sum()
            tx, tn = g.Tmax - dtx, g.Tmin - dtn
            rs = main_rs(g, lat, elev, coef, tx, tn)
            idx = [pos[i] for i in g.index]
            cols["B1"][idx] = eto_series(tx, tn, g.ea, g.u10, rs, g.target, lat, elev)
            cols["B2"][idx] = eto_series(tx, tn, g.ea, (g.u10 - du).clip(lower=0.1), rs, g.target, lat, elev)
            cols["B3"][idx] = g.ETo_main * ratio
            if "B4" in cols:
                a4, b4, c4 = fit_rs_forecast(t)
                rs4 = rs_s3(g.Tmax, g.Tmin, g.rain_flag, g.Ra.values, g.Rso.values, a4, b4, c4)
                cols["B4"][idx] = eto_series(g.Tmax, g.Tmin, g.ea, g.u10, rs4, g.target, lat, elev)
    names = {"ETo_main": "보정 없음(원자료)", "B1": "기온 보정", "B2": "기온+풍속 보정", "B3": "ETo 비율 보정",
             "B4": "Rs 계수 재보정(예보 입력)"}
    out = []
    for c, name in names.items():
        if c != "ETo_main":
            if c not in cols:
                continue
            d[c] = cols[c]
        m = lead_metrics(d, c)
        m.insert(0, "method", name)
        _, c3 = cum3(d, c)
        if len(c3):
            m = m.merge(c3[["run_name", "RMSE"]].rename(columns={"RMSE": "RMSE_3d"}), on="run_name", how="left")
        else:              # 보정값을 못 구한 경우(예: S4인데 대상월이 2개뿐이라 중첩 교차검증의 학습 자료가 없음)
            m["RMSE_3d"] = np.nan
        out.append(m)
    bias = (df.assign(dTmax=df.Tmax - df.Tmax_obs, dTmin=df.Tmin - df.Tmin_obs, du10=df.u10 - df.u10_obs)
              .groupby(["run_name", "lead_day"], sort=False)[["dTmax", "dTmin", "du10"]].mean().reset_index())
    return pd.concat(out, ignore_index=True), bias, sorted(d.fold.unique()), d[list(cols)]


def method_comparison(df, methods=None):
    """같은 대상일에서 Rs 추정 방법별 예보 ETo 지표(발표 × 선행일) — 주 방법과 비교 방법의 차이를 본다"""
    methods = methods or [m for m in ("S4", "S3", "S1") if f"ETo_{m}" in df.columns]
    return pd.concat([lead_metrics(df, f"ETo_{m}").assign(method=m) for m in methods], ignore_index=True)


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
def coef_file(path):
    """계수 파일 경로. 상대경로가 현재 폴더에 없으면 이 스크립트 폴더의 같은 이름 파일을 쓴다
       (다른 폴더에서 실행해도 저장소의 rs_coef.csv를 찾도록)"""
    if os.path.isabs(path) or os.path.exists(path):
        return path
    alt = os.path.join(os.path.dirname(os.path.abspath(__file__)), path)
    return alt if os.path.exists(alt) else path


def load_obs(path):
    obs, meta = load_station_workbook(path)
    obs = add_eto_obs(obs, meta["lat"], meta["elev"], meta.get("anem", 10.0))
    return obs, meta


def run_verify(fcst_paths, obs_path, stn, coef_path="rs_coef.csv", analysis=True, compare_paths=None,
               s4_coef_path=None, grid=None):
    """검증 전체 계산. analysis=True면 오차분해·편향보정 탐색까지.
       compare_paths: 비교할 다른 격자의 과거 예보(같은 관측·계수·Kc로 계산해 격자 차이를 봄)
       s4_coef_path : 다른 해 운영 S4 계수 파일(rs_sky_coef.csv). 주면 교차검증 대신 그 계수를 고정 적용(독립 연도 검증)
       grid         : 파일에 격자 정보가 없을 때(요소별 KST CSV) 쓸 격자 'nx_ny'"""
    obs, meta = load_obs(obs_path)
    lon = STATION_LON.get(str(stn))
    arch = load_archive(fcst_paths)
    grid_given = None
    if grid and not arch.location:              # 요소별 KST CSV에는 격자 정보가 없음 → 사용자가 준 격자
        arch.location.add(grid)
        grid_given = grid
    rep = check_archive(arch)
    st = service_table(arch, lat=meta["lat"], lon=lon)
    coef = load_coef(stn, coef_file(coef_path))
    kp = kc_params(meta["settings"])
    s4_fixed = load_s4_fixed(stn, s4_coef_path) if s4_coef_path else None
    ft = forecast_table(st, obs, meta["lat"], meta["elev"], coef, kp, s4_fixed)
    s4_table = ft.attrs.get("s4_table")
    df, dropped = split_verifiable(ft)
    name, grid = STATIONS.get(str(stn), ("", ""))
    c3, _ = cum3(df)
    h1_start = pd.Timestamp(kp["bud"]) if kp.get("bud") else df.target.min()
    res = dict(stn=str(stn), stn_name=name, stn_grid=grid, arch=arch, check=rep, table=df, obs=obs, meta=meta,
               coef=coef, kp=kp, dropped=dropped, skipped_runs=list(getattr(arch, "skipped_runs", [])),
               cum3_runs=set(zip(c3.run_name, c3.run)) if len(c3) else set(), main=main_method(df),
               h1_start=h1_start, h1_end=min(obs.date.max(), pd.Timestamp(h1_start.year, 9, 30)),   # H1은 생육기(~9/30)
               s4_fixed=s4_fixed, s4_coef_path=s4_coef_path, grid_given=grid_given)
    if res["main"] == "S4":
        res["s4_table"] = s4_table
        res["s4_all"] = s4_fit_all(df)
        res["methods"] = method_comparison(df)
    if compare_paths:
        arch2 = load_archive(compare_paths)
        df2, _ = split_verifiable(forecast_table(service_table(arch2, lat=meta["lat"], lon=lon), obs, meta["lat"],
                                                 meta["elev"], coef, kp, s4_fixed))
        res["compare"] = dict(check=check_archive(arch2), table=df2)
        res["grid_cmp"] = grid_comparison(df, df2)
    if analysis:
        res["attr"], res["attr_names"] = error_attribution(df, meta["lat"], meta["elev"], coef)
        res["bc"], res["bc_bias"], res["bc_folds"], bc_rows = bias_correction_cv(df, meta["lat"], meta["elev"], coef)
        res["month"] = month_metrics(df)
        res["month_verdict"] = month_verdict(res["month"])
        res["boot"] = bootstrap_h2(df)
        # 탐색: 교차검증한 보정을 적용했을 때 판정 여유가 얼마나 늘어나는지 (G4 보정 채택 판단 근거)
        res["boot_bc"] = {name: bootstrap_h2(df.assign(**{c: bc_rows[c]}), fcol=c)
                          for c, name in (("B2", "기온+풍속 보정"), ("B3", "ETo 비율 보정")) if bc_rows[c].notna().any()}
        if res["main"] == "S4":   # 비교: 같은 대상일에서 하늘상태 없이(S3) 계산했을 때의 판정 불확실성
            res["boot_s3"] = bootstrap_h2(df, fcol="ETo_S3")
        res["findings"] = findings(res)
    return res


def grid_comparison(df, df2):
    """두 격자 예보의 비교: 선행시간별 성능·입력 편향, 같은 발표·대상일의 직접 차이(df − df2)"""
    met = pd.concat([lead_metrics(d).assign(grid=g) for g, d in (("A", df), ("B", df2))], ignore_index=True)
    diag = pd.concat([input_diagnostics(d).assign(grid=g) for g, d in (("A", df), ("B", df2))], ignore_index=True)
    j = df.merge(df2, on=["run", "lead_day"], suffixes=("", "_B"))
    rows = []
    for c, label in (("Tmax", "최고기온 (℃)"), ("Tmin", "최저기온 (℃)"), ("ea", "ea (kPa)"), ("u10", "풍속 u10 (m/s)"),
                     ("rain", "강수 (mm/일)"), ("Rs_S3", "추정 Rs (MJ/m²/일)"), ("ETo_S3", "예보 ETo (mm/일)")):
        d = j[c] - j[f"{c}_B"]
        rows.append(dict(item=label, mean=d.mean(), sd=d.std(), frac_diff=float((d.abs() > 1e-9).mean())))
    diff = pd.DataFrame(rows)
    j["month"] = j.target.dt.month
    by_month = j.groupby("month").apply(lambda x: pd.Series(dict(
        n=len(x), dTmax=(x.Tmax - x.Tmax_B).mean(), dTmin=(x.Tmin - x.Tmin_B).mean(),
        du10=(x.u10 - x.u10_B).mean(), dETo=(x.ETo_S3 - x.ETo_S3_B).mean()))).reset_index()
    agree = float((j.rain_flag == j.rain_flag_B).mean())
    return dict(metrics=met, diag=diag, diff=diff, by_month=by_month, rain_agree=agree, n=len(j))


def s4_label(res):
    """S4 계수 방식 설명(해석 문장·CLI용)"""
    if res.get("s4_fixed") is not None:
        f = res["s4_fixed"]
        per = f"{f.fit_start.min()}~{f.fit_end.max()}" if "fit_start" in f and str(f.fit_start.min()) else ""
        return f"다른 해 운영 계수 고정({os.path.basename(str(res.get('s4_coef_path') or 'rs_sky_coef.csv'))}" \
               + (f", 적합 {per})" if per else ")")
    return "월 단위 교차검증 계수"


def findings(res):
    """요약 시트의 해석 문장(④~). 수식으로 연결할 수 없는 Python 분석 결과와 자료 조건 경고.
       순서: 판정 불확실성 → 월별 약점 → 오차 원인 → 보정 탐색 → 격자·자료 조건"""
    out, df = [], res["table"]
    if res.get("main") == "S4" and res.get("methods") is not None:
        m = res["methods"]
        r1 = lambda meth, rn: float(m[(m.method == meth) & (m.run_name == rn) & (m.lead_day == 1)].RMSE.iloc[0])
        d1 = df[df.lead_day == 1]
        mon = lambda c: d1.groupby(d1.target.dt.month).apply(lambda g: math.sqrt(((g[c] - g.ETo_obs) ** 2).mean()))
        m4, m3 = mon("ETo_S4"), mon("ETo_S3")
        better = [int(x) for x in m4.index if m4[x] < m3[x] - 0.02]
        worse = [int(x) for x in m4.index if m4[x] > m3[x] + 0.02]
        lab = s4_label(res)
        txt = (f"하늘상태 반영(S4, {lab}) — 같은 대상일의 S3(하늘상태 없음)와 비교: D+1 RMSE 아침 "
               f"{r1('S3', '아침'):.2f}→{r1('S4', '아침'):.2f}, 저녁 {r1('S3', '저녁'):.2f}→{r1('S4', '저녁'):.2f} mm/일")
        if better:
            txt += f". D+1이 좋아진 달 {'·'.join(map(str, better))}월"
        if worse:
            txt += f", 나빠진 달 {'·'.join(map(str, worse))}월"
        out.append(txt + " (요약 방법 비교 표, 월별 시트)")
    bt = res.get("boot")
    if bt:
        rng_ = ", ".join(f"{rn} {v['rmse_d1'][0]:.2f}~{v['rmse_d1'][2]:.2f}" for rn, v in bt["runs"].items())
        frac = ", ".join(f"{rn} {v['pass_frac']:.0%}" for rn, v in bt["runs"].items())
        txt = (f"판정 불확실성({bt['block']}일 블록 부트스트랩 {bt['n_boot']:,}회): D+1 RMSE 90% 구간 {rng_} mm/일, "
               f"두 기준을 모두 충족한 비율 {frac}")
        b3 = res.get("boot_s3")
        if b3:
            txt += " (같은 대상일 S3: " + ", ".join(f"{rn} {v['pass_frac']:.0%}" for rn, v in b3["runs"].items()) + ")"
        weak = [rn for rn, v in bt["runs"].items() if v["pass_frac"] < 0.9]
        if weak:
            txt += f" → {'·'.join(weak)} 발표는 기준과의 여유가 표본 변동보다 작음"
            bb = res.get("boot_bc") or {}
            if bb:
                name = max(bb, key=lambda n: min(bb[n]["runs"][w]["pass_frac"] for w in weak))
                txt += (f". 교차검증한 '{name}'을 적용하면 "
                        + ", ".join(f"{w} {bb[name]['runs'][w]['pass_frac']:.0%}" for w in weak))
        out.append(txt + " (오차분해 ④)")
    mm = res.get("month")
    if mm is not None and len(mm):
        m1 = mm[(mm.lead_day == 1) & (mm.n >= 10)]
        bad = m1[m1.RMSE > H2_RMSE_D1_MAX]
        if len(bad):
            parts = []
            for rn, g in bad.groupby("run_name", sort=False):
                w = g.loc[g.RMSE.idxmax()]
                parts.append(f"{rn} {'·'.join(str(int(x)) for x in g.month)}월(최대 {int(w.month)}월 {w.RMSE:.2f})")
            b = m1[m1.month.isin(sorted(set(bad.month)))]
            txt = f"월별(D+1): RMSE가 기준 {H2_RMSE_D1_MAX:.1f}을 넘는 달 — {', '.join(parts)}. "
            if b.dRs.max() < 0:
                txt += f"이 달들의 예보 Rs는 관측보다 {-b.dRs.max():.1f}~{-b.dRs.min():.1f} MJ/m²/일 작음"
            else:
                txt += f"이 달들의 예보 Rs 편향 {b.dRs.min():+.1f}~{b.dRs.max():+.1f} MJ/m²/일"
            over = b.assign(ex=b.rain_fcst - b.rain_obs).groupby("month").ex.mean()
            if over.max() >= 0.15:
                mo = int(over.idxmax()); x = b[b.month == mo]
                if res.get("main") == "S4" and S4_RAIN in df:
                    pm = df[(df.lead_day == 1) & (df.target.dt.month == mo)][S4_RAIN].mean()
                    txt += (f". {mo}월은 예보상 비 오는 날이 {x.rain_fcst.mean():.0%}(실제 {x.rain_obs.mean():.0%})로 많고 "
                            f"강수확률 하루 최대도 평균 {pm:.0%}여서 강수 항이 Rs를 더 낮춤")
                else:
                    txt += (f". {mo}월은 예보상 비 오는 날이 {x.rain_fcst.mean():.0%}(실제 {x.rain_obs.mean():.0%})로 많아 "
                            f"강수유무 보정이 Rs를 더 낮춤")
            out.append(txt + " (월별 시트)")
        else:
            out.append(f"월별(D+1): 모든 달에서 RMSE ≤ {H2_RMSE_D1_MAX:.1f} mm/일 (월별 시트)")
    att, names = res["attr"], res["attr_names"]
    d1 = att[att.lead_day == 1]
    base = d1[names[0]].mean()
    cand = {n: base - d1[n].mean() for n in names[1:5]}
    top = max(cand, key=cand.get)
    struct = d1[names[5]].mean()
    txt = (f"오차 분해(D+1, 두 발표 평균 RMSE {base:.2f}): 예보 입력 중에는 '{top}' 교체 시 {cand[top]:.2f} mm/일 감소로 "
           f"가장 큼. 예보가 완벽해도 남는 구조오차(Rs 추정)는 {struct:.2f} mm/일")
    if len(names) > 6:
        rs_only = d1[names[6]].mean()
        txt += f", 추정 Rs만 관측 Rs로 바꾸면 {rs_only:.2f} mm/일"
        if base - rs_only > cand[top]:
            txt += " → 일사 추정이 가장 큰 오차원"
    out.append(txt + " (오차분해 ①)")
    bc = res["bc"]
    r1 = lambda m: bc[(bc.method == m) & (bc.lead_day == 1)].RMSE.mean()
    raw = r1("보정 없음(원자료)")
    allv = {m: r1(m) for m in dict.fromkeys(bc.method) if m != "보정 없음(원자료)"}
    vals = {m: v for m, v in allv.items() if pd.notna(v)}
    na = [m for m in allv if m not in vals]       # 교차검증 학습 자료가 부족해 값을 못 구한 보정
    if vals:
        best = min(vals, key=vals.get)
        gain = 1 - vals[best] / raw
        out.append(f"보정 탐색(월 단위 교차검증, D+1 두 발표 평균 RMSE {raw:.2f}): "
                   + ", ".join(f"{m} {v:.2f}" for m, v in vals.items())
                   + (f" mm/일 — 가장 좋은 방법은 '{best}'({gain:.0%} 감소)" if gain >= 0.005 else
                      " mm/일 — 어느 보정도 오차를 줄이지 못함(다른 달 자료로 추정한 보정값이 맞지 않음)")
                   + (f". 대상월이 적어 계산하지 못한 보정: {', '.join(na)}" if na else "")
                   + ". 편향 보정은 쓰지 않음 — 한 해에서 구한 보정값이 다른 해로 옮겨 가지 않음(VALIDATION #10) (오차분해 ②)")
    loc = res["check"]["location"]
    if res.get("grid_cmp"):
        gc, loc2 = res["grid_cmp"], res["compare"]["check"]["location"]
        d = gc["diff"].set_index("item")["mean"]
        m = gc["metrics"]
        r1 = lambda g, rn: float(m[(m.grid == g) & (m.run_name == rn) & (m.lead_day == 1)].RMSE.iloc[0])
        out.append(f"격자 비교({', '.join(loc)} − {', '.join(loc2)}, 같은 발표·대상일): 최고기온 {d['최고기온 (℃)']:+.1f}℃, "
                   f"최저기온 {d['최저기온 (℃)']:+.1f}℃, 예보 ETo {d['예보 ETo (mm/일)']:+.2f} mm/일. "
                   f"D+1 RMSE 아침 {r1('B', '아침'):.2f}→{r1('A', '아침'):.2f}, 저녁 {r1('B', '저녁'):.2f}→{r1('A', '저녁'):.2f} (격자비교 시트)")
    if res.get("stn_grid") and loc != [res["stn_grid"]]:
        out.append(f"주의: 예보 격자 {', '.join(loc)}가 ASOS {res['stn']} 격자({res['stn_grid']})와 다릅니다. "
                   f"대표성 오차(특히 기온 편향)가 달라질 수 있어 {res['stn_grid']} 자료로 재확인이 필요합니다")
    gaps = run_gaps(res.get("skipped_runs", []))
    if gaps:
        txt = ", ".join(f"{a:%m/%d %H}시~{b:%m/%d %H}시({n}회)" if n > 1 else f"{a:%m/%d %H}시" for a, b, n, _ in gaps)
        dr = res.get("dropped")
        why = dr.drop_reason.value_counts().to_dict() if dr is not None and len(dr) else {}
        out.append(f"자료 공백: 요소가 빠진 서비스 발표 {sum(g[2] for g in gaps)}회 제외({txt})"
                   + (". 행 제외: " + ", ".join(f"{k} {v}행" for k, v in why.items()) if why else "") + " (방법 시트)")
    dr = res.get("dropped")
    if dr is not None and len(dr) and (dr.drop_reason == "하늘상태·강수확률 없음").any():
        last = df.run.max()
        out.append(f"기간: 하늘상태·강수확률 자료가 있는 발표({last:%m/%d %H}시까지)만 판정에 썼습니다 → 대상일 "
                   f"{df.target.min():%m/%d}~{df.target.max():%m/%d}. 뒤쪽 자료를 받으면 다시 계산합니다")
    months = sorted(df.target.dt.month.unique())
    if not set(range(4, 10)) <= set(months):
        out.append(f"기간: 대상일 {df.target.min():%Y-%m-%d}~{df.target.max():%Y-%m-%d}만 포함한 "
                   f"중간 결과입니다. 생육기(4~9월) 전체 판정에는 나머지 달의 예보 자료가 필요합니다")
    return [f"{'④⑤⑥⑦⑧⑨⑩⑪'[i]} {t}" for i, t in enumerate(out)]


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
    v.add_argument("--compare", nargs="+", default=None, help="비교할 다른 격자의 과거 예보 CSV 폴더(또는 파일들)")
    v.add_argument("--s4-coef", default=None,
                   help="다른 해 운영 S4 계수 파일(rs_sky_coef.csv) — 주면 교차검증 대신 고정 적용(독립 연도 검증)")
    v.add_argument("--grid", default=None, help="파일에 격자 정보가 없을 때(요소별 KST CSV) 격자 nx_ny, 예: 73_134")
    k = sub.add_parser("calib-sky", help="하늘상태 포함 과거 예보 + 관측으로 운영용 S4 계수(선행일별) → rs_sky_coef.csv")
    k.add_argument("--fcst", nargs="+", required=True); k.add_argument("--obs", required=True)
    k.add_argument("--stn", required=True); k.add_argument("--coef", default="rs_coef.csv")
    k.add_argument("--sky-coef", default="rs_sky_coef.csv")
    # ── G4: 예보 물수지 ──
    def _common(q, s4_default=None):
        q.add_argument("--fcst", nargs="+", required=True); q.add_argument("--obs", required=True)
        q.add_argument("--stn", required=True); q.add_argument("--coef", default="rs_coef.csv")
        q.add_argument("--s4-coef", default=s4_default, help="S4 계수 파일(주면 고정 적용)")
        q.add_argument("--grid", default=None, help="파일에 격자 정보가 없을 때 격자 nx_ny")
        q.add_argument("--err", default="fcst_error_table.csv", help="예보 ETo 오차표(월·발표·선행일별)")
    e = sub.add_parser("errtable", help="과거 예보 + 관측으로 예보 ETo 오차표 행(지점·해)을 만들어 fcst_error_table.csv에 기록")
    _common(e)
    w = sub.add_parser("wbverify", help="예보 물수지(관수 필요 예상일) 검증 엑셀 — G4")
    _common(w)
    w.add_argument("--irrig", default=None, help="관수 기록 CSV(날짜, 관수량_mm) — 관측 물수지에 반영")
    w.add_argument("--auto-irrigate", action="store_true", help="관수 규칙 시나리오: 전날 끝 Dr ≥ RAW면 Dr만큼 관수")
    w.add_argument("--out", default=None)
    s = sub.add_parser("service", help="한 서비스 발표(아침 02시·저녁 17시)의 관수 전망 엑셀")
    _common(s, s4_default="rs_sky_coef.csv")
    s.add_argument("--run", default=None, help="발표시각 'YYYY-MM-DD HH' (없으면 자료의 가장 최근 서비스 발표)")
    s.add_argument("--irrig", default=None, help="관수 기록 CSV(날짜, 관수량_mm)")
    s.add_argument("--out", default=None)
    a = ap.parse_args(argv)

    if a.cmd in ("errtable", "wbverify", "service"):
        return _main_wb(a)

    if a.cmd == "calib-sky":
        res = run_verify(a.fcst, a.obs, a.stn, a.coef, analysis=False)
        if res["main"] != "S4":
            raise SystemExit("[중단] 하늘상태(SKY) 예보가 없어 S4 계수를 맞출 수 없습니다")
        tab = res["s4_all"]
        save_sky_coef(a.stn, tab, a.sky_coef, note=f"{os.path.basename(a.obs)} 예보 입력 → 관측 Rs, 선행일별",
                      rain_input=S4_RAIN)
        print(tab.round(4).to_string(index=False))
        print(f"[저장] {a.sky_coef}")
        return

    if a.cmd == "calib":
        m0, m1 = (int(x) for x in a.months.split("-"))
        obs, meta = load_obs(a.obs)
        cf = fit(obs, meta["lat"], meta["elev"], months=range(m0, m1 + 1))
        save_coef(a.stn, cf, a.coef, note=f"{os.path.basename(a.obs)} {m0}~{m1}월 관측 Rs로 보정")
        print(f"[계수] stn={a.stn} {cf}")
        print(f"[저장] {a.coef}")
        return

    res = run_verify(a.fcst, a.obs, a.stn, a.coef, compare_paths=a.compare, s4_coef_path=a.s4_coef, grid=a.grid)
    df = res["table"]
    met = lead_metrics(df)
    chk = res["check"]
    print(f"[예보] {'+'.join(chk.get('formats', []))} 격자 {chk['location']}, 발표 {chk.get('issues')}회, "
          f"누락 {len(chk.get('missing_issues', []))}회, 일부만 있는 발표 {len({t for t, *_ in chk.get('short_issues', [])})}회")
    print(f"[Rs 계수] {res['coef']['source']}")
    print(f"[주 방법] {res['main']}" + (f" (하늘상태 포함, {s4_label(res)})" if res["main"] == "S4" else ""))
    for e in chk.get("missing_elements", []):
        print(f"[경고] 필수 요소 {e} 파일이 없습니다 → 모든 행이 '예보 입력 결측'으로 빠집니다")
    for f, e in chk.get("ignored_files", []):
        print(f"[참고] 쓰지 않는 요소({e}) 파일을 건너뜀: {f}")
    if res.get("grid_given"):
        print(f"[참고] 파일에 격자 정보가 없어 --grid {res['grid_given']}로 기록")
    if res["coef"].get("a") is None:
        print(f"[경고] {a.coef}에 지점 {a.stn}의 S3 계수가 없어 FAO-56 기본값(S1)으로 계산했습니다. calib를 먼저 실행하세요")
    print(met.round(3).to_string(index=False))
    _, h1 = h1_table(res["obs"], res["meta"]["lat"], res["meta"]["elev"], res["coef"], res["h1_start"], res["h1_end"],
                     res["meta"].get("anem", 10.0))
    for m, (s, _) in h1.items():
        print(f"[H1 {m}] {res['h1_start']:%Y-%m-%d}~{res['h1_end']:%Y-%m-%d} RMSE {s['RMSE']:.3f}, 합계오차 {s['SUMERR']:+.1%}")
    for rn, v in h2_verdict(met).items():
        print(f"[H2] {rn}: {'통과' if v['pass_'] else '미달'} (최소 개선율 {v['min_skill']:.0%}, D+1 RMSE {v['rmse_d1']:.2f})")
    if res.get("month") is not None:
        m1 = res["month"][res["month"].lead_day == 1]
        print("[월별 D+1 RMSE]")
        print(m1.pivot(index="run_name", columns="month", values="RMSE").round(2).to_string())
    if res.get("boot"):
        b = res["boot"]
        for rn, v in b["runs"].items():
            print(f"[불확실성] {rn}: D+1 RMSE 90% {v['rmse_d1'][0]:.2f}~{v['rmse_d1'][2]:.2f}, "
                  f"최소 개선율 90% {v['min_skill'][0]:.0%}~{v['min_skill'][2]:.0%}, 기준 충족 {v['pass_frac']:.0%}")
    for ln in res.get("findings", []):
        print(ln)
    grid = "-".join(res["check"]["location"]) or "grid"
    out = a.out or f"output/fcst_verify({a.stn})_{grid}_{df.target.min():%Y%m%d}_{df.target.max():%Y%m%d}.xlsx"
    from fcst_report import build_verify_workbook
    saved = build_verify_workbook(res, out)
    print(f"[완료] {saved}")


def _main_wb(a):
    """G4 하위 명령: errtable / wbverify / service (fcst_wb.py, fcst_wb_report.py)"""
    import fcst_wb as W
    from fao56_core import load_irrigation_log
    irrig = load_irrigation_log(a.irrig) if getattr(a, "irrig", None) else None
    if a.cmd == "errtable":
        p = W.prepare(a.fcst, a.obs, a.stn, a.coef, a.s4_coef, a.grid)
        year = int(pd.to_datetime(p["ft"].target).min().year)
        rows = W.eto_error_rows(p["ft"], a.stn, year)
        path = coef_file(a.err) if os.path.exists(coef_file(a.err)) else a.err
        W.save_error_table(rows, path)
        d = rows[(rows.kind == "day") & (rows.month == 0)]
        print(f"[오차표] 지점 {a.stn} {year}년 {len(rows)}행 → {path}")
        print(d[["run_name", "lead_day", "n", "rmse", "mbe", "obs_mean"]].round(3).to_string(index=False))
        return
    if a.cmd == "wbverify":
        res = W.run_wbverify(a.fcst, a.obs, a.stn, a.coef, a.s4_coef, a.grid, a.err, irrig=irrig, auto_irrigate=a.auto_irrigate)
        if not len(res["err"]):
            print(f"[경고] {a.err}에 검증 연도({res['year']})를 뺀 다른 해 오차가 없어 범위에 기본 상대 오차 {W.ERR_DEFAULT_REL:.0%}를 씀")
        soil = res["soil"]
        print(f"[물수지] TAW {soil['taw']:.0f} mm, RAW {soil['raw']:.0f} mm, 시나리오 {'관수 규칙' if a.auto_irrigate else '무관수'}"
              + (f", 관수 기록 {len(irrig)}일" if irrig else ""))
        print(res["lead"][["run_name", "order", "lead_day", "n", "RMSE_Dr_center", "MBE_Dr_center", "RMSE_Dr_obs",
                           "RMSE_Dr_pers", "skill_Dr_center"]].round(2).to_string(index=False))
        print(res["fe3_sum"].to_string(index=False))
        for ln in res["findings"]:
            print(ln)
        grid = "-".join(res["check"]["location"]) or "grid"
        t = pd.to_datetime(res["runs"][res["runs"].ok].target)
        tag = "_irrig" if a.auto_irrigate else ""
        out = a.out or f"output/fcst_wbverify({a.stn})_{grid}_{t.min():%Y%m%d}_{t.max():%Y%m%d}{tag}.xlsx"
        from fcst_wb_report import build_wbverify_workbook
        print(f"[완료] {build_wbverify_workbook(res, out)}")
        return
    sv = W.service_prepare(a.fcst, a.obs, a.stn, a.run, a.coef, a.s4_coef, a.grid, a.err, irrig=irrig)
    ol = sv["outlook"]
    raw = sv["soil"]["raw"]
    lab = lambda k: ("지금 필요" if k == 0 else (f"D+{int(ol['days'].lead_day.iloc[k - 1])} ({ol['days'].target.iloc[k - 1]:%m/%d})"
                                              + (" 참고" if k > W.MAIN_DAYS else "")) if k else "기간 안 없음")
    print(f"[발표] {ol['run_name']} {ol['run']:%Y-%m-%d %H시}, 관측 마지막 날 {sv['obs_last']:%Y-%m-%d}, 출발 Dr {ol['start']:.1f} mm (RAW {raw:.0f})")
    print(f"[3일 ETc] {ol['cum3'][0]:.1f} ± {ol['cum3'][1]:.1f} mm, 관수 필요 예상일 {lab(ol['need']['center'])} "
          f"(빠르면 {lab(ol['need']['early'])}, 늦으면 {lab(ol['need']['late'])})")
    out = a.out or f"output/fcst_service({a.stn})_{ol['run']:%Y%m%d_%H}.xlsx"
    from fcst_wb_report import build_service_workbook
    print(f"[완료] {build_service_workbook(sv, out)}")



if __name__ == "__main__":
    main()
