#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fcst_archive.py — 기상자료개방포털 '과거 단기예보(격자)' CSV 파서와 발표별 일 집계 (02-Cycle 1단계)

[입력 파일 형식] (포털 내려받기, 요소·월별 1파일)
  format: day(UTC),hour(UTC),forecast,value,day(KST),hour(KST)  location:73_135 Start : 20260401
  1,0200,+6,17.000000,1,1100
  Start : 20260402
  ...
  - day/hour(KST) : 발표 일·시각 (UTC 열과 같은 시점, 9시간 차이)
  - 1시간 요소(TMP·REH·WSD·PCP): forecast = 발표 후 경과 시간(lead, h). 예보시각 = 발표시각 + lead
      포털 자료는 lead +6 h부터 들어 있음(02시 발표 → 08시부터).
      마지막 날(연장기간)은 3시간 간격이고 WSD·PCP는 코드값(1~3).
  - 일 요소(TMX·TMN): forecast = +6, +7, … 은 대상일 순번(첫 대상일 = +6)
      첫 대상일: TMX는 02·05·08·11시 발표 → 오늘, 14·17·20·23시 → 내일
                 TMN은 02시 발표 → 오늘, 그 외 → 내일          (활용가이드 2026-06-23판 표와 일치)

[일 집계 규칙]  (THEORY.md 9장, ARCHITECTURE.md 7장)
  - 대상 시각마다 '서비스 발표 시각 이전의 가장 최근 발표' 값을 사용
    → 02시 발표의 오늘 00~07시는 전날 17·20·23시 발표 값으로 채움(발표 시점에 알 수 있던 정보만 사용)
  - ea = 시간별 e°(TMP)·REH/100 의 평균, u10 = WSD 평균, 강수 = PCP 합
  - 연장기간 코드: WSD 1 → 같은 발표의 직전 정량일 평균(최대 3.9), 2 → 6.5, 3 → 11 m/s
                   PCP 1 → 1.5, 2 → 9, 3 → 20 mm/h (3시간 칸이므로 ×3)
"""
import csv, datetime as dt, math, os, re
from collections import defaultdict

import pandas as pd

ELEMENTS = ("TMX", "TMN", "TMP", "REH", "WSD", "PCP")
# 파일명 키워드 (포털 기본 파일명: 지역_요소명_시작일_종료일.csv). 순서가 중요: '최고기온'·'최저기온'을 '기온'보다 먼저 검사
_KEYWORDS = [("TMX", ("최고기온",)), ("TMN", ("최저기온",)), ("PCP", ("강수량",)), ("REH", ("습도",)),
             ("WSD", ("풍속",)), ("TMP", ("1시간기온", "기온"))]
WSD_CODE = {2: 6.5, 3: 11.0}          # 코드 1은 직전 정량일 평균(최대 3.9)으로 대체
WSD_CODE1_CAP = 3.9
PCP_CODE_MMH = {1: 1.5, 2: 9.0, 3: 20.0}
RAIN_FLAG_MM = 1.0


def _svp(t):
    return 0.6108 * math.exp(17.27 * t / (t + 237.3))


# ── 파일 읽기 ─────────────────────────────────────────────────────────
def read_portal_csv(path):
    """포털 CSV 1개 → (location 'nx_ny', DataFrame[issue(KST), forecast, value])"""
    loc, recs = None, []
    with open(path, encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            s = line.strip()
            if not s:
                continue
            if s.startswith("format"):
                m = re.search(r"location:\s*(\d+_\d+)", s)
                loc = m.group(1) if m else None
                m = re.search(r"Start\s*:\s*(\d{8})", s)
                ym = m.group(1)[:6] if m else None
                continue
            if s.startswith("Start"):
                continue
            p = [x.strip() for x in s.split(",")]
            if len(p) < 6:
                continue
            day_u, hour_u = int(p[0]), p[1].zfill(4)
            # UTC 일자를 기준으로 발표시각(KST)을 계산: 파일의 연월 + UTC 일(말일 다음 '32일' 등도 안전하게 처리)
            base = dt.datetime.strptime(ym + "01", "%Y%m%d") + dt.timedelta(days=day_u - 1)
            issue_utc = base + dt.timedelta(hours=int(hour_u[:2]))
            recs.append((issue_utc + dt.timedelta(hours=9), int(p[2].replace("+", "")), float(p[3])))
    df = pd.DataFrame(recs, columns=["issue", "forecast", "value"])
    return loc, df


def detect_element(path, df):
    """파일명 키워드 → 없으면 값 분포로 요소 판별"""
    name = os.path.basename(path)
    for elem, keys in _KEYWORDS:
        if any(k in name for k in keys):
            return elem
    per_issue = df.groupby("issue").size()
    v = df["value"]
    if per_issue.max() <= 4:                             # 일 요소: 대상일 수(3행) 패턴으로 구분
        three = set(per_issue[per_issue == 3].index.hour)
        if three == {14}:
            return "TMX"                                 # 일최고: 14시 발표만 오늘 값이 없음
        if three == {5, 8, 11, 14}:
            return "TMN"                                 # 일최저: 02시 발표만 오늘 값 포함, 17~23시는 4일
        return "TMX" if v.mean() >= 20 else "TMN"
    if (v.round(6) % 5 == 0).mean() > 0.99 and v.min() >= 0 and v.max() <= 100 and v.mean() > 20:
        return "REH"
    if (v == 0).mean() > 0.8:
        return "PCP"
    if v.max() <= 30 and (v.round(1) != v.round(0)).mean() > 0.3:
        return "WSD"
    return "TMP"


# ── 아카이브 ───────────────────────────────────────────────────────────
class Archive:
    """요소별 과거 예보 묶음. hourly[elem]: issue, lead, target, value, code(bool); daily[elem]: issue, target_date, value"""

    def __init__(self):
        self.hourly, self.daily, self.location, self.files = {}, {}, set(), defaultdict(list)

    @property
    def issues(self):
        s = set()
        for d in list(self.hourly.values()) + list(self.daily.values()):
            s |= set(d["issue"])
        return sorted(s)

    def add_file(self, path, element=None):
        loc, df = read_portal_csv(path)
        elem = element or detect_element(path, df)
        if loc:
            self.location.add(loc)
        self.files[elem].append(os.path.basename(path))
        if elem in ("TMX", "TMN"):
            new = _daily_targets(df, elem)
            old = self.daily.get(elem)
            self.daily[elem] = new if old is None else pd.concat([old, new]).drop_duplicates(["issue", "target_date"], keep="last")
        else:
            new = _hourly_targets(df)
            old = self.hourly.get(elem)
            self.hourly[elem] = new if old is None else pd.concat([old, new]).drop_duplicates(["issue", "lead"], keep="last")
        return elem


def _daily_targets(df, elem):
    out = df.copy()
    h = out["issue"].dt.hour
    if elem == "TMX":
        first = (~h.isin([2, 5, 8, 11])).astype(int)
    else:
        first = (h != 2).astype(int)
    out["target_date"] = out["issue"].dt.normalize() + pd.to_timedelta(first + out["forecast"] - 6, unit="D")
    return out[["issue", "target_date", "value"]]


def _hourly_targets(df):
    out = df.rename(columns={"forecast": "lead"}).sort_values(["issue", "lead"]).copy()
    out["target"] = out["issue"] + pd.to_timedelta(out["lead"], unit="h")
    # 3시간 간격 구간(연장기간) 표시: 발표별로 lead 간격이 처음 3이 되는 지점부터
    step = out.groupby("issue")["lead"].diff()
    first3 = out.loc[step == 3].groupby("issue")["lead"].min()
    out["code"] = out["lead"] >= out["issue"].map(first3).fillna(10 ** 6) if len(first3) else False
    return out[["issue", "lead", "target", "value", "code"]]


def load_archive(paths, elements=None):
    """paths: CSV 경로 목록(폴더면 *.csv 전체). elements: {경로: 요소}로 판별 결과를 덮어쓸 때 사용"""
    files = []
    for p in paths:
        if os.path.isdir(p):
            files += sorted(os.path.join(p, f) for f in os.listdir(p) if f.lower().endswith(".csv"))
        else:
            files.append(p)
    arch = Archive()
    for f in files:
        arch.add_file(f, (elements or {}).get(f))
    return arch


def check_archive(arch):
    """G1 점검용 구조 확인 결과(dict). 발표 누락, 요소별 발표당 행수 패턴, 격자 일관성."""
    rep = {"location": sorted(arch.location), "elements": {e: len(arch.files[e]) for e in arch.files}}
    iss = arch.issues
    if iss:
        full = pd.date_range(iss[0], iss[-1], freq="3h")
        rep["issues"] = len(iss)
        rep["missing_issues"] = [str(t) for t in full if t not in set(iss)]
    if "TMX" in arch.daily:
        c = arch.daily["TMX"].groupby("issue").size()
        rep["TMX_rows_by_hour"] = c.groupby(c.index.hour).agg(lambda s: sorted(set(s))).to_dict()
    if "TMN" in arch.daily:
        c = arch.daily["TMN"].groupby("issue").size()
        rep["TMN_rows_by_hour"] = c.groupby(c.index.hour).agg(lambda s: sorted(set(s))).to_dict()
    for e, d in arch.hourly.items():
        rep[f"{e}_first_lead"] = int(d["lead"].min())
        rep[f"{e}_code_values"] = sorted(d.loc[d["code"], "value"].unique().tolist())[:10]
    return rep


# ── 발표별 일 집계 ─────────────────────────────────────────────────────
def _latest_as_of(hdf, run):
    """발표시각 run 이전(포함) 자료 중 대상시각별 가장 최근 발표 값"""
    sub = hdf[hdf["issue"] <= run].sort_values("issue")
    return sub.drop_duplicates("target", keep="last").set_index("target")


def daily_inputs(arch, run, targets):
    """서비스 발표 run(Timestamp)의 대상일 목록 → 일 입력 DataFrame
       열: target, lead_day, Tmax, Tmin, ea, u10, rain, rain_flag, hours, ext, filled"""
    run = pd.Timestamp(run)
    sel = {e: _latest_as_of(arch.hourly[e], run) for e in ("TMP", "REH", "WSD", "PCP") if e in arch.hourly}
    dsel = {}
    for e in ("TMX", "TMN"):
        d = arch.daily[e]
        d = d[d["issue"] == run]
        dsel[e] = dict(zip(d["target_date"], d["value"]))
    rows, prev_wsd = [], None
    for tday in targets:
        tday = pd.Timestamp(tday).normalize()
        win = lambda s: s[(s.index >= tday) & (s.index < tday + pd.Timedelta(days=1))]
        tmp, reh, wsd, pcp = (win(sel[e]) if e in sel else None for e in ("TMP", "REH", "WSD", "PCP"))
        ext = bool(wsd is not None and wsd["code"].any())
        # ea: TMP·REH가 모두 있는 시각
        ea = None
        if tmp is not None and reh is not None:
            j = tmp[["value"]].join(reh[["value"]], lsuffix="_t", rsuffix="_h", how="inner")
            if len(j):
                ea = float((j["value_t"].map(_svp) * j["value_h"] / 100).mean())
        # 풍속: 연장기간 코드 변환
        u10 = None
        if wsd is not None and len(wsd):
            vals = []
            for v, c in zip(wsd["value"], wsd["code"]):
                if c:
                    k = int(round(v))
                    vals.append(min(prev_wsd, WSD_CODE1_CAP) if (k == 1 and prev_wsd is not None) else
                                (WSD_CODE1_CAP if k == 1 else WSD_CODE.get(k, v)))
                else:
                    vals.append(v)
            u10 = float(sum(vals) / len(vals))
            if not ext:
                prev_wsd = u10
        # 강수: 1시간 값은 mm, 코드 칸은 3시간 × 대표강도
        rain = None
        if pcp is not None and len(pcp):
            rain = float(sum((3 * PCP_CODE_MMH.get(int(round(v)), 0.0)) if c else v
                             for v, c in zip(pcp["value"], pcp["code"])))
        hours = int(len(tmp)) if tmp is not None else 0
        filled = int((tmp["issue"] < run).sum()) if tmp is not None else 0
        rows.append(dict(target=tday, lead_day=int((tday - run.normalize()).days), Tmax=dsel["TMX"].get(tday),
                         Tmin=dsel["TMN"].get(tday), ea=ea, u10=u10, rain=rain,
                         rain_flag=(None if rain is None else int(rain >= RAIN_FLAG_MM)),
                         hours=hours, ext=ext, filled=filled))
    return pd.DataFrame(rows)


SERVICE_RUNS = {"아침": (2, range(0, 4)), "저녁": (17, range(1, 5))}   # 발표시각, 대상일(D+k)


def service_table(arch, runs=SERVICE_RUNS):
    """모든 서비스 발표(02시·17시)에 대해 대상일별 일 입력을 한 표로"""
    out = []
    for name, (hour, leads) in runs.items():
        for run in [t for t in arch.issues if t.hour == hour]:
            df = daily_inputs(arch, run, [run.normalize() + pd.Timedelta(days=k) for k in leads])
            df.insert(0, "run", run); df.insert(0, "run_name", name)
            out.append(df)
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame()
