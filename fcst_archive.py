#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fcst_archive.py — 과거 단기예보 CSV 파서(기상자료개방포털 · OpenAPI 응답)와 발표별 일 집계 (02-Cycle 1단계)

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

[선택 요소] 하늘상태(SKY)·강수확률(POP) — 있으면 일 집계에 넣는다(발표 완전성 판단에는 쓰지 않음)
  - SKY 코드 1 맑음, 3 구름많음, 4 흐림(그 밖의 값은 결측). 낮 시간 일사 비중(태양고도 사인)으로 가중해
    '구름많음 비율'·'흐림 비율'(0~1)을 만든다. POP는 같은 가중의 낮 평균과 하루 최대(0~1).
  - 서비스 발표 자체에 SKY가 없으면 이전 발표로 채우지 않고 비워 둔다.

[결측]
  - 값이 ±900 이상(예: −999.9)이면 결측으로 버린다.
  - 서비스 발표는 필수 6요소가 모두 있는 발표만 쓴다. 빠진 발표는 service_table()이 arch.skipped_runs에 기록.

[여러 달 파일] 'Start : YYYYMMDD'(UTC 날짜) 행마다 연월을 갱신한다(행의 day(UTC)는 그 달의 일자).

[OpenAPI 응답 CSV] 단기예보 조회서비스(VilageFcstInfoService_2.0 getVilageFcst) 응답 항목을 모은 CSV도 읽는다.
  baseDate,baseTime,category,fcstDate,fcstTime,fcstValue,nx,ny
  20250401,200,TMP,20250401,300,0,73,134
  - 한 파일에 모든 요소(category)가 들어 있다. 발표·예보 일시는 KST. 발표는 하루 8회(02·05·…·23시)
  - lead = 예보시각 − 발표시각(h). API는 발표 1시간 뒤부터 준다(포털은 6시간 뒤부터)
      → 02시 발표의 오늘 00~02시만 전날 23시 발표로 채운다(03~07시는 그 발표 자체의 값)
  - TMX·TMN: 대상일 = fcstDate (TMN 06시, TMX 15시 칸)
  - PCP 문자열: "강수없음" 0, "1mm 미만" PCP_LT1_MM, "30.0~50.0mm" 40, "50.0mm 이상" 50, "3.0mm" 3.
    숫자로만 온 값(글피 1시간 칸의 0.1~ 소수, 연장기간 코드 0~3)은 그대로
  - 빈 줄(',,,,,,,')·반복된 머리행은 건너뛰고, 같은 발표·요소·예보시각이 두 번 있으면 뒤의 값을 쓴다

[요소별 KST CSV] 요소마다 1파일, 일시는 KST, 값 표기는 OpenAPI와 같다(강수 문자열, lead +1부터).
  발표일,발표시각,예보일,예보시각,값
  20250401,0200,20250401,0300,4
  - 요소: 파일명 키워드 → 없으면 값으로 판별(TMN은 예보시각 06시만, TMX는 15시만, 강수는 '강수없음' 등 문자열,
    음수가 섞인 소수는 바람성분 UUU·VVV). 쓰지 않는 요소(UUU·VVV·VEC·PTY·SNO·WAV) 파일은 건너뛰고 기록한다
[인코딩] UTF-8(BOM 포함)이 아니면 CP949(엑셀 저장본)로 읽는다
"""
import csv, datetime as dt, math, os, re
from collections import defaultdict

import numpy as np
import pandas as pd

ELEMENTS = ("TMX", "TMN", "TMP", "REH", "WSD", "PCP")   # 필수 6요소
OPTIONAL = ("SKY", "POP")                              # 선택: 하늘상태, 강수확률
SKY_VALID = (1.0, 3.0, 4.0)                            # 하늘상태 코드: 1 맑음, 3 구름많음, 4 흐림
# 파일명 키워드 (포털 기본 파일명: 지역_요소명_시작일_종료일.csv). 순서가 중요: '최고기온'·'최저기온'을 '기온'보다 먼저 검사
_KEYWORDS = [("TMX", ("최고기온",)), ("TMN", ("최저기온",)), ("PCP", ("강수량",)), ("POP", ("강수확률",)),
             ("SKY", ("하늘상태",)), ("REH", ("습도",)), ("WSD", ("풍속",)),
             ("UUU", ("동서바람성분",)), ("VVV", ("남북바람성분",)), ("VEC", ("풍향",)), ("PTY", ("강수형태",)),
             ("SNO", ("적설",)), ("WAV", ("파고",)), ("TMP", ("1시간기온", "기온"))]
USED = ("TMX", "TMN", "TMP", "REH", "WSD", "PCP", "SKY", "POP")   # 읽어 쓰는 요소(필수 6 + 선택 2)
WSD_CODE = {2: 6.5, 3: 11.0}          # 코드 1은 직전 정량일 평균(최대 3.9)으로 대체
WSD_CODE1_CAP = 3.9
PCP_CODE_MMH = {1: 1.5, 2: 9.0, 3: 20.0}
RAIN_FLAG_MM = 1.0
MISSING_ABS = 900.0                   # 활용가이드: +900 이상 / −900 이하는 결측 (포털 CSV는 −999.9)
# OpenAPI 강수 문자열 → mm/h (ARCHITECTURE 7장 '강수 문자열'). 포털 과거자료는 정수 mm라 '1mm 미만'이 0으로 보임
API_COLS = ("baseDate", "baseTime", "category", "fcstDate", "fcstTime", "fcstValue")
KST_COLS = ("발표일", "발표시각", "예보일", "예보시각", "값")
PCP_LT1_MM = 0.5                      # "1mm 미만" (0.1~0.9 mm)
PCP_30_50_MM = 40.0                   # "30.0~50.0mm"
PCP_GE50_MM = 50.0                    # "50.0mm 이상"


def _svp(t):
    return 0.6108 * math.exp(17.27 * t / (t + 237.3))


# ── 파일 읽기 ─────────────────────────────────────────────────────────
def read_portal_csv(path):
    """포털 CSV 1개 → (location 'nx_ny', DataFrame[issue(KST), forecast, value])
       결측값(|값| ≥ 900, 예: −999.9) 행은 버리고 개수를 df.attrs["n_missing"]에 남긴다."""
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
                # 여러 달을 한 파일로 받은 경우: 'Start : YYYYMMDD'(UTC 날짜)마다 연월을 갱신한다.
                # 행의 day(UTC)는 그 연월의 일자다.
                m = re.search(r"Start\s*:\s*(\d{8})", s)
                if m:
                    ym = m.group(1)[:6]
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
    bad = df["value"].abs() >= MISSING_ABS
    df = df[~bad].reset_index(drop=True)
    df.attrs["n_missing"] = int(bad.sum())
    return loc, df


def keyword_element(path):
    """파일명 키워드로 요소 판별(없으면 None). 업로드 과정에서 한글 파일명이 '_'로 바뀌면 판별하지 못함"""
    name = os.path.basename(path)
    for elem, keys in _KEYWORDS:
        if any(k in name for k in keys):
            return elem
    return None


def detect_element(path, df):
    """파일명 키워드 → 없으면 값 분포로 요소 판별"""
    return keyword_element(path) or _detect_values(df)


def _detect_values(df):
    """값 분포로 요소 판별. df: issue, forecast, value"""
    per_issue = df.groupby("issue").size()
    v = df["value"]
    if per_issue.max() <= 4:                             # 일 요소: 대상일 수(3행) 패턴으로 구분
        three = set(per_issue[per_issue == 3].index.hour)
        if three == {14}:
            return "TMX"                                 # 일최고: 14시 발표만 오늘 값이 없음
        if three == {5, 8, 11, 14}:
            return "TMN"                                 # 일최저: 02시 발표만 오늘 값 포함, 17~23시는 4일
        return "TMX" if v.mean() >= 20 else "TMN"
    r = v.round(6)
    if r.isin(SKY_VALID).mean() > 0.99 and r.isin(SKY_VALID + (0.0, 2.0)).all():
        return "SKY"                                     # 하늘상태: 코드 1·3·4 (드물게 잘못된 0)
    nz = r[r > 0]
    if len(nz) and v.min() == 0 and v.max() <= 100 and v.max() >= 20 and (v > 0).mean() > 0.05 \
            and (nz % 10 == 0).mean() > 0.9:
        return "POP"                                     # 강수확률: 0 포함, 0이 아닌 값은 거의 10% 단위(API 글피 칸에 66 등 드물게)
    if v.min() >= 0 and v.max() > 100:
        return "VEC"                                     # 풍향: 0~360°
    if (v.round(6) % 5 == 0).mean() > 0.99 and v.min() >= 0 and v.max() <= 100 and v.mean() > 20:
        return "REH"                                     # 습도: 5% 단위, 0~100
    if v.min() >= 0 and v.median() == 0:
        return "PCP"                                     # 강수: 음수 없음, 비 오는 시각은 절반 미만(장마철 포함)
    if v.max() <= 30 and (v.round(1) != v.round(0)).mean() > 0.3:
        return "WSD"                                     # 풍속: 0.1 m/s 단위 소수
    return "TMP"                                         # 기온: 정수 ℃


def _encoding(path):
    """UTF-8(BOM 포함)로 읽히면 'utf-8-sig', 아니면 'cp949'(한글 엑셀에서 CSV로 저장한 파일)"""
    with open(path, "rb") as f:
        b = f.read(1 << 16)
    try:
        b.decode("utf-8")
    except UnicodeDecodeError as e:
        if e.start < len(b) - 3:                         # 끝에서 잘린 글자가 아니면 UTF-8이 아님
            return "cp949"
    return "utf-8-sig"


def csv_format(path):
    """첫 줄(빈 줄 제외)로 파일 형식 판별: 'openapi'(응답 열 이름) / 'element'(발표일·…·값) / 'portal'"""
    with open(path, encoding=_encoding(path), errors="replace") as f:
        for line in f:
            s = line.strip().lstrip("\ufeff").replace('"', "")
            if s:
                cols = [c.strip() for c in s.split(",")]
                if all(c in cols for c in API_COLS):
                    return "openapi"
                if all(c in cols for c in KST_COLS):
                    return "element"
                return "portal"
    return "portal"


def is_openapi_csv(path):
    """첫 줄(빈 줄 제외)이 OpenAPI 응답 열 이름(baseDate, …, fcstValue)이면 True"""
    return csv_format(path) == "openapi"


def pcp_mm(s):
    """OpenAPI 강수(PCP) 값 → mm/h. 읽을 수 없으면 NaN"""
    s = str(s).strip()
    if s in ("강수없음", "-", ""):
        return 0.0
    if "미만" in s:
        return PCP_LT1_MM
    if "~" in s:
        return PCP_30_50_MM
    if "이상" in s:
        return PCP_GE50_MM
    try:
        return float(s.replace("mm", ""))
    except ValueError:
        return float("nan")


def read_openapi_csv(path, elements=ELEMENTS + OPTIONAL):
    """OpenAPI 응답 CSV 1개 → (격자 집합 {'nx_ny'}, {요소: DataFrame})
       1시간 요소: DataFrame[issue, forecast(= lead h), value] — read_portal_csv와 같은 모양
       TMX·TMN : DataFrame[issue, target_date, value] (대상일 = fcstDate)
       결측(|값| ≥ 900, 읽을 수 없는 값) 행은 버리고 개수를 df.attrs["n_missing"]에 남긴다."""
    d = pd.read_csv(path, encoding=_encoding(path), dtype=str, skipinitialspace=True)
    d.columns = [c.strip().lstrip("\ufeff") for c in d.columns]
    d = d.dropna(subset=["baseDate", "category"])
    d = d[d["baseDate"].str.strip() != "baseDate"]                 # 파일을 이어 붙일 때 반복된 머리행
    d = d.assign(**{c: d[c].str.strip() for c in API_COLS})
    d = d.drop_duplicates(["baseDate", "baseTime", "category", "fcstDate", "fcstTime"], keep="last")
    locs = set((d["nx"].str.strip() + "_" + d["ny"].str.strip()).unique()) if {"nx", "ny"} <= set(d.columns) else set()
    issue = pd.to_datetime(d["baseDate"] + d["baseTime"].str.zfill(4), format="%Y%m%d%H%M")
    target = pd.to_datetime(d["fcstDate"] + d["fcstTime"].str.zfill(4), format="%Y%m%d%H%M")
    out = {}
    for e in elements:
        m = (d["category"] == e).values
        if not m.any():
            continue
        raw = d.loc[m, "fcstValue"]
        if e == "PCP":
            conv = {s: pcp_mm(s) for s in raw.unique()}
            v = raw.map(conv).astype(float)
        else:
            v = pd.to_numeric(raw, errors="coerce")
        bad = (v.isna() | (v.abs() >= MISSING_ABS)).values
        iss, tgt, val = issue[m][~bad], target[m][~bad], v[~bad]
        if e in ("TMX", "TMN"):
            df = pd.DataFrame({"issue": iss.values, "target_date": tgt.dt.normalize().values, "value": val.values})
        else:
            lead = ((tgt - iss) / pd.Timedelta(hours=1)).round().astype(int)
            df = pd.DataFrame({"issue": iss.values, "forecast": lead.values, "value": val.values})
        df.attrs["n_missing"] = int(bad.sum())
        out[e] = df.reset_index(drop=True)
    return locs, out


def _detect_kst(raw, num, df, target):
    """요소별 KST CSV의 요소를 값으로 판별. raw: 문자열 값, num: 숫자로 바꾼 값, df: issue·forecast·value"""
    h = target.dt.hour
    if h.nunique() == 1 and h.iloc[0] in (6, 15):
        return "TMN" if h.iloc[0] == 6 else "TMX"         # 하루 1칸: 최저 06시, 최고 15시
    txt = raw[num.isna()]
    if txt.str.contains("강수|mm").any():
        return "PCP"                                     # '강수없음', '1mm 미만', '2.0mm' …
    if txt.str.contains("적설|cm").any():
        return "SNO"
    v = num.dropna()
    if not len(v) or (v.abs() >= MISSING_ABS).mean() > 0.99:
        return "WAV"                                     # 전부 결측(−999): 육지 격자의 파고
    v = v[v.abs() < MISSING_ABS]
    if (v < 0).mean() > 0.05 and (v.round(0) != v).mean() > 0.2:
        return "UUU/VVV"                                 # 음수가 섞인 소수: 동서·남북 바람성분(값으로는 둘을 구분 못 함)
    if v.isin([0, 1, 2, 3, 4]).all() and (v == 0).mean() > 0.5:
        return "PTY"                                     # 강수형태 코드 0~4 (강수는 문자열이라 여기 오지 않음)
    return _detect_values(df.assign(value=num).dropna(subset=["value"]))


def read_element_csv(path, element=None):
    """요소별 KST CSV 1개(발표일,발표시각,예보일,예보시각,값) → (요소, DataFrame)
       1시간 요소: [issue, forecast(= lead h), value], TMX·TMN: [issue, target_date, value].
       요소는 element → 파일명 키워드 → 값 분포 순으로 정한다. 결측(|값| ≥ 900, 읽을 수 없는 값)은 버리고 개수를 attrs에."""
    d = pd.read_csv(path, encoding=_encoding(path), dtype=str, skipinitialspace=True)
    d.columns = [c.strip().lstrip("\ufeff") for c in d.columns]
    d = d.dropna(subset=["발표일", "값"])
    d = d[d["발표일"].str.strip() != "발표일"]                      # 이어 붙인 파일의 반복된 머리행
    d = d.assign(**{c: d[c].str.strip() for c in KST_COLS})
    d = d.drop_duplicates(["발표일", "발표시각", "예보일", "예보시각"], keep="last")
    issue = pd.to_datetime(d["발표일"] + d["발표시각"].str.zfill(4), format="%Y%m%d%H%M")
    target = pd.to_datetime(d["예보일"] + d["예보시각"].str.zfill(4), format="%Y%m%d%H%M")
    raw = d["값"]
    num = pd.to_numeric(raw, errors="coerce")
    lead = ((target - issue) / pd.Timedelta(hours=1)).round().astype(int)
    elem = element or keyword_element(path) or _detect_kst(
        raw, num, pd.DataFrame({"issue": issue.values, "forecast": lead.values}), target)
    v = raw.map({x: pcp_mm(x) for x in raw.unique()}).astype(float) if elem == "PCP" else num
    bad = (v.isna() | (v.abs() >= MISSING_ABS)).values
    if elem in ("TMX", "TMN"):
        df = pd.DataFrame({"issue": issue[~bad].values, "target_date": target[~bad].dt.normalize().values,
                           "value": v[~bad].values})
    else:
        df = pd.DataFrame({"issue": issue[~bad].values, "forecast": lead[~bad].values, "value": v[~bad].values})
    df.attrs["n_missing"] = int(bad.sum())
    return elem, df.reset_index(drop=True)


# ── 아카이브 ───────────────────────────────────────────────────────────
class Archive:
    """요소별 과거 예보 묶음. hourly[elem]: issue, lead, target, value, code(bool); daily[elem]: issue, target_date, value"""

    def __init__(self):
        self.hourly, self.daily, self.location, self.files = {}, {}, set(), defaultdict(list)
        self.n_missing = defaultdict(int)        # 요소별 결측값(±900) 행 수
        self.spans = defaultdict(list)           # 요소별 파일의 (첫 발표, 마지막 발표, 파일명) — 판별 중복 확인용
        self.formats = set()                     # 읽은 파일 형식: "portal"(기상자료개방포털) / "openapi"(조회서비스 응답) / "element"(요소별 KST)
        self.ignored = []                        # 쓰지 않는 요소라 건너뛴 파일: (파일명, 요소)

    @property
    def issues(self):
        s = set()
        for d in list(self.hourly.values()) + list(self.daily.values()):
            s |= set(d["issue"])
        return sorted(s)

    def issue_set(self, elem):
        """요소별 발표시각 집합(캐시)"""
        c = self.__dict__.setdefault("_iss_cache", {})
        if elem not in c:
            d = self.hourly.get(elem) if elem in self.hourly else self.daily.get(elem)
            c[elem] = set(d["issue"]) if d is not None else set()
        return c[elem]

    @property
    def required_issues(self):
        """필수 6요소 중 하나에라도 있는 발표시각"""
        s = set()
        for e, d in list(self.hourly.items()) + list(self.daily.items()):
            if e in ELEMENTS:
                s |= set(d["issue"])
        return sorted(s)

    @property
    def complete_issues(self):
        """필수 6요소에 모두 값이 있는 발표시각 (서비스 발표는 이 중에서만 고른다. 선택 요소 SKY·POP은 보지 않음)"""
        sets = [set(d["issue"]) for e, d in list(self.hourly.items()) + list(self.daily.items()) if e in ELEMENTS]
        return sorted(set.intersection(*sets)) if sets else []

    def add_file(self, path, element=None):
        """CSV 1개를 더한다. 반환: 요소 이름(요소별 파일) 또는 요소 이름 튜플(OpenAPI 파일 — 한 파일에 여러 요소).
           쓰지 않는 요소(UUU·VVV 등)의 파일은 저장하지 않고 self.ignored에 남긴다."""
        self._iss_cache = {}
        fmt = csv_format(path)
        self.formats.add(fmt)
        if fmt == "openapi":
            locs, tables = read_openapi_csv(path)
            self.location |= locs
            for elem, df in tables.items():
                self._store(elem, df, path)
            return tuple(tables)
        if fmt == "element":
            elem, df = read_element_csv(path, element)
        else:
            loc, df = read_portal_csv(path)
            elem = element or detect_element(path, df)
            if loc:
                self.location.add(loc)
        if elem not in USED:
            self.ignored.append((os.path.basename(path), elem))
            return elem
        self._store(elem, df, path)
        return elem

    def _store(self, elem, df, path):
        """요소 1개의 표(DataFrame[issue, forecast, value] 또는 TMX·TMN의 [issue, target_date, value])를 아카이브에 합친다"""
        self.files[elem].append(os.path.basename(path))
        self.n_missing[elem] += df.attrs.get("n_missing", 0)
        if elem == "SKY":                        # 코드표에 없는 값(예: 0)은 결측
            bad = ~df["value"].round(6).isin(SKY_VALID)
            self.n_missing[elem] += int(bad.sum())
            df = df[~bad].reset_index(drop=True)
        if len(df):
            self.spans[elem].append((df["issue"].min(), df["issue"].max(), os.path.basename(path)))
        if elem in ("TMX", "TMN"):
            new = df[["issue", "target_date", "value"]] if "target_date" in df else _daily_targets(df, elem)
            old = self.daily.get(elem)
            self.daily[elem] = new if old is None else pd.concat([old, new]).drop_duplicates(["issue", "target_date"], keep="last")
        else:
            new = _hourly_targets(df)
            old = self.hourly.get(elem)
            self.hourly[elem] = new if old is None else pd.concat([old, new]).drop_duplicates(["issue", "lead"], keep="last")


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
    rep = {"location": sorted(arch.location), "elements": {e: len(arch.files[e]) for e in arch.files},
           "formats": sorted(getattr(arch, "formats", set())),
           "missing_elements": [e for e in ELEMENTS if e not in arch.hourly and e not in arch.daily],
           "ignored_files": list(getattr(arch, "ignored", []))}
    iss = arch.required_issues
    if iss:
        full = pd.date_range(iss[0], iss[-1], freq="3h")
        comp = set(arch.complete_issues)
        rep["first_issue"], rep["last_issue"] = str(iss[0]), str(iss[-1])
        rep["issues"] = len(comp)                                        # 6요소가 모두 있는 발표 수
        rep["missing_issues"] = [str(t) for t in full if t not in comp]  # 한 요소라도 없는 발표
        cov = {}
        for e in ELEMENTS + tuple(x for x in OPTIONAL if x in arch.hourly):
            d = arch.hourly.get(e) if e in arch.hourly else arch.daily.get(e)
            have = set(d["issue"]) if d is not None else set()
            cov[e] = dict(files=len(arch.files.get(e, [])), first=str(min(have)) if have else "",
                          last=str(max(have)) if have else "", issues=len(have),
                          missing_issues=sum(1 for t in full if t not in have), missing_values=arch.n_missing.get(e, 0))
        rep["coverage"] = cov
        # 같은 요소로 판별된 파일의 발표 범위가 겹치면 요소 판별 오류일 수 있음
        warn = []
        for e, sp in arch.spans.items():
            sp = sorted(sp)
            for (a0, a1, fa), (b0, b1, fb) in zip(sp, sp[1:]):
                if b0 <= a1:
                    warn.append(f"{e}: {fa} ↔ {fb} 발표 범위 겹침(요소 판별 또는 중복 파일 확인 — 겹친 칸은 나중 파일 값)")
        rep["warnings"] = warn
        # 일부만 받은 발표: 같은 발표시각(시)의 보통 행 수(최빈값)보다 적은 발표 — 내려받기 중단·편집 흔적 확인용
        short = []
        for e in ELEMENTS + OPTIONAL:
            d = arch.hourly.get(e) if e in arch.hourly else arch.daily.get(e)
            if d is None or not len(d):
                continue
            c = d.groupby("issue").size()
            mode = c.groupby(c.index.hour).agg(lambda s: int(s.mode().iloc[0]))
            for t, n in c.items():
                if n < mode[t.hour]:
                    short.append((str(t), e, int(n), int(mode[t.hour])))
        rep["short_issues"] = sorted(short)
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


def sun_weights(times, lat=None, lon=None):
    """시각별 태양고도의 사인값(밤은 0) — 낮 시간 하늘상태·강수확률을 일사 비중으로 평균할 때의 가중치.
       KST(동경 135°) 시각 → 태양시 = 시각 + (경도 − 135)/15 + 균시차[FAO-56 식32]. 위도가 없으면 06~18시 같은 가중치."""
    t = pd.DatetimeIndex(times)
    if lat is None:
        return np.where((t.hour >= 6) & (t.hour <= 18), 1.0, 0.0)
    J = t.dayofyear.values
    b = 2 * np.pi * (J - 81) / 364
    sc = 0.1645 * np.sin(2 * b) - 0.1255 * np.cos(b) - 0.025 * np.sin(b)
    solar = t.hour.values + t.minute.values / 60 + ((135.0 if lon is None else lon) - 135.0) / 15 + sc
    omega = np.pi / 12 * (solar - 12)
    phi, dec = np.radians(lat), 0.409 * np.sin(2 * np.pi * J / 365 - 1.39)
    return np.maximum(np.sin(phi) * np.sin(dec) + np.cos(phi) * np.cos(dec) * np.cos(omega), 0.0)


def daily_inputs(arch, run, targets, lat=None, lon=None):
    """서비스 발표 run(Timestamp)의 대상일 목록 → 일 입력 DataFrame
       열: target, lead_day, Tmax, Tmin, ea, u10, rain, rain_flag, hours, ext, filled,
           (SKY·POP이 있으면) sky_cloudy, sky_overcast, pop, pop_max"""
    run = pd.Timestamp(run)
    sel = {e: _latest_as_of(arch.hourly[e], run) for e in ("TMP", "REH", "WSD", "PCP") if e in arch.hourly}
    # 선택 요소는 서비스 발표 자체에 있을 때만 쓴다(없는 발표를 이전 발표로 통째로 대신하지 않음)
    osel = {e: _latest_as_of(arch.hourly[e], run) for e in OPTIONAL
            if e in arch.hourly and run in arch.issue_set(e)}
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
        row = dict(target=tday, lead_day=int((tday - run.normalize()).days), Tmax=dsel["TMX"].get(tday),
                   Tmin=dsel["TMN"].get(tday), ea=ea, u10=u10, rain=rain,
                   rain_flag=(None if rain is None else int(rain >= RAIN_FLAG_MM)),
                   hours=hours, ext=ext, filled=filled)
        # 하늘상태: 낮 시간 일사 비중으로 가중한 구름많음·흐림 비율 / 강수확률: 같은 가중의 낮 평균, 하루 최대
        for e in OPTIONAL:
            if e not in arch.hourly:
                continue
            x = win(osel[e]) if e in osel else None
            w = sun_weights(x.index, lat, lon) if x is not None and len(x) else np.zeros(0)
            ok = w.sum() > 0
            if e == "SKY":
                v = x["value"].round().values if ok else None
                row["sky_cloudy"] = float((w * (v == 3)).sum() / w.sum()) if ok else None
                row["sky_overcast"] = float((w * (v == 4)).sum() / w.sum()) if ok else None
            else:
                row["pop"] = float((w * x["value"].values).sum() / w.sum() / 100) if ok else None
                row["pop_max"] = float(x["value"].max() / 100) if ok else None
                # 기대 강수량(G4 비교용): 시각마다 강수량 × 강수확률. 예보 강수량은 비가 온다면의 양이고(강수확률 60% 이상인 시각에만 값이 있음)
                # 강수확률을 곱하면 기댓값이 된다. 강수확률이 없는 시각은 1로 둔다
                if pcp is not None and len(pcp):
                    amt = np.where(pcp["code"].values, 3 * np.array([PCP_CODE_MMH.get(int(round(v)), 0.0) for v in pcp["value"]]),
                                   pcp["value"].values.astype(float))
                    pp = (x["value"].reindex(pcp.index).values / 100) if x is not None else np.full(len(pcp), np.nan)
                    row["rain_exp"] = float(np.sum(amt * np.where(np.isnan(pp), 1.0, pp)))
                else:
                    row["rain_exp"] = None
        rows.append(row)
    return pd.DataFrame(rows)


SERVICE_RUNS = {"아침": (2, range(0, 4)), "저녁": (17, range(1, 5))}   # 발표시각, 대상일(D+k)


def service_table(arch, runs=SERVICE_RUNS, lat=None, lon=None):
    """모든 서비스 발표(02시·17시)에 대해 대상일별 일 입력을 한 표로.
       한 요소라도 값이 없는 발표는 쓰지 않고 arch.skipped_runs에 (구분, 발표시각, 빠진 요소)를 남긴다.
       (다른 발표로 대신 채우면 '그 발표의 예보' 검증이 아니게 되므로)"""
    out, comp, iss = [], set(arch.complete_issues), arch.required_issues
    arch.skipped_runs = []
    full = pd.date_range(iss[0], iss[-1], freq="3h") if iss else []
    have = {e: set(d["issue"]) for e, d in list(arch.hourly.items()) + list(arch.daily.items())}
    for name, (hour, leads) in runs.items():
        for run in [t for t in full if t.hour == hour]:
            if run not in comp:
                arch.skipped_runs.append((name, run, [e for e in ELEMENTS if run not in have.get(e, set())]))
                continue
            df = daily_inputs(arch, run, [run.normalize() + pd.Timedelta(days=k) for k in leads], lat, lon)
            df.insert(0, "run", run); df.insert(0, "run_name", name)
            out.append(df)
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame()
