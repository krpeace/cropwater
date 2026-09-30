#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kma_fcst.py — 운영 수집기의 API·저장 계층 (02-Cycle 5단계, G5 · H3)

[하는 일]
  1) 단기예보 조회서비스(VilageFcstInfoService_2.0 getVilageFcst) 한 발표를 페이지별로 받는다(fetch_vilage)
     - 응답 코드를 나눔: 00 정상 / 03 자료 없음(아직 제공 전 → 다시 요청) / 서버·시간 초과(다시 요청) /
       인증키·요청 오류(10~33, HTTP 401·403 → 다시 요청해도 소용없음)
  2) 받은 발표의 완결성 점검(check_issue): 발표시각별 1시간 요소 행 수·최고/최저기온 대상일 수 (THEORY 9장 ◆ 운영 수집)
  3) 원자료 보관: 응답 본문을 받은 그대로 gzip으로 저장(data/ops/raw/...)
  4) 완결된 발표만 OpenAPI 형식 월별 CSV에 쌓음(fcst_archive가 그대로 읽는 형식). 하늘상태·강수확률이 덜 차면 두 요소는 빼고 쌓음(→ S3)
  5) ASOS 일자료(AsosDalyInfoService getWthrDataList) 최근 며칠을 받아 관측 캐시 CSV에 덮어씀(fetch_asos_days, Store.upsert_obs)
  6) 수집 로그(시도마다 한 줄, collect_log.csv)

HTTP는 http_get(url, timeout) → (상태코드, 본문 bytes) 함수로 주입한다. 기본은 requests. 시험에서는 모의 API(ops_replay.py)를 넣는다.
근거: docs/THEORY.md 9장 ◆ 운영 수집, 설계: docs/ARCHITECTURE.md 7장, 판정: docs/VALIDATION.md G5
"""
import csv
import datetime as dt
import gzip
import json
import os
import re
import time
import urllib.parse

import pandas as pd

from fcst_archive import ELEMENTS, MISSING_ABS, OPTIONAL, SKY_VALID, pcp_mm

VILAGE_URL = "http://apis.data.go.kr/1360000/VilageFcstInfoService_2.0/getVilageFcst"
ASOS_URL = "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"
NUM_ROWS = 1000                  # 한 페이지 행 수 (발표당 약 800~1,100행 → 1~2페이지)
TIMEOUT_S = 30
ISSUE_HOURS = (2, 5, 8, 11, 14, 17, 20, 23)
# 발표시각별 정상 행 구조 (2025·2026년 OpenAPI 자료, 모든 발표가 같음 — THEORY 9장 ◆ 운영 수집)
HOURLY_N = {2: 78, 5: 75, 8: 72, 11: 69, 14: 66, 17: 87, 20: 84, 23: 81}
TMX_N = {2: 4, 5: 4, 8: 4, 11: 4, 14: 3, 17: 4, 20: 4, 23: 4}
TMN_N = {2: 4, 5: 3, 8: 3, 11: 3, 14: 3, 17: 4, 20: 4, 23: 4}
HOURLY = ("TMP", "REH", "WSD", "PCP", "SKY", "POP")
FATAL_CODES = {"10", "11", "12", "20", "21", "30", "31", "32", "33"}   # 요청·인증 오류: 다시 요청해도 소용없음
NODATA_CODES = {"03"}
API_COLS = ["baseDate", "baseTime", "category", "fcstDate", "fcstTime", "fcstValue", "nx", "ny"]
ASOS_FIELDS = ["tm", "maxTa", "minTa", "avgTa", "avgRhm", "minRhm", "avgWs", "avgPv", "avgTd", "avgPa",
               "sumGsr", "sumSsHr", "sumLrgEv", "sumRn"]
LOG_COLS = ["logged_at", "slot", "kind", "target", "attempt", "http", "code", "msg", "n_items", "total", "status",
            "required_ok", "complete8", "elapsed_s", "note"]


# ── 시각 (KST, 한국은 서머타임 없음) ──────────────────────────────────────
def kst_now():
    """현재 KST(시간대 정보 없는 datetime). 컴퓨터 시간대 설정과 무관하게 UTC + 9시간"""
    return (dt.datetime.now(dt.timezone.utc) + dt.timedelta(hours=9)).replace(tzinfo=None, microsecond=0)


class Clock:
    """실제 시계. 모의 재생에서는 ops_replay.SimClock으로 바꾼다"""
    def now(self):
        return kst_now()

    def sleep(self, sec):
        if sec > 0:
            time.sleep(sec)


# ── HTTP ─────────────────────────────────────────────────────────────────
def http_get(url, timeout=TIMEOUT_S):
    import requests
    r = requests.get(url, timeout=timeout)
    return r.status_code, r.content


def service_key(key):
    """공공데이터포털 키: 인코딩 키(% 포함)는 그대로, 디코딩 키는 URL 인코딩해서 쓴다"""
    key = (key or "").strip()
    return key if "%" in key else urllib.parse.quote(key, safe="")


def vilage_url(key, issue, nx, ny, page=1, rows=NUM_ROWS):
    issue = pd.Timestamp(issue)
    return (f"{VILAGE_URL}?serviceKey={service_key(key)}&pageNo={page}&numOfRows={rows}&dataType=JSON"
            f"&base_date={issue:%Y%m%d}&base_time={issue:%H}00&nx={int(nx)}&ny={int(ny)}")


def asos_url(key, stn, start, end, rows=50):
    return (f"{ASOS_URL}?serviceKey={service_key(key)}&pageNo=1&numOfRows={rows}&dataType=JSON"
            f"&dataCd=ASOS&dateCd=DAY&startDt={pd.Timestamp(start):%Y%m%d}&endDt={pd.Timestamp(end):%Y%m%d}&stnIds={stn}")


def mask_key(url):
    return re.sub(r"serviceKey=[^&]*", "serviceKey=***", url)


def parse_response(body):
    """응답 본문(bytes) → dict(code, msg, items, total). JSON이 아니면(인증 오류는 XML로 오기도 함) XML에서 코드·메시지를 찾는다"""
    text = body.decode("utf-8", "replace") if isinstance(body, (bytes, bytearray)) else str(body)
    try:
        js = json.loads(text)
        resp = js.get("response", js)
        hdr = resp.get("header", {}) or {}
        b = resp.get("body") or {}
        items = (b.get("items") or {})
        items = items.get("item", []) if isinstance(items, dict) else []
        items = items if isinstance(items, list) else [items]
        return dict(code=str(hdr.get("resultCode", "?")), msg=str(hdr.get("resultMsg", "")), items=items,
                    total=int(b.get("totalCount", len(items)) or 0))
    except (ValueError, AttributeError):
        code = re.search(r"<(?:resultCode|returnReasonCode)>\s*([^<\s]+)\s*<", text)
        msg = re.search(r"<(?:resultMsg|returnAuthMsg|errMsg)>\s*([^<]+?)\s*<", text)
        return dict(code=code.group(1) if code else "?", msg=(msg.group(1) if msg else text[:120].replace("\n", " ")),
                    items=[], total=0)


def classify(http, code):
    """HTTP 상태·응답 코드 → 'ok' / 'nodata'(아직 없음 → 다시) / 'retry'(일시 오류 → 다시) / 'fatal'(다시 해도 소용없음)"""
    if http in (401, 403):
        return "fatal"
    if http is None or http >= 500 or http != 200:
        return "retry"
    code = str(code).zfill(2) if str(code).isdigit() else str(code)
    if code == "00":
        return "ok"
    if code in NODATA_CODES:
        return "nodata"
    if code in FATAL_CODES:
        return "fatal"
    return "retry"


def fetch_pages(url_of_page, http=http_get, timeout=TIMEOUT_S, rows=NUM_ROWS):
    """페이지를 이어 받아 dict(status, http, code, msg, items, total, bodies[(page, bytes)], error, elapsed_s)"""
    t0 = time.monotonic()
    out = dict(status="retry", http=None, code="", msg="", items=[], total=0, bodies=[], error="")
    page = 1
    while True:
        url = url_of_page(page)
        try:
            status, body = http(url, timeout)
        except Exception as e:                      # 시간 초과·연결 오류
            out.update(status="retry", error=f"{type(e).__name__}: {str(e)[:120]}")
            break
        out["http"] = status
        if status == 200:
            out["bodies"].append((page, body))
        if status != 200:
            out.update(status=classify(status, ""), error=f"HTTP {status}")
            break
        p = parse_response(body)
        out.update(code=p["code"], msg=p["msg"])
        st = classify(200, p["code"])
        if st != "ok":
            out["status"] = st
            break
        out["items"] += p["items"]
        out["total"] = p["total"]
        if p["total"] == 0:
            out["status"] = "nodata"
            break
        if len(out["items"]) >= p["total"] or not p["items"] or page * rows >= p["total"]:
            out["status"] = "ok"
            break
        page += 1
    out["elapsed_s"] = round(time.monotonic() - t0, 2)
    return out


def fetch_vilage(key, issue, nx, ny, http=http_get, timeout=TIMEOUT_S):
    """단기예보 한 발표(issue: 발표시각) → fetch_pages 결과"""
    return fetch_pages(lambda p: vilage_url(key, issue, nx, ny, p), http, timeout)


def fetch_asos_days(key, stn, start, end, http=http_get, timeout=TIMEOUT_S):
    """ASOS 일자료 start~end → fetch_pages 결과(items = 날짜별 dict)"""
    return fetch_pages(lambda p: asos_url(key, stn, start, end), http, timeout, rows=50)


# ── 완결성 점검 ──────────────────────────────────────────────────────────
def items_frame(items):
    """응답 항목 목록 → DataFrame[API_COLS] (값은 문자열 그대로)"""
    if not items:
        return pd.DataFrame(columns=API_COLS)
    d = pd.DataFrame(items)
    for c in API_COLS:
        if c not in d:
            d[c] = ""
    d = d[API_COLS].astype(str)
    d["baseTime"] = d["baseTime"].str.zfill(4)
    d["fcstTime"] = d["fcstTime"].str.zfill(4)
    return d.drop_duplicates(["baseDate", "baseTime", "category", "fcstDate", "fcstTime"], keep="last").reset_index(drop=True)


def _valid(cat, s):
    """값 문자열이 쓸 수 있는 값인지(결측 ±900, 하늘상태 코드표 밖 값은 빠진 칸)"""
    v = s.map(pcp_mm) if cat == "PCP" else pd.to_numeric(s, errors="coerce")
    ok = v.notna() & (v.abs() < MISSING_ABS)
    if cat == "SKY":
        ok &= v.round(6).isin(SKY_VALID)
    return ok


def check_issue(df, issue):
    """받은 발표 1개의 완결성. 반환 dict(required_ok, optional_ok, complete8, counts{요소: 유효 행 수}, expected, missing[문구])"""
    issue = pd.Timestamp(issue)
    h = issue.hour
    d = df[(df.baseDate == f"{issue:%Y%m%d}") & (df.baseTime == f"{issue:%H}00")]
    counts, missing = {}, []
    exp = {e: HOURLY_N.get(h, 1) for e in HOURLY}
    exp.update(TMX=TMX_N.get(h, 1), TMN=TMN_N.get(h, 1))
    for cat in ELEMENTS + OPTIONAL:
        s = d.loc[d.category == cat, "fcstValue"]
        n = int(_valid(cat, s).sum()) if len(s) else 0
        counts[cat] = n
        if n < exp[cat]:
            missing.append(f"{cat} {n}/{exp[cat]}")
    req = all(counts[e] >= exp[e] for e in ELEMENTS)
    opt = all(counts[e] >= exp[e] for e in OPTIONAL)
    return dict(required_ok=req, optional_ok=opt, complete8=req and opt, counts=counts, expected=exp, missing=missing)


# ── 저장 ─────────────────────────────────────────────────────────────────
class Store:
    """운영 자료 폴더 (기본 data/ops)
         raw/fcst/<격자>/<YYYY>/<MM>/<발표 YYYYmmdd_HHMM>__<받은 시각>_p<n>.json.gz   응답 그대로
         raw/asos/<지점>/<YYYY>/<MM>/<시작>_<끝>__<받은 시각>.json.gz
         fcst/<격자>/vilage_<격자>_<YYYYMM>.csv     완결 발표만 (OpenAPI 형식)
         fcst/<격자>/issues.csv                     받은 발표 목록(발표, 받은 시각, 8요소 완결 여부, 행 수)
         obs/asos_<지점>_daily.csv                   ASOS 일자료(날짜마다 가장 최근 조회 값)
         log/collect_log.csv, log/slot_log.csv"""

    def __init__(self, root="data/ops"):
        self.root = root

    def p(self, *parts):
        path = os.path.join(self.root, *parts)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        return path

    # 원자료
    def save_raw(self, kind, key, name, bodies, fetched_at):
        """bodies: [(page, bytes)] → 경로 목록 (받은 그대로 gzip)"""
        out = []
        stamp = pd.Timestamp(fetched_at).strftime("%Y%m%dT%H%M%S")
        for page, body in bodies:
            path = self.p("raw", kind, key, name[:4], name[4:6], f"{name}__{stamp}_p{page}.json.gz")
            with gzip.open(path, "wb") as f:
                f.write(body if isinstance(body, (bytes, bytearray)) else str(body).encode("utf-8"))
            out.append(path)
        return out

    # 예보 누적
    def fcst_dir(self, grid):
        return os.path.join(self.root, "fcst", grid)

    def issues(self, grid):
        path = os.path.join(self.fcst_dir(grid), "issues.csv")
        if not os.path.exists(path):
            return pd.DataFrame(columns=["issue", "stored_at", "complete8", "rows"])
        t = pd.read_csv(path, encoding="utf-8-sig")
        t["issue"] = pd.to_datetime(t["issue"])
        return t

    def has_issue(self, grid, issue):
        t = self.issues(grid)
        return bool((t.issue == pd.Timestamp(issue)).any())

    def issue_complete8(self, grid, issue):
        t = self.issues(grid)
        m = t[t.issue == pd.Timestamp(issue)]
        return bool(len(m) and str(m.complete8.iloc[-1]).lower() in ("true", "1"))

    def store_issue(self, grid, issue, df, chk, stored_at):
        """완결 점검을 통과한(필수 6요소) 발표를 월별 CSV에 붙인다. 하늘상태·강수확률이 덜 찼으면 두 요소는 뺀다(→ S3)"""
        issue = pd.Timestamp(issue)
        d = df[(df.baseDate == f"{issue:%Y%m%d}") & (df.baseTime == f"{issue:%H}00")]
        if not chk["optional_ok"]:
            d = d[~d.category.isin(OPTIONAL)]
        path = self.p("fcst", grid, f"vilage_{grid}_{issue:%Y%m}.csv")
        new = not os.path.exists(path)
        d.to_csv(path, mode="a", header=new, index=False, encoding="utf-8-sig" if new else "utf-8")
        ip = self.p("fcst", grid, "issues.csv")
        newi = not os.path.exists(ip)
        with open(ip, "a", encoding="utf-8-sig" if newi else "utf-8", newline="") as f:
            w = csv.writer(f)
            if newi:
                w.writerow(["issue", "stored_at", "complete8", "rows"])
            w.writerow([f"{issue:%Y-%m-%d %H:%M}", f"{pd.Timestamp(stored_at):%Y-%m-%d %H:%M:%S}", chk["complete8"], len(d)])
        return path

    def fcst_files(self, grid, since=None):
        """계산에 쓸 월별 CSV 목록(since가 있으면 그 달부터)"""
        d = self.fcst_dir(grid)
        if not os.path.isdir(d):
            return []
        fs = sorted(f for f in os.listdir(d) if f.startswith(f"vilage_{grid}_") and f.endswith(".csv"))
        if since is not None:
            ym = f"{pd.Timestamp(since):%Y%m}"
            fs = [f for f in fs if f[-10:-4] >= ym]
        return [os.path.join(d, f) for f in fs]

    # 관측 캐시
    def obs_path(self, stn):
        return self.p("obs", f"asos_{stn}_daily.csv")

    def upsert_obs(self, stn, items, fetched_at):
        """ASOS 일자료 항목 → 캐시 CSV(날짜마다 가장 최근 조회 값으로 덮어씀). 반환: 캐시 DataFrame"""
        path = self.obs_path(stn)
        old = pd.read_csv(path, encoding="utf-8-sig", dtype=str) if os.path.exists(path) else pd.DataFrame(columns=ASOS_FIELDS + ["fetched_at"])
        new = pd.DataFrame([{k: ("" if it.get(k) is None else str(it.get(k))) for k in ASOS_FIELDS} for it in items])
        if len(new):
            new["fetched_at"] = f"{pd.Timestamp(fetched_at):%Y-%m-%d %H:%M:%S}"
            old = pd.concat([old[~old.tm.isin(new.tm)], new], ignore_index=True)
        old = old.sort_values("tm")
        old.to_csv(path, index=False, encoding="utf-8-sig")
        return old

    # 로그
    def log(self, name, row, cols):
        path = self.p("log", name)
        new = not os.path.exists(path)
        with open(path, "a", encoding="utf-8-sig" if new else "utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
            if new:
                w.writeheader()
            w.writerow({c: row.get(c, "") for c in cols})

    def read_log(self, name):
        path = os.path.join(self.root, "log", name)
        return pd.read_csv(path, encoding="utf-8-sig", dtype=str) if os.path.exists(path) else pd.DataFrame()


def obs_rows_complete(item):
    """ASOS 한 날의 ETo 필수 입력(최고·최저기온, 풍속, 습도류, 일사)이 있는지 (obs_daily.add_eto_obs와 같은 조건)"""
    g = lambda k: str(item.get(k, "") or "").strip() not in ("", "-", "None", "nan")
    return g("maxTa") and g("minTa") and g("avgWs") and (g("avgPv") or g("avgTd") or g("avgRhm")) and g("sumGsr")


def asos_cache_frame(cache):
    """관측 캐시(문자열) → obs_daily 원데이터와 같은 열(date, Tmax, …, rain)"""
    from obs_daily import _COLS
    m = dict(zip(ASOS_FIELDS, _COLS))
    d = cache.rename(columns=m)[_COLS].copy()
    d["date"] = pd.to_datetime(d["date"])
    for c in _COLS[1:]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d["rain"] = d["rain"].fillna(0.0)
    return d
