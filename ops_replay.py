#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ops_replay.py — 운영 수집기 모의 재생(오프라인 검증, G5)

과거 단기예보(OpenAPI로 받은 CSV)와 01-Cycle 관측 워크북으로 기상청 API를 흉내 내고, 가짜 시계로 운영 경로
(cropwater_ops.run_slot)를 슬롯마다 그대로 돌린다. 네트워크 없이 수집기 논리를 검정한다(THEORY 9장 ◆ 운영 수집 검증 설계).

  MockVilage : 단기예보 조회서비스 흉내. 발표 10분 뒤부터 제공, 페이지 나눔, 장애 주입
               (지연·발표 누락·HTTP 500·시간 초과·잘린 응답·하늘상태/강수확률 누락·인증 오류)
  MockAsos   : ASOS 일자료 조회서비스 흉내. 날짜 D의 자료는 D+1일 avail_h시부터 제공
  SimClock   : sleep이 시간을 앞으로 돌리는 시계
  replay()   : 기간의 모든 슬롯을 차례로 처리

  python ops_replay.py --fcst data/fcst_101_2026 --obs output/eto101_apple_20260101_20260928.xlsx \
         --start 2026-04-01 --end 2026-09-28 --data data/replay_2026 [--faults faults.json] [--asos-hour 1]
"""
import argparse
import json
import random
import urllib.parse

import pandas as pd

from fcst_archive import ELEMENTS, OPTIONAL, csv_format, read_element_csv


class SimClock:
    def __init__(self, t):
        self.t = pd.Timestamp(t)

    def now(self):
        return self.t

    def sleep(self, sec):
        self.t = self.t + pd.Timedelta(seconds=max(float(sec), 0.0))

    def set(self, t):
        self.t = max(self.t, pd.Timestamp(t))


def load_raw_items(paths, grid="73_134"):
    """과거 예보 CSV(요소별 KST 또는 OpenAPI 응답) → {발표시각: [API 항목 dict(값은 문자열 그대로)]}"""
    import os
    files = []
    for p in paths:
        files += sorted(os.path.join(p, f) for f in os.listdir(p) if f.lower().endswith(".csv")) if os.path.isdir(p) else [p]
    nx, ny = grid.split("_")
    parts = []
    for f in files:
        fmt = csv_format(f)
        if fmt == "element":
            elem, _ = read_element_csv(f)
            if elem not in ELEMENTS + OPTIONAL:
                continue
            d = pd.read_csv(f, encoding="utf-8-sig", dtype=str)
            d.columns = [c.strip().lstrip("﻿") for c in d.columns]
            parts.append(pd.DataFrame({"baseDate": d["발표일"].str.strip(), "baseTime": d["발표시각"].str.strip().str.zfill(4),
                                       "category": elem, "fcstDate": d["예보일"].str.strip(),
                                       "fcstTime": d["예보시각"].str.strip().str.zfill(4), "fcstValue": d["값"].str.strip()}))
        elif fmt == "openapi":
            d = pd.read_csv(f, encoding="utf-8-sig", dtype=str).dropna(subset=["baseDate", "category"])
            d = d[d.category.isin(ELEMENTS + OPTIONAL)]
            parts.append(d[["baseDate", "baseTime", "category", "fcstDate", "fcstTime", "fcstValue"]].assign(
                baseTime=lambda x: x.baseTime.str.zfill(4), fcstTime=lambda x: x.fcstTime.str.zfill(4)))
    a = pd.concat(parts, ignore_index=True)
    a["nx"], a["ny"] = int(nx), int(ny)
    a["issue"] = pd.to_datetime(a.baseDate + a.baseTime, format="%Y%m%d%H%M")
    out = {}
    for iss, g in a.groupby("issue"):
        out[iss] = g.drop(columns="issue").to_dict("records")
    return out


def _json(code, msg, items=None, total=0, page=1, rows=1000):
    body = {"response": {"header": {"resultCode": code, "resultMsg": msg}}}
    if items is not None:
        body["response"]["body"] = {"dataType": "JSON", "items": {"item": items}, "pageNo": page, "numOfRows": rows,
                                    "totalCount": total}
    return json.dumps(body, ensure_ascii=False).encode("utf-8")


XML_AUTH = (b"<OpenAPI_ServiceResponse><cmmMsgHeader><errMsg>SERVICE ERROR</errMsg>"
            b"<returnAuthMsg>SERVICE_KEY_IS_NOT_REGISTERED_ERROR</returnAuthMsg><returnReasonCode>30</returnReasonCode>"
            b"</cmmMsgHeader></OpenAPI_ServiceResponse>")


class MockApi:
    """단기예보·ASOS 일자료 모의 API. faults(dict):
         delay_min   {발표 'YYYY-mm-dd HH:MM': 분}   제공이 발표 + 10분 + 지연부터
         missing     [발표 …]                        끝내 제공 안 함(NODATA)
         truncate    {발표: 횟수}                     처음 n번은 1시간 기온 행 일부가 빠진 응답
         drop_opt    [발표 …]                        하늘상태·강수확률 없이 제공
         fatal       [['시작', '끝'] …]               그 시각 범위의 모든 요청에 인증 오류(코드 30, XML)
         p_http500, p_timeout                        요청마다 서버 오류·시간 초과 확률
         seed                                        난수 시드"""

    def __init__(self, fcst_items, obs_items, clock, faults=None, asos_hour=1):
        self.f, self.o, self.clock = fcst_items, obs_items, clock
        fa = faults or {}
        ts = lambda x: pd.Timestamp(x)
        self.delay = {ts(k): v for k, v in fa.get("delay_min", {}).items()}
        self.missing = {ts(k) for k in fa.get("missing", [])}
        self.trunc = {ts(k): v for k, v in fa.get("truncate", {}).items()}
        self.drop_opt = {ts(k) for k in fa.get("drop_opt", [])}
        self.fatal = [(ts(a), ts(b)) for a, b in fa.get("fatal", [])]
        self.p500, self.pto = fa.get("p_http500", 0.0), fa.get("p_timeout", 0.0)
        self.rng = random.Random(fa.get("seed", 20260930))
        self.asos_hour = asos_hour
        self.calls = 0

    def __call__(self, url, timeout=30):
        self.calls += 1
        now = self.clock.now()
        q = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(url).query))
        if any(a <= now <= b for a, b in self.fatal):
            return 200, XML_AUTH
        r = self.rng.random()
        if r < self.pto:
            self.clock.sleep(timeout)
            raise TimeoutError("모의 시간 초과")
        if r < self.pto + self.p500:
            return 500, b"<html>500 Internal Server Error</html>"
        self.clock.sleep(1.0)
        if "VilageFcst" in url:
            return self._vilage(q, now)
        return self._asos(q, now)

    def _vilage(self, q, now):
        iss = pd.Timestamp(q["base_date"] + q["base_time"])
        avail = iss + pd.Timedelta(minutes=10 + self.delay.get(iss, 0))
        items = self.f.get(iss)
        if items is None or iss in self.missing or now < avail:
            return 200, _json("03", "NO_DATA")
        if iss in self.drop_opt:
            items = [x for x in items if x["category"] not in OPTIONAL]
        if self.trunc.get(iss, 0) > 0 and int(q.get("pageNo", 1)) == 1:
            self.trunc[iss] -= 1
            tmp = [i for i, x in enumerate(items) if x["category"] == "TMP"]
            cut = set(tmp[len(tmp) // 2:])
            items = [x for i, x in enumerate(items) if i not in cut]
        rows, page = int(q.get("numOfRows", 1000)), int(q.get("pageNo", 1))
        chunk = items[(page - 1) * rows: page * rows]
        return 200, _json("00", "NORMAL_SERVICE", chunk, len(items), page, rows)

    def _asos(self, q, now):
        d0, d1 = pd.Timestamp(q["startDt"]), pd.Timestamp(q["endDt"])
        got = [self.o[d] for d in pd.date_range(d0, d1)
               if d in self.o and now >= d + pd.Timedelta(days=1, hours=self.asos_hour)]
        if not got:
            return 200, _json("03", "NO_DATA")
        return 200, _json("00", "NORMAL_SERVICE", got, len(got), 1, int(q.get("numOfRows", 50)))


def obs_items_from_workbook(path):
    """01-Cycle 워크북 원데이터 → {날짜: ASOS 응답 항목(dict, 값은 문자열)}"""
    from obs_daily import load_station_workbook
    from kma_fcst import ASOS_FIELDS
    df, _ = load_station_workbook(path)
    cols = ["date", "Tmax", "Tmin", "Tavg", "RHmean", "RHmin", "u10", "pv", "td", "pa", "Rs", "ss", "epan", "rain"]
    out = {}
    for r in df[cols].itertuples(index=False):
        it = {}
        for k, v in zip(ASOS_FIELDS, r):
            if k == "tm":
                it[k] = f"{pd.Timestamp(v):%Y-%m-%d}"
            else:
                it[k] = "" if pd.isna(v) else repr(float(v))
        out[pd.Timestamp(r.date)] = it
    return out


def slots_between(start, end):
    from cropwater_ops import SLOTS
    out = []
    for d in pd.date_range(pd.Timestamp(start).normalize(), pd.Timestamp(end).normalize()):
        for name, h in SLOTS.items():
            out.append((name, d + pd.Timedelta(hours=h)))
    return sorted(out, key=lambda x: x[1])


def replay(cfg, fcst_paths, start, end, faults=None, asos_hour=1, mode="light", build_slots=(), progress=True):
    """기간의 모든 슬롯을 모의 API로 처리. mode: 'light'(수집·출처 판단만, 빠름) / 'full'(슬롯마다 서비스 엑셀)
       build_slots: light 모드에서도 엑셀을 만들 슬롯 발표시각 목록. 반환: 슬롯 로그 행 목록"""
    import kma_fcst as K
    from cropwater_ops import run_slot
    clock = SimClock(pd.Timestamp(start) - pd.Timedelta(hours=1))
    api = MockApi(load_raw_items(fcst_paths, cfg["grid"]), obs_items_from_workbook(cfg["obs_source"]), clock, faults, asos_hour)
    store = K.Store(cfg["data"])
    rows = []
    build_slots = {pd.Timestamp(x) for x in build_slots}
    for i, (name, issue) in enumerate(slots_between(start, end)):
        clock.set(issue + pd.Timedelta(minutes=10))
        full = mode == "full" or issue in build_slots
        row = run_slot(cfg, name, issue, http=api, clock=clock, key="TESTKEY", write=full, store=store,
                       build=("full" if full else "light"))
        rows.append(row)
        if progress and (i % 20 == 0):
            print(f"  {issue:%Y-%m-%d %H} {row['result']} (API 호출 {api.calls})", flush=True)
    return rows, api


def main(argv=None):
    ap = argparse.ArgumentParser(description="운영 수집기 모의 재생(오프라인)")
    ap.add_argument("--fcst", nargs="+", required=True); ap.add_argument("--obs", required=True)
    ap.add_argument("--start", required=True); ap.add_argument("--end", required=True)
    ap.add_argument("--data", default="data/replay"); ap.add_argument("--out", default="output/replay")
    ap.add_argument("--stn", default="101"); ap.add_argument("--grid", default="73_134")
    ap.add_argument("--faults", default=None, help="장애 주입 JSON(MockApi 설명 참고)")
    ap.add_argument("--asos-hour", type=float, default=1, help="전날 ASOS 일자료가 제공되는 시각(시)")
    ap.add_argument("--obs-until", default=None, help="관측 워크북을 이 날짜까지만 쓰고 뒤는 모의 ASOS로(관측 경로 검정)")
    ap.add_argument("--mode", default="light", choices=["light", "full"])
    a = ap.parse_args(argv)
    from cropwater_ops import load_config
    cfg = load_config(None, stn=a.stn, grid=a.grid, obs=a.obs, data=a.data, out=a.out, obs_until=a.obs_until)
    cfg["obs_source"] = a.obs
    faults = json.load(open(a.faults, encoding="utf-8")) if a.faults else None
    rows, api = replay(cfg, a.fcst, a.start, a.end, faults, a.asos_hour, a.mode)
    res = pd.Series([r["result"] for r in rows]).value_counts()
    print(res.to_string())
    print(f"[API 호출] {api.calls}회")


if __name__ == "__main__":
    main()
