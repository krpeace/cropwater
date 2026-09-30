#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cropwater_ops.py — 운영 수집기와 관수 전망 자동 생성 (02-Cycle 5단계, G5 · H3)

[실행] 스케줄러(Windows 작업 스케줄러·cron)가 하루 두 번 부른다
  python cropwater_ops.py run --config ops_config.json            # 02:10·17:10 — 지금 시각의 슬롯을 처리
  python cropwater_ops.py run --config ops_config.json --slot "2026-10-01 02"   # 특정 슬롯 다시(지연 수집)
  python cropwater_ops.py probe-asos --config ops_config.json     # 전날 ASOS 일자료 조회 가능 시각 점검(매시)
  python cropwater_ops.py report --config ops_config.json [--since 2026-10-01] [--until 2026-11-30]   # H3 보고서 엑셀
  python cropwater_ops.py grid 37.90262 127.73570                  # 농장 좌표 → 격자

[한 슬롯에서 하는 일] (THEORY 9장 ◆ 운영 수집)
  1) 서비스 발표(아침 02시·저녁 17시)를 발표 10분 뒤부터 받는다. 못 받으면 +5·+10·+15·+20·+30분에 다시(마감 = 발표 + 40분)
  2) 첫 시도 뒤, 지난 24시간 안에 아직 받지 않은 발표를 한 번씩 받는다(02시 발표의 00~02시 채움, 백업)
  3) ASOS 전날 일자료를 조회하고(없으면 기록만) 빠진 최근 날을 채운다
  4) 운영 결측 규칙(#12)으로 서비스 표 → 관수 전망 엑셀(fcst_wb_report.build_service_workbook)
       서비스 발표가 없으면 직전 발표로 대신, 하늘상태·강수확률이 없으면 S3, 전날 관측이 없으면 전날 아침 발표 D+0 예보로 채움
  5) 슬롯 결과를 slot_log.csv에 남긴다: 정시 성공 / 백업: S3 / 백업: 직전 발표 / 지연 수집 / 실패

[설정] ops_config.json (ops_config.example.json 참고) — 명령행 인자가 있으면 그 값이 우선
근거: docs/THEORY.md 9장 ◆ 운영 수집, 설계: docs/ARCHITECTURE.md 7장, 판정: docs/VALIDATION.md G5
"""
import argparse
import json
import os
import sys

import numpy as np
import pandas as pd

import kma_fcst as K

SLOTS = {"아침": 2, "저녁": 17}
AVAIL_MIN = 10                       # 발표 10분 뒤부터 제공(활용가이드)
DEADLINE_MIN = 40                    # 마감 = 발표 + 40분 (02:40, 17:40)
RETRY_MIN = (0, 5, 10, 15, 20, 30)   # 슬롯 시작(발표 + 10분)부터의 재시도 시각(분)
CATCHUP_H = 24                       # 함께 받는 지난 발표 범위(시간)
BACKUP_MAX_H = 24                    # 직전 발표로 대신할 수 있는 범위(시간)
OBS_BACKFILL_DAYS = 30               # 관측 캐시: 전날 외에 워크북 뒤 최근 며칠의 빠진 날을 채움
WINDOW_DAYS = 40                     # 서비스 계산에 읽는 예보 기간(편향 점검 30일 + 여유)
SLOT_COLS = ["slot_run", "run_name", "started_at", "finished_at", "deadline", "service_received_at", "on_time",
             "result", "src_run", "backup", "rs_method", "obs_prev", "obs_prev_fetched_at", "attempts", "excel", "error"]
DEFAULTS = dict(stn="101", grid=None, lat=None, lon=None, obs=None, data="data/ops", out="output/service",
                apikey="apikey.txt", coef="rs_coef.csv", s4_coef="rs_sky_coef.csv", err="fcst_error_table.csv",
                irrig=None, history=[], window_days=WINDOW_DAYS)


# ── 설정 ─────────────────────────────────────────────────────────────────
def load_config(path=None, **over):
    """ops_config.json + 명령행 값 → dict. 격자가 없으면 lat·lon(농장 좌표)으로, 그것도 없으면 지점 표(STATIONS)로"""
    cfg = dict(DEFAULTS)
    if path:
        with open(path, encoding="utf-8-sig") as f:
            cfg.update({k: v for k, v in json.load(f).items() if not k.startswith("_")})
    cfg.update({k: v for k, v in over.items() if v is not None})
    if not cfg.get("grid"):
        if cfg.get("lat") is not None and cfg.get("lon") is not None:
            from kma_grid import latlon_to_grid
            cfg["grid"] = "%d_%d" % latlon_to_grid(float(cfg["lat"]), float(cfg["lon"]))
        else:
            from cropwater_fcst import STATIONS
            cfg["grid"] = STATIONS.get(str(cfg["stn"]), ("", ""))[1]
    if not cfg.get("grid"):
        raise SystemExit("[설정] 격자(grid 'nx_ny') 또는 농장 좌표(lat, lon)가 필요합니다")
    cfg["stn"] = str(cfg["stn"])
    if isinstance(cfg.get("history"), str):
        cfg["history"] = [cfg["history"]]
    return cfg


def api_key(cfg):
    from fao56_core import load_apikeys
    k = load_apikeys(cfg["apikey"]).get("DATA_GO_KR", "")
    if not k or "여기에" in k:
        raise SystemExit(f"[인증키] {cfg['apikey']}에 DATA_GO_KR 키가 없습니다(apikey.txt.example 참고)")
    return k


# ── 슬롯 ─────────────────────────────────────────────────────────────────
def slot_times(issue):
    issue = pd.Timestamp(issue)
    return issue + pd.Timedelta(minutes=AVAIL_MIN), issue + pd.Timedelta(minutes=DEADLINE_MIN)


def current_slot(now):
    """지금 시각의 슬롯 = 시작 시각(발표 + 10분)이 지금 이전인 가장 최근 슬롯 → (구분, 발표시각)"""
    now = pd.Timestamp(now)
    cands = []
    for back in (0, 1):
        d = now.normalize() - pd.Timedelta(days=back)
        for name, h in SLOTS.items():
            iss = d + pd.Timedelta(hours=h)
            if slot_times(iss)[0] <= now + pd.Timedelta(minutes=1):      # 스케줄러가 1분 일찍 불러도 같은 슬롯
                cands.append((iss, name))
    iss, name = max(cands)
    return name, iss


def parse_slot(s):
    t = pd.Timestamp(s if len(str(s)) > 13 else f"{s}:00")
    name = {h: n for n, h in SLOTS.items()}.get(t.hour)
    if name is None:
        raise SystemExit("[슬롯] 서비스 발표는 02시·17시입니다(예: '2026-10-01 02')")
    return name, t


# ── 수집 ─────────────────────────────────────────────────────────────────
class Collector:
    def __init__(self, cfg, store=None, http=K.http_get, clock=None, key=None):
        self.cfg, self.http = cfg, http
        self.store = store or K.Store(cfg["data"])
        self.clock = clock or K.Clock()
        self.key = key if key is not None else api_key(cfg)
        self.nx, self.ny = (int(x) for x in cfg["grid"].split("_"))
        self.fatal = ""

    def _log(self, slot, kind, target, attempt, res, chk=None, status=None, note=""):
        self.store.log("collect_log.csv", dict(
            logged_at=f"{self.clock.now():%Y-%m-%d %H:%M:%S}", slot=f"{pd.Timestamp(slot):%Y-%m-%d %H:%M}", kind=kind,
            target=target, attempt=attempt, http=res.get("http"), code=res.get("code"), msg=res.get("msg"),
            n_items=len(res.get("items", [])), total=res.get("total"), status=status or res.get("status"),
            required_ok="" if chk is None else chk["required_ok"], complete8="" if chk is None else chk["complete8"],
            elapsed_s=res.get("elapsed_s"), note=(note or res.get("error", ""))[:200]), K.LOG_COLS)

    def get_issue(self, slot, issue, attempt=1):
        """발표 1개를 한 번 받아 원자료 보관·완결성 점검·(필수 6요소가 차면) 저장. 반환: (상태, 점검)"""
        issue = pd.Timestamp(issue)
        now = self.clock.now()
        res = K.fetch_vilage(self.key, issue, self.nx, self.ny, self.http)
        if res["bodies"]:
            self.store.save_raw("fcst", self.cfg["grid"], f"{issue:%Y%m%d_%H%M}", res["bodies"], now)
        chk, status = None, res["status"]
        if res["status"] == "ok":
            chk = K.check_issue(K.items_frame(res["items"]), issue)
            status = "ok" if chk["complete8"] else ("partial" if chk["required_ok"] else "incomplete")
            if chk["required_ok"] and not self.store.issue_complete8(self.cfg["grid"], issue):
                self.store.store_issue(self.cfg["grid"], issue, K.items_frame(res["items"]), chk, now)
        if status == "fatal":
            self.fatal = f"[{res.get('code') or res.get('http')}] {res.get('msg') or res.get('error')}"
        self._log(slot, "fcst", f"{issue:%Y-%m-%d %H:%M}", attempt, res, chk, status,
                  note=("; ".join(chk["missing"]) if chk and chk["missing"] else ""))
        return status, chk

    def catch_up(self, slot):
        """지난 CATCHUP_H 시간 안에 아직 받지 않은 발표를 한 번씩(가장 최근부터)"""
        slot = pd.Timestamp(slot)
        got = 0
        for back in range(3, CATCHUP_H + 1, 3):
            iss = slot - pd.Timedelta(hours=back)
            if iss.hour not in K.ISSUE_HOURS or self.store.has_issue(self.cfg["grid"], iss) or self.fatal:
                continue
            st, _ = self.get_issue(slot, iss, attempt=1)
            got += st in ("ok", "partial")
        return got

    def collect_service(self, name, issue):
        """서비스 발표를 마감까지 받기. 반환 dict(received_at, complete8, required_ok, attempts, last)"""
        issue = pd.Timestamp(issue)
        start, deadline = slot_times(issue)
        grid = self.cfg["grid"]
        out = dict(received_at=None, complete8=False, required_ok=False, attempts=0, last="", late=False, rerun=False)
        if self.store.issue_complete8(grid, issue):           # 이미 받음(같은 슬롯을 다시 실행) — 정시 판단은 처음 실행 기록으로
            out.update(complete8=True, required_ok=True, last="이미 받음", rerun=True)
            self.catch_up(issue)
            return out
        caught = False
        now = self.clock.now()
        offsets = [o for o in RETRY_MIN if start + pd.Timedelta(minutes=o) <= deadline]
        if now > deadline:                                    # 마감 뒤에 다시 실행: 한 번만(지연 수집)
            offsets, out["late"] = [None], True
        for o in offsets:
            if o is not None:
                t_try = start + pd.Timedelta(minutes=o)
                now = self.clock.now()
                if now < t_try:
                    self.clock.sleep((t_try - now).total_seconds())
                elif now > deadline:
                    break
            out["attempts"] += 1
            st, chk = self.get_issue(issue, issue, out["attempts"])
            out["last"] = st
            if chk is not None and chk["required_ok"]:
                out["required_ok"] = True
                if out["received_at"] is None:
                    out["received_at"] = self.clock.now()
            if st == "ok":
                out.update(complete8=True, received_at=self.clock.now())
            if not caught:                                     # 첫 시도 뒤: 백업 발표를 먼저 확보
                self.catch_up(issue)
                caught = True
            if st in ("ok", "fatal"):
                break
        if not caught:
            self.catch_up(issue)
        return out

    def update_obs(self, slot, D):
        """ASOS 전날(D−1) 일자료 조회 + 최근 빠진 날 채우기. 반환 dict(prev_ok, prev_fetched_at)"""
        stn, D = self.cfg["stn"], pd.Timestamp(D).normalize()
        prev = D - pd.Timedelta(days=1)
        now = self.clock.now()
        res = K.fetch_asos_days(self.key, stn, prev, prev, self.http)
        if res["bodies"]:
            self.store.save_raw("asos", stn, f"{prev:%Y%m%d}_{prev:%Y%m%d}", res["bodies"], now)
        items = [it for it in res["items"] if str(it.get("tm", ""))[:10] == f"{prev:%Y-%m-%d}"]
        ok = res["status"] == "ok" and len(items) > 0 and K.obs_rows_complete(items[0])
        if items:
            self.store.upsert_obs(stn, items, now)
        self._log(slot, "asos_d1", f"{prev:%Y-%m-%d}", 1, res, status=("ok" if ok else ("partial" if items else res["status"])))
        # 관측 워크북 뒤의 빠진 날(캐시에 없거나 ETo 입력이 빠진 날) 채우기 — 워크북 안의 날짜는 다시 받지 않음
        have, wb_last = known_obs_dates(self.cfg, self.store)
        need = [d for d in pd.date_range(D - pd.Timedelta(days=OBS_BACKFILL_DAYS), prev - pd.Timedelta(days=1))
                if d > wb_last and d not in have]
        if need and not self.fatal:
            r2 = K.fetch_asos_days(self.key, stn, min(need), max(need), self.http)
            if r2["bodies"]:
                self.store.save_raw("asos", stn, f"{min(need):%Y%m%d}_{max(need):%Y%m%d}", r2["bodies"], now)
            if r2["items"]:
                self.store.upsert_obs(stn, r2["items"], now)
            self._log(slot, "asos_fill", f"{min(need):%Y-%m-%d}~{max(need):%Y-%m-%d}", 1, r2)
        return dict(prev_ok=ok, prev_fetched_at=now if ok else None)


# ── 관측 ─────────────────────────────────────────────────────────────────
_WB_CACHE = {}


def _workbook(path, until=None):
    """01-Cycle 워크북 읽기(파일 시각이 같으면 다시 읽지 않음). until: 이 날짜까지만(모의 재생에서 관측 경로 검정용)"""
    from obs_daily import load_station_workbook
    key = (os.path.abspath(path), os.path.getmtime(path))
    if key not in _WB_CACHE:
        _WB_CACHE.clear()
        _WB_CACHE[key] = load_station_workbook(path)
    df, meta = _WB_CACHE[key]
    df, meta = df.copy(), dict(meta)
    if until is not None:
        df = df[df.date <= pd.Timestamp(until)].reset_index(drop=True)
    return df, meta


def load_obs_ops(cfg, store):
    """관측 = 01-Cycle 워크북 원데이터 + ASOS 캐시(워크북에 없는 날짜). 반환 (obs, meta) — cropwater_fcst.load_obs와 같은 모양"""
    from obs_daily import add_eto_obs
    df, meta = _workbook(cfg["obs"], cfg.get("obs_until"))
    meta["wb_last"] = pd.Timestamp(df.date.max())
    path = store.obs_path(cfg["stn"])
    if os.path.exists(path):
        cache = K.asos_cache_frame(pd.read_csv(path, encoding="utf-8-sig", dtype=str))
        cache = cache[~cache.date.isin(set(df.date))]
        if len(cache):
            df = pd.concat([df, cache], ignore_index=True).sort_values("date").reset_index(drop=True)
    return add_eto_obs(df, meta["lat"], meta["elev"], meta.get("anem", 10.0)), meta


def known_obs_dates(cfg, store):
    """(ETo 입력이 모두 있는 관측 날짜 집합(워크북 + 캐시), 워크북 마지막 날)"""
    obs, meta = load_obs_ops(cfg, store)
    return set(pd.to_datetime(obs.loc[obs.ETo_obs.notna(), "date"])), meta["wb_last"]


# ── 서비스 계산 ─────────────────────────────────────────────────────────
def ops_prepare(cfg, store, issue, window_days=None):
    """운영 예보표: 최근 window_days일의 받은 발표(+ history) → ops_service_table → forecast_table(S4 운영 계수 고정, S3 대체)"""
    from cropwater_fcst import STATION_LON, coef_file, forecast_table, load_s4_fixed
    from fcst_archive import check_archive, load_archive, ops_service_table
    from obs_daily import kc_params
    from rs_model import load_coef
    issue = pd.Timestamp(issue)
    wd = int(window_days or cfg.get("window_days") or WINDOW_DAYS)
    t0 = issue.normalize() - pd.Timedelta(days=wd)
    paths = store.fcst_files(cfg["grid"], since=t0 - pd.Timedelta(days=1)) + list(cfg.get("history") or [])
    if not paths:
        raise RuntimeError("받은 예보가 없습니다")
    arch = load_archive(paths)
    if not arch.location:
        arch.location.add(cfg["grid"])
    obs, meta = load_obs_ops(cfg, store)
    st = ops_service_table(arch, lat=meta["lat"], lon=STATION_LON.get(cfg["stn"]), max_age_h=BACKUP_MAX_H,
                           start=t0, end=issue)
    if st.empty:
        raise RuntimeError("서비스 표가 비었습니다(받은 예보 없음)")
    kp = kc_params(meta["settings"])
    s4 = load_s4_fixed(cfg["stn"], cfg["s4_coef"])
    ft = forecast_table(st, obs, meta["lat"], meta["elev"], load_coef(cfg["stn"], coef_file(cfg["coef"])), kp, s4, s3_fallback=True)
    ft.attrs = {}
    return dict(ft=ft, obs=obs, meta=meta, kp=kp, check=check_archive(arch), arch=arch, s4_fixed=s4, stn=cfg["stn"],
                failed_slots=list(getattr(arch, "failed_slots", [])))


def build_service(cfg, store, name, issue, obs_info=None, write=True):
    """한 슬롯의 관수 전망. 반환 dict(ok, src_run, backup, rs_method, obs_prev, excel, error, sv)"""
    from cropwater_fcst import coef_file
    import fcst_wb as W
    issue = pd.Timestamp(issue)
    p = ops_prepare(cfg, store, issue)
    ft = p["ft"]
    rows = ft[(ft.run == issue) & (ft.run_name == name)]
    if rows.empty:
        return dict(ok=False, error="24시간 안에 쓸 수 있는 발표 없음", src_run=None, backup=None, rs_method="", obs_prev="")
    main = rows.sort_values("lead_day").head(W.MAIN_DAYS)
    if main["ETc_main"].isna().any():
        return dict(ok=False, error="주 지표 3일 중 예보 ETc가 빈 날 있음", src_run=rows.src_run.iloc[0], backup=bool(rows.backup.iloc[0]),
                    rs_method="", obs_prev="")
    drop = rows.index[rows["ETc_main"].isna()]                 # 대체 발표가 담지 못한 마지막 날(참고)
    p["ft"] = ft.drop(index=drop)
    irrig = None
    if cfg.get("irrig"):
        from fao56_core import load_irrigation_log
        irrig = load_irrigation_log(cfg["irrig"])
    sv = W.service_from(p, cfg["stn"], issue, coef_file(cfg["err"]), irrig)
    prev = issue.normalize() - pd.Timedelta(days=1)
    o = sv["owb"].set_index("date")
    obs_prev = "관측" if prev in o.index and o.loc[prev, "ETo_src"] == "관측" else "예보로 채움"
    r = rows.dropna(subset=["ETc_main"])
    rs = "+".join(sorted(set(r["rs_method"].dropna()), reverse=True))
    src = pd.Timestamp(r.src_run.iloc[0])
    sv["ops"] = dict(src_run=src, backup=bool(r.backup.iloc[0]), rs_method=rs, obs_prev=obs_prev,
                     n_days=len(r), dropped=len(drop), made_at=(obs_info or {}).get("made_at"))
    out = os.path.join(cfg["out"], f"fcst_service({cfg['stn']})_{issue:%Y%m%d_%H}.xlsx")
    excel = None
    if write:
        from fcst_wb_report import build_service_workbook
        excel = build_service_workbook(sv, out)
    return dict(ok=True, src_run=src, backup=bool(r.backup.iloc[0]), rs_method=rs, obs_prev=obs_prev, excel=excel, error="", sv=sv)


def slot_result(col, svc):
    """슬롯 결과 분류 (THEORY 9장 ◆ 운영 수집 H3)"""
    if not svc.get("ok"):
        return "실패"
    if col.get("rerun"):
        return "다시 만듦"
    if col.get("late") and col.get("required_ok") and not svc.get("backup"):
        return "지연 수집"
    if col.get("complete8") and not svc.get("backup") and svc.get("rs_method") == "S4":
        return "정시 성공"
    if svc.get("backup"):
        return "백업: 직전 발표" + (" (S3)" if "S3" in svc.get("rs_method", "") else "")
    return "백업: S3"


def light_service(cfg, store, name, issue):
    """모의 재생용 빠른 판단(예보 계산 없이): 받은 발표 목록으로 출처(서비스 발표 / 직전 발표)와 Rs 방법, 전날 관측 여부만.
       build_service와 같은 규칙(24시간 안의 가장 최근 발표, 8요소 완결이면 S4). 대체 발표는 주 지표 3일을 늘 담는다(THEORY 9장 #12)"""
    issue = pd.Timestamp(issue)
    t = store.issues(cfg["grid"])
    ok = t[(t.issue <= issue) & (t.issue >= issue - pd.Timedelta(hours=BACKUP_MAX_H))]
    if not len(ok):
        return dict(ok=False, error="24시간 안에 쓸 수 있는 발표 없음", src_run=None, backup=None, rs_method="", obs_prev="")
    src = ok.issue.max()
    c8 = str(ok[ok.issue == src].complete8.iloc[-1]).lower() in ("true", "1")
    have, _ = known_obs_dates(cfg, store)
    prev = issue.normalize() - pd.Timedelta(days=1)
    return dict(ok=True, src_run=src, backup=bool(src != issue), rs_method="S4" if c8 else "S3",
                obs_prev="관측" if prev in have else "예보로 채움", excel=None, error="")


def run_slot(cfg, name=None, issue=None, http=K.http_get, clock=None, key=None, write=True, store=None, build="full"):
    """한 슬롯 처리(수집 → 관측 → 서비스 엑셀 → 슬롯 로그). 반환 dict(슬롯 로그 행)
       build: 'full'(관수 전망 계산, write면 엑셀) / 'light'(모의 재생용: 출처·방법만 판단)"""
    clock = clock or K.Clock()
    store = store or K.Store(cfg["data"])
    if issue is None:
        name, issue = current_slot(clock.now())
    issue = pd.Timestamp(issue)
    started = clock.now()
    col_ = Collector(cfg, store, http, clock, key)
    col = col_.collect_service(name, issue)
    obs_info = col_.update_obs(issue, issue.normalize()) if not col_.fatal else dict(prev_ok=False, prev_fetched_at=None)
    obs_info["made_at"] = clock.now()
    try:
        svc = build_service(cfg, store, name, issue, obs_info, write=write) if build == "full" else light_service(cfg, store, name, issue)
    except Exception as e:                                    # 계산 오류도 슬롯 실패로 남긴다
        svc = dict(ok=False, error=f"{type(e).__name__}: {e}"[:200], src_run=None, backup=None, rs_method="", obs_prev="")
    _, deadline = slot_times(issue)
    on_time = bool(col["complete8"] and col["received_at"] is not None and pd.Timestamp(col["received_at"]) <= deadline
                   and not col.get("late"))
    if col.get("rerun"):                                     # 같은 슬롯 재실행: 정시 여부는 처음 실행 기록으로(보고서는 슬롯의 첫 행을 씀)
        on_time = None
    row = dict(slot_run=f"{issue:%Y-%m-%d %H:%M}", run_name=name, started_at=f"{started:%Y-%m-%d %H:%M:%S}",
               finished_at=f"{clock.now():%Y-%m-%d %H:%M:%S}", deadline=f"{deadline:%Y-%m-%d %H:%M}",
               service_received_at=f"{pd.Timestamp(col['received_at']):%Y-%m-%d %H:%M:%S}" if col["received_at"] is not None else "",
               on_time=on_time, result=slot_result(col, svc),
               src_run=f"{pd.Timestamp(svc['src_run']):%Y-%m-%d %H:%M}" if svc.get("src_run") is not None else "",
               backup=svc.get("backup"), rs_method=svc.get("rs_method", ""), obs_prev=svc.get("obs_prev", ""),
               obs_prev_fetched_at=f"{obs_info['prev_fetched_at']:%Y-%m-%d %H:%M:%S}" if obs_info.get("prev_fetched_at") is not None else "",
               attempts=col["attempts"], excel=svc.get("excel") or "", error=(col_.fatal or svc.get("error") or "")[:200])
    store.log("slot_log.csv", row, SLOT_COLS)
    row["_svc"] = svc
    return row


def probe_asos(cfg, http=K.http_get, clock=None, key=None, store=None):
    """전날 ASOS 일자료 조회 가능 시각 점검: 지금 D−1을 한 번 조회해 결과를 collect_log.csv(kind asos_probe)에 남김"""
    clock = clock or K.Clock()
    store = store or K.Store(cfg["data"])
    col_ = Collector(cfg, store, http, clock, key)
    now = clock.now()
    prev = now.normalize() - pd.Timedelta(days=1)
    res = K.fetch_asos_days(col_.key, cfg["stn"], prev, prev, http)
    items = [it for it in res["items"] if str(it.get("tm", ""))[:10] == f"{prev:%Y-%m-%d}"]
    ok = res["status"] == "ok" and len(items) > 0 and K.obs_rows_complete(items[0])
    if res["bodies"]:
        store.save_raw("asos", cfg["stn"], f"{prev:%Y%m%d}_{prev:%Y%m%d}", res["bodies"], now)
    if items:
        store.upsert_obs(cfg["stn"], items, now)
    col_._log(now, "asos_probe", f"{prev:%Y-%m-%d}", 1, res, status=("ok" if ok else ("partial" if items else res["status"])))
    return ok


# ── H3 보고서 ───────────────────────────────────────────────────────────
def h3_summary(slot_log, collect_log, since=None, until=None):
    """슬롯 로그 → H3 지표. 슬롯마다 첫 행(예정 시각의 실행)으로 정시 여부를 판단하고, 실행 기록이 없는 슬롯은 '실행 안 됨'(실패)"""
    s = slot_log.copy()
    if not len(s):
        return dict(n_expected=0), pd.DataFrame(), pd.DataFrame()
    s["slot_run"] = pd.to_datetime(s["slot_run"])
    s["started_at"] = pd.to_datetime(s["started_at"])
    t0 = pd.Timestamp(since) if since else s.slot_run.min().normalize()
    t1 = pd.Timestamp(until) + pd.Timedelta(hours=23, minutes=59) if until else s.slot_run.max()
    exp = [d + pd.Timedelta(hours=h) for d in pd.date_range(t0.normalize(), t1.normalize()) for h in SLOTS.values()]
    exp = [t for t in exp if t0 <= t <= t1]
    # 슬롯의 처음 실행·마지막 실행 행 그대로(groupby.first는 열마다 빈 값을 건너뛰어 다른 행의 값이 섞이므로 쓰지 않음)
    srt = s.sort_values("started_at", kind="stable")
    first = srt.drop_duplicates("slot_run", keep="first").set_index("slot_run")
    last = srt.drop_duplicates("slot_run", keep="last").set_index("slot_run")
    rows = []
    for t in exp:
        if t in first.index:
            f, l = first.loc[t], last.loc[t]
            ok = str(f.on_time).lower() == "true"
            rows.append(dict(slot_run=t, run_name=f.run_name, first_result=f.result, on_time=ok,
                             backup_excel=(not ok) and bool(str(f.excel).strip()) and f.result != "실패",
                             final_result=l.result, **{k: f.get(k, "") for k in ("src_run", "rs_method", "obs_prev", "attempts", "error")}))
        else:
            rows.append(dict(slot_run=t, run_name={h: n for n, h in SLOTS.items()}[t.hour], first_result="실행 안 됨",
                             on_time=False, backup_excel=False, final_result="실행 안 됨"))
    tab = pd.DataFrame(rows)
    n = len(tab)
    n_ok = int(tab.on_time.sum())
    fail = tab[~tab.on_time]
    summ = dict(n_expected=n, n_on_time=n_ok, rate=n_ok / n if n else np.nan, n_not_on_time=len(fail),
                n_backup_auto=int(fail.backup_excel.sum()),
                backup_rate=(float(fail.backup_excel.mean()) if len(fail) else np.nan),
                pass_rate=(n_ok / n >= 0.99) if n else False,
                pass_backup=(bool(fail.backup_excel.all()) if len(fail) else True),
                enough=n >= 100, since=t0, until=t1,
                counts=tab.first_result.value_counts().to_dict())
    summ["verdict"] = ("충족" if summ["pass_rate"] and summ["pass_backup"] else "미달") + ("" if summ["enough"] else f" (표본 {n}회 < 100회: 중간 점검)")
    # 전날 관측: 날짜별 처음 조회된 시각(수집·점검 로그), 아침 슬롯 시점의 조회 여부
    obs = pd.DataFrame()
    c = collect_log.copy() if len(collect_log) else pd.DataFrame()
    if len(c):
        c = c[c.kind.isin(["asos_d1", "asos_probe"])].copy()
        if len(c):
            c["logged_at"] = pd.to_datetime(c["logged_at"])
            c["target"] = pd.to_datetime(c["target"])
            okc = c[c.status == "ok"]
            first_ok = okc.groupby("target").logged_at.min()
            mor = c[(c.kind == "asos_d1") & (c.logged_at.dt.hour < 12)].groupby("target").status.first()
            obs = pd.DataFrame({"first_ok": first_ok, "morning_status": mor}).reset_index().rename(columns={"index": "target"})
            obs["lag_h"] = (obs.first_ok - obs.target) / pd.Timedelta(hours=1) - 24     # D−1 끝(자정)부터 몇 시간 뒤
            summ["obs_morning_ok"] = float((obs.morning_status == "ok").mean()) if obs.morning_status.notna().any() else np.nan
            summ["obs_lag_median_h"] = float(obs.lag_h.median()) if obs.lag_h.notna().any() else np.nan
    return summ, tab, obs


def build_report(cfg, since=None, until=None, out=None):
    store = K.Store(cfg["data"])
    sl, cl = store.read_log("slot_log.csv"), store.read_log("collect_log.csv")
    summ, tab, obs = h3_summary(sl, cl, since, until)
    from ops_report import build_h3_workbook
    out = out or os.path.join(cfg["out"], f"ops_h3_report({cfg['stn']})_{summ['since']:%Y%m%d}_{summ['until']:%Y%m%d}.xlsx")
    return summ, build_h3_workbook(summ, tab, obs, cl, out, cfg)


# ── CLI ──────────────────────────────────────────────────────────────────
def main(argv=None):
    ap = argparse.ArgumentParser(description="단기예보 운영 수집기·관수 전망 자동 생성 (02-Cycle G5)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(q):
        q.add_argument("--config", default=None, help="설정 파일(ops_config.json)")
        q.add_argument("--stn"); q.add_argument("--grid"); q.add_argument("--obs"); q.add_argument("--data")
        q.add_argument("--out"); q.add_argument("--apikey")
    r = sub.add_parser("run", help="지금 시각의 슬롯(또는 --slot)을 처리: 수집 → 관측 → 관수 전망 엑셀")
    common(r)
    r.add_argument("--slot", default=None, help="'YYYY-MM-DD HH' (02 또는 17). 없으면 지금 시각의 슬롯")
    r.add_argument("--no-wait", action="store_true", help="재시도를 기다리지 않고 한 번만 시도(수동 확인용)")
    q = sub.add_parser("probe-asos", help="전날 ASOS 일자료 조회 가능 시각 점검(매시 실행)")
    common(q)
    p = sub.add_parser("report", help="수집 로그 → H3 보고서 엑셀")
    common(p)
    p.add_argument("--since"); p.add_argument("--until"); p.add_argument("--report-out", default=None)
    g = sub.add_parser("grid", help="위경도 → 단기예보 격자")
    g.add_argument("lat", type=float); g.add_argument("lon", type=float)
    a = ap.parse_args(argv)
    if a.cmd == "grid":
        from kma_grid import latlon_to_grid
        print("%d_%d" % latlon_to_grid(a.lat, a.lon))
        return 0
    cfg = load_config(a.config, stn=a.stn, grid=a.grid, obs=a.obs, data=a.data, out=a.out, apikey=a.apikey)
    if a.cmd == "probe-asos":
        ok = probe_asos(cfg)
        print(f"[전날 관측] {'조회됨' if ok else '아직 없음'}")
        return 0
    if a.cmd == "report":
        summ, path = build_report(cfg, a.since, a.until, a.report_out)
        print(f"[H3] {summ['since']:%Y-%m-%d}~{summ['until']:%Y-%m-%d} 예정 {summ['n_expected']}회, 정시 성공 {summ['n_on_time']}회 "
              f"({summ['rate']:.1%}), 백업 자동 전환 {summ['n_backup_auto']}/{summ['n_not_on_time']} → {summ['verdict']}")
        print(f"[완료] {path}")
        return 0
    if not cfg.get("obs"):
        raise SystemExit("[설정] obs(01-Cycle 관측 워크북)가 필요합니다")
    name, issue = parse_slot(a.slot) if a.slot else (None, None)
    if a.no_wait:
        global RETRY_MIN
        RETRY_MIN = (0,)
    row = run_slot(cfg, name, issue)
    print(f"[슬롯] {row['run_name']} {row['slot_run']} → {row['result']} (발표 {row['src_run'] or '-'}, Rs {row['rs_method'] or '-'}, "
          f"전날 {row['obs_prev'] or '-'}, 시도 {row['attempts']}회)")
    if row["error"]:
        print(f"[오류] {row['error']}")
    if row["excel"]:
        print(f"[완료] {row['excel']}")
    return 0 if row["result"] != "실패" else 2


if __name__ == "__main__":
    sys.exit(main())
