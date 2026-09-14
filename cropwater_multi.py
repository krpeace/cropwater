#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cropwater_multi.py — 다지점 ETo·ETc·물수지 비교, FAO-56 PM vs 대형증발접시(Kp×Epan)

[함께 필요한 파일]
  fao56_core.py — 같은 폴더에 있어야 함

[두 개의 인증키를 쓴다]
  1) 기상 데이터(ASOS 일자료) : 공공데이터포털(data.go.kr) serviceKey
  2) 지점정보(위도·고도·풍속계높이) : 기상청 API허브(apihub.kma.go.kr) authKey

[실행 예시]
  python cropwater_multi.py --stns 101,119,216 --start 20260401 --end 20260831
  python cropwater_multi.py --stns 101,216 --start 20260701 --end 20260912 --zr 0.50 --pdep 0.40

[물수지 파라미터 (기본값: 양토·기준작물)]
  --zr   근권심도 m         (기본 0.50, FAO-56 Table 22 잔디)
  --fc   포장용수량 m³/m³   (기본 0.22, 양토 Table 19)
  --wp   위조점 m³/m³       (기본 0.10, 양토 Table 19)
  --pdep 고갈계수 p         (기본 0.40, 잔디 Table 22)
  → TAW = 1000×(FC−WP)×Zr, RAW = p×TAW
  물수지는 기준작물(Kc=1) 기준, 무관수(자연강우) 가정.

[apikey.txt 형식]
  DATA_GO_KR=<공공데이터포털 키>
  KMA_HUB=<기상청 API허브 키>
"""
import argparse, math, datetime as dt, os
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from fao56_core import (svp, slope_svp as slope, extra_radiation as extra_rad,
                         daylight_hours as daylight, wind_2m as wind2m,
                         kp_class_a as kp_pan, eto_penman_monteith as eto_pm,
                         num, fetch_station_table, fetch_asos, load_apikeys,
                         safe_save, load_station_backup, save_station_backup)

# ── 주요 지점 예시 ──
MAJOR=[(90,"속초"),(101,"춘천"),(93,"북춘천"),(105,"강릉"),(114,"원주"),(216,"태백"),(212,"홍천"),
       (108,"서울"),(112,"인천"),(119,"수원"),(232,"천안"),(127,"충주"),(131,"청주"),(133,"대전"),(129,"서산"),
       (136,"안동"),(137,"상주"),(138,"포항"),(143,"대구"),(278,"의성"),(283,"경주"),
       (146,"전주"),(172,"고창"),(156,"광주"),(165,"목포"),(168,"여수"),
       (159,"부산"),(152,"울산"),(192,"진주"),(253,"김해"),(184,"제주"),(189,"서귀포")]
DEFAULT_STNS=[101,93,216,119,232,127,136,192]
STN_NAMES={c:n for c,n in MAJOR}

FONT="맑은 고딕"
GREEN="2E6A4C"; BLUE="2B6E86"; BROWN="8B6914"
LIGHT="EFF1EC"; YEL="FBEED2"
OK_FILL="D4EDDA"; NG_FILL="F8D7DA"
th=Side("thin",color="D0D3CC"); BD=Border(th,th,th,th)

# ════════════════════════════════════════════════════════════════
# 계산 함수
# ════════════════════════════════════════════════════════════════

def compute_station(rows, lat, elev, anem, fetch=100.0):
    """ASOS 일자료 → ETo_PM, ETo_pan 계산. 물수지 필드(DP, Peff, Dr, Ks, irr)는 아직 없음."""
    out=[]; miss_si=miss_ev=0
    for x in rows:
        d=x.get("tm"); J=dt.datetime.strptime(d,"%Y-%m-%d").timetuple().tm_yday
        Tmax=num(x.get("maxTa")); Tmin=num(x.get("minTa")); RH=num(x.get("avgRhm")); ws=num(x.get("avgWs"))
        pv=num(x.get("avgPv")); td=num(x.get("avgTd")); pa=num(x.get("avgPa"))
        gsr=num(x.get("sumGsr")); ss=num(x.get("sumSsHr")); ev=num(x.get("sumLrgEv")); rn=num(x.get("sumRn")) or 0
        rec=dict(date=d, PM=None, pan=None, Epan=ev, rain=rn)
        if None in (Tmax,Tmin,ws): out.append(rec); continue
        es=(svp(Tmax)+svp(Tmin))/2
        ea=(pv*0.1 if pv is not None else (svp(td) if td is not None else (RH/100*es if RH else None)))
        u2=wind2m(ws, anem or 10)
        if gsr is not None: Rs=gsr
        elif ss is not None: Rs=(0.25+0.5*ss/daylight(lat,J))*extra_rad(lat,J); miss_si+=1
        else: Rs=None
        if Rs is not None and ea is not None: rec["PM"]=eto_pm(Tmax,Tmin,Rs,u2,ea,elev,lat,J,pa)
        if ev is not None and RH: rec["pan"]=kp_pan(u2,RH,fetch)*ev
        else: miss_ev+=1
        out.append(rec)
    return out, miss_si, miss_ev


def compute_water_balance(recs, taw, raw):
    """FAO-56 식(85) 일별 근권 물수지 추적.
       기준작물(Kc=1) · 무관수(자연강우만) 가정.
       recs 리스트에 DP, Peff, Dr, Ks, irr, In 필드를 추가하여 반환.
    """
    Dr = 0.0  # 초기 고갈량 = 0 (포장용수량 출발)
    for rec in recs:
        P   = rec.get("rain") or 0
        ETo = rec.get("PM")   # None이면 결측

        # DP [식88]: 강수가 현재 고갈량 초과분 → 심층침투
        DP = max(P - Dr, 0)
        Dr_after = max(Dr - P, 0)   # 강수 후 고갈량

        # Ks [식84]: 스트레스 계수
        if ETo is not None:
            Ks = 1.0 if Dr_after <= raw else max((taw - Dr_after) / (taw - raw), 0.0)
            ETc_adj = Ks * ETo   # Kc=1(기준작물)
        else:
            Ks = None
            ETc_adj = 0.0        # 결측일은 보수적으로 ETc=0 처리

        # Dr,i [식85]: 0 ≤ Dr ≤ TAW
        Dr_end = min(Dr_after + ETc_adj, taw)
        irr    = (Dr_end >= raw)

        rec["DP"]   = round(DP,      2)
        rec["Peff"] = round(P - DP,  2)
        rec["Dr"]   = round(Dr_end,  2)
        rec["Ks"]   = round(Ks,      3) if Ks is not None else None
        rec["irr"]  = irr
        rec["In"]   = round(Dr_end,  2) if irr else 0.0  # 필요 순관수량

        Dr = Dr_end
    return recs


def aggregate_monthly(recs):
    """월별 물수지 집계. 반환: {월번호: dict}"""
    m = defaultdict(lambda: dict(P=0.0, Peff=0.0, DP=0.0, ETo=0.0,
                                  n_eto=0, irr_days=0, stress_days=0, In=0.0))
    for r in recs:
        mo = int(r["date"][5:7])
        m[mo]["P"]    += r.get("rain") or 0
        m[mo]["Peff"] += r.get("Peff") or 0
        m[mo]["DP"]   += r.get("DP")   or 0
        m[mo]["In"]   += r.get("In")   or 0
        if r.get("PM") is not None:
            m[mo]["ETo"]   += r["PM"]
            m[mo]["n_eto"] += 1
        if r.get("irr"):
            m[mo]["irr_days"] += 1
        if r.get("Ks") is not None and r["Ks"] < 1.0:
            m[mo]["stress_days"] += 1
    return dict(m)


# ════════════════════════════════════════════════════════════════
# 엑셀 셀 헬퍼
# ════════════════════════════════════════════════════════════════

def H(ws, r, c, v, fill=LIGHT, white=False, sz=10):
    x = ws.cell(r, c, v)
    x.font      = Font(name=FONT, bold=True, size=sz, color=("FFFFFF" if white else "1A1A1A"))
    x.fill      = PatternFill("solid", fgColor=fill)
    x.alignment = Alignment("center", "center", wrap_text=True)
    x.border    = BD
    return x

def C(ws, r, c, v, fmt=None, fill=None):
    x = ws.cell(r, c, v)
    x.font      = Font(name=FONT, size=10)
    x.border    = BD
    x.alignment = Alignment("center", "center")
    if fmt and isinstance(v, (int, float)): x.number_format = fmt
    if fill: x.fill = PatternFill("solid", fgColor=fill)
    return x


# ════════════════════════════════════════════════════════════════
# 대화형 헬퍼
# ════════════════════════════════════════════════════════════════

def ask(p, d):
    try: v = input(f"{p} [{d}]: ").strip()
    except EOFError: v = ""
    return v or d

def getkey(arg, keys, label, prompt):
    if arg: return arg
    if keys.get(label): return keys[label]
    return input(prompt).strip()


# ════════════════════════════════════════════════════════════════
# 관수필요 달력 히트맵 + Dr 선형 차트
# ════════════════════════════════════════════════════════════════

def _hex_blend(c1, c2, t):
    """두 hex 색상(6자리) 사이를 t(0~1)로 선형 보간."""
    r = int(int(c1[0:2],16) + t*(int(c2[0:2],16)-int(c1[0:2],16)))
    g = int(int(c1[2:4],16) + t*(int(c2[2:4],16)-int(c1[2:4],16)))
    b = int(int(c1[4:6],16) + t*(int(c2[4:6],16)-int(c1[4:6],16)))
    return f"{max(0,min(255,r)):02X}{max(0,min(255,g)):02X}{max(0,min(255,b)):02X}"

def _dr_color(dr, raw, taw):
    """Dr 값 → 배경색 hex.  안전(녹)→주의(노)→위험(주황)→심각(적)."""
    if dr is None: return "D0D3CC"          # 회색(결측)
    if dr <= 0:    return "FFFFFF"           # 흰색(포장용수량)
    if dr < raw:
        t = dr / raw                         # 0→1
        return _hex_blend("D4EDDA","FFF3CD", t)   # 연녹 → 연노
    else:
        t = min(1.0, (dr - raw) / max(taw - raw, 1))
        return _hex_blend("FFF3CD","C82020", t)   # 연노 → 진적

def _dr_font_color(dr, raw, taw):
    """Dr이 TAW*0.6 이상이면 흰 글씨, 아니면 검정."""
    if dr is not None and dr >= taw * 0.6:
        return "FFFFFF"
    return "1A1A1A"

def build_calendar_sheet(wb, allrows, stns_list, stn_names, taw, raw):
    """
    관수필요_달력 시트 생성.
    - 상단: 날짜(행) × 지점(열) Dr 히트맵 (녹→적 4단계 색상)
    - 하단: Dr 선형 차트 (3개 지점 + RAW 기준선)
    """
    ws = wb.create_sheet("관수필요_달력")

    # ── 데이터 재구성: date_str → {stn: rec} ──
    date_stn = defaultdict(dict)
    for stn, nm, rec in allrows:
        date_stn[rec["date"]][stn] = rec
    dates = sorted(date_stn.keys())
    if not dates: return

    n_stns = len(stns_list)
    # 컬럼 레이아웃: A=일자, B=월, C~(C+n-1)=지점별Dr, 마지막=RAW
    COL_DATE  = 1
    COL_MONTH = 2
    COL_STN0  = 3                           # 첫 지점 Dr
    COL_RAW   = COL_STN0 + n_stns          # RAW 상수열
    CHART_DATA_ROW = 1                      # 헤더 행
    DATA_START = 2                          # 데이터 시작 행

    # ── 타이틀 ──
    title_cols = COL_RAW + 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=title_cols)
    tc = ws.cell(1, 1, f"관수필요 달력 히트맵  |  TAW={taw}mm, RAW={raw}mm  |  "
                       f"■ 안전(Dr<RAW)  ■ 위험(Dr≥RAW)  ■ 심각(Dr≥TAW×0.6)")
    tc.font      = Font(name=FONT, bold=True, size=11, color="2F5496")
    tc.alignment = Alignment(horizontal="left", vertical="center")

    # ── 헤더 행 ──
    HDR = 2
    H(ws, HDR, COL_DATE,  "일자",     BROWN, white=True, sz=9)
    H(ws, HDR, COL_MONTH, "월",       BROWN, white=True, sz=9)
    for i, stn in enumerate(stns_list):
        nm = stn_names.get(stn, str(stn))
        H(ws, HDR, COL_STN0+i, f"{nm}\nDr(mm)", BROWN, white=True, sz=9)
    H(ws, HDR, COL_RAW, f"RAW\n({raw}mm)", GREEN, white=True, sz=9)

    # ── 데이터 행 ──
    prev_month = None
    for ri, d in enumerate(dates, DATA_START + 1):  # +1 for HDR row
        d_dt = dt.datetime.strptime(d, "%Y-%m-%d") if isinstance(d, str) else d
        mo = d_dt.month

        def _bdr(is_month_start):
            top = Side("medium" if is_month_start else "thin", color="D0D3CC")
            return Border(left=Side("thin",color="D0D3CC"), right=Side("thin",color="D0D3CC"),
                          top=top, bottom=Side("thin",color="D0D3CC"))

        is_start = (mo != prev_month)
        prev_month = mo
        _al = Alignment("center","center")

        # 일자 셀
        dc = ws.cell(ri, COL_DATE, d_dt)
        dc.number_format = "m/d"; dc.font = Font(name=FONT, size=8)
        dc.border = _bdr(is_start); dc.alignment = _al

        # 월 셀
        mc = ws.cell(ri, COL_MONTH, mo)
        mc.number_format = "0"; mc.font = Font(name=FONT, size=8)
        mc.border = _bdr(is_start); mc.alignment = _al

        # 지점별 Dr 셀
        for i, stn in enumerate(stns_list):
            rec = date_stn[d].get(stn, {})
            dr  = rec.get("Dr")
            col = COL_STN0 + i
            bg  = _dr_color(dr, raw, taw)
            fc  = _dr_font_color(dr, raw, taw)
            cell = ws.cell(ri, col, round(dr,1) if dr is not None else None)
            cell.font      = Font(name=FONT, size=8, bold=(dr is not None and dr >= raw), color=fc)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.border    = _bdr(is_start)
            cell.alignment = _al
            if dr is not None: cell.number_format = "0.0"

        # RAW 상수열 (차트 기준선용)
        rc = ws.cell(ri, COL_RAW, raw)
        rc.number_format = "0.0"; rc.font = Font(name=FONT, size=8, color="888888")
        rc.border = _bdr(is_start); rc.alignment = _al

    last_data_row = DATA_START + len(dates)  # 마지막 데이터 행

    # ── 컬럼 너비 ──
    ws.column_dimensions[get_column_letter(COL_DATE)].width  = 6
    ws.column_dimensions[get_column_letter(COL_MONTH)].width = 3
    for i in range(n_stns):
        ws.column_dimensions[get_column_letter(COL_STN0+i)].width = 8
    ws.column_dimensions[get_column_letter(COL_RAW)].width = 7
    ws.column_dimensions[get_column_letter(COL_RAW)].hidden = True  # 차트 참조용, 표시 불필요
    ws.freeze_panes = f"{get_column_letter(COL_STN0)}{DATA_START+1}"
    ws.row_dimensions[HDR].height = 30

    # ── Dr 선형 차트 ──
    chart = LineChart()
    chart.title  = "일별 근권 고갈량(Dr) 추이 — 무관수·기준작물(Kc=1) 가정"
    chart.y_axis.title = "Dr (mm)"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = float(taw) * 1.05
    chart.x_axis.title = "날짜"
    chart.x_axis.numFmt = "m/d"
    chart.x_axis.majorUnit = 14         # 2주 눈금
    chart.style   = 10
    chart.width   = max(22, n_stns * 6)
    chart.height  = 12

    COLORS = ["2B6E86","E07000","2E6A4C","8B1A1A","6A2E8B","1A5C8B"]
    cats = Reference(ws, min_col=COL_DATE, min_row=DATA_START+1, max_row=last_data_row)

    for i, stn in enumerate(stns_list):
        nm  = stn_names.get(stn, str(stn))
        ref = Reference(ws, min_col=COL_STN0+i, min_row=HDR, max_row=last_data_row)
        chart.add_data(ref, titles_from_data=True)
        s = chart.series[-1]
        s.graphicalProperties.line.solidFill = COLORS[i % len(COLORS)]
        s.graphicalProperties.line.width = 18000
        s.smooth = False

    # RAW 기준선 (점선 적색)
    raw_ref = Reference(ws, min_col=COL_RAW, min_row=HDR, max_row=last_data_row)
    chart.add_data(raw_ref, titles_from_data=True)
    s_raw = chart.series[-1]
    s_raw.graphicalProperties.line.solidFill = "FF0000"
    s_raw.graphicalProperties.line.width = 12000
    s_raw.graphicalProperties.line.dashStyle = "dash"

    chart.set_categories(cats)

    # 차트 위치: 히트맵 아래
    chart_row = last_data_row + 2
    ws.add_chart(chart, f"A{chart_row}")

    # 범례 안내 (차트 위)
    legend_row = last_data_row + 1
    ws.merge_cells(start_row=legend_row, start_column=1,
                   end_row=legend_row, end_column=title_cols)
    lg = ws.cell(legend_row, 1,
        "▲ 차트: 실선=지점별 Dr | 적색점선=RAW(관수 트리거)  "
        "▲ 히트맵 색상: 흰색→연녹(Dr↑) → 연노(→RAW) → 연주황→진적(→TAW)")
    lg.font = Font(name=FONT, size=9, color="555555", italic=True)
    lg.alignment = Alignment(horizontal="left")


# ════════════════════════════════════════════════════════════════
# main
# ════════════════════════════════════════════════════════════════

def main():
    _today     = dt.date.today()
    _yesterday = (_today - dt.timedelta(days=1)).strftime("%Y%m%d")
    _year      = _today.year
    _default_start = f"{_year}0101"
    _default_stns  = "119,101,131,129,146,156,136,192,189"

    ap = argparse.ArgumentParser()
    ap.add_argument("--key",    default=None)
    ap.add_argument("--hubkey", default=None)
    ap.add_argument("--start",  default=None)
    ap.add_argument("--end",    default=None)
    ap.add_argument("--stns",   default=None)
    ap.add_argument("--fetch",  type=float, default=100.0)
    ap.add_argument("--out",    default=None)
    # ── 물수지 토양 파라미터 ──
    ap.add_argument("--zr",   type=float, default=0.50, help="근권심도 m (기본 0.50, 잔디)")
    ap.add_argument("--fc",   type=float, default=0.22, help="포장용수량 (기본 0.22, 양토)")
    ap.add_argument("--wp",   type=float, default=0.10, help="위조점 (기본 0.10, 양토)")
    ap.add_argument("--pdep", type=float, default=0.40, help="고갈계수 p (기본 0.40, 잔디)")
    a = ap.parse_args()

    # 토양·물수지 파라미터 계산
    TAW = round(1000 * (a.fc - a.wp) * a.zr, 1)   # mm
    RAW = round(a.pdep * TAW, 1)                    # mm
    Ea  = 0.95                                       # 점적관수 효율(고정)

    keys   = load_apikeys()
    key    = getkey(a.key,    keys, "DATA_GO_KR", "공공데이터포털 serviceKey: ")
    hubkey = getkey(a.hubkey, keys, "KMA_HUB",    "기상청 API허브 authKey: ")

    print("\n[지점정보 조회] stn_inf.php …")
    try:
        info, names = fetch_station_table(hubkey)
        save_station_backup(info, names)
        info_source = "API"
        print(f"  성공: {len(info)}개 지점 좌표 확보")
    except Exception as e:
        info, names = load_station_backup()
        info_source = "BACKUP"
        print(f"  API 연결 실패({e}) → 로컬 백업파일 사용 ({len(info)}개 지점)")

    if a.stns:
        stns = [int(s) for s in a.stns.split(",")]
    else:
        stns_input = ask(
            "분석 지점 (쉼표 구분, 119=수원·101=춘천·131=청주·129=서산·146=전주·156=광주·136=안동·192=진주·189=서귀포)",
            _default_stns)
        stns = [int(s.strip()) for s in stns_input.split(",")]

    start = a.start or ask("조회 시작일 YYYYMMDD", _default_start)
    end   = a.end   or ask("조회 종료일 YYYYMMDD", _yesterday)

    if a.out:
        out_path = a.out
    else:
        stn_str  = ", ".join(str(s) for s in stns)
        out_path = f"output/eto({stn_str})_{start}_{end}.xlsx"

    print(f"\n[분석] {len(stns)}개 지점, {start}~{end}")
    print(f"[물수지] TAW={TAW}mm, RAW={RAW}mm (Zr={a.zr}m, FC={a.fc}, WP={a.wp}, p={a.pdep}) | 기준작물(Kc=1) · 무관수 가정")

    # ════════════════════════════════
    # 시트 1: 지점비교
    # ════════════════════════════════
    wb = Workbook(); ws = wb.active; ws.title = "지점비교"

    heads_eto = ["지점","지점명","위도","고도(m)","풍속계(m)","정보출처",
                 "유효일수","일사결측","증발량결측",
                 "ETo_PM\n평균(mm/d)","ETo_pan\n평균(mm/d)","pan/PM","6월","7월","8월"]
    heads_wb  = ["ΣP\n(mm)","ΣPeff\n(mm)","유효강수율\n(%)","ΣETo\n(mm)",
                 "물수지\nΣPeff-ΣETo\n(mm)","스트레스\n일수","관수필요\n발생일수"]
    all_heads = heads_eto + heads_wb

    for c, h in enumerate(all_heads, 1):
        col_fill = GREEN if c <= len(heads_eto) else BROWN
        H(ws, 1, c, h, col_fill, white=True)

    allrows   = []          # (stn, nm, rec) 전체 일별 레코드
    all_monthly = {}        # {stn: monthly_dict}
    ri = 2

    for stn in stns:
        if stn not in info:
            print(f"  {stn}: 좌표 없음 → 건너뜀"); continue
        lat, lon, elev, anem = info[stn]
        nm  = names.get(stn) or STN_NAMES.get(stn, str(stn))
        src = info_source
        try:
            rows = fetch_asos(key, stn, start, end)
        except Exception as e:
            print(f"  {stn} {nm}: 데이터 조회 실패 {e}"); continue

        recs, msi, mev = compute_station(rows, lat, elev, anem, a.fetch)
        recs = compute_water_balance(recs, TAW, RAW)    # ← 물수지 추가

        valid = [r for r in recs if r["PM"] is not None and r["pan"] is not None]
        pm    = [r["PM"]  for r in valid]
        pan   = [r["pan"] for r in valid]
        pmavg  = (sum(pm)  / len(pm))  if pm  else None
        panavg = (sum(pan) / len(pan)) if pan else None
        ratio  = (sum(pan) / sum(pm))  if pm  else None

        def mr(mm):
            v = [r for r in valid if int(r["date"][5:7]) == mm]
            return (sum(x["pan"] for x in v) / sum(x["PM"] for x in v)) if v else None

        # ETo 컬럼
        C(ws,ri,1,stn,"0");     C(ws,ri,2,nm)
        C(ws,ri,3,round(lat,4),"0.0000");  C(ws,ri,4,round(elev,1),"0.0")
        C(ws,ri,5,round(anem,1),"0.0");    C(ws,ri,6,src)
        C(ws,ri,7,len(valid),"0");         C(ws,ri,8,msi,"0"); C(ws,ri,9,mev,"0")
        C(ws,ri,10,round(pmavg,2)  if pmavg  else None,"0.00")
        C(ws,ri,11,round(panavg,2) if panavg else None,"0.00")
        C(ws,ri,12,round(ratio,3)  if ratio  else None,"0.000")
        for k, mm in enumerate([6,7,8]):
            v = mr(mm); C(ws,ri,13+k,round(v,3) if v else None,"0.000")

        # 물수지 컬럼 (16~22)
        sum_P    = sum(r.get("rain") or 0 for r in recs)
        sum_Peff = sum(r.get("Peff") or 0 for r in recs)
        sum_DP   = sum(r.get("DP")   or 0 for r in recs)
        sum_ETo  = sum(r["PM"] for r in recs if r.get("PM") is not None)
        balance  = sum_Peff - sum_ETo
        stress   = sum(1 for r in recs if r.get("Ks") is not None and r["Ks"] < 1.0)
        irr_days = sum(1 for r in recs if r.get("irr"))
        peff_pct = (sum_Peff / sum_P * 100) if sum_P > 0 else None

        C(ws,ri,16,round(sum_P,1),  "0.0")
        C(ws,ri,17,round(sum_Peff,1),"0.0")
        C(ws,ri,18,round(peff_pct,1) if peff_pct else None,"0.0")
        C(ws,ri,19,round(sum_ETo,1), "0.0")
        bal_fill = OK_FILL if balance >= 0 else NG_FILL
        C(ws,ri,20,round(balance,1), "+0.0;-0.0", fill=bal_fill)
        C(ws,ri,21,stress,   "0", fill=(YEL if stress > 0 else None))
        C(ws,ri,22,irr_days, "0", fill=(NG_FILL if irr_days > 0 else None))

        print(f"  {stn} {nm:5}[{src:7}] 유효{len(valid)}일 | "
              f"ΣP={sum_P:.0f} ΣPeff={sum_Peff:.0f} ΣETo={sum_ETo:.0f} "
              f"물수지={balance:+.0f}mm | 관수필요{irr_days}일 스트레스{stress}일")

        for r in recs: allrows.append((stn, nm, r))
        all_monthly[stn] = {"nm": nm, "monthly": aggregate_monthly(recs)}
        ri += 1

    ws.freeze_panes = "C2"
    col_widths = [6,8,9,8,9,10,9,9,10,10,10,9,8,8,8, 8,8,9,8,10,9,9]
    for c, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ════════════════════════════════
    # 시트 2: 일별_전지점
    # ════════════════════════════════
    ws = wb.create_sheet("일별_전지점")
    day_heads = ["지점","지점명","일자",
                 "ETo_PM\n(mm)","ETo_pan\n(mm)","대형증발량\n(mm)","강수P\n(mm)",
                 "Peff\n(mm)","DP\n(mm)","Dr\n(mm)","Ks","관수\n필요"]
    for c, h in enumerate(day_heads, 1):
        col_fill = GREEN if c <= 7 else BROWN
        H(ws, 1, c, h, col_fill, white=True)

    ri = 2
    for stn, nm, r in allrows:
        C(ws,ri,1,stn,"0"); C(ws,ri,2,nm)
        cc = C(ws,ri,3,dt.datetime.strptime(r["date"],"%Y-%m-%d")); cc.number_format="yyyy-mm-dd"
        C(ws,ri,4, round(r["PM"],2)   if r.get("PM")   is not None else None,"0.00")
        C(ws,ri,5, round(r["pan"],2)  if r.get("pan")  is not None else None,"0.00")
        C(ws,ri,6, round(r["Epan"],1) if r.get("Epan") is not None else None,"0.0")
        C(ws,ri,7, round(r["rain"],1),"0.0")
        C(ws,ri,8, round(r.get("Peff") or 0,1),"0.0")
        C(ws,ri,9, round(r.get("DP")   or 0,1),"0.0")
        Dr_v = r.get("Dr"); Ks_v = r.get("Ks")
        dr_fill = NG_FILL if r.get("irr") else None
        C(ws,ri,10,round(Dr_v,1) if Dr_v is not None else None,"0.0",fill=dr_fill)
        C(ws,ri,11,round(Ks_v,3) if Ks_v is not None else None,"0.000",
          fill=(YEL if (Ks_v is not None and Ks_v < 1.0) else None))
        C(ws,ri,12,"●" if r.get("irr") else "",fill=dr_fill)
        ri += 1

    ws.freeze_panes = "D2"
    for c, w in enumerate([6,8,12,10,10,10,9,9,9,9,8,7], 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ════════════════════════════════
    # 시트 3: 월별_물수지
    # ════════════════════════════════
    ws = wb.create_sheet("월별_물수지")
    ws.merge_cells("A1:L1")
    ws.cell(1,1, f"월별 물수지 — 기준작물(Kc=1) · 무관수 가정 | TAW={TAW}mm, RAW={RAW}mm "
                 f"(Zr={a.zr}m, FC={a.fc}, WP={a.wp}, p={a.pdep})"
    ).font = Font(name=FONT, bold=True, size=11, color="2F5496")

    mb_heads = ["지점","지점명","월",
                "ΣP\n(mm)","ΣPeff\n(mm)","유효강수율\n(%)","ΣDP\n(mm)",
                "ΣETo\n(mm)","물수지\nΣPeff-ΣETo\n(mm)",
                "스트레스\n일수","관수필요\n발생일수","필요 순관수량\nΣIn(mm)"]
    for c, h in enumerate(mb_heads, 1):
        H(ws, 2, c, h, BROWN, white=True)

    ri = 3
    for stn in stns:
        if stn not in all_monthly: continue
        nm      = all_monthly[stn]["nm"]
        monthly = all_monthly[stn]["monthly"]
        for mo in sorted(monthly.keys()):
            d = monthly[mo]
            P    = d["P"];  Peff = d["Peff"]; DP = d["DP"]
            ETo  = d["ETo"]; balance = Peff - ETo
            pct  = (Peff / P * 100) if P > 0 else None

            C(ws,ri,1,stn,"0");  C(ws,ri,2,nm)
            C(ws,ri,3,mo,"0")
            C(ws,ri,4, round(P,1),    "0.0")
            C(ws,ri,5, round(Peff,1), "0.0")
            C(ws,ri,6, round(pct,1)   if pct else None,"0.0")
            C(ws,ri,7, round(DP,1),   "0.0")
            C(ws,ri,8, round(ETo,1),  "0.0")
            bal_fill = OK_FILL if balance >= 0 else NG_FILL
            C(ws,ri,9, round(balance,1),"+0.0;-0.0",fill=bal_fill)
            C(ws,ri,10,d["stress_days"],"0",fill=(YEL if d["stress_days"] > 0 else None))
            C(ws,ri,11,d["irr_days"],  "0",fill=(NG_FILL if d["irr_days"] > 0 else None))
            C(ws,ri,12,round(d["In"],1),"0.0",fill=(NG_FILL if d["irr_days"] > 0 else None))
            ri += 1

    ws.freeze_panes = "D3"
    for c, w in enumerate([6,8,5,8,8,9,8,8,10,9,9,10], 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ════════════════════════════════
    # 시트 4: 관수필요_달력 (히트맵 + Dr 차트)
    # ════════════════════════════════
    build_calendar_sheet(wb, allrows, stns, names, TAW, RAW)

    # ════════════════════════════════
    # 시트 5: 설명
    # ════════════════════════════════
    ws = wb.create_sheet("설명")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 90
    doc = [
        ("다지점 ETo·물수지 비교","",True),
        ("인증키","기상데이터=공공데이터포털 serviceKey / 지점정보=기상청 API허브 authKey (별개 포털·별개 키)"),
        ("지점정보","stn_inf.php(지점일람표)로 위도·고도·풍속계높이 자동조회, 성공 시 stations_backup.csv에 저장."),
        ("ETo_PM","FAO-56 Penman-Monteith [식6]. 일사 결측 시 일조시간(Ångström)으로 대체."),
        ("ETo_pan","증발접시법. Class A 계수 Kp(Allen-Pruitt) × 대형증발량."),
        ("판정","pan/PM이 전 지점 0.7~0.8 → 계통적 편차 / 편차 크면 지점특성 / 결측 많으면 대체경로 필요"),
        ("물수지 근거 (FAO-56 Ch.8)","",True),
        ("기준작물·무관수","물수지는 기준작물(Kc=1) 기준. 실측 관수 데이터 없이 자연강우만 반영."),
        (f"TAW [식82]",f"총유효수분 = 1000×(FC−WP)×Zr. 기본값 {TAW}mm (Zr={a.zr}m, FC={a.fc}, WP={a.wp})."),
        (f"RAW [식83]",f"쉽게이용가능수분 = p×TAW. 기본값 {RAW}mm (p={a.pdep}). Dr≥RAW이면 관수필요."),
        ("DP [식88]","심층침투 = max(P − Dr,i-1, 0). 강수가 고갈량 초과분은 근권 아래로 손실."),
        ("Ks [식84]","수분스트레스계수. Dr≤RAW이면 1.0, 초과 시 (TAW−Dr)/(TAW−RAW)."),
        ("Dr [식85]","일별 근권 고갈량. Dr,i = Dr,i-1 − P + ETc_adj(=Ks×ETo) + DP. 0 ≤ Dr ≤ TAW."),
        ("관수필요","Dr ≥ RAW이면 ●. Dr 미리셋(무관수 가정이므로 관수 후 초기화 없음)."),
        ("필요 순관수량 In","net irrigation depth. 관수필요 시점 Dr 값. 이론적 필요량, 실측값 아님."),
        ("파라미터 변경","--zr, --fc, --wp, --pdep 로 토양 조건 변경 가능. 작물별 Zr은 FAO-56 Table 22 참조."),
    ]
    for k,v,*hd in doc:
        head = bool(hd and hd[0])
        rr   = ws.max_row + (1 if ws.max_row > 1 or ws.cell(1,1).value else 0)
        H(ws,rr,1,k,(GREEN if head else LIGHT),white=head,sz=(11 if head else 10))
        cb = ws.cell(rr,2,v)
        cb.font      = Font(name=FONT, bold=head, size=(11 if head else 10),
                            color=("FFFFFF" if head else "1A1A1A"))
        cb.alignment = Alignment(wrap_text=True, vertical="top")
        if head: cb.fill = PatternFill("solid", fgColor=GREEN)

    saved_path = safe_save(wb, out_path)
    print(f"\n[완료] {saved_path} → 시트: 지점비교 / 일별_전지점 / 월별_물수지 / 관수필요_달력 / 설명")

if __name__ == "__main__": main()
