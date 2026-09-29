#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fcst_report.py — 02-Cycle H2 검증 엑셀 (라이브 수식)

시트: 요약 / 일별비교 / 3일누적 / 입력진단 / 오차분해 / Rs계수 / 관측 / 설정 / 방법 / 차트자료
  - 값으로 넣는 것: 예보 일 입력(fcst_archive 집계), ASOS 관측 일자료, 오차분해(Python 계산)
  - 나머지는 엑셀 수식: Ra·Rs·PM ETo·Kc·ETc·기준선·오차·지표·판정
  - 노란 칸(설정·Rs계수·합격 기준)을 바꾸면 전체가 다시 계산된다
"""
import datetime as dt

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

from fao56_core import safe_save

FONT = "맑은 고딕"; GREEN = "2E6A4C"; BLUE = "2B6E86"; BROWN = "A8681B"; LIGHT = "EFF1EC"; YEL = "FBEED2"
PASS_FILL = "D8EAD3"; FAIL_FILL = "F4C7B8"
thin = Side(style="thin", color="D0D3CC"); BORDER = Border(thin, thin, thin, thin)
F2, F3, PCT, DATE, DTM = "0.00", "0.000", "0%", "yyyy-mm-dd", "yyyy-mm-dd hh:mm"


def H(ws, r, c, v, fill=GREEN, white=True, size=10, wrap=True):
    x = ws.cell(r, c, v)
    x.font = Font(name=FONT, bold=True, size=size, color=("FFFFFF" if white else "1A1A1A"))
    x.fill = PatternFill("solid", fgColor=fill); x.border = BORDER
    x.alignment = Alignment("center", "center", wrap_text=wrap)
    return x


def C(ws, r, c, v, fmt=None, left=False, fill=None, bold=False, color="1A1A1A", size=10, wrap=False):
    x = ws.cell(r, c, v)
    x.font = Font(name=FONT, size=size, bold=bold, color=color); x.border = BORDER
    x.alignment = Alignment("left" if left else "center", "center", wrap_text=wrap)
    if fmt: x.number_format = fmt
    if fill: x.fill = PatternFill("solid", fgColor=fill)
    return x


def T(ws, r, c, v, size=10, bold=False, color="1A1A1A", wrap=False):
    x = ws.cell(r, c, v)
    x.font = Font(name=FONT, size=size, bold=bold, color=color)
    x.alignment = Alignment("left", "center", wrap_text=wrap)
    return x


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def pm(tmax, tmin, ea, u2, rs, rso, t, es, d, g):
    """FAO-56 식(6) 엑셀 수식 (01-Cycle 계산과정 시트와 같은 형태, G=0)"""
    rnl = (f"0.000000004903*((({tmax}+273.16)^4+({tmin}+273.16)^4)/2)*(0.34-0.14*SQRT({ea}))"
           f"*(1.35*MIN({rs}/{rso},1)-0.35)")
    return f"=(0.408*{d}*(0.77*{rs}-{rnl})+{g}*(900/({t}+273))*{u2}*({es}-{ea}))/({d}+{g}*(1+0.34*{u2}))"


E0 = lambda x: f"0.6108*EXP(17.27*{x}/({x}+237.3))"


# ── 설정 ────────────────────────────────────────────────────────────────
S_ROWS = {}   # 이름 → 셀 주소(설정!$B$n)


def sheet_settings(wb, res):
    ws = wb.create_sheet("설정")
    meta, kp, chk = res["meta"], res["kp"], res["check"]
    st = meta["settings"]
    T(ws, 1, 1, "설정 — 노란 칸을 바꾸면 모든 시트가 다시 계산됩니다", 12, True, GREEN)
    H(ws, 2, 1, "항목"); H(ws, 2, 2, "값"); H(ws, 2, 3, "근거·출처")
    rows = [
        ("[지점]", None, None, None),
        ("site", "지점", f"ASOS {res['stn']} {res.get('stn_name', '')}".strip(),
         f"검증 기준 관측소. 이 지점의 단기예보 격자: {res.get('stn_grid') or '미등록'} (ARCHITECTURE 7장)"),
        ("lat", "위도 (°)", meta["lat"], "Ra·Rso 계산. 01-Cycle 워크북 설정 시트 값 (기상청 지점정보)"),
        ("elev", "고도 (m)", meta["elev"], "Rso, 예보 기압(식7) 계산. 01-Cycle 워크북 설정 시트 값"),
        ("anem", "ASOS 풍속계 높이 (m)", meta.get("anem", 10.0), "관측 u2 환산(식47). 01-Cycle 워크북 설정 시트 값"),
        ("[예보]", None, None, None),
        ("grid", "예보 격자 (nx_ny)", ", ".join(chk["location"]), "과거 단기예보 CSV의 location (지점 격자와 다르면 대표성 오차가 달라짐)"),
        ("fanem", "예보 풍속 높이 (m)", 10.0, "단기예보 WSD는 10 m 풍속으로 가정 → u2 = 0.748 × u10"),
        ("pelev", "예보 기압 P (kPa)", "=101.3*((293-0.0065*$B$6)/293)^5.26", "FAO-56 식(7). 예보에 기압이 없어 고도로 추정"),
        ("rthr", "강수유무 기준 (mm/일)", 1.0, "일강수 ≥ 기준이면 1 (Rs 강수유무 보정의 입력)"),
        ("krs", "FAO-56 kRs (S1)", 0.16, "FAO-56 식(50) 내륙 기본값 (해안 0.19)"),
        ("[작물계수 Kc — 01-Cycle 워크북과 동일]", None, None, None),
        ("crop", "작물", st.get("작물", "사과"), "01-Cycle 워크북 설정 시트"),
        ("scn", "시나리오 번호", kp["scenario"], "FAO-56 Table 12 (참고 표시). 표값은 아래 3칸을 직접 수정"),
        ("kci_t", "표값 Kc_ini", kp["table"][0], "선택 시나리오의 표값"),
        ("kcm_t", "표값 Kc_mid", kp["table"][1], "선택 시나리오의 표값"),
        ("kce_t", "표값 Kc_end", kp["table"][2], "선택 시나리오의 표값"),
        ("bud", "생육 시작일", kp["bud"], "01-Cycle 워크북 설정 시트(발아일)"),
        ("li", "L_ini (일)", kp["L"][0], "FAO-56 Table 11"),
        ("ld", "L_dev (일)", kp["L"][1], "FAO-56 Table 11"),
        ("lm", "L_mid (일)", kp["L"][2], "FAO-56 Table 11"),
        ("ll", "L_late (일)", kp["L"][3], "FAO-56 Table 11"),
        ("h", "생육중기 수고 h (m)", kp["h"], "식(62)·(65)의 (h/3)^0.3 항"),
        ("u2m", "중기 평균 u2 (m/s)", kp["u2_mid"], "01-Cycle 자동집계값(2026 관측, 생육중기)"),
        ("rhm", "중기 평균 RHmin (%)", kp["rh_mid"], "01-Cycle 자동집계값"),
        ("u2e", "후기 평균 u2 (m/s)", kp["u2_end"], "01-Cycle 값(후기 관측 부족 시 기본값 1.5)"),
        ("rhe", "후기 평균 RHmin (%)", kp["rh_end"], "01-Cycle 값(후기 관측 부족 시 기본값 55)"),
        ("mulch", "멀칭 보정계수", kp["mulch"], "전면 비닐멀칭 0.70~0.90, 없으면 1.00"),
        ("kci", "★ 적용 Kc_ini", None, "Kc_ini × 멀칭 (현지기상 보정 대상 아님)"),
        ("kcm", "★ 적용 Kc_mid", None, "FAO-56 식(62): Kc_mid + [0.04(u2−2) − 0.004(RHmin−45)](h/3)^0.3, u2·RHmin 범위 제한"),
        ("kce", "★ 적용 Kc_end", None, "FAO-56 식(65): 표값 > 0.45일 때만 보정"),
    ]
    r = 3
    for key, label, val, note in rows:
        if label is None:
            H(ws, r, 1, key, fill=LIGHT, white=False); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
            r += 1; continue
        S_ROWS[key] = f"설정!$B${r}"
        C(ws, r, 1, label, left=True, fill=LIGHT)
        editable = key in ("lat", "elev", "anem", "fanem", "rthr", "krs", "kci_t", "kcm_t", "kce_t", "bud",
                           "li", "ld", "lm", "ll", "h", "u2m", "rhm", "u2e", "rhe", "mulch")
        fmt = DATE if key == "bud" else (F3 if isinstance(val, float) else None)
        C(ws, r, 2, val, fmt=fmt, fill=(YEL if editable else None))
        C(ws, r, 3, note, left=True, color="555555", size=9)
        r += 1
    assert S_ROWS["elev"] == "설정!$B$6", S_ROWS["elev"]          # pelev 수식이 $B$6을 참조
    s = lambda k: S_ROWS[k].replace("설정!", "")
    ws[s("kci")] = f"={s('kci_t')}*{s('mulch')}"
    ws[s("kcm")] = (f"=({s('kcm_t')}+(0.04*(MEDIAN(1,{s('u2m')},6)-2)-0.004*(MEDIAN(20,{s('rhm')},80)-45))"
                    f"*({s('h')}/3)^0.3)*{s('mulch')}")
    ws[s("kce")] = (f"=IF({s('kce_t')}>0.45,{s('kce_t')}+(0.04*(MEDIAN(1,{s('u2e')},6)-2)-0.004*(MEDIAN(20,{s('rhe')},80)-45))"
                    f"*({s('h')}/3)^0.3,{s('kce_t')})*{s('mulch')}")
    for k in ("kci", "kcm", "kce", "pelev"):
        ws[s(k)].number_format = F3
    widths(ws, {"A": 26, "B": 18, "C": 80})
    ws.freeze_panes = "A3"
    return ws


# ── Rs계수 ──────────────────────────────────────────────────────────────
R_ROWS = {}


def sheet_rscoef(wb, res, obs_rows):
    ws = wb.create_sheet("Rs계수")
    cf = res["coef"]; fitinfo = cf
    T(ws, 1, 1, "Rs 추정 계수와 H1 재검증 (G2) — 계수는 검증 연도(2026)와 겹치지 않는 해의 관측으로 정함", 12, True, GREEN)
    H(ws, 3, 1, "항목"); H(ws, 3, 2, "값"); H(ws, 3, 3, "설명")
    items = [
        ("a", "a (S3)", cf.get("a"), "Rs/Ra = a + b·√(Tmax−Tmin) + c·강수유무  [THEORY 9장]"),
        ("b", "b (S3)", cf.get("b"), ""),
        ("c", "c (S3)", cf.get("c"), "비 오는 날(≥ 1 mm) 일사 감소분"),
        ("krs2", "지점 kRs (S2)", cf.get("krs"), "식(50)의 kRs를 같은 자료로 맞춘 값"),
        ("src", "출처", cf.get("source", ""), "rs_coef.csv (cropwater_fcst.py calib)"),
        ("fit", "보정 자료", f"{fitinfo.get('fit_start', '')} ~ {fitinfo.get('fit_end', '')} ({fitinfo.get('n', '')}일)",
         fitinfo.get("note", "")),
        ("fitrmse", "보정 자료 Rs RMSE (MJ/m²/일)", fitinfo.get("rmse_rs"), "S3 적합 잔차"),
    ]
    r = 4
    for key, label, val, note in items:
        R_ROWS[key] = f"Rs계수!$B${r}"
        C(ws, r, 1, label, left=True, fill=LIGHT)
        C(ws, r, 2, val, fmt=("0.0000" if isinstance(val, float) else None),
          fill=(YEL if key in ("a", "b", "c", "krs2") else None), left=isinstance(val, str))
        C(ws, r, 3, note, left=True, color="555555", size=9)
        r += 1
    r += 1
    T(ws, r, 1, "H1 재검증 — 관측 입력(기온·습도·풍속·기압)에 Rs만 추정 → 관측 ETo와 비교", 11, True, GREEN); r += 1
    C(ws, r, 1, "검증 시작일", left=True, fill=LIGHT); C(ws, r, 2, res["h1_start"].date(), DATE, fill=YEL); R_ROWS["h1s"] = f"$B${r}"; r += 1
    C(ws, r, 1, "검증 종료일", left=True, fill=LIGHT); C(ws, r, 2, res["h1_end"].date(), DATE, fill=YEL); R_ROWS["h1e"] = f"$B${r}"; r += 1
    C(ws, r, 1, "H1 기준: RMSE 상한 (mm/일)", left=True, fill=LIGHT); C(ws, r, 2, 0.6, F2, fill=YEL); R_ROWS["h1rm"] = f"$B${r}"; r += 1
    C(ws, r, 1, "H1 기준: 합계오차 한계 (±)", left=True, fill=LIGHT); C(ws, r, 2, 0.05, PCT, fill=YEL); R_ROWS["h1se"] = f"$B${r}"; r += 1
    r += 1
    hdr = ["방법", "n", "관측 ETo 평균", "추정 ETo 평균", "MBE", "RMSE", "합계오차", "Rs RMSE", "H1 판정"]
    for j, h in enumerate(hdr, 1):
        H(ws, r, j, h)
    r += 1
    a0, a1 = obs_rows
    rng = lambda col: f"관측!${col}${a0}:${col}${a1}"
    msk = f"(({rng('A')}>={R_ROWS['h1s']})*({rng('A')}<={R_ROWS['h1e']}))"
    R_ROWS["h1_first"] = r
    for name, eto, rs in (("S1 식(50) kRs 0.16", "AC", "Z"), ("S2 식(50) 지점 kRs", "AD", "AA"),
                          ("S3 식(50)+강수유무 보정 ★", "AE", "AB")):
        C(ws, r, 1, name, left=True, fill=LIGHT)
        ws.cell(r, 2, f"=SUMPRODUCT({msk}*1)")
        ws.cell(r, 3, f"=SUMPRODUCT({msk}*{rng('X')})/B{r}")
        ws.cell(r, 4, f"=SUMPRODUCT({msk}*{rng(eto)})/B{r}")
        ws.cell(r, 5, f"=D{r}-C{r}")
        ws.cell(r, 6, f"=SQRT(SUMPRODUCT({msk}*({rng(eto)}-{rng('X')})^2)/B{r})")
        ws.cell(r, 7, f"=D{r}/C{r}-1")
        ws.cell(r, 8, f"=SQRT(SUMPRODUCT({msk}*({rng(rs)}-{rng('I')})^2)/B{r})")
        ws.cell(r, 9, f'=IF(AND(F{r}<={R_ROWS["h1rm"]},ABS(G{r})<={R_ROWS["h1se"]}),"충족","미달")')
        for j, fmt in zip(range(2, 10), ["0", F2, F2, F2, F2, "0.0%", F2, None]):
            x = ws.cell(r, j); x.number_format = fmt or "General"; x.font = Font(name=FONT, size=10)
            x.border = BORDER; x.alignment = Alignment("center", "center")
        r += 1
    r += 1
    T(ws, r, 1, "월별 편향 (S3 추정 ETo − 관측 ETo, mm/일)", 11, True, GREEN); r += 1
    for j, h in enumerate(["월", "일수", "관측 ETo 평균", "S3 편향", "S1 편향"], 1):
        H(ws, r, j, h)
    r += 1
    for m in range(res["h1_start"].month, res["h1_end"].month + 1):
        mm = f"(({rng('K')}={m})*({rng('A')}>={R_ROWS['h1s']})*({rng('A')}<={R_ROWS['h1e']}))"
        C(ws, r, 1, m)
        ws.cell(r, 2, f"=SUMPRODUCT({mm}*1)")
        ws.cell(r, 3, f"=IF(B{r}>0,SUMPRODUCT({mm}*{rng('X')})/B{r},0)")
        ws.cell(r, 4, f"=IF(B{r}>0,SUMPRODUCT({mm}*({rng('AE')}-{rng('X')}))/B{r},0)")
        ws.cell(r, 5, f"=IF(B{r}>0,SUMPRODUCT({mm}*({rng('AC')}-{rng('X')}))/B{r},0)")
        for j, fmt in zip(range(2, 6), ["0", F2, F2, F2]):
            x = ws.cell(r, j); x.number_format = fmt; x.font = Font(name=FONT, size=10); x.border = BORDER
            x.alignment = Alignment("center", "center")
        r += 1
    widths(ws, {"A": 30, "B": 16, "C": 16, "D": 16, "E": 10, "F": 10, "G": 10, "H": 10, "I": 10})
    ws.column_dimensions["C"].width = 16
    return ws


# ── 관측 ────────────────────────────────────────────────────────────────
OBS_COLS = ["일자", "최고기온 Tmax(℃)", "최저기온 Tmin(℃)", "평균습도(%)", "평균풍속 u10(m/s)", "평균증기압(hPa)",
            "평균이슬점(℃)", "현지기압(hPa)", "일사 Rs(MJ/m²)", "일강수(mm)", "월", "연중일 J", "ea(kPa)", "P(kPa)",
            "u2(m/s)", "적위 δ", "일몰각 ωs", "Ra(MJ/m²)", "Rso(MJ/m²)", "T평균(℃)", "es(kPa)", "Δ(kPa/℃)", "γ(kPa/℃)",
            "★ 관측 ETo(mm)", "강수유무", "Rs S1", "Rs S2", "Rs S3", "ETo S1(관측입력)", "ETo S2(관측입력)",
            "ETo S3(관측입력)", "Kc", "관측 ETc(mm)", "직전 7일 평균 ETo"]


def sheet_obs(wb, obs):
    ws = wb.create_sheet("관측")
    for j, h in enumerate(OBS_COLS, 1):
        H(ws, 1, j, h, fill=(GREEN if j <= 10 else BLUE))
    S = S_ROWS; R = R_ROWS
    r0 = 2
    for i, row in enumerate(obs.itertuples()):
        r = r0 + i
        vals = [row.date.date(), row.Tmax, row.Tmin, row.RHmean, row.u10, row.pv, row.td, row.pa, row.Rs, row.rain]
        for j, v in enumerate(vals, 1):
            v = None if (not isinstance(v, dt.date) and pd.isna(v)) else v
            x = ws.cell(r, j, v); x.number_format = DATE if j == 1 else "0.0" if j != 9 else F2
        f = {
            "K": f"=MONTH(A{r})",
            "L": f"=A{r}-DATE(YEAR(A{r}),1,1)+1",
            "M": f'=IF(F{r}<>"",F{r}/10,IF(G{r}<>"",{E0(f"G{r}")},D{r}/100*({E0(f"B{r}")}+{E0(f"C{r}")})/2))',
            "N": f'=IF(H{r}<>"",H{r}/10,{S["pelev"]})',
            "O": f"=E{r}*4.87/LN(67.8*{S['anem']}-5.42)",
            "P": f"=0.409*SIN(2*PI()*L{r}/365-1.39)",
            "Q": f"=ACOS(-TAN(RADIANS({S['lat']}))*TAN(P{r}))",
            "R": f"=(24*60/PI())*0.082*(1+0.033*COS(2*PI()*L{r}/365))*(Q{r}*SIN(RADIANS({S['lat']}))*SIN(P{r})"
                 f"+COS(RADIANS({S['lat']}))*COS(P{r})*SIN(Q{r}))",
            "S": f"=(0.75+0.00002*{S['elev']})*R{r}",
            "T": f"=(B{r}+C{r})/2",
            "U": f"=({E0(f'B{r}')}+{E0(f'C{r}')})/2",
            "V": f"=4098*{E0(f'T{r}')}/(T{r}+237.3)^2",
            "W": f"=0.000665*N{r}",
            "X": pm(f"B{r}", f"C{r}", f"M{r}", f"O{r}", f"I{r}", f"S{r}", f"T{r}", f"U{r}", f"V{r}", f"W{r}"),
            "Y": f"=IF(J{r}>={S['rthr']},1,0)",
            "Z": f"=MIN(MAX({S['krs']}*SQRT(MAX(B{r}-C{r},0))*R{r},0.05*R{r}),S{r})",
            "AA": f"=MIN(MAX({R['krs2']}*SQRT(MAX(B{r}-C{r},0))*R{r},0.05*R{r}),S{r})",
            "AB": f"=MIN(MAX(({R['a']}+{R['b']}*SQRT(MAX(B{r}-C{r},0))+{R['c']}*Y{r})*R{r},0.05*R{r}),S{r})",
            "AC": pm(f"B{r}", f"C{r}", f"M{r}", f"O{r}", f"Z{r}", f"S{r}", f"T{r}", f"U{r}", f"V{r}", f"W{r}"),
            "AD": pm(f"B{r}", f"C{r}", f"M{r}", f"O{r}", f"AA{r}", f"S{r}", f"T{r}", f"U{r}", f"V{r}", f"W{r}"),
            "AE": pm(f"B{r}", f"C{r}", f"M{r}", f"O{r}", f"AB{r}", f"S{r}", f"T{r}", f"U{r}", f"V{r}", f"W{r}"),
        }
        d = f"(A{r}-{S['bud']})"
        li, ld, lm, ll = S["li"], S["ld"], S["lm"], S["ll"]
        ki, km, ke = S["kci"], S["kcm"], S["kce"]
        f["AF"] = (f"=IF({d}<0,0,IF({d}<{li},{ki},IF({d}<{li}+{ld},{ki}+({km}-{ki})*({d}-{li}+1)/{ld},"
                   f"IF({d}<{li}+{ld}+{lm},{km},IF({d}<{li}+{ld}+{lm}+{ll},{km}+({ke}-{km})*({d}-{li}-{ld}-{lm}+1)/{ll},0)))))")
        f["AG"] = f"=AF{r}*X{r}"
        f["AH"] = f"=AVERAGE(X{r - 7}:X{r - 1})" if i >= 7 else ""
        for col, formula in f.items():
            x = ws[f"{col}{r}"]; x.value = formula if formula else None
            x.number_format = "0" if col in ("K", "L", "Y") else F3 if col in ("M", "N", "P", "Q", "U", "V", "AF") else \
                "0.00000" if col == "W" else F2
    r1 = r0 + len(obs) - 1
    for j in range(1, len(OBS_COLS) + 1):
        ws.column_dimensions[CL(j)].width = 11
    ws.column_dimensions["A"].width = 12
    ws.freeze_panes = "B2"
    return ws, (r0, r1)


# ── 일별비교 ─────────────────────────────────────────────────────────────
DB_GROUPS = [("A", "F", "발표·대상일", GREEN), ("G", "N", "예보 일 입력 (fcst_archive 집계값)", GREEN),
             ("O", "X", "Rs 추정·PM 중간값 (수식)", BLUE), ("Y", "AC", "ETo (mm/일)", BROWN),
             ("AD", "AG", "ETo 오차 (−관측)", BROWN), ("AH", "AO", "Kc·ETc (mm/일)", BLUE),
             ("AP", "BB", "입력 진단 (관측값·예보 오차)", GREEN)]
DB_COLS = ["발표번호", "구분", "발표시각", "발표일", "선행일 k", "대상일",
           "마지막날(3시간·코드)", "예보 Tmax(℃)", "예보 Tmin(℃)", "예보 ea(kPa)", "예보 u10(m/s)", "예보 강수(mm)",
           "시각 수", "이전 발표로 채운 시각 수",
           "강수유무", "Ra", "Rso", "Rs S3", "Rs S1", "u2(m/s)", "T평균", "es", "Δ", "γ",
           "★ 예보 ETo S3", "예보 ETo S1", "관측 ETo", "지속성", "7일평균",
           "오차 S3", "오차 S1", "오차 지속성", "오차 7일평균",
           "Kc", "★ 예보 ETc", "관측 ETc", "지속성 ETc", "7일평균 ETc", "ETc 오차", "지속성 ETc 오차", "7일평균 ETc 오차",
           "관측 Tmax", "관측 Tmin", "관측 ea", "관측 u10", "관측 Rs", "관측 강수유무",
           "Tmax 오차", "Tmin 오차", "일교차 오차", "ea 오차", "u10 오차", "Rs 오차(S3)", "강수 판정"]


def sheet_daily(wb, df, obs_rows):
    ws = wb.create_sheet("일별비교")
    for c0, c1, label, fill in DB_GROUPS:
        H(ws, 1, _ci(c0), label, fill=fill)
        ws.merge_cells(f"{c0}1:{c1}1")
    for j, h in enumerate(DB_COLS, 1):
        H(ws, 2, j, h, fill=LIGHT, white=False, size=9)
    S, R = S_ROWS, R_ROWS
    a0, a1 = obs_rows
    ob = lambda col: f"관측!${col}${a0}:${col}${a1}"
    look = lambda col, key: f"INDEX({ob(col)},MATCH({key},{ob('A')},0))"
    r0 = 3
    run_ids = {run: i + 1 for i, run in enumerate(sorted(df.run.unique(), key=lambda t: (t.hour != 2, t)))}
    df = df.sort_values(["run_name", "run", "lead_day"], key=lambda s: s.map({"아침": 0, "저녁": 1}) if s.name == "run_name" else s)
    for i, row in enumerate(df.itertuples()):
        r = r0 + i
        vals = [run_ids[row.run], row.run_name, row.run.to_pydatetime(), row.run.date(), int(row.lead_day),
                row.target.date(), ("○" if row.ext else ""), row.Tmax, row.Tmin, float(row.ea), float(row.u10),
                float(row.rain), int(row.hours), int(row.filled)]
        for j, v in enumerate(vals, 1):
            x = ws.cell(r, j, v)
            x.number_format = {3: DTM, 4: DATE, 6: DATE, 10: F3, 11: F3}.get(j, "General")
        f = {
            "O": f"=IF(L{r}>={S['rthr']},1,0)",
            "P": f"={look('R', f'F{r}')}",
            "Q": f"={look('S', f'F{r}')}",
            "R": f"=MIN(MAX(({R['a']}+{R['b']}*SQRT(MAX(H{r}-I{r},0))+{R['c']}*O{r})*P{r},0.05*P{r}),Q{r})",
            "S": f"=MIN(MAX({S['krs']}*SQRT(MAX(H{r}-I{r},0))*P{r},0.05*P{r}),Q{r})",
            "T": f"=K{r}*4.87/LN(67.8*{S['fanem']}-5.42)",
            "U": f"=(H{r}+I{r})/2",
            "V": f"=({E0(f'H{r}')}+{E0(f'I{r}')})/2",
            "W": f"=4098*{E0(f'U{r}')}/(U{r}+237.3)^2",
            "X": f"=0.000665*{S['pelev']}",
            "Y": pm(f"H{r}", f"I{r}", f"J{r}", f"T{r}", f"R{r}", f"Q{r}", f"U{r}", f"V{r}", f"W{r}", f"X{r}"),
            "Z": pm(f"H{r}", f"I{r}", f"J{r}", f"T{r}", f"S{r}", f"Q{r}", f"U{r}", f"V{r}", f"W{r}", f"X{r}"),
            "AA": f"={look('X', f'F{r}')}",
            "AB": f"={look('X', f'D{r}-1')}",
            "AC": f"={look('AH', f'D{r}')}",
            "AD": f"=Y{r}-AA{r}", "AE": f"=Z{r}-AA{r}", "AF": f"=AB{r}-AA{r}", "AG": f"=AC{r}-AA{r}",
            "AH": f"={look('AF', f'F{r}')}",
            "AI": f"=AH{r}*Y{r}", "AJ": f"=AH{r}*AA{r}", "AK": f"=AH{r}*AB{r}", "AL": f"=AH{r}*AC{r}",
            "AM": f"=AI{r}-AJ{r}", "AN": f"=AK{r}-AJ{r}", "AO": f"=AL{r}-AJ{r}",
            "AP": f"={look('B', f'F{r}')}", "AQ": f"={look('C', f'F{r}')}", "AR": f"={look('M', f'F{r}')}",
            "AS": f"={look('E', f'F{r}')}", "AT": f"={look('I', f'F{r}')}", "AU": f"={look('Y', f'F{r}')}",
            "AV": f"=H{r}-AP{r}", "AW": f"=I{r}-AQ{r}", "AX": f"=(H{r}-I{r})-(AP{r}-AQ{r})", "AY": f"=J{r}-AR{r}",
            "AZ": f"=K{r}-AS{r}", "BA": f"=R{r}-AT{r}",
            "BB": f'=IF(O{r}=1,IF(AU{r}=1,"적중","오보"),IF(AU{r}=1,"놓침","무강수 일치"))',
        }
        for col, formula in f.items():
            x = ws[f"{col}{r}"]; x.value = formula
            x.number_format = "0" if col in ("O", "AU") else F3 if col in ("V", "W", "AH", "AR", "AY") else \
                "0.00000" if col == "X" else "General" if col == "BB" else F2
    r1 = r0 + len(df) - 1
    for j in range(1, len(DB_COLS) + 1):
        ws.column_dimensions[CL(j)].width = 10
    for col, w in {"B": 6, "C": 16, "D": 11, "F": 11, "A": 7, "E": 7}.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[2].height = 42
    ws.freeze_panes = "G3"
    ws.auto_filter.ref = f"A2:{CL(len(DB_COLS))}{r1}"
    return ws, (r0, r1), run_ids


def _ci(col):
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(col)


# ── 3일누적 ─────────────────────────────────────────────────────────────
def sheet_cum3(wb, df, db_rows, run_ids):
    ws = wb.create_sheet("3일누적")
    T(ws, 1, 1, "발표별 첫 3개 대상일 누적 (아침 D+0~D+2, 저녁 D+1~D+3) — FAO-56은 추정 Rs 기반 ETo를 여러 날 합계로 쓰도록 권고", 11, True, GREEN)
    hdr = ["발표번호", "구분", "발표시각", "대상기간", "첫 선행일", "예보 ETo 3일합", "관측 ETo 3일합", "지속성 3일합",
           "ETo 오차", "지속성 오차", "예보 ETc 3일합", "관측 ETc 3일합", "지속성 ETc 3일합", "ETc 오차", "지속성 ETc 오차"]
    for j, h in enumerate(hdr, 1):
        H(ws, 2, j, h, size=9)
    d0, d1 = db_rows
    db = lambda col: f"일별비교!${col}${d0}:${col}${d1}"
    runs = df.drop_duplicates("run")[["run", "run_name"]].copy()
    runs["id"] = runs.run.map(run_ids)
    runs = runs.sort_values("id")
    r = 3
    for row in runs.itertuples():
        k0 = int(df[df.run == row.run].lead_day.min())
        C(ws, r, 1, row.id); C(ws, r, 2, row.run_name); C(ws, r, 3, row.run.to_pydatetime(), DTM)
        C(ws, r, 4, f"D+{k0}~D+{k0 + 2}"); C(ws, r, 5, k0)
        s = lambda col: f'=SUMIFS({db(col)},{db("A")},$A{r},{db("E")},">="&$E{r},{db("E")},"<="&($E{r}+2))'
        for j, col in zip(range(6, 9), ("Y", "AA", "AB")):
            C(ws, r, j, s(col), F2)
        C(ws, r, 9, f"=F{r}-G{r}", F2); C(ws, r, 10, f"=H{r}-G{r}", F2)
        for j, col in zip(range(11, 14), ("AI", "AJ", "AK")):
            C(ws, r, j, s(col), F2)
        C(ws, r, 14, f"=K{r}-L{r}", F2); C(ws, r, 15, f"=M{r}-L{r}", F2)
        r += 1
    widths(ws, {CL(j): 11 for j in range(1, 16)}); ws.column_dimensions["C"].width = 16
    ws.freeze_panes = "D3"
    return ws, (3, r - 1)


# ── 입력진단 ─────────────────────────────────────────────────────────────
def sheet_diag(wb, db_rows, groups):
    ws = wb.create_sheet("입력진단")
    T(ws, 1, 1, "입력 진단 — 예보 입력 − 관측 (편향 = 평균, RMSE). 기온 정의 차이: 예보 TMX 09~18시·TMN 03~09시, ASOS 0~24시", 11, True, GREEN)
    hdr = ["구분", "선행일", "n", "Tmax 편향", "Tmax RMSE", "Tmin 편향", "Tmin RMSE", "일교차 편향", "일교차 RMSE",
           "ea 편향", "ea RMSE", "u10 편향", "u10 RMSE", "Rs 편향", "Rs RMSE", "적중", "놓침", "오보", "무강수 일치",
           "POD", "FAR", "CSI"]
    for j, h in enumerate(hdr, 1):
        H(ws, 2, j, h, size=9)
    d0, d1 = db_rows
    db = lambda col: f"일별비교!${col}${d0}:${col}${d1}"
    r = 3
    for rn, k in groups:
        C(ws, r, 1, rn); C(ws, r, 2, k)
        crit = f'{db("B")},$A{r},{db("E")},$B{r}'
        m = f"({db('B')}=$A{r})*({db('E')}=$B{r})"
        C(ws, r, 3, f"=COUNTIFS({crit})", "0")
        for j, col in zip(range(4, 16, 2), ("AV", "AW", "AX", "AY", "AZ", "BA")):
            fmt = F3 if col == "AY" else F2
            C(ws, r, j, f"=AVERAGEIFS({db(col)},{crit})", fmt)
            C(ws, r, j + 1, f"=SQRT(SUMPRODUCT({m}*{db(col)}^2)/$C{r})", fmt)
        for j, lab in zip(range(16, 20), ("적중", "놓침", "오보", "무강수 일치")):
            C(ws, r, j, f'=COUNTIFS({crit},{db("BB")},"{lab}")', "0")
        C(ws, r, 20, f"=IF(P{r}+Q{r}>0,P{r}/(P{r}+Q{r}),0)", F2)
        C(ws, r, 21, f"=IF(P{r}+R{r}>0,R{r}/(P{r}+R{r}),0)", F2)
        C(ws, r, 22, f"=IF(P{r}+Q{r}+R{r}>0,P{r}/(P{r}+Q{r}+R{r}),0)", F2)
        r += 1
    r += 1
    notes = ["단위: 기온 ℃, ea kPa, u10 m/s, Rs MJ/m²/일. 편향 = 예보 − 관측 (+면 예보가 큼).",
             "Rs 예보는 관측이 아니라 식(50)+강수유무 보정으로 추정한 값(S3)입니다.",
             "강수: 예보 PCP 일합계와 ASOS 일강수를 1 mm 기준으로 판정. POD = 적중/(적중+놓침), FAR = 오보/(적중+오보), CSI = 적중/(적중+놓침+오보).",
             "아침 D+k와 전날 저녁 D+(k+1)은 같은 대상일입니다. 최고·최저기온은 17시 발표 값이 다음 날 02시 발표까지 유지되는 경우가 많아 두 발표의 기온 진단이 거의 같습니다."]
    for n in notes:
        T(ws, r, 1, n, 9, color="555555"); r += 1
    widths(ws, {CL(j): 9 for j in range(1, 23)})
    return ws


# ── 오차분해 (Python 계산값) ─────────────────────────────────────────────
def sheet_attr(wb, res):
    ws = wb.create_sheet("오차분해")
    T(ws, 1, 1, "오차 분해와 편향 보정 탐색 — Python(cropwater_fcst.py) 계산값 (수식 아님)", 12, True, GREEN)
    T(ws, 2, 1, "입력을 관측값으로 바꿔 PM을 다시 풀어야 해서 값으로 넣었습니다. 설정·계수를 바꿔도 이 시트는 다시 계산되지 않습니다.", 9, color=BROWN)
    att, names = res["attr"], res["attr_names"]
    r = 4
    T(ws, r, 1, "① 예보 입력을 하나씩 관측값으로 바꿨을 때의 ETo RMSE (mm/일, S3)", 11, True, GREEN); r += 1
    hdr = ["구분", "선행일"] + names
    for j, h in enumerate(hdr, 1):
        H(ws, r, j, h, size=9)
    r += 1
    for row in att.itertuples(index=False):
        for j, v in enumerate(row, 1):
            C(ws, r, j, v, F3 if j > 2 else None)
        r += 1
    notes = ["읽는 법: '예보 입력 그대로'보다 많이 줄어드는 입력일수록 그 입력의 예보오차가 ETo 오차에 크게 기여합니다.",
             "'관측 입력 전부 (Rs만 추정)'은 예보가 완벽해도 남는 구조오차(H1 조건, 기압은 고도 추정)입니다.",
             "입력끼리 상관이 있어 줄어든 양의 합이 전체 오차와 같지는 않습니다."]
    for n in notes:
        T(ws, r, 1, n, 9, color="555555"); r += 1
    r += 1
    T(ws, r, 1, "② 지점 편향 보정 탐색 — 월 단위 교차검증 (보정값은 검증 달을 뺀 나머지 달로 추정)", 11, True, GREEN); r += 1
    T(ws, r, 1, f"교차검증 묶음: {', '.join(res['bc_folds'])} (마지막 달 대상일이 10일 미만이면 앞 달에 포함). H2 판정에는 쓰지 않는 탐색 결과입니다.", 9, color="555555"); r += 1
    bc = res["bc"]
    methods = list(dict.fromkeys(bc.method))
    for metric, label, fmt in (("RMSE", "RMSE (mm/일)", F3), ("MBE", "MBE (mm/일)", F3), ("skill_pers", "지속성 대비 개선율", PCT),
                               ("skill_7d", "7일평균 대비 개선율", PCT)):
        H(ws, r, 1, label, fill=LIGHT, white=False); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        for j, m in enumerate(methods, 3):
            H(ws, r, j, m, size=9)
        r += 1
        pv = bc.pivot_table(index=["run_name", "lead_day"], columns="method", values=metric, sort=False)
        for (rn, k), vals in pv.iterrows():
            C(ws, r, 1, rn); C(ws, r, 2, int(k))
            for j, m in enumerate(methods, 3):
                C(ws, r, j, float(vals[m]), fmt)
            r += 1
    H(ws, r, 1, "3일 누적 RMSE (mm/3일)", fill=LIGHT, white=False); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    for j, m in enumerate(methods, 3):
        H(ws, r, j, m, size=9)
    r += 1
    for rn in dict.fromkeys(bc.run_name):
        C(ws, r, 1, rn); C(ws, r, 2, "첫 3일")
        for j, m in enumerate(methods, 3):
            C(ws, r, j, float(bc[(bc.method == m) & (bc.run_name == rn)].RMSE_3d.iloc[0]), F2)
        r += 1
    r += 1
    T(ws, r, 1, "③ 전체 기간 입력 편향 (예보 − 관측) — 보정값의 크기", 11, True, GREEN); r += 1
    for j, h in enumerate(["구분", "선행일", "Tmax (℃)", "Tmin (℃)", "u10 (m/s)"], 1):
        H(ws, r, j, h, size=9)
    r += 1
    for row in res["bc_bias"].itertuples(index=False):
        for j, v in enumerate(row, 1):
            C(ws, r, j, v, F2 if j > 2 else None)
        r += 1
    widths(ws, {"A": 12, "B": 8, **{CL(j): 15 for j in range(3, 11)}})
    return ws


# ── 차트자료 ─────────────────────────────────────────────────────────────
def sheet_chartdata(wb, df, db_rows, obs_rows):
    ws = wb.create_sheet("차트자료")
    H(ws, 1, 1, "대상일"); H(ws, 1, 2, "관측 ETo"); H(ws, 1, 3, "아침 D+1 예보"); H(ws, 1, 4, "저녁 D+1 예보")
    d0, d1 = db_rows; a0, a1 = obs_rows
    db = lambda col: f"일별비교!${col}${d0}:${col}${d1}"
    m = set(df[(df.run_name == "아침") & (df.lead_day == 1)].target) & set(df[(df.run_name == "저녁") & (df.lead_day == 1)].target)
    dates = sorted(m)
    for i, d in enumerate(dates):
        r = 2 + i
        C(ws, r, 1, d.date(), DATE)
        C(ws, r, 2, f"=INDEX(관측!$X${a0}:$X${a1},MATCH(A{r},관측!$A${a0}:$A${a1},0))", F2)
        C(ws, r, 3, f'=SUMIFS({db("Y")},{db("B")},"아침",{db("E")},1,{db("F")},A{r})', F2)
        C(ws, r, 4, f'=SUMIFS({db("Y")},{db("B")},"저녁",{db("E")},1,{db("F")},A{r})', F2)
    widths(ws, {"A": 12, "B": 11, "C": 13, "D": 13})
    return ws, (2, 1 + len(dates))


# ── 요약 ────────────────────────────────────────────────────────────────
def sheet_summary(wb, res, db_rows, c3_rows, groups, chart_rows):
    ws = wb["요약"]
    chk, df = res["check"], res["table"]
    d0, d1 = db_rows
    db = lambda col: f"일별비교!${col}${d0}:${col}${d1}"
    T(ws, 1, 1, f"단기예보 기반 ETo·ETc 예측 검증 (H2) — ASOS {res['stn']} · {res['meta']['settings'].get('작물', '사과')}", 14, True, GREEN)
    T(ws, 2, 1, f"02-Cycle 3단계 · 작성 {dt.date.today():%Y-%m-%d} · cropwater_fcst.py verify · 근거 docs/THEORY.md 9장, 판정 기록 docs/VALIDATION.md", 9, color="555555")
    iss = res["arch"].issues
    n_m = int((df.run_name == "아침").sum() / max(df[df.run_name == "아침"].lead_day.nunique(), 1))
    n_e = int((df.run_name == "저녁").sum() / max(df[df.run_name == "저녁"].lead_day.nunique(), 1))
    info = [
        ("예보 자료", f"기상자료개방포털 과거 단기예보(격자 {', '.join(chk['location'])}) · TMX·TMN·TMP·REH·WSD·PCP · "
                    f"발표 {chk['issues']}회 ({iss[0]:%Y-%m-%d %H}시 ~ {iss[-1]:%Y-%m-%d %H}시), 누락 {len(chk['missing_issues'])}회"),
        ("서비스 발표", f"아침 02시 {n_m}회 → 오늘~D+3 · 저녁 17시 {n_e}회 → 내일~D+4 (다른 발표는 이전 시각 채움에만 사용)"),
        ("대상일", f"{df.target.min():%Y-%m-%d} ~ {df.target.max():%Y-%m-%d}"),
        ("관측 기준", f"ASOS {res['stn']} 일자료 → FAO-56 PM ETo (관측 Rs, 01-Cycle 규칙과 동일, 차이 < 1e-12 mm)"),
        ("Rs 추정", f"식(50) + 강수유무 보정(S3), 계수 {res['coef'].get('source', '')} — Rs계수 시트"),
        ("작물계수", f"{res['meta']['settings'].get('작물', '')} 시나리오 {res['kp']['scenario']}, 생육 시작 {res['kp']['bud']} — 01-Cycle 워크북과 같은 Kc"),
    ]
    r = 4
    for k, v in info:
        C(ws, r, 1, k, left=True, fill=LIGHT, bold=True); C(ws, r, 2, v, left=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=14); r += 1
    r += 1
    # 합격 기준 (수정 가능)
    T(ws, r, 1, "H2 합격 기준 (노란 칸 수정 가능, 2026-09-29 확정값)", 11, True, GREEN); r += 1
    C(ws, r, 1, "지속성 대비 개선율 ≥", left=True, fill=LIGHT); C(ws, r, 2, 0.30, PCT, fill=YEL)
    C(ws, r, 3, "D+1~D+3 모두 (두 발표 각각)", left=True, color="555555"); ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    skill_cell = f"$B${r}"; r += 1
    C(ws, r, 1, "D+1 RMSE ≤ (mm/일)", left=True, fill=LIGHT); C(ws, r, 2, 1.0, F2, fill=YEL)
    C(ws, r, 3, "두 발표 각각", left=True, color="555555"); ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
    rmse_cell = f"$B${r}"; r += 2
    verdict_hdr = r
    for j, h in enumerate(["구분", "D+1~D+3 최소 개선율", "D+1 RMSE", "개선율 기준", "D+1 기준", "판정"], 1):
        H(ws, r, j, h)
    r += 1
    verdict_rows = {}
    for rn in dict.fromkeys(g[0] for g in groups):
        verdict_rows[rn] = r; r += 1
    overall_row = r; r += 2
    # 선행시간별 ETo 지표
    T(ws, r, 1, "선행시간별 예보 ETo 성능 (S3, mm/일) — 모든 값은 일별비교 시트를 참조하는 수식", 11, True, GREEN); r += 1
    met_hdr = ["", "구분", "선행일", "n", "관측 평균", "예보 평균", "MBE", "RMSE", "R²", "지속성 RMSE", "개선율(지속성)",
               "7일평균 RMSE", "개선율(7일평균)", "H2 판정 대상"]
    eto_first = r + 1
    rows_eto = _metric_table(ws, r, met_hdr, groups, db, "Y", "AA", "AD", "AF", "AG")
    r = rows_eto[-1] + 2
    T(ws, r, 1, "선행시간별 예보 ETc (Kc × ETo, mm/일)", 11, True, GREEN); r += 1
    rows_etc = _metric_table(ws, r, met_hdr, groups, db, "AI", "AJ", "AM", "AN", "AO")
    r = rows_etc[-1] + 2
    lead_row = {g: rr for g, rr in zip(groups, rows_eto)}
    for rn, vr in verdict_rows.items():
        ks = [lead_row[(rn, k)] for k in (1, 2, 3) if (rn, k) in lead_row]
        C(ws, vr, 1, rn)
        C(ws, vr, 2, "=MIN(" + ",".join(f"K{x}" for x in ks) + ")", PCT)
        C(ws, vr, 3, f"=H{lead_row[(rn, 1)]}", F2)
        C(ws, vr, 4, f'=IF(B{vr}>={skill_cell},"충족","미달")')
        C(ws, vr, 5, f'=IF(C{vr}<={rmse_cell},"충족","미달")')
        C(ws, vr, 6, f'=IF(AND(D{vr}="충족",E{vr}="충족"),"통과","미달")', bold=True)
    C(ws, overall_row, 1, f"H2 종합 (대상일 {df.target.min():%m/%d}~{df.target.max():%m/%d})", bold=True, fill=LIGHT)
    ws.merge_cells(start_row=overall_row, start_column=1, end_row=overall_row, end_column=5)
    conds = ",".join(f'F{vr}="통과"' for vr in verdict_rows.values())
    C(ws, overall_row, 6, f'=IF(AND({conds}),"통과","미달")', bold=True)
    rng_v = f"D{verdict_hdr + 1}:F{overall_row}"
    ws.conditional_formatting.add(rng_v, FormulaRule(formula=[f'OR(D{verdict_hdr + 1}="통과",D{verdict_hdr + 1}="충족")'],
                                                     fill=PatternFill("solid", fgColor=PASS_FILL)))
    ws.conditional_formatting.add(rng_v, FormulaRule(formula=[f'D{verdict_hdr + 1}="미달"'],
                                                     fill=PatternFill("solid", fgColor=FAIL_FILL)))
    # 3일 누적
    T(ws, r, 1, "3일 누적 (발표별 첫 3개 대상일) — 3일누적 시트 참조", 11, True, GREEN); r += 1
    c0, c1 = c3_rows
    cr = lambda col: f"'3일누적'!${col}${c0}:${col}${c1}"
    hdr3 = ["", "구분", "대상기간", "n", "관측 3일합 평균", "ETo MBE", "ETo RMSE", "상대 RMSE", "지속성 RMSE", "개선율(지속성)",
            "ETc MBE", "ETc RMSE", "ETc 지속성 RMSE", "ETc 개선율"]
    for j, h in enumerate(hdr3, 1):
        if h: H(ws, r, j, h, size=9)
    r += 1
    c3_row = {}
    for rn in dict.fromkeys(g[0] for g in groups):
        c3_row[rn] = r
        m = f"({cr('B')}=$B{r})"
        C(ws, r, 2, rn)
        C(ws, r, 3, "D+0~D+2" if rn == "아침" else "D+1~D+3")
        C(ws, r, 4, f"=COUNTIF({cr('B')},$B{r})", "0")
        C(ws, r, 5, f"=SUMPRODUCT({m}*{cr('G')})/D{r}", F2)
        C(ws, r, 6, f"=SUMPRODUCT({m}*{cr('I')})/D{r}", F2)
        C(ws, r, 7, f"=SQRT(SUMPRODUCT({m}*{cr('I')}^2)/D{r})", F2)
        C(ws, r, 8, f"=G{r}/E{r}", PCT)
        C(ws, r, 9, f"=SQRT(SUMPRODUCT({m}*{cr('J')}^2)/D{r})", F2)
        C(ws, r, 10, f"=1-G{r}/I{r}", PCT)
        C(ws, r, 11, f"=SUMPRODUCT({m}*{cr('N')})/D{r}", F2)
        C(ws, r, 12, f"=SQRT(SUMPRODUCT({m}*{cr('N')}^2)/D{r})", F2)
        C(ws, r, 13, f"=SQRT(SUMPRODUCT({m}*{cr('O')}^2)/D{r})", F2)
        C(ws, r, 14, f"=1-L{r}/M{r}", PCT)
        r += 1
    r += 1
    # 주요 발견 (수식으로 연결된 문장 + 해석)
    T(ws, r, 1, "주요 발견", 11, True, GREEN); r += 1
    va, ve = verdict_rows["아침"], verdict_rows["저녁"]
    di = lambda rn, k, col: f"입력진단!{col}{3 + groups.index((rn, k))}"
    lines = [
        f'="① 판정: D+1~D+3 지속성 대비 개선율 최소 아침 "&TEXT(B{va},"0%")&" · 저녁 "&TEXT(B{ve},"0%")'
        f'&", D+1 RMSE 아침 "&TEXT(C{va},"0.00")&" · 저녁 "&TEXT(C{ve},"0.00")&" mm/일 → "&F{overall_row}',
        f'="② 편향: 예보 ETo가 관측보다 평균 "&TEXT(AVERAGE(G{rows_eto[0]}:G{rows_eto[-1]}),"+0.00;-0.00")'
        f'&" mm/일 (−면 과소). 3일 누적 편향은 아침 "&TEXT(F{c3_row["아침"]},"+0.0;-0.0")&" mm, 저녁 "&TEXT(F{c3_row["저녁"]},"+0.0;-0.0")&" mm"',
        f'="③ 입력 편향(아침 D+1, 예보−관측): 최고기온 "&TEXT({di("아침", 1, "D")},"+0.0;-0.0")&"℃, 최저기온 "'
        f'&TEXT({di("아침", 1, "F")},"+0.0;-0.0")&"℃, 풍속 "&TEXT({di("아침", 1, "L")},"+0.00;-0.00")&" m/s, ea "'
        f'&TEXT({di("아침", 1, "J")},"+0.00;-0.00")&" kPa (입력진단 시트)"',
    ]
    for txt in res.get("findings", []):
        lines.append(txt)
    for ln in lines:
        T(ws, r, 1, ln, 10); ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14); r += 1
    # 차트 (주요 발견 아래)
    r += 1
    bar = BarChart(); bar.type = "col"; bar.grouping = "clustered"
    bar.title = "선행시간별 RMSE (ETo, mm/일)"; bar.y_axis.title = "RMSE (mm/일)"
    cats = Reference(ws, min_col=1, min_row=rows_eto[0], max_row=rows_eto[-1])
    for col, name, color in ((8, "예보(S3)", GREEN), (10, "지속성", BROWN), (12, "7일평균", "9AA39A")):
        bar.add_data(Reference(ws, min_col=col, min_row=rows_eto[0], max_row=rows_eto[-1]), titles_from_data=False)
        srs = bar.series[-1]; srs.tx = SeriesLabel(v=name)
        srs.graphicalProperties.solidFill = color; srs.graphicalProperties.line.solidFill = color
    bar.set_categories(cats); bar.height = 8; bar.width = 17; bar.legend.position = "b"
    bar.y_axis.majorGridlines = None
    ws.add_chart(bar, f"A{r}")
    ln = LineChart(); ln.title = "관측 ETo와 D+1 예보 ETo (mm/일)"; ln.y_axis.title = "ETo (mm/일)"
    cw = wb["차트자료"]; q0, q1 = chart_rows
    for col, name, color, w in ((2, "관측", "1A1A1A", 22000), (3, "아침 D+1 예보", BLUE, 15000), (4, "저녁 D+1 예보", BROWN, 15000)):
        ln.add_data(Reference(cw, min_col=col, min_row=q0, max_row=q1), titles_from_data=False)
        srs = ln.series[-1]; srs.tx = SeriesLabel(v=name); srs.smooth = False
        srs.graphicalProperties.line.solidFill = color; srs.graphicalProperties.line.width = w
    ln.set_categories(Reference(cw, min_col=1, min_row=q0, max_row=q1))
    ln.x_axis.number_format = "mm-dd"; ln.height = 8; ln.width = 24; ln.legend.position = "b"
    ws.add_chart(ln, f"G{r}")
    widths(ws, {"A": 20, "B": 16, "C": 13, "D": 12, "E": 12, "F": 12, "G": 10, "H": 10, "I": 10, "J": 12, "K": 13,
                "L": 12, "M": 14, "N": 12})
    return ws


def _metric_table(ws, r, hdr, groups, db, fcol, ocol, ecol, pcol, mcol):
    for j, h in enumerate(hdr, 1):
        if h: H(ws, r, j, h, size=9)
    r += 1
    rows = []
    for rn, k in groups:
        crit = f'{db("B")},$B{r},{db("E")},$C{r}'
        m = f"({db('B')}=$B{r})*({db('E')}=$C{r})"
        f, o = db(fcol), db(ocol)
        C(ws, r, 1, f"{rn} D+{k}", fill=LIGHT)
        C(ws, r, 2, rn); C(ws, r, 3, k)
        C(ws, r, 4, f"=COUNTIFS({crit})", "0")
        C(ws, r, 5, f"=AVERAGEIFS({o},{crit})", F2)
        C(ws, r, 6, f"=AVERAGEIFS({f},{crit})", F2)
        C(ws, r, 7, f"=AVERAGEIFS({db(ecol)},{crit})", F2)
        C(ws, r, 8, f"=SQRT(SUMPRODUCT({m}*{db(ecol)}^2)/D{r})", F2)
        C(ws, r, 9, (f"=(D{r}*SUMPRODUCT({m}*{f}*{o})-SUMPRODUCT({m}*{f})*SUMPRODUCT({m}*{o}))^2"
                     f"/((D{r}*SUMPRODUCT({m}*{f}^2)-SUMPRODUCT({m}*{f})^2)*(D{r}*SUMPRODUCT({m}*{o}^2)-SUMPRODUCT({m}*{o})^2))"), F2)
        C(ws, r, 10, f"=SQRT(SUMPRODUCT({m}*{db(pcol)}^2)/D{r})", F2)
        C(ws, r, 11, f"=1-H{r}/J{r}", PCT)
        C(ws, r, 12, f"=SQRT(SUMPRODUCT({m}*{db(mcol)}^2)/D{r})", F2)
        C(ws, r, 13, f"=1-H{r}/L{r}", PCT)
        C(ws, r, 14, "○" if k in (1, 2, 3) else "참고")
        rows.append(r); r += 1
    return rows


# ── 방법 ────────────────────────────────────────────────────────────────
def sheet_method(wb, res):
    ws = wb.create_sheet("방법")
    T(ws, 1, 1, "방법과 정의 (근거: docs/THEORY.md 9장, 설계: docs/ARCHITECTURE.md 7장)", 12, True, GREEN)
    items = [
        ("자료", "예보", "기상자료개방포털 '단기예보(격자)' 과거자료 CSV. 요소 TMX·TMN·TMP·REH·WSD·PCP. 발표일시 = CSV의 UTC 일·시 + 9시간."),
        ("", "관측", "ASOS 일자료(01-Cycle cropwater_station.py 출력 워크북 원데이터 시트). 관측 ETo는 01-Cycle 계산과정 시트와 같은 식."),
        ("서비스", "발표", "아침 02시 발표 → 오늘(D+0)~D+3, 저녁 17시 발표 → 내일(D+1)~D+4. 검증에는 이 두 발표만 사용(다른 발표를 섞으면 성능이 부풀려짐)."),
        ("일 입력", "Tmax·Tmin", "해당 발표의 TMX(낮최고 09:01~18:00)·TMN(아침최저 03:01~09:00). ASOS 일최고·최저(0~24시)와 정의가 다름."),
        ("", "ea", "시간별 e°(TMP) × REH/100 의 일평균 (ASOS 평균증기압과 같은 개념)."),
        ("", "u2", "WSD 일평균(10 m) × 4.87/ln(67.8×10−5.42) = 0.748 × u10  [식47]"),
        ("", "강수", "PCP 일합계. 기준(설정, 1 mm) 이상이면 강수유무 1."),
        ("", "지나간 시각 채움", "포털 과거자료는 발표 6시간 뒤부터 들어 있음(02시 발표 → 08시부터). 발표 시점에 이미 지난 시각은 그 이전의 가장 최근 발표 값으로 채움 → 아침 D+0의 00~07시 = 전날 17·20·23시 발표. (API는 발표 1시간 뒤부터 제공되므로 운영에서는 00~02시만 채움)"),
        ("", "마지막 날", "아침 D+3·저녁 D+4는 00시(1시간) + 03~21시(3시간 간격) 8개 시각. 풍속·강수는 코드값 → WSD 1: 같은 발표 직전 정량일 평균(최대 3.9), 2: 6.5, 3: 11 m/s / PCP 1: 1.5, 2: 9, 3: 20 mm/h × 3시간."),
        ("예보 ETo", "PM", "FAO-56 식(6), G = 0. 기압은 식(7) 고도 추정. Ra·Rso는 지점 위도·고도로 계산."),
        ("", "Rs", "S3: Rs/Ra = a + b√(Tmax−Tmin) + c·강수유무, [0.05Ra, Rso]로 제한. 계수는 검증 연도와 겹치지 않는 해(2025)의 관측으로 결정. 비교용 S1: kRs 0.16."),
        ("ETc", "Kc", "01-Cycle 워크북과 같은 Kc(시나리오 표값 + 식62·65 현지기상 보정, 식66 선형 보간). 예보·관측·기준선 모두 대상일 Kc를 곱함."),
        ("기준선", "지속성", "발표일 전날(가장 최근의 완결된 관측일) 관측 ETo를 모든 선행일에 사용. 02시에는 전날 자료가 아직 공개 전일 수 있어 실제 운영보다 기준선에 유리한 가정(판정에 보수적)."),
        ("", "7일평균", "발표일 전 7일(D−7~D−1) 관측 ETo 평균."),
        ("지표", "MBE·RMSE·R²", "오차 = 예보 − 관측. MBE +면 과대추정. R²는 예보·관측 상관계수의 제곱."),
        ("", "개선율", "1 − RMSE_예보 / RMSE_기준선. 0보다 크면 기준선보다 좋음."),
        ("", "3일 누적", "발표별 첫 3개 대상일 합계(아침 D+0~D+2, 저녁 D+1~D+3). FAO-56은 추정 Rs 기반 ETo를 여러 날 합계로 쓰도록 권고."),
        ("판정", "H2", "두 발표 각각 D+1~D+3 개선율 ≥ 기준, D+1 RMSE ≤ 상한 (요약 시트 노란 칸). 아침 D+0·저녁 D+4는 참고."),
        ("한계", "격자", "이 자료의 격자는 73_135(신북읍)로 ASOS 101 격자(73,134, 신사우동 등)와 다름 → 대표성 오차가 다를 수 있음."),
        ("", "기간", "2026년 4~6월 대상일(봄~초여름)만 포함. 장마·한여름 성능은 7~9월 자료로 확인."),
        ("", "강수", "예보 강수는 강수유무(Rs 보정)에만 쓰였고, 유효강수·물수지 영향은 다음 단계(G4)에서 검증."),
    ]
    H(ws, 3, 1, "구분"); H(ws, 3, 2, "항목"); H(ws, 3, 3, "내용")
    r = 4
    for a, b, c in items:
        C(ws, r, 1, a, fill=(LIGHT if a else None), bold=bool(a)); C(ws, r, 2, b, left=True)
        C(ws, r, 3, c, left=True, wrap=True); ws.row_dimensions[r].height = 30 if len(c) > 70 else 18
        r += 1
    widths(ws, {"A": 10, "B": 16, "C": 120})
    return ws


# ── 빌드 ────────────────────────────────────────────────────────────────
def build_verify_workbook(res, out):
    df = res["table"]
    groups = [(rn, int(k)) for rn in ("아침", "저녁") for k in sorted(df[df.run_name == rn].lead_day.unique())]
    first_run = df.run.min().normalize()
    obs = res["obs"]
    obs = obs[(obs.date >= first_run - pd.Timedelta(days=7))].reset_index(drop=True)
    wb = Workbook()
    wb.active.title = "요약"
    sheet_settings(wb, res)
    # Rs계수는 관측 행 범위를 알아야 하므로, 관측 시트 행 범위를 먼저 계산
    obs_rows = (2, 1 + len(obs))
    sheet_rscoef(wb, res, obs_rows)
    _, obs_rows2 = sheet_obs(wb, obs)
    assert obs_rows2 == obs_rows
    _, db_rows, run_ids = sheet_daily(wb, df, obs_rows)
    _, c3_rows = sheet_cum3(wb, df, db_rows, run_ids)
    sheet_diag(wb, db_rows, groups)
    sheet_attr(wb, res)
    _, chart_rows = sheet_chartdata(wb, df, db_rows, obs_rows)
    sheet_summary(wb, res, db_rows, c3_rows, groups, chart_rows)
    sheet_method(wb, res)
    order = ["요약", "일별비교", "3일누적", "입력진단", "오차분해", "Rs계수", "관측", "설정", "방법", "차트자료"]
    wb._sheets = [wb[n] for n in order]
    for n in ("요약", "Rs계수", "설정", "방법", "오차분해", "입력진단", "3일누적"):
        wb[n].sheet_view.showGridLines = False
    wb["요약"].sheet_properties.tabColor = GREEN
    for n in ("요약", "Rs계수", "입력진단", "오차분해", "방법", "설정"):
        w = wb[n]; w.page_setup.orientation = "landscape"; w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
    return safe_save(wb, out)
