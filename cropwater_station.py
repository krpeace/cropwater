#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cropwater_station.py — 단일 지점 기준증발산(ETo)·작물증발산(ETc)·일별 물수지 분석, Excel 출력
  (구 파일명: eto_singlestation.py → eto_chuncheon_excel_v3.py)
   다룰 수 있어 이름을 변경함. 계산 로직 변경 없음)

  시트: 설정 / 원데이터(ASOS) / 계산과정(FAO-56, 라이브 수식) / 결과요약 / 계산근거

[함께 필요한 파일]
  fao56_core.py — 같은 폴더에 있어야 함(공통 함수 모듈: num·fetch_asos·fetch_station_table·load_apikeys)

특징
  1) 계산과정 시트는 '값'이 아니라 '엑셀 수식'으로 작성 → 셀 클릭 시 계산식이 보이고
     설정값(위도·고도 등)을 바꾸면 자동 재계산된다.
  2) 설정 시트 C열에 각 기본값의 근거·출처를 표기.
  3) 지점번호를 넣으면 위도·고도·풍속계높이를 기상청 지점일람표(stn_inf.php)로 자동조회.
  4) 작물계수 Kc(설정!B9)로 ETc(=Kc×ETo)까지 산출.
  5) 출력 파일명에 지점코드 표기: eto(101)_20260801_20260831.xlsx

[실행]
  pip install requests openpyxl
  python cropwater_station.py           ← 실행하면 지점번호·기간을 물어봄
  (인증키는 apikey.txt 파일에 아래 2줄로 저장해두면 매번 안 물어봄)
    DATA_GO_KR=<공공데이터포털 키, 기상데이터용>
    KMA_HUB=<기상청 API허브 키, 지점정보 자동조회용>
"""
import argparse, datetime as dt, os, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fao56_core import (num, fetch_asos as _fetch_asos_raw, fetch_station_table, load_apikeys, safe_save,
                         load_station_backup, save_station_backup,
                         KC_SCENARIOS, KCB_SCENARIOS, kc_climate_adjust,
                         stage_of_date, kc_of_date,
                         load_crop_library, crops_sorted,
                         wind_2m)   # ← 공통 모듈(같은 폴더 필요)

FONT="맑은 고딕"; GREEN="2E6A4C"; BLUE="2B6E86"; BROWN="A8681B"; LIGHT="EFF1EC"; YEL="FBEED2"; WARN="F4C7B8"
thin=Side(style="thin",color="D0D3CC"); BORDER=Border(thin,thin,thin,thin)
def H(ws,r,c,v,fill=LIGHT,white=False,size=10):
    x=ws.cell(r,c,v); x.font=Font(name=FONT,bold=True,size=size,color=("FFFFFF" if white else "1A1A1A"))
    x.fill=PatternFill("solid",fgColor=fill); x.alignment=Alignment("center","center",wrap_text=True); x.border=BORDER; return x
def C(ws,r,c,v,fmt=None,left=False,fill=None):
    x=ws.cell(r,c,v); x.font=Font(name=FONT,size=10); x.border=BORDER
    x.alignment=Alignment("left" if left else "center","center")
    if fmt: x.number_format=fmt
    if fill: x.fill=PatternFill("solid",fgColor=fill)
    return x

def fetch_asos(key,stn,start,end):
    """조회(HTTP 요청은 fao56_core.fetch_asos가 담당) + 이 스크립트가 쓰는 형태로 파싱."""
    rows=_fetch_asos_raw(key,stn,start,end)
    out=[]
    for x in rows:
        out.append(dict(
            tm=dt.datetime.strptime(x.get("tm"),"%Y-%m-%d").date(),
            maxTa=num(x.get("maxTa")),minTa=num(x.get("minTa")),avgTa=num(x.get("avgTa")),
            avgRhm=num(x.get("avgRhm")),minRhm=num(x.get("minRhm")),avgWs=num(x.get("avgWs")),
            avgPv=num(x.get("avgPv")),avgTd=num(x.get("avgTd")),avgPa=num(x.get("avgPa")),
            sumGsr=num(x.get("sumGsr")),sumSsHr=num(x.get("sumSsHr")),sumLrgEv=num(x.get("sumLrgEv")),
            sumRn=num(x.get("sumRn"))))
    return out

def build_workbook(rows, p, out):
    n=len(rows); last=n+1; S="원데이터"
    wb=Workbook()

    # ===== 설정 =====
    ws=wb.active; ws.title="설정"
    H(ws,1,1,"항목",GREEN,True); H(ws,1,2,"값 (노란칸은 수정 가능)",GREEN,True); H(ws,1,3,"기본값 근거·출처",GREEN,True)
    ws.merge_cells("E1:F1"); H(ws,1,5,"※ 위도·고도·풍속계높이·fetch를 바꾸면 계산과정이 자동 재계산됩니다.",YEL)
    # ── 메타정보(위도·고도·풍속계높이) 출처별 근거 문구 ──
    SRC_LABEL = {
        "API":     "기상청 API 허브(stn_inf.php) 실시간 자동조회 실측값",
        "BACKUP":  "기상청 API 허브 연결 실패로 관측지점정보 백업 파일의 메타정보 적용",
        "DEFAULT": "지점정보 백업 파일에도 없어 최종 안전값(춘천 근사) 사용 — 실제와 다를 수 있으니 --lat/--elev/--anem로 직접 지정 권장",
        "MANUAL":  "사용자가 --lat/--elev/--anem 인자로 직접 지정",
    }
    src = SRC_LABEL.get(p.get("meta_source"), SRC_LABEL["DEFAULT"])

    rowsset=[
        ("위도 (°)",p["lat"],"0.0000",
         f"천문일사(Ra) 산정에 사용. [출처] {src}"),
        ("고도 (m)",p["elev"],"0.0",
         f"기압 추정식·청천일사(Rso) 계산의 기준값(단, PM에 미치는 영향은 작음). [출처] {src}"),
        ("풍속계 높이 (m)",p["anem"],"0.0",
         f"3개 메타값 중 PM에 가장 민감한 항목(지점마다 다름: 춘천 10·태백 16·수원 18.7m 등). [출처] {src}"),
        ("증발접시 상풍거리 fetch (m)",p["fetch"],"0",
         "FAO-56 Table 8 기본 가정값. 증발접시 주변이 녹지(잔디 등)이고 상풍(바람이 불어오는 방향) 노출거리 100m 기준 — 관측소 주변 환경 실사 전까지의 표준 근사치"),
        ("지점번호 (참고용)",p["stn"],None,
         f"실행 시 입력한 ASOS 지점번호. 관측지점 메타정보 출처: {src}"),
        ("조회 시작일 (참고용)",p["start"],None,
         "실행 시 입력한 조회기간 시작일. 계산식에는 사용되지 않고 원본 확인용으로만 표시"),
        ("조회 종료일 (참고용)",p["end"],None,
         "실행 시 입력한 조회기간 종료일. 계산식에는 사용되지 않고 원본 확인용으로만 표시"),
    ]
    for i,(k,v,fmt,note) in enumerate(rowsset,2):
        H(ws,i,1,k,LIGHT)
        editable=i<=5
        C(ws,i,2,v,fmt,fill=(YEL if editable else None))
        nc=ws.cell(i,3,note); nc.font=Font(name=FONT,size=9,color="555555"); nc.border=BORDER
        nc.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
    ws.column_dimensions["A"].width=24; ws.column_dimensions["B"].width=16; ws.column_dimensions["C"].width=70
    for col in "EF": ws.column_dimensions[col].width=30
    for i in range(2,len(rowsset)+2): ws.row_dimensions[i].height=32

    # ═══════ 5+6+7. 작물계수 Kc 시스템 (시나리오 · 생육단계 · 현지기상 보정) ═══════
    # 여기서 설정한 값들이 계산과정의 일별 Kc·ETc를 자동으로 결정한다.
    # 셀 좌표를 변수로 관리(다른 시트에서 참조).
    KR = {}          # 이름 → 셀 좌표('$B$14' 등)
    def _sec(rr, title):
        H(ws,rr,1,title,GREEN,white=True,size=10)
        ws.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=3)
    def _row(rr, label, value, fmt, note, editable=False, name=None, warn=False):
        H(ws,rr,1,label,LIGHT)
        C(ws,rr,2,value,fmt,fill=(WARN if warn else (YEL if editable else None)))
        nc=ws.cell(rr,3,note); nc.font=Font(name=FONT,size=9,color=("A32D2D" if warn else "555555"),bold=warn); nc.border=BORDER
        nc.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
        ws.row_dimensions[rr].height=(44 if warn else 32)
        if name: KR[name]=f"설정!$B${rr}"
    r=len(rowsset)+2

    # ── [5] Kc 시나리오 (작물 라이브러리 기반) ──
    crop = p.get("crop") or {}   # main()에서 라이브러리에서 로드해 넘겨줌
    scenarios = crop.get("scenarios") or [(0.50, 1.20, 0.95)]   # 폴백: 사과 시나리오3
    has_scn   = crop.get("has_scenarios", False) and len(scenarios) > 1
    crop_label= f"{crop.get('name_ko','(사과)')} — {crop.get('name_en','Apple')}" if crop else "(작물 정보 없음)"
    fao_ref   = "FAO-56 Table 12"
    _sec(r,f"[5] Kc 시나리오 — {fao_ref} ({crop.get('category','')})"); r+=1
    _row(r,"작물",crop_label,"@",
         f"crop_id={crop.get('crop_id','')}. 시나리오 {len(scenarios)}개 사용 가능. "
         f"작물을 바꾸려면 실행 시 --crop 인자를 사용(예: --crop cabbage). "
         f"라이브러리는 crops_library.csv, 사용자 수정은 crops_overrides.csv"); r+=1
    if has_scn:
        scn_notes = ("1=청경·서리 (0.45/0.95/0.70) | 2=청경·무서리 (0.60/0.95/0.75) | "
                     "3=초생재배·서리 (0.50/1.20/0.95, 한국 과수원 일반) | 4=초생재배·무서리 (0.80/1.20/0.85). "
                     "이 값을 바꾸면 아래 표값 3개와 최종 Kc가 자동 재계산됨.")
        default_scn = p.get("kc_scn", 3 if len(scenarios) >= 3 else 1)
    else:
        scn_notes = f"이 작물은 단일 시나리오만 있습니다({scenarios[0][0]}/{scenarios[0][1]}/{scenarios[0][2]}). 시나리오 번호는 1로 고정."
        default_scn = 1
    _row(r,"시나리오 번호 (1~4)",default_scn,"0",scn_notes,editable=has_scn,name="SCN"); r+=1
    # 동적 CHOOSE: 작물의 시나리오 개수만큼 값을 나열 (없는 시나리오는 마지막 값으로 채워 CHOOSE 오류 방지)
    def _choose_kc(idx):
        vals = [s[idx] for s in scenarios]
        while len(vals) < 4: vals.append(vals[-1])
        return ",".join(f"{v:.3f}" for v in vals)
    _row(r,"표값 Kc_ini",f"=CHOOSE({KR['SCN']},{_choose_kc(0)})","0.000",
         "선택된 시나리오의 초기 Kc 표값 (현지기상 보정 없음, 멀칭 보정만 적용)",name="TAB_INI"); r+=1
    _row(r,"표값 Kc_mid",f"=CHOOSE({KR['SCN']},{_choose_kc(1)})","0.000",
         "선택된 시나리오의 생육중기 Kc 표값 (현지기상 보정 대상)",name="TAB_MID"); r+=1
    _row(r,"표값 Kc_end",f"=CHOOSE({KR['SCN']},{_choose_kc(2)})","0.000",
         "선택된 시나리오의 생육후기 Kc 표값 (0.45 초과일 때만 현지기상 보정)",name="TAB_END"); r+=1

    # ── [6] 생육단계 구분 (FAO-56 Table 11) ──
    _sec(r,"[6] 생육단계 구분 — FAO-56 Table 11 (Deciduous Orchard, High Latitudes)"); r+=1
    import datetime as _dt
    _bud_warn = p.get("is_short_cycle", False) and not p.get("bud_date_manual", False)
    if _bud_warn:
        _bud_note = (f"⚠ 확인 필요: 이 작물은 생육기간 총 {p.get('L_total','?')}일의 단기/1년생 작물로 추정됩니다. "
                     f"기본값(2026-04-01)은 검증되지 않은 임시값입니다 — 실제 정식·파종일로 반드시 수정하세요 "
                     f"(또는 다음 실행 시 --bud-date YYYYMMDD 지정). 틀리면 생육단계·Kc·ETc 전체가 어긋납니다.")
    else:
        _bud_note = ("사과 눈이 움직이기 시작하는 날 등, 작물의 생육 시작 시점(정식·파종·발아일). "
                     "지역·기후·품종에 따라 다름. 이 날짜부터 스테이지 카운트 시작.")
    _row(r,"생육 시작일 (발아기/정식일)",p.get("bud_date",_dt.date(2026,4,1)),"yyyy-mm-dd",
         _bud_note,editable=True,name="BUD",warn=_bud_warn); r+=1
    _row(r,"L_ini (초기, 일)",p.get("L_ini") or crop.get("L_ini") or 20,"0",
         "발아 후 Kc가 Kc_ini에 머무는 기간. 작물 라이브러리 기본값(FAO-56 Table 11). 지역·품종 편차 클 때 실측 조정",editable=True,name="LINI"); r+=1
    _row(r,"L_dev (발육, 일)",p.get("L_dev") or crop.get("L_dev") or 70,"0",
         "Kc가 Kc_ini→Kc_mid로 선형 증가하는 기간",editable=True,name="LDEV"); r+=1
    _row(r,"L_mid (중기, 일)",p.get("L_mid") or crop.get("L_mid") or 90,"0",
         "Kc가 Kc_mid에 머무는 기간(대개 생식 성장기~성숙 직전)",editable=True,name="LMID"); r+=1
    _row(r,"L_late (후기, 일)",p.get("L_late") or crop.get("L_late") or 30,"0",
         "Kc가 Kc_mid→Kc_end로 선형 감소하는 기간(성숙~수확·낙엽)",editable=True,name="LLATE"); r+=1
    _row(r,"총 생육기간 (일)",f"={KR['LINI']}+{KR['LDEV']}+{KR['LMID']}+{KR['LLATE']}","0",
         "L_ini + L_dev + L_mid + L_late",name="LTOT"); r+=1
    _row(r,"생육중기 시작일",f"={KR['BUD']}+{KR['LINI']}+{KR['LDEV']}","yyyy-mm-dd",
         "발아일 + L_ini + L_dev (현지기상 보정용 평균 산정 구간의 시작)",name="MID_S"); r+=1
    _row(r,"생육중기 종료일",f"={KR['BUD']}+{KR['LINI']}+{KR['LDEV']}+{KR['LMID']}-1","yyyy-mm-dd",
         "생육중기 종료일(포함)",name="MID_E"); r+=1
    _row(r,"생육후기 종료일 (생육 종료)",f"={KR['BUD']}+{KR['LTOT']}-1","yyyy-mm-dd",
         "생육 종료일. 이 이후는 Kc=0 처리",name="END_E"); r+=1

    # ── [7] Kc 현지기상 보정 (FAO-56 식 62/65) ──
    _sec(r,"[7] Kc 현지기상 보정 — FAO-56 식(62) p.121, 식(65) p.125"); r+=1
    _row(r,"생육중기 초목 수고 h (m)",p.get("h_m") or crop.get("h_m") or 3.2,"0.0",
         "생육중기 캐노피 평균 높이. 작물 라이브러리 기본값. 현지기상 보정식의 (h/3)^0.3 항에 사용",editable=True,name="H"); r+=1
    _u2m=p.get("u2_mid"); _rhm=p.get("rh_mid")
    _u2mid_note=(f"조회기간 중 생육중기 구간의 실측 평균값(자동집계). 강원도원 워크북과 정합화 기준."
                 if _u2m is not None else
                 "생육중기 기간의 지상 2m 풍속 평균. 실측 자동집계 실패(중기 구간 <10일) → 기본값 1.5 사용. 자동제한 [1,6]")
    _rhmid_note=(f"조회기간 중 생육중기 구간의 실측 평균값(자동집계, 최소상대습도 minRhm 기준)."
                 if _rhm is not None else
                 "생육중기 기간의 일 최저상대습도 평균. 실측 자동집계 실패 → 기본값 55 사용. 자동제한 [20,80]")
    _row(r,"중기 평균 u2 (m/s)",_u2m if _u2m is not None else 1.5,"0.00",
         _u2mid_note,editable=True,name="U2_MID"); r+=1
    _row(r,"중기 평균 RHmin (%)",_rhm if _rhm is not None else 55,"0",
         _rhmid_note,editable=True,name="RH_MID"); r+=1
    _u2e=p.get("u2_end"); _rhe=p.get("rh_end")
    _row(r,"후기 평균 u2 (m/s)",_u2e if _u2e is not None else 1.5,"0.00",
         ("조회기간 중 생육후기 구간의 실측 평균값(자동집계)." if _u2e is not None
          else "생육후기 기간 실측 부족 → 기본값 1.5 사용"),editable=True,name="U2_END"); r+=1
    _row(r,"후기 평균 RHmin (%)",_rhe if _rhe is not None else 55,"0",
         ("조회기간 중 생육후기 구간의 실측 평균값(자동집계)." if _rhe is not None
          else "생육후기 기간 실측 부족 → 기본값 55 사용"),editable=True,name="RH_END"); r+=1
    _row(r,"멀칭 Kc 보정계수",p.get("mulch",1.0),"0.00",
         "전면 비닐멀칭 시 Kc를 10~30% 낮춤(0.70~0.90). 없으면 1.00",editable=True,name="MULCH"); r+=1
    # ── 최종 Kc 3값 (수식으로 자동 계산) ──
    _sec(r,"[5+7] 최종 적용 Kc (자동 계산)"); r+=1
    # Kc_ini: 현지기상 보정 없음, 멀칭 보정만
    _row(r,"★ 적용 Kc_ini",f"={KR['TAB_INI']}*{KR['MULCH']}","0.000",
         "FAO-56 p.121: Kc_ini는 현지기상 보정 대상 아님. 멀칭 보정만 반영",name="APP_INI"); r+=1
    # Kc_mid: 식(62) 현지기상 보정 + 멀칭. u2·RHmin은 MEDIAN으로 [1,6]·[20,80] 자동 제한
    _row(r,"★ 적용 Kc_mid (현지기상 보정 후)",
         f"=({KR['TAB_MID']}+(0.04*(MEDIAN(1,{KR['U2_MID']},6)-2)-0.004*(MEDIAN(20,{KR['RH_MID']},80)-45))*({KR['H']}/3)^0.3)*{KR['MULCH']}",
         "0.000","FAO-56 식(62): 현지기상(u2·RHmin) 보정. Kc_mid + [0.04(u2−2)−0.004(RHmin−45)](h/3)^0.3, 멀칭 보정 포함",name="APP_MID"); r+=1
    # Kc_end: 표값>0.45일 때만 보정
    _row(r,"★ 적용 Kc_end (현지기상 보정 후)",
         f"=IF({KR['TAB_END']}>0.45,({KR['TAB_END']}+(0.04*(MEDIAN(1,{KR['U2_END']},6)-2)-0.004*(MEDIAN(20,{KR['RH_END']},80)-45))*({KR['H']}/3)^0.3),{KR['TAB_END']})*{KR['MULCH']}",
         "0.000","FAO-56 식(65): Kc_end>0.45인 경우만 현지기상 보정, 그 외는 표값 유지. 멀칭 보정 곱함",name="APP_END"); r+=1

    # ═══════ [8] 토양 파라미터 (일별 물수지 계산용, FAO-56 Ch.8) ═══════
    _sec(r,"[8] 토양·근권 파라미터 (물수지 계산용)"); r+=1
    _crop_zr = p["crop"].get("zr") or 1.0
    _crop_p  = p["crop"].get("p")  or 0.50
    _row(r,"포장용수량 θFC (m³/m³)", 0.22, "0.00",
         "FAO-56 Table 19, Loam(양토) 대표값. 토성에 따라 모래 0.10, 점토 0.36 범위",
         editable=True, name="FC"); r+=1
    _row(r,"위조점 θWP (m³/m³)", 0.10, "0.00",
         "FAO-56 Table 19, Loam(양토) 대표값. 모래 0.04, 점토 0.20 범위",
         editable=True, name="WP"); r+=1
    _row(r,"근권심도 Zr (m)", _crop_zr, "0.00",
         f"FAO-56 Table 22. {p['crop'].get('name_ko','작물')} 관수 스케줄링용 값. 천수답이면 큰 값 사용",
         editable=True, name="ZR"); r+=1
    _row(r,"토양수분고갈계수 p", _crop_p, "0.00",
         f"FAO-56 Table 22. ETc≈5mm/d 기준. 고온건조 시 10-25% 감소, 저ETc 시 20% 증가",
         editable=True, name="PDEP"); r+=1
    _row(r,"★ TAW (mm)", f"=1000*({KR['FC']}-{KR['WP']})*{KR['ZR']}", "0.0",
         "총유효수분 = 1000·(θFC−θWP)·Zr. FAO-56 식(82)", name="TAW"); r+=1
    _row(r,"★ RAW (mm)", f"={KR['PDEP']}*{KR['TAW']}", "0.0",
         "쉽게이용가능수분 = p×TAW. Dr이 이 값에 도달하면 관수 필요. FAO-56 식(83)", name="RAW"); r+=1
    _row(r,"초기 고갈량 Dr₀ (mm)", 0, "0.0",
         "전일 충분한 강우/관수 후 포장용수량 출발 가정(Dr=0). 건조 조건이면 양수값 입력",
         editable=True, name="DR0"); r+=1
    _row(r,"관수효율 Ea", 0.95, "0.00",
         "점적관수 기준 0.95. 스프링클러 0.75~0.85. 총관수량 Ig = In/Ea",
         editable=True, name="EA"); r+=1

    for rr in range(len(rowsset)+2, r): ws.row_dimensions[rr].height=32

    # ===== 원데이터 =====
    ws=wb.create_sheet(S)
    cols=[("일자","tm",None),("최고기온(℃)","maxTa","0.0"),("최저기온(℃)","minTa","0.0"),
          ("평균기온(℃)","avgTa","0.0"),("평균습도(%)","avgRhm","0"),("최소습도(%)","minRhm","0"),
          ("평균풍속(m/s)","avgWs","0.0"),("평균증기압(hPa)","avgPv","0.0"),("평균이슬점(℃)","avgTd","0.0"),
          ("평균현지기압(hPa)","avgPa","0.0"),("합계일사(MJ/m²)","sumGsr","0.00"),
          ("합계일조(hr)","sumSsHr","0.0"),("대형증발량(mm)","sumLrgEv","0.0"),
          ("일강수량(mm)","sumRn","0.0")]   # 참조용(ETo 계산에는 미사용). 맨끝 열이라 계산과정 참조 불변
    for c,(h,_,_) in enumerate(cols,1): H(ws,1,c,h,GREEN,True)
    for i,r in enumerate(rows,2):
        for c,(_,k,fmt) in enumerate(cols,1):
            v=r.get(k); cell=C(ws,i,c,v,fmt)
            if k=="tm": cell.number_format="yyyy-mm-dd"
    ws.freeze_panes="B2"; ws.column_dimensions["A"].width=12
    for c in range(2,len(cols)+1): ws.column_dimensions[get_column_letter(c)].width=11

    # ===== 계산과정 (라이브 수식) =====
    ws=wb.create_sheet("계산과정")
    heads=["일자","연중일 DOY","평균기온 T(℃)","e°(Tmax)","e°(Tmin)","포화증기압 es(kPa)",
           "실제증기압 ea(kPa)","기울기 Δ(kPa/℃)","기압 P(kPa)","건습계 γ(kPa/℃)","풍속 u₂(m/s)",
           "dr","적위 δ(rad)","일몰각 ωs(rad)","천문일사 Ra(MJ/m²)","청천일사 Rso(MJ/m²)","일사 Rs(MJ/m²)",
           "순단파 Rns(MJ/m²)","순장파 Rnl(MJ/m²)","순복사 Rn(MJ/m²)","★ ETo_PM(mm)",
           "생육단계","일별 Kc",                                              # ← 신규(V, W)
           "증발접시계수 Kp","대형증발량(mm)","★ ETo_pan(mm)","★ ETc_PM(mm)","★ ETc_pan(mm)"]
    fmts=[None,"0","0.0","0.000","0.000","0.000","0.000","0.0000","0.0","0.0000","0.00","0.000",
          "0.0000","0.0000","0.00","0.00","0.00","0.00","0.00","0.00","0.00",
          "@","0.000",                                                        # 생육단계=텍스트, Kc=0.000
          "0.000","0.0","0.00","0.00","0.00"]
    for c,h in enumerate(heads,1):
        fill=GREEN if h.startswith("★") else (BROWN if h in ("생육단계","일별 Kc") else (BLUE if c<=20 else BLUE))
        H(ws,1,c,h,fill,white=True)
    # 설정 시트 좌표: 위도=B2, 고도=B3, 풍속계=B4, fetch=B5 (기존 그대로).
    # Kc 관련 셀은 build 시점에 KR 딕셔너리로 관리되지만, 계산과정 수식은 그 좌표(문자열)를
    # 그대로 참조. Kc 시스템 셀 좌표는 하단에서 참조.
    for i in range(2,last+1):
        r=i  # 원데이터와 행 정렬 동일
        F={
         1:f"={S}!A{r}", 2:f"={S}!A{r}-DATE(YEAR({S}!A{r}),1,1)+1",
         3:f"=({S}!B{r}+{S}!C{r})/2",
         4:f"=0.6108*EXP(17.27*{S}!B{r}/({S}!B{r}+237.3))",
         5:f"=0.6108*EXP(17.27*{S}!C{r}/({S}!C{r}+237.3))",
         6:f"=(D{r}+E{r})/2",
         7:f'=IF({S}!H{r}<>"",{S}!H{r}/10,IF({S}!I{r}<>"",0.6108*EXP(17.27*{S}!I{r}/({S}!I{r}+237.3)),{S}!E{r}/100*F{r}))',
         8:f"=4098*(0.6108*EXP(17.27*C{r}/(C{r}+237.3)))/(C{r}+237.3)^2",
         9:f'=IF({S}!J{r}<>"",{S}!J{r}/10,101.3*((293-0.0065*설정!$B$3)/293)^5.26)',
         10:f"=0.000665*I{r}",
         11:f"={S}!G{r}*4.87/LN(67.8*설정!$B$4-5.42)",
         12:f"=1+0.033*COS(2*PI()*B{r}/365)",
         13:f"=0.409*SIN(2*PI()*B{r}/365-1.39)",
         14:f"=ACOS(-TAN(RADIANS(설정!$B$2))*TAN(M{r}))",
         15:f"=(24*60/PI())*0.082*L{r}*(N{r}*SIN(RADIANS(설정!$B$2))*SIN(M{r})+COS(RADIANS(설정!$B$2))*COS(M{r})*SIN(N{r}))",
         16:f"=(0.75+0.00002*설정!$B$3)*O{r}",
         17:f'=IF({S}!K{r}<>"",{S}!K{r},(0.25+0.5*{S}!L{r}/(24/PI()*N{r}))*O{r})',
         18:f"=0.77*Q{r}",
         19:f"=0.000000004903*((({S}!B{r}+273.16)^4+({S}!C{r}+273.16)^4)/2)*(0.34-0.14*SQRT(G{r}))*(1.35*MIN(Q{r}/P{r},1)-0.35)",
         20:f"=R{r}-S{r}",
         21:f"=(0.408*H{r}*T{r}+J{r}*(900/(C{r}+273))*K{r}*(F{r}-G{r}))/(H{r}+J{r}*(1+0.34*K{r}))",
         # 22: 생육단계 (경과일에 따라 문자열 판정) — LET 대신 중첩 IF (LibreOffice·구형 Excel 호환)
         22:(f'=IF({S}!A{r}-{KR["BUD"]}<0,"발아전",'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]},"초기",'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]}+{KR["LDEV"]},"발육",'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]}+{KR["LDEV"]}+{KR["LMID"]},"중기",'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]}+{KR["LDEV"]}+{KR["LMID"]}+{KR["LLATE"]},"후기","후기이후")))))'),
         # 23: 일별 Kc (FAO-56 식 66 곡선 보간)
         23:(f'=IF({S}!A{r}-{KR["BUD"]}<0,0,'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]},{KR["APP_INI"]},'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]}+{KR["LDEV"]},'
                f'{KR["APP_INI"]}+({KR["APP_MID"]}-{KR["APP_INI"]})*({S}!A{r}-{KR["BUD"]}-{KR["LINI"]}+1)/{KR["LDEV"]},'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]}+{KR["LDEV"]}+{KR["LMID"]},{KR["APP_MID"]},'
             f'IF({S}!A{r}-{KR["BUD"]}<{KR["LINI"]}+{KR["LDEV"]}+{KR["LMID"]}+{KR["LLATE"]},'
                f'{KR["APP_MID"]}+({KR["APP_END"]}-{KR["APP_MID"]})*({S}!A{r}-{KR["BUD"]}-{KR["LINI"]}-{KR["LDEV"]}-{KR["LMID"]}+1)/{KR["LLATE"]},'
             f'0)))))'),
         # 24: 증발접시 Kp
         24:f"=0.108-0.0286*K{r}+0.0422*LN(설정!$B$5)+0.1434*LN({S}!E{r})-0.000631*LN(설정!$B$5)^2*LN({S}!E{r})",
         25:f"={S}!M{r}",                    # 대형증발량
         26:f"=X{r}*Y{r}",                    # ETo_pan = Kp × Epan  (Kp는 X24, Epan은 Y25 → 계산과정 열 X, Y)
         27:f"=U{r}*W{r}",                    # ETc_PM  = ETo_PM × 일별 Kc
         28:f"=Z{r}*W{r}",                    # ETc_pan = ETo_pan × 일별 Kc  (ETo_pan은 Z26)
        }
        for c in range(1,29):
            cell=C(ws,i,c,F[c],fmts[c-1])
            if c==1: cell.number_format="yyyy-mm-dd"
    ws.freeze_panes="B2"; ws.column_dimensions["A"].width=12
    for c in range(2,29): ws.column_dimensions[get_column_letter(c)].width=11
    ws.column_dimensions[get_column_letter(22)].width=9   # 생육단계는 좁게

    # ===== 결과요약 =====
    ws=wb.create_sheet("결과요약")
    H(ws,1,1,f"지점 {p['stn']} 기준증발산량 결과요약  {p['start']}~{p['end']}",GREEN,True,11)
    ws.merge_cells("A1:D1")
    def kv(r,k,formula,fmt=None):
        H(ws,r,1,k,LIGHT); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=4)
        C(ws,r,2,formula,fmt,left=True)
    kv(3,"ETo_PM  월합계 (mm)",f"=SUM(계산과정!U2:U{last})","0.0")
    kv(4,"ETo_PM  일평균 (mm/d)",f"=AVERAGE(계산과정!U2:U{last})","0.00")
    kv(5,"ETo_pan 월합계 (mm)",f"=SUM(계산과정!Z2:Z{last})","0.0")
    kv(6,"ETo_pan 일평균 (mm/d)",f"=AVERAGE(계산과정!Z2:Z{last})","0.00")
    kv(7,"대형증발량 월합계 (mm)",f"=SUM(계산과정!Y2:Y{last})","0.0")
    kv(8,"두 방법 상관 r",f'=IFERROR(CORREL(계산과정!U2:U{last},계산과정!Z2:Z{last}),"—")',"0.000")
    kv(9,"편향 ME (pan−PM, mm/d)",f"=AVERAGE(계산과정!Z2:Z{last})-AVERAGE(계산과정!U2:U{last})","0.00")
    kv(10,"RMSE (mm/d)",f'=IFERROR(SQRT(SUMPRODUCT((계산과정!Z2:Z{last}-계산과정!U2:U{last})^2)/{n}),"—")',"0.00")
    kv(11,"월누적 차이 (pan/PM−1)",f'=IFERROR(SUM(계산과정!Z2:Z{last})/SUM(계산과정!U2:U{last})-1,"—")',"0.0%")
    kv(12,"ETc_PM (=Kc×ETo) 월합계 (mm)",f"=SUM(계산과정!AA2:AA{last})","0.0")
    kv(13,"ETc_PM 일평균 (mm/d)",f"=AVERAGE(계산과정!AA2:AA{last})","0.00")
    # 물수지 요약
    kv(14,"","")  # 빈 행
    H(ws,15,1,"물수지 요약 — 무관수(자연강우) 가정, FAO-56 Ch.8",GREEN,True,11); ws.merge_cells("A15:D15")
    kv(16,"기간 총강수 ΣP (mm)",f"=SUM(물수지!F2:F{last})","0.0")
    kv(17,"유효강수 (근권 흡수분, mm)",f"=SUM(물수지!F2:F{last})-SUM(물수지!H2:H{last})","0.0")
    kv(18,"심층침투 ΣDP (mm)",f"=SUM(물수지!H2:H{last})","0.0")
    kv(19,"유효강수율 (%)",f'=IFERROR((SUM(물수지!F2:F{last})-SUM(물수지!H2:H{last}))/SUM(물수지!F2:F{last})*100,"—")',"0.0")
    kv(20,"관수 필요 발생 일수",f"=COUNTIF(물수지!M2:M{last},\"●\")","0")
    kv(21,"필요 순관수량 ΣIn (mm)",f"=SUM(물수지!N2:N{last})","0.0")
    kv(22,"필요 총관수량 ΣIg (mm)",f"=SUM(물수지!O2:O{last})","0.0")
    kv(23,"TAW (mm)",f"={KR['TAW']}","0.0")
    kv(24,"RAW (mm)",f"={KR['RAW']}","0.0")
    # 일별표
    rr=26
    for c,h in enumerate(["일자","ETo_PM(mm)","ETo_pan(mm)","대형증발량(mm)","ETc_PM(mm)","강수(mm)","Dr(mm)","Ks","관수"],1): H(ws,rr,c,h,GREEN,True)
    for i in range(2,last+1):
        rr+=1
        C(ws,rr,1,f"=계산과정!A{i}"); ws.cell(rr,1).number_format="yyyy-mm-dd"
        C(ws,rr,2,f"=계산과정!U{i}","0.00"); C(ws,rr,3,f"=계산과정!Z{i}","0.00")
        C(ws,rr,4,f"=계산과정!Y{i}","0.0"); C(ws,rr,5,f"=계산과정!AA{i}","0.00")
        C(ws,rr,6,f"=물수지!F{i}","0.0"); C(ws,rr,7,f"=물수지!L{i}","0.0")
        C(ws,rr,8,f"=물수지!J{i}","0.00"); C(ws,rr,9,f"=물수지!M{i}")
    for c,w in zip(range(1,10),[24,12,12,12,12,9,9,8,7]): ws.column_dimensions[get_column_letter(c)].width=w

    # ===== 물수지 (FAO-56 Ch.8, 식82~88, 라이브 수식) =====
    WB = "물수지"
    ws=wb.create_sheet(WB)
    # ※ 무관수(자연강우만) 가정 — 실측 관수 데이터 없음.
    #   In/Ig는 FAO-56 식(85) 역산한 이론적 필요량이며 실제 관수 후 Dr 리셋 없음.
    #   다음 버전에서 실측 관수량 입력 시 식(85)에 직접 반영 예정.
    wb_heads=["일자","생육단계","Kc","ETo\n(mm)","ETc\n(mm)","강수P\n(mm)",
              "Dr,i-1\n(mm)","DP\n(mm)","Dr비후\n(mm)","Ks","ETc_adj\n(mm)",
              "Dr,i\n(mm)","관수\n필요","필요 순관수량\nIn (mm)","필요 총관수량\nIg (mm)"]
    wb_fmts=[None,"@","0.000","0.00","0.00","0.0","0.0","0.0","0.0","0.00","0.00","0.0",None,"0.0","0.0"]
    SOIL_BROWN = "8B6914"
    for c,h in enumerate(wb_heads,1):
        H(ws,1,c,h,SOIL_BROWN,white=True)
    # 셀 참조: 설정 시트의 토양 파라미터
    # KR 값은 이미 '설정!$B$41' 형식이므로 그대로 사용 (설정! 이중 참조 방지)
    _TAW = KR['TAW']
    _RAW = KR['RAW']
    _DR0 = KR['DR0']
    _EA  = KR['EA']
    for i in range(2, last+1):
        r = i
        # G: Dr,i-1 — 첫 행은 초기값, 이후는 전일 Dr,i(L열)
        dr_prev = _DR0 if i == 2 else f"L{r-1}"
        WF = {
            1: f"=계산과정!A{r}",                     # 일자
            2: f"=계산과정!V{r}",                     # 생육단계
            3: f"=계산과정!W{r}",                     # Kc
            4: f"=계산과정!U{r}",                     # ETo
            5: f"=D{r}*C{r}",                         # ETc = ETo × Kc
            6: f"={S}!N{r}",                          # 강수량 P (원데이터 N열)
            7: f"={dr_prev}",                         # Dr,i-1
            8: f"=MAX(F{r}-G{r}, 0)",                 # DP = max(P - Dr,i-1, 0) [식88 간이]
            9: f"=MAX(G{r}-F{r}, 0)",                 # Dr_비후 = max(Dr,i-1 - P, 0)
            10: f"=IF(I{r}<={_RAW}, 1, ({_TAW}-I{r})/({_TAW}-{_RAW}))",  # Ks [식84]
            11: f"=J{r}*E{r}",                        # ETc_adj = Ks × ETc
            12: f"=MIN(I{r}+K{r}, {_TAW})",           # Dr,i [식85] 0≤Dr≤TAW
            13: f'=IF(L{r}>={_RAW},"●","")',          # 관수필요 판정
            14: f'=IF(M{r}="●",L{r},0)',              # 순관수 In = Dr(근권 전량보충)
            15: f'=IF(N{r}>0,N{r}/{_EA},0)',           # 총관수 Ig = In/Ea
        }
        for c in range(1, 16):
            cell = C(ws, r, c, WF[c], wb_fmts[c-1])
            if c == 1: cell.number_format = "yyyy-mm-dd"
            # Dr,i 열: RAW 초과 시 경고색
            if c == 12:
                # 조건부 서식은 openpyxl에서 번거로우므로 값 기반 서식은 엑셀에서 수동 설정 안내
                pass
    ws.freeze_panes = "B2"
    wb_widths = [12, 8, 7, 8, 8, 8, 8, 8, 8, 7, 8, 8, 6, 9, 9]
    for c, w in enumerate(wb_widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    # ===== 계산근거 =====
    ws=wb.create_sheet("계산근거"); ws.column_dimensions["A"].width=18; ws.column_dimensions["B"].width=90
    doc=[("ETo 산출 근거 (FAO-56)","",True),
         ("기준","Allen et al., FAO Irrigation & Drainage Paper 56 (1998)"),
         ("원데이터","기상청 ASOS 일자료(공공데이터포털 getWthrDataList) — 공개 OpenAPI만 사용"),
         ("es","포화수증기압 (e°(Tmax)+e°(Tmin))/2, e°(T)=0.6108·EXP(17.27·T/(T+237.3))  [식11]"),
         ("ea","실제수증기압. 우선순위: 측정증기압(avgPv÷10) → 이슬점 e°(avgTd) → 상대습도(RH/100·es)"),
         ("Δ","포화수증기압곡선 기울기 4098·e°(T)/(T+237.3)²  [식13]"),
         ("P, γ","P=현지기압(avgPa÷10) 또는 고도식, γ=0.000665·P  [식7,8]"),
         ("u₂","2m 풍속 u_z·4.87/LN(67.8·z−5.42), z=풍속계높이  [식47]"),
         ("Ra","천문일사: 위도·연중일(DOY)로 산정 (dr·δ·ωs 경유)  [식21]"),
         ("Rso","청천일사 (0.75+2e-5·고도)·Ra  [식37]"),
         ("Rs","일사: 측정 합계일사(sumGsr) 우선, 결측 시 일조기반 Ångström (0.25+0.50·n/N)·Ra  [식35]"),
         ("Rn","순복사 Rns−Rnl, Rns=0.77·Rs, Rnl=σ·((Tmax⁴+Tmin⁴)/2)·(0.34−0.14·√ea)·(1.35·Rs/Rso−0.35)  [식38~40]"),
         ("ETo_PM","기준증발산량(PM). 분자 0.408·Δ·Rn + γ·900/(T+273)·u₂·(es−ea), 분모 Δ + γ·(1+0.34·u₂)  [식6, G=0]"),
         ("Kp","Class A 증발접시 계수(녹지 fetch, Allen-Pruitt): u₂·상대습도·상풍거리(fetch)의 함수"),
         ("ETo_pan","증발접시법 기준증발산량 = Kp × 대형증발량(sumLrgEv)  [FAO-56]"),
         ("Kc 시나리오","FAO-56 Table 12(사과·배·체리) 4개 시나리오. 축1: 청경(초생 없음) vs 초생재배(active ground cover) / 축2: 서리 발생 vs 무서리. 한국 사과원은 대개 시나리오 3(초생·서리, 0.50/1.20/0.95). 설정 시트의 '시나리오 번호'를 바꾸면 표값 3개(Kc_ini/mid/end)가 자동 갱신됨."),
         ("생육단계 [식66]","발아일 + L_ini/dev/mid/late로 5구간(발아전·초기·발육·중기·후기·후기이후) 판정. 일별 Kc 곡선: 초기=Kc_ini 유지 → 발육=선형 증가 → 중기=Kc_mid 유지 → 후기=Kc_end로 선형 감소 → 후기이후=0."),
         ("Kc 현지기상 보정 [식62/65]","Kc_mid_adj = Kc_mid_table + [0.04·(u2−2) − 0.004·(RHmin−45)]·(h/3)^0.3.  u2는 MEDIAN(1,u2,6), RHmin은 MEDIAN(20,RHmin,80)으로 자동 제한. Kc_end는 표값>0.45일 때만 같은 식으로 보정(식65). Kc_ini는 현지기상 보정 대상 아님(멀칭 보정만)."),
         ("멀칭 보정","전면 비닐멀칭 시 Kc를 10~30% 낮춤(FAO-56 p.196). 설정의 '멀칭 Kc 보정계수'로 Kc_ini/mid/end 모두에 곱함."),
         ("ETc","일별 ETc = 일별 Kc × ETo. 계산과정 시트에서 매일 자동 산출. Kc_pan 방식은 참고용(같은 일별 Kc를 ETo_pan에 곱함)."),
         ("물수지 근거 (FAO-56 Ch.8)","",True),
         ("TAW [식82]","총유효수분 TAW = 1000·(θFC−θWP)·Zr (mm). θFC/θWP는 FAO-56 Table 19, Zr은 Table 22 기준값."),
         ("RAW [식83]","쉽게이용가능수분 RAW = p·TAW (mm). Dr이 RAW에 도달하면 작물이 스트레스를 받기 시작. p는 Table 22 기준값."),
         ("Ks [식84]","수분스트레스계수. Dr≤RAW이면 Ks=1(무스트레스). Dr>RAW이면 Ks=(TAW−Dr)/(TAW−RAW), 0~1 범위."),
         ("Dr [식85]","일별 근권 고갈량. Dr,i = Dr,i-1 − P + ETc,adj + DP. 범위: 0 ≤ Dr ≤ TAW."),
         ("DP [식88]","심층침투. DP = max(P − Dr,i-1, 0). 강수가 현재 고갈량보다 많으면 근권을 넘쳐 아래로 배수."),
         ("ETc_adj [식81]","스트레스 보정 후 증발산. ETc_adj = Ks · Kc · ETo. Ks<1이면 실제 소비량이 잠재 ETc보다 줄어듦."),
         ("관수필요 판정","Dr,i ≥ RAW이면 관수 필요(●). 무관수(자연강우만) 가정이므로 ● 이후에도 Dr 리셋 없이 계속 누적됨. 실측 관수 데이터는 다음 버전에서 식(85)에 직접 반영 예정."),
         ("필요 순관수량 In","net irrigation depth. 관수필요(●) 시점의 Dr 값. 근권을 포장용수량(Dr=0)까지 보충하는 데 필요한 순수량(mm). 이론적 필요량이며 실측 관수량이 아님."),
         ("필요 총관수량 Ig","gross irrigation depth. In을 관수효율 Ea로 나눈 값(Ig = In/Ea). 손실(미도달·증발 등)을 포함한 실제 공급 필요량. 점적관수 Ea=0.95 기준."),
         ("계산과정 시트","모든 셀이 라이브 수식. 셀 클릭 시 수식 확인 가능. 설정값 변경 시 자동 재계산됨."),
         ("물수지 시트","라이브 수식. 설정의 토양 파라미터(θFC/θWP/Zr/p) 변경 시 전체 자동 재계산됨."),
         ("주의","일사·대형증발량은 지점별 관측 여부 상이. 물수지는 RO(지표유출)=0, CR(모관상승)=0, 실측 관수량=0(무관수 가정)으로 단순화.")]
    rr=0
    for k,v,*hd in doc:
        rr+=1; head=bool(hd and hd[0])
        if v and str(v).startswith("="): v=" "+v      # '=' 로 시작하는 텍스트 수식오인 방지
        H(ws,rr,1,k,(GREEN if head else LIGHT),white=head,size=(11 if head else 10))
        cb=ws.cell(rr,2,v); cb.font=Font(name=FONT,bold=head,size=(11 if head else 10),color=("FFFFFF" if head else "1A1A1A"))
        cb.alignment=Alignment(wrap_text=True,vertical="top")
        if head: cb.fill=PatternFill("solid",fgColor=GREEN)

    saved_path=safe_save(wb,out)
    return n, saved_path

def ask(prompt, default):
    try: v=input(f"{prompt} [{default}]: ").strip()
    except EOFError: v=""
    return v or default

def main():
    import datetime as _dt_now
    _today=_dt_now.date.today()
    _yesterday=(_today-_dt_now.timedelta(days=1)).strftime("%Y%m%d")
    _year=_today.year
    _default_bud=f"{_year}0401"    # 조회년도 4/1
    _default_start=f"{_year}0101"  # 조회년도 1/1
    _default_end=_yesterday        # 어제

    ap=argparse.ArgumentParser()
    ap.add_argument("--key",default=None); ap.add_argument("--hubkey",default=None)
    ap.add_argument("--stn",default=None)
    ap.add_argument("--start",default=None); ap.add_argument("--end",default=None)
    ap.add_argument("--lat",type=float,default=None); ap.add_argument("--elev",type=float,default=None)
    ap.add_argument("--anem",type=float,default=None); ap.add_argument("--fetch",type=float,default=100.0)
    ap.add_argument("--crop",default=None,
                    help="작물 crop_id (기본: apple). 사용 가능한 목록은 crops_library.csv 참조")
    ap.add_argument("--bud-date",dest="bud_date",default=None,
                    help="정식/파종/발아일 YYYYMMDD. 작물마다 실제 캘린더가 다르므로 특히 1년생 작물(배추 등)은 필수 확인 권장.")
    ap.add_argument("--out",default="eto_상세.xlsx")
    a=ap.parse_args()

    # 인증키: 인자 > apikey.txt > 직접입력
    keys=load_apikeys()
    key=a.key or keys.get("DATA_GO_KR") or input("공공데이터포털 인증키(Encoding)를 붙여넣으세요: ").strip()
    hubkey=a.hubkey or keys.get("KMA_HUB")

    # ── 대화형 입력 (--인자 없을 때만 물어봄) ──
    stn  =a.stn   or ask("지점번호 (146=전주, 101=춘천, 216=태백)", "146")
    crop_id=a.crop or ask("작물 crop_id (apple/kimchi_cabbage/cabbage/tomato 등)", "apple")
    bud_date_str=a.bud_date or ask("발아·정식일 YYYYMMDD", _default_bud)
    start=a.start or ask("조회 시작일 YYYYMMDD", _default_start)
    end  =a.end   or ask("조회 종료일 YYYYMMDD", _default_end)

    # ── 지점정보 자동조회 (stn_inf.php): 위도·고도·풍속계높이 ──
    # 우선순위: --lat/--elev/--anem 직접지정 > API 자동조회(성공시 백업파일에 저장) >
    #          로컬 백업파일(stations_backup.csv) > 최종 안전값(춘천 근사)
    lat,elev,anem = a.lat,a.elev,a.anem
    stn_name=None
    if a.lat is not None or a.elev is not None or a.anem is not None:
        meta_source="MANUAL"
    else:
        meta_source=None
        if hubkey:
            try:
                info,names=fetch_station_table(hubkey)
                save_station_backup(info,names)   # 성공한 김에 백업파일 갱신(다음 실패 대비 캐싱)
                if int(stn) in info:
                    lat,lon,elev,anem=info[int(stn)]
                    stn_name=names.get(int(stn))
                    meta_source="API"
                    print(f"[지점정보] {stn} {stn_name or ''}: 위도 {lat}, 고도 {elev}m, 풍속계 {anem}m (API 자동조회, 백업파일 갱신됨)")
                else:
                    print(f"[지점정보] {stn}이 지점일람표 응답에 없음 → 백업파일/기본값 사용")
            except Exception as e:
                print(f"[지점정보] API 허브 연결 실패({e}) → 백업파일에서 조회")
        else:
            print("[지점정보] API허브 인증키(KMA_HUB) 없음 → 백업파일에서 조회. 정확도를 높이려면 apikey.txt에 KMA_HUB 추가 권장")
        # API 미사용/실패 시 로컬 백업파일 확인 (엉뚱한 다른 지점 좌표가 쓰이는 것을 방지)
        if meta_source is None:
            binfo,bnames=load_station_backup()
            if int(stn) in binfo:
                lat,lon,elev,anem=binfo[int(stn)]
                stn_name=bnames.get(int(stn))
                meta_source="BACKUP"
                print(f"  → 지점 {stn} {stn_name or ''} 백업파일 값 사용: 위도 {lat}, 고도 {elev}m, 풍속계 {anem}m")
            else:
                meta_source="DEFAULT"
                print(f"  ⚠ 지점 {stn}은 백업파일에도 없어 최종 안전값(춘천 근사)을 씁니다 — 실제와 다를 수 있으니 --lat/--elev/--anem로 직접 지정 권장")
    # 최종 안전 기본값(춘천) — 위 어느 경로로도 못 채운 값만 채움
    if lat  is None: lat=37.9026
    if elev is None: elev=77.7
    if anem is None: anem=10.0

    print(f"[조회] 지점 {stn}, {start}~{end} …")
    rows=fetch_asos(key,stn,start,end)
    if not rows: print("데이터가 없습니다. 지점/기간/인증키를 확인하세요."); sys.exit(1)
    # ── 작물 라이브러리에서 선택한 작물 로드 ──
    crop_lib=load_crop_library()
    if crop_id not in crop_lib:
        available=", ".join(c for c in [x["crop_id"] for x in crops_sorted(crop_lib)][:20])+" ..."
        print(f"[작물] '{crop_id}'가 라이브러리에 없음 → 기본값(apple)로 진행.")
        print(f"  사용 가능: {available}")
        crop=crop_lib.get("apple",{})
        crop_id="apple"
    else:
        crop=crop_lib[crop_id]
        print(f"[작물] {crop_id} ({crop.get('name_ko','')}, {crop.get('category','')}) 로드 — "
              f"시나리오 {len(crop['scenarios'])}개, L={crop.get('L_ini')}/{crop.get('L_dev')}/{crop.get('L_mid')}/{crop.get('L_late')}, h={crop.get('h_m')}m")

    # ── 정식/파종/발아일 (--bud-date) 및 단기작물(1년생 등) 경고 ──
    import datetime as _dt2
    SHORT_CYCLE_DAYS=200   # 이보다 생육기간 총합이 짧으면 "1년생/단기작물"로 간주해 경고
    L_ini_v=crop.get("L_ini") or 20; L_dev_v=crop.get("L_dev") or 70
    L_mid_v=crop.get("L_mid") or 90; L_late_v=crop.get("L_late") or 30
    L_total_v=L_ini_v+L_dev_v+L_mid_v+L_late_v
    is_short_cycle=L_total_v < SHORT_CYCLE_DAYS
    bud_date_manual = bud_date_str != _default_bud   # 기본값과 다르면 사용자가 직접 지정한 것
    try:
        bud_date=_dt2.datetime.strptime(bud_date_str,"%Y%m%d").date()
        if bud_date_manual:
            print(f"[생육시작일] 사용자 지정: {bud_date}")
    except ValueError:
        print(f"[생육시작일] 형식 오류('{bud_date_str}', YYYYMMDD 필요) → 기본값 {_default_bud} 사용")
        bud_date=_dt2.datetime.strptime(_default_bud,"%Y%m%d").date()
        bud_date_manual=False
    if is_short_cycle and not bud_date_manual:
        print(f"  ⚠⚠ 경고: '{crop.get('name_ko',crop_id)}'는 생육기간 총 {L_total_v}일의 단기/1년생 작물로 추정됩니다.")
        print(f"     정식·파종일을 확인하지 않고 기본값({_default_bud})을 적용하면 생육단계·Kc·ETc가 실제와 크게 어긋날 수 있습니다.")
        print(f"     --bud-date YYYYMMDD 로 실제 정식/파종일을 반드시 지정하세요. (예: --bud-date 20260603)")

    # ── 생육중기·후기 구간의 실측 u2·RHmin 자동 집계 (Kc 현지기상 보정식 62/65 입력용) ──
    # 강원도원 워크북과 정합화하기 위해, 조회기간 중 생육중기/후기에 해당하는 날짜의
    # 실측 u2(풍속계높이→2m 환산)·minRhm 평균을 계산해 p에 넣는다. 최소 10일 이상 있을 때만 적용.
    mid_start=bud_date+_dt2.timedelta(days=L_ini_v+L_dev_v)
    mid_end  =bud_date+_dt2.timedelta(days=L_ini_v+L_dev_v+L_mid_v-1)
    late_start=mid_end+_dt2.timedelta(days=1)
    late_end  =bud_date+_dt2.timedelta(days=L_total_v-1)
    def _stage_agg(rows_,d_start,d_end):
        u2s=[]; rhs=[]
        for x in rows_:
            d=x["tm"]
            d=d if isinstance(d,_dt2.date) else _dt2.datetime.strptime(str(d)[:10],"%Y-%m-%d").date()
            if d_start<=d<=d_end:
                ws=x.get("avgWs"); rh=x.get("minRhm")
                if ws is not None: u2s.append(wind_2m(ws, anem or 10.0))
                if rh is not None: rhs.append(rh)
        return (sum(u2s)/len(u2s) if len(u2s)>=10 else None,
                sum(rhs)/len(rhs) if len(rhs)>=10 else None,
                len(u2s))
    u2_mid_v,rh_mid_v,n_mid=_stage_agg(rows,mid_start,mid_end)
    u2_end_v,rh_end_v,n_end=_stage_agg(rows,late_start,late_end)
    if u2_mid_v is not None:
        print(f"[Kc 자동집계] 중기({mid_start}~{mid_end}, {n_mid}일 실측): u2={u2_mid_v:.3f} m/s, RHmin={rh_mid_v:.1f}%")
    else:
        print(f"[Kc 자동집계] 중기 구간 실측 부족(<10일) → 기본값(u2=1.5, RHmin=55) 사용. 정확한 Kc 현지기상 보정을 위해서는 중기 포함 기간 조회 권장")
    if u2_end_v is not None:
        print(f"[Kc 자동집계] 후기({late_start}~{late_end}, {n_end}일 실측): u2={u2_end_v:.3f} m/s, RHmin={rh_end_v:.1f}%")

    p=dict(lat=lat,elev=elev,anem=anem,fetch=a.fetch,stn=stn,start=start,end=end,meta_source=meta_source,crop=crop,
           bud_date=bud_date,bud_date_manual=bud_date_manual,is_short_cycle=is_short_cycle,L_total=L_total_v,
           u2_mid=u2_mid_v,rh_mid=rh_mid_v,u2_end=u2_end_v,rh_end=rh_end_v)
    out=a.out if a.out!="eto_상세.xlsx" else f"output/eto({stn})_{crop_id}_{start}_{end}.xlsx"
    n,saved_path=build_workbook(rows,p,out)
    print(f"[완료] {saved_path}  ({n}일)  → 시트: 설정 / 원데이터 / 계산과정(수식) / 결과요약 / 물수지 / 계산근거")
    print("  엑셀에서 열면 자동 계산됩니다. '설정' 시트의 노란칸(위도·고도 등)을 바꾸면 재계산돼요.")
    if is_short_cycle and not bud_date_manual:
        print("  ⚠ 설정 시트의 '생육 시작일' 행이 경고색으로 표시되어 있습니다 — 반드시 확인 후 수정하세요.")

if __name__=="__main__":
    main()
