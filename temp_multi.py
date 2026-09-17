"""
temp_multi.py — 다지점 온도·적산온도 종합 분석 엑셀 생성

하나의 엑셀 파일 안에 시트로 모든 분석 결과를 출력한다:
  1. 설정          — 조회 조건 요약 (노란칸 편집 가능)
  2. 권역별         — 권역 선택 콤보 + 동기간 비교
  3. 지점별         — 지점 선택 콤보 + 동기간 비교
  4. 기간별(월별)   — 권역=열, 구분=행, 월별 블록 구성
  5. 년도별         — 슬라이딩 윈도우 (올해-N년 ~ 올해+5년)
  6. 적산온도 도달   — 권역별 임계값별 도달일자 + 평년
  7. 원데이터       — 지점별 일별 전체 데이터 (맨 뒤)

사용법:
  python temp_multi.py                             # 전체 66개 지점, Tbase=10
  python temp_multi.py --crop apple                # 사과 기준온도(4℃) 자동 적용
  python temp_multi.py --tbase 5                   # 기준온도 직접 지정
  python temp_multi.py --region 강원영서             # 특정 권역만
  python temp_multi.py --method avg                # avgTa 방법 사용
"""

import argparse
import sys
from datetime import date, timedelta
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment

import stations
import solar_core
import temp_core
from excel_util import (
    H, C, FONT_NAME, GREEN, BLUE, BROWN, LIGHT, YEL, WARN,
    LIGHT_BLUE, LIGHT_YEL, BORDER, FMT_NUM1, FMT_NUM2, FMT_INT,
    WHITE, set_col_widths, write_header_row, write_data_row, apply_title,
)


# ---------------------------------------------------------------------------
# 동기간 비교 유틸
# ---------------------------------------------------------------------------

def _filter_same_period(parsed_rows, start_md, end_md):
    filtered = []
    for r in parsed_rows:
        md = (r["month"], r["day"])
        if start_md <= md <= end_md:
            filtered.append(r)
    return filtered


def _period_md_from_dates(start_date, end_date):
    return (start_date.month, start_date.day), (end_date.month, end_date.day)


# ---------------------------------------------------------------------------
# 집계 유틸 (temp_core 래핑)
# ---------------------------------------------------------------------------

def _daily_average_rows(rows):
    """다중 지점 → 일별 지점평균."""
    keys = ["avg_ta", "max_ta", "min_ta", "gdd"]
    daily = defaultdict(lambda: defaultdict(list))
    for r in rows:
        day_key = (r["year"], r["month"], r["day"])
        for key in keys:
            v = r.get(key)
            if v is not None:
                daily[day_key][key].append(v)
    result = []
    for (year, month, day), metrics in sorted(daily.items()):
        row = {"year": year, "month": month, "day": day}
        for key, vals in metrics.items():
            row[key] = sum(vals) / len(vals)
        result.append(row)
    return result


def _monthly_metric_by_year(rows, metric_key, agg_type):
    buckets = defaultdict(lambda: defaultdict(list))
    for r in rows:
        v = r.get(metric_key)
        if v is not None:
            buckets[r["year"]][r["month"]].append(v)
    result = {}
    for year, months in buckets.items():
        result[year] = {}
        for month, vals in months.items():
            if agg_type == "avg":
                result[year][month] = sum(vals) / len(vals)
            elif agg_type == "cumsum":
                result[year][month] = sum(vals)
            elif agg_type == "max":
                result[year][month] = max(vals)
            elif agg_type == "min":
                result[year][month] = min(vals)
    return result


def _yearly_agg(valid_vals, agg_type):
    """월별 값 리스트 → 연간 값 집계. agg_type: avg/cumsum/max/min"""
    if not valid_vals:
        return None
    if agg_type == "cumsum":
        return sum(valid_vals)
    elif agg_type == "max":
        return max(valid_vals)
    elif agg_type == "min":
        return min(valid_vals)
    else:  # avg
        return sum(valid_vals) / len(valid_vals)


def _calc_normal_monthly(monthly_by_year, n_start, n_end):
    normal = {}
    for month in range(1, 13):
        yr_vals = []
        for yr in range(n_start, n_end + 1):
            v = monthly_by_year.get(yr, {}).get(month)
            if v is not None:
                yr_vals.append(v)
        normal[month] = (sum(yr_vals) / len(yr_vals)) if yr_vals else None
    return normal


# ---------------------------------------------------------------------------
# 읽기전용 비교기간 표시
# ---------------------------------------------------------------------------
def _write_readonly_period(ws, row, col):
    H(ws, row, col, "비교기간")
    cell = C(ws, row, col + 1, None, left=True)
    cell.value = '=설정!B8&" ~ "&설정!B9'
    cell.font = Font(name=FONT_NAME, size=10, color="1A1A1A")
    cell.fill = PatternFill("solid", fgColor=WHITE)


# ═══════════════════════════════════════════════════════════════════════════
# 시트 1: 설정
# ═══════════════════════════════════════════════════════════════════════════
def _sheet_settings(wb, start_date, end_date, ref_year, normal_years,
                    target_regions, target_stations,
                    t_base, t_upper, method, crop_id):
    ws = wb.active
    ws.title = "설정"
    set_col_widths(ws, [18, 40, 20])

    apply_title(ws, 1, 1, "온도·적산온도 분석 설정", merge_end_col=3, size=12)
    C(ws, 2, 1, "(노란칸은 수정 가능)", left=True)

    row = 4
    H(ws, row, 1, "올해");              C(ws, row, 2, ref_year, fill=YEL)
    row += 1
    H(ws, row, 1, "작년");              C(ws, row, 2, None)
    ws.cell(row, 2).value = "=B4-1";    ws.cell(row, 2).number_format = "0"
    row += 1
    H(ws, row, 1, "평년 기간(년)");      C(ws, row, 2, normal_years, fill=YEL)
    row += 1
    H(ws, row, 1, "평년 범위");          C(ws, row, 2, None, left=True)
    ws.cell(row, 2).value = '=B4-B6&" ~ "&B4-1&" ("&B6&"년 평균)"'
    row += 1

    # 분석 기간 — 행 8, 9 고정 (다른 시트에서 설정!B8, 설정!B9 참조)
    H(ws, row, 1, "분석 시작일")         # row=8
    C(ws, row, 2, start_date.strftime("%m-%d"))
    ws.cell(row, 2).number_format = "@"
    C(ws, row, 3, "← 변경 시 스크립트 재실행 필요", left=True)
    ws.cell(row, 3).font = Font(name=FONT_NAME, size=9, color="999999")
    row += 1
    H(ws, row, 1, "분석 종료일")         # row=9
    C(ws, row, 2, end_date.strftime("%m-%d"))
    ws.cell(row, 2).number_format = "@"
    row += 1

    # 적산온도 파라미터 — 행 10, 11, 12
    row += 1
    H(ws, row, 1, "기준온도 Tbase (℃)") # row=11
    C(ws, row, 2, t_base)
    row += 1
    H(ws, row, 1, "상한온도 Tupper (℃)")
    C(ws, row, 2, t_upper if t_upper else "")
    row += 1
    H(ws, row, 1, "계산 방법")
    method_label = {"maxmin": "(Tmax+Tmin)/2", "avg": "avgTa", "aquacrop": "AquaCrop"}
    C(ws, row, 2, method_label.get(method, method), left=True)
    row += 1
    H(ws, row, 1, "작물")
    C(ws, row, 2, crop_id if crop_id else "(사용자 지정)", left=True)
    row += 1

    # 고정 정보
    row += 1
    items = [
        ("데이터 소스", "기상청 ASOS 일자료 (공공데이터포털)"),
        ("대상 권역", ", ".join(target_regions) if target_regions else "전체"),
        ("대상 지점 수", str(len(target_stations))),
        ("생성일", date.today().isoformat()),
    ]
    for label, value in items:
        H(ws, row, 1, label)
        C(ws, row, 2, value, left=True)
        row += 1

    # 지점 목록
    row += 1
    H(ws, row, 1, "지점번호"); H(ws, row, 2, "지점명"); H(ws, row, 3, "권역")
    for stn_id in target_stations:
        row += 1
        C(ws, row, 1, stn_id)
        C(ws, row, 2, stations.station_name(stn_id), left=True)
        C(ws, row, 3, stations.region_of(stn_id) or "", left=True)


# ═══════════════════════════════════════════════════════════════════════════
# 시트 7 (맨 뒤): 원데이터
# ═══════════════════════════════════════════════════════════════════════════
def _sheet_raw_data(wb, all_parsed, t_base, method):
    ws = wb.create_sheet("원데이터")
    headers = [
        "날짜", "지점번호", "지점명", "권역",
        "평균기온(℃)", "최고기온(℃)", "최저기온(℃)",
        f"GDD(Tbase={t_base}℃)",
    ]
    write_header_row(ws, 1, headers)
    set_col_widths(ws, [12, 9, 8, 8, 12, 12, 12, 14])
    ws.freeze_panes = "A2"

    row = 2
    for stn_id in sorted(all_parsed.keys()):
        region = stations.region_of(stn_id) or ""
        for p in all_parsed[stn_id]:
            values = [
                p["date"].isoformat(),
                stn_id,
                stations.station_name(stn_id),
                region,
                p["avg_ta"], p["max_ta"], p["min_ta"],
                p.get("gdd"),
            ]
            write_data_row(ws, row, values, fmt=FMT_NUM1, left_cols={0, 2, 3})
            ws.cell(row, 2).number_format = "0"
            row += 1


# ═══════════════════════════════════════════════════════════════════════════
# 시트 2. 권역별 집계 — VLOOKUP 기반
# ═══════════════════════════════════════════════════════════════════════════

def _build_region_ref_sheet(wb, region_data, ref_year, normal_years):
    """숨긴 _권역별REF 시트: 복합키-값 형태로 온도·GDD 집계 저장."""
    ws = wb.create_sheet("_권역별REF")

    metrics = [
        ("평균기온", "avg_ta", "avg"),
        ("최고기온", "max_ta", "max"),
        ("최저기온", "min_ta", "min"),
        ("최고기온평균", "max_ta", "avg"),
        ("최저기온평균", "min_ta", "avg"),
        ("적산온도", "gdd", "cumsum"),
    ]

    row = 1
    ws.cell(row, 1, "키"); ws.cell(row, 2, "값")
    row = 2

    for metric_label, metric_key, agg_type in metrics:
        for region in stations.REGION_ORDER:
            reg_rows = region_data.get(region, [])
            if not reg_rows:
                continue
            avg_rows = _daily_average_rows(reg_rows)
            monthly = _monthly_metric_by_year(avg_rows, metric_key, agg_type)
            all_years = sorted(monthly.keys())

            for year in all_years:
                vals = monthly[year]
                for m in range(1, 13):
                    key = f"{metric_label}_{region}_{year}_{m}"
                    v = vals.get(m)
                    ws.cell(row, 1, key)
                    ws.cell(row, 2, round(v, 2) if v is not None else "")
                    row += 1
                # 연간
                valid = [vals[m] for m in range(1, 13) if vals.get(m) is not None]
                yearly_val = _yearly_agg(valid, agg_type)
                key = f"{metric_label}_{region}_{year}_연"
                ws.cell(row, 1, key)
                ws.cell(row, 2, round(yearly_val, 2) if yearly_val is not None else "")
                row += 1

            # 평년
            for base_year in all_years:
                n_start = base_year - normal_years
                n_end = base_year - 1
                if not any(y in monthly for y in range(n_start, n_end + 1)):
                    continue
                normal_vals = _calc_normal_monthly(monthly, n_start, n_end)
                for m in range(1, 13):
                    key = f"{metric_label}_{region}_평년{base_year}_{m}"
                    v = normal_vals.get(m)
                    ws.cell(row, 1, key)
                    ws.cell(row, 2, round(v, 2) if v is not None else "")
                    row += 1
                valid = [normal_vals[m] for m in range(1, 13) if normal_vals.get(m) is not None]
                yearly_val = _yearly_agg(valid, agg_type)
                key = f"{metric_label}_{region}_평년{base_year}_연"
                ws.cell(row, 1, key)
                ws.cell(row, 2, round(yearly_val, 2) if yearly_val is not None else "")
                row += 1

    ws.sheet_state = "hidden"
    return ws


def _write_region_vlookup_block(ws, row_start, title, metric_label):
    """권역별 VLOOKUP 블록 — 올해/작년/평년/차이 5행."""
    apply_title(ws, row_start, 1, title, merge_end_col=14, size=11)
    headers = ["비교"] + [f"{m}월" for m in range(1, 13)] + ["연간"]
    write_header_row(ws, row_start + 1, headers)

    row = row_start + 2
    compare_rows = [
        ('="올해("&설정!$B$4&")"', '설정!$B$4'),
        ('="작년("&설정!$B$5&")"', '설정!$B$5'),
        ('="평년("&(설정!$B$4-설정!$B$6)&"~"&설정!$B$5&")"', '"평년"&설정!$B$4'),
    ]

    row_refs = {}
    for i, (label_formula, year_ref) in enumerate(compare_rows):
        comp_name = ["올해", "작년", "평년"][i]
        row_refs[comp_name] = row

        cell_a = ws.cell(row, 1)
        cell_a.value = label_formula
        cell_a.font = Font(name=FONT_NAME, size=10)
        cell_a.alignment = Alignment("left", "center")
        cell_a.border = BORDER

        for m in range(1, 13):
            col_idx = m + 1
            formula = (
                f'=IFERROR(VLOOKUP("{metric_label}_"&$B$1&"_"&{year_ref}&"_{m}",'
                f"'_권역별REF'!$A:$B,2,FALSE),\"\")"
            )
            cell = ws.cell(row, col_idx)
            cell.value = formula
            cell.number_format = FMT_NUM1
            cell.border = BORDER

        formula = (
            f'=IFERROR(VLOOKUP("{metric_label}_"&$B$1&"_"&{year_ref}&"_연",'
            f"'_권역별REF'!$A:$B,2,FALSE),\"\")"
        )
        ws.cell(row, 14).value = formula
        ws.cell(row, 14).number_format = FMT_NUM1
        ws.cell(row, 14).border = BORDER
        row += 1

    this_r = row_refs["올해"]
    last_r = row_refs["작년"]
    normal_r = row_refs["평년"]

    for diff_label, ref_r in [("차이(전년)", last_r), ("차이(평년)", normal_r)]:
        C(ws, row, 1, diff_label, left=True)
        ws.cell(row, 1).border = BORDER
        for col_idx in range(2, 15):
            cl = get_column_letter(col_idx)
            formula = (
                f'=IF(OR({cl}{this_r}="",{cl}{ref_r}=""),"",'
                f'{cl}{this_r}-{cl}{ref_r})'
            )
            ws.cell(row, col_idx).value = formula
            ws.cell(row, col_idx).number_format = FMT_NUM1
            ws.cell(row, col_idx).border = BORDER
        row += 1

    return row


def _sheet_region_summary(wb, all_parsed, ref_year, normal_years,
                          start_date, end_date, target_regions_list,
                          t_base, t_upper, method):
    ws = wb.create_sheet("권역별")

    start_md, end_md = _period_md_from_dates(start_date, end_date)

    # 데이터 준비: GDD 계산 + 권역별 묶기 + 동기간 필터
    region_data = defaultdict(list)
    for stn_id, parsed in all_parsed.items():
        temp_core.add_gdd_to_rows(parsed, t_base, t_upper, method)
        region = stations.region_of(stn_id) or "기타"
        filtered = _filter_same_period(parsed, start_md, end_md)
        region_data[region].extend(filtered)

    # 전국
    all_rows = []
    for rows in region_data.values():
        all_rows.extend(rows)
    region_data["전국"] = all_rows

    # 숨긴 REF 시트
    _build_region_ref_sheet(wb, region_data, ref_year, normal_years)

    # 콤보 박스
    H(ws, 1, 1, "권역 선택")
    C(ws, 1, 2, "전국", fill=YEL)
    region_list = '"' + ",".join(stations.REGION_ORDER) + '"'
    dv = DataValidation(type="list", formula1=region_list, allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws["B1"])

    _write_readonly_period(ws, 1, 3)

    # 지표 블록
    metrics = [
        (f"< 평균기온 (℃) >", "평균기온"),
        (f"< 최고기온 (℃) >", "최고기온"),
        (f"< 최저기온 (℃) >", "최저기온"),
        (f"< 최고기온 평균 (℃) >", "최고기온평균"),
        (f"< 최저기온 평균 (℃) >", "최저기온평균"),
        (f"< 적산온도 (℃·일, Tbase={t_base}℃) >", "적산온도"),
    ]

    row = 3
    for title_text, metric_label in metrics:
        row = _write_region_vlookup_block(ws, row, title_text, metric_label)
        row += 1

    set_col_widths(ws, [22, 9] + [9] * 12 + [10])


# ═══════════════════════════════════════════════════════════════════════════
# 시트 3: 지점별 집계
# ═══════════════════════════════════════════════════════════════════════════

def _build_station_ref_sheet(wb, station_data, ref_year, normal_years):
    ws = wb.create_sheet("_지점별REF")

    metrics = [
        ("평균기온", "avg_ta", "avg"),
        ("최고기온", "max_ta", "max"),
        ("최저기온", "min_ta", "min"),
        ("최고기온평균", "max_ta", "avg"),
        ("최저기온평균", "min_ta", "avg"),
        ("적산온도", "gdd", "cumsum"),
    ]

    ws.cell(1, 1, "키"); ws.cell(1, 2, "값")
    ws.cell(1, 4, "지점표시명"); ws.cell(1, 5, "지점번호")

    stn_list_row = 2
    first_display = ""
    for stn_id in sorted(station_data.keys()):
        display = f"{stations.station_name(stn_id)}({stn_id})"
        ws.cell(stn_list_row, 4, display)
        ws.cell(stn_list_row, 5, stn_id)
        if stn_list_row == 2:
            first_display = display
        stn_list_row += 1
    n_stations = stn_list_row - 2

    row = 2
    for metric_label, metric_key, agg_type in metrics:
        for stn_id in sorted(station_data.keys()):
            parsed = station_data.get(stn_id, [])
            if not parsed:
                continue
            monthly = _monthly_metric_by_year(parsed, metric_key, agg_type)
            all_years = sorted(monthly.keys())

            for year in all_years:
                vals = monthly[year]
                for m in range(1, 13):
                    ws.cell(row, 1, f"{metric_label}_{stn_id}_{year}_{m}")
                    v = vals.get(m)
                    ws.cell(row, 2, round(v, 2) if v is not None else "")
                    row += 1
                valid = [vals[m] for m in range(1, 13) if vals.get(m) is not None]
                yearly_val = _yearly_agg(valid, agg_type)
                ws.cell(row, 1, f"{metric_label}_{stn_id}_{year}_연")
                ws.cell(row, 2, round(yearly_val, 2) if yearly_val is not None else "")
                row += 1

            for base_year in all_years:
                n_start = base_year - normal_years
                n_end = base_year - 1
                if not any(y in monthly for y in range(n_start, n_end + 1)):
                    continue
                nv = _calc_normal_monthly(monthly, n_start, n_end)
                for m in range(1, 13):
                    ws.cell(row, 1, f"{metric_label}_{stn_id}_평년{base_year}_{m}")
                    v = nv.get(m)
                    ws.cell(row, 2, round(v, 2) if v is not None else "")
                    row += 1
                valid = [nv[m] for m in range(1, 13) if nv.get(m) is not None]
                yearly_val = _yearly_agg(valid, agg_type)
                ws.cell(row, 1, f"{metric_label}_{stn_id}_평년{base_year}_연")
                ws.cell(row, 2, round(yearly_val, 2) if yearly_val is not None else "")
                row += 1

    ws.sheet_state = "hidden"
    return ws, n_stations, first_display


def _write_station_vlookup_block(ws, row_start, title, metric_label):
    apply_title(ws, row_start, 1, title, merge_end_col=14, size=11)
    headers = ["비교"] + [f"{m}월" for m in range(1, 13)] + ["연간"]
    write_header_row(ws, row_start + 1, headers)

    row = row_start + 2
    compare_rows = [
        ('="올해("&설정!$B$4&")"', '설정!$B$4'),
        ('="작년("&설정!$B$5&")"', '설정!$B$5'),
        ('="평년("&(설정!$B$4-설정!$B$6)&"~"&설정!$B$5&")"', '"평년"&설정!$B$4'),
    ]

    row_refs = {}
    for i, (label_formula, year_ref) in enumerate(compare_rows):
        comp = ["올해", "작년", "평년"][i]
        row_refs[comp] = row
        cell_a = ws.cell(row, 1)
        cell_a.value = label_formula
        cell_a.font = Font(name=FONT_NAME, size=10)
        cell_a.alignment = Alignment("left", "center")
        cell_a.border = BORDER

        for m in range(1, 13):
            ci = m + 1
            formula = (
                f'=IFERROR(VLOOKUP("{metric_label}_"&$B$2&"_"&{year_ref}&"_{m}",'
                f"'_지점별REF'!$A:$B,2,FALSE),\"\")"
            )
            ws.cell(row, ci).value = formula
            ws.cell(row, ci).number_format = FMT_NUM1
            ws.cell(row, ci).border = BORDER

        formula = (
            f'=IFERROR(VLOOKUP("{metric_label}_"&$B$2&"_"&{year_ref}&"_연",'
            f"'_지점별REF'!$A:$B,2,FALSE),\"\")"
        )
        ws.cell(row, 14).value = formula
        ws.cell(row, 14).number_format = FMT_NUM1
        ws.cell(row, 14).border = BORDER
        row += 1

    this_r = row_refs["올해"]
    last_r = row_refs["작년"]
    normal_r = row_refs["평년"]

    for diff_label, ref_r in [("차이(전년)", last_r), ("차이(평년)", normal_r)]:
        C(ws, row, 1, diff_label, left=True)
        ws.cell(row, 1).border = BORDER
        for ci in range(2, 15):
            cl = get_column_letter(ci)
            ws.cell(row, ci).value = (
                f'=IF(OR({cl}{this_r}="",{cl}{ref_r}=""),"",'
                f'{cl}{this_r}-{cl}{ref_r})')
            ws.cell(row, ci).number_format = FMT_NUM1
            ws.cell(row, ci).border = BORDER
        row += 1
    return row


def _sheet_station_summary(wb, all_parsed, ref_year, normal_years,
                           start_date, end_date,
                           t_base, t_upper, method):
    ws = wb.create_sheet("지점별")
    start_md, end_md = _period_md_from_dates(start_date, end_date)

    station_data = {}
    for stn_id, parsed in all_parsed.items():
        temp_core.add_gdd_to_rows(parsed, t_base, t_upper, method)
        filtered = _filter_same_period(parsed, start_md, end_md)
        station_data[stn_id] = filtered

    ref_ws, n_stations, first_display = _build_station_ref_sheet(
        wb, station_data, ref_year, normal_years)

    # 지점 선택 UI
    H(ws, 1, 1, "지점 선택")
    C(ws, 1, 2, first_display, fill=YEL)
    dv_range = f"'_지점별REF'!$D$2:$D${n_stations + 1}"
    dv = DataValidation(type="list", formula1=dv_range, allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws["B1"])

    H(ws, 2, 1, "지점번호")
    ws.cell(2, 2).value = f"=VLOOKUP(B1,'_지점별REF'!$D:$E,2,FALSE)"
    ws.cell(2, 2).number_format = "0"

    _write_readonly_period(ws, 1, 3)

    metrics = [
        (f"< 평균기온 (℃) >", "평균기온"),
        (f"< 최고기온 (℃) >", "최고기온"),
        (f"< 최저기온 (℃) >", "최저기온"),
        (f"< 최고기온 평균 (℃) >", "최고기온평균"),
        (f"< 최저기온 평균 (℃) >", "최저기온평균"),
        (f"< 적산온도 (℃·일, Tbase={t_base}℃) >", "적산온도"),
    ]

    row = 4
    for title_text, metric_label in metrics:
        row = _write_station_vlookup_block(ws, row, title_text, metric_label)
        row += 1

    set_col_widths(ws, [22, 9] + [9] * 12 + [10])


# ═══════════════════════════════════════════════════════════════════════════
# 시트 6-7: 적산온도 도달 — REF + VLOOKUP 동적 구조
#
# 설계:
#   _지역달성REF / _지점달성REF 숨긴 시트에
#   "도달일_{ms}℃_{subject}_{year}"  → "M-D" 문자열
#   "도달DOY_{ms}℃_{subject}_{year}" → 정수 DOY  (차이 계산용)
#   형태로 전 연도·전 base_year를 사전 저장.
#
#   표시 시트에서 VLOOKUP 키에 설정!$B$4 (올해) 를 참조하므로
#   올해 값을 바꾸면 올해/작년/평년/차이가 자동 갱신된다.
# ═══════════════════════════════════════════════════════════════════════════

def _store_reaching_to_ref(ws, row, ms_label, subject, reaching, normal_years, ref_year):
    """reaching 딕셔너리를 REF 시트에 기록 — 도달일 + DOY + 평년.

    Args:
        ws       : REF 워크시트
        row      : 현재 기록 행 번호
        ms_label : "500℃" 형식 문자열
        subject  : 권역명 또는 지점번호(str)
        reaching : {year: {milestone_int: "M-D"|None}}
        normal_years : 평년 기간
        ref_year : 기준 연도 (슬라이딩 윈도우 상한)

    Returns:
        int: 다음 빈 행 번호
    """
    max_base = ref_year + 5
    milestone = int(ms_label.replace("℃", ""))

    # ── 연도별 도달일 + DOY ──
    for year in sorted(reaching.keys()):
        date_str = reaching[year].get(milestone)

        ws.cell(row, 1, f"도달일_{ms_label}_{subject}_{year}")
        ws.cell(row, 2, date_str if date_str else "")
        row += 1

        ws.cell(row, 1, f"도달DOY_{ms_label}_{subject}_{year}")
        if date_str:
            try:
                m, d = map(int, date_str.split("-"))
                ws.cell(row, 2, date(year, m, d).timetuple().tm_yday)
            except (ValueError, OverflowError):
                ws.cell(row, 2, "")
        else:
            ws.cell(row, 2, "")
        row += 1

    # ── 평년(base_year별) 도달일 + DOY ──
    data_start = min(reaching.keys()) if reaching else 2016
    for base_year in range(data_start, max_base + 1):
        normal = temp_core.gdd_reaching_dates_normal(
            reaching, base_year, normal_years)
        date_str = normal.get(milestone)

        ws.cell(row, 1, f"도달일_{ms_label}_{subject}_평년{base_year}")
        ws.cell(row, 2, date_str if date_str else "")
        row += 1

        ws.cell(row, 1, f"도달DOY_{ms_label}_{subject}_평년{base_year}")
        if date_str:
            try:
                m, d = map(int, date_str.split("-"))
                ws.cell(row, 2, date(base_year, m, d).timetuple().tm_yday)
            except (ValueError, OverflowError):
                ws.cell(row, 2, "")
        else:
            ws.cell(row, 2, "")
        row += 1

    return row


def _build_reaching_ref_region(wb, region_data, milestones,
                                t_base, t_upper, method,
                                normal_years, ref_year):
    """숨긴 _지역달성REF 시트 생성."""
    ws = wb.create_sheet("_지역달성REF")
    ws.cell(1, 1, "키"); ws.cell(1, 2, "값")
    row = 2

    for region in stations.REGION_ORDER:
        reg_rows = region_data.get(region, [])
        if not reg_rows:
            continue
        avg_rows = _daily_average_rows(reg_rows)
        temp_core.add_gdd_to_rows(avg_rows, t_base, t_upper, method)
        reaching = temp_core.gdd_reaching_dates(
            avg_rows, t_base, t_upper, method, milestones)

        for ms in milestones:
            row = _store_reaching_to_ref(
                ws, row, f"{ms}℃", region, reaching, normal_years, ref_year)

    ws.sheet_state = "hidden"


def _build_reaching_ref_station(wb, all_parsed, milestones,
                                 t_base, t_upper, method,
                                 normal_years, ref_year):
    """숨긴 _지점달성REF 시트 생성 + 지점목록(D/E열) 저장."""
    ws = wb.create_sheet("_지점달성REF")
    ws.cell(1, 1, "키"); ws.cell(1, 2, "값")
    ws.cell(1, 4, "지점표시명"); ws.cell(1, 5, "지점번호")
    row = 2

    stn_list_row = 2
    first_display = ""
    for stn_id in sorted(all_parsed.keys()):
        display = f"{stations.station_name(stn_id)}({stn_id})"
        ws.cell(stn_list_row, 4, display)
        ws.cell(stn_list_row, 5, stn_id)
        if stn_list_row == 2:
            first_display = display
        stn_list_row += 1

    for stn_id in sorted(all_parsed.keys()):
        parsed = all_parsed[stn_id]
        if not parsed:
            continue
        temp_core.add_gdd_to_rows(parsed, t_base, t_upper, method)
        reaching = temp_core.gdd_reaching_dates(
            parsed, t_base, t_upper, method, milestones)

        for ms in milestones:
            row = _store_reaching_to_ref(
                ws, row, f"{ms}℃", str(stn_id), reaching, normal_years, ref_year)

    ws.sheet_state = "hidden"
    return stn_list_row - 2, first_display  # n_stations, first_display


def _write_reaching_vlookup_block(ws, row_start, subject_label, subject_key,
                                   milestones, header_row, ref_sheet):
    """올해/작년/평년/차이 5행 블록 — VLOOKUP 동적.

    VLOOKUP 키: "도달일_{ms}℃_{subject_key}_{year}"
    설정!$B$4 를 참조하므로 올해 변경 시 자동 갱신.
    차이 계산: DOY 정수를 직접 빼서 일수 차이 산출.
    """
    n_ms = len(milestones)

    # 소제목
    apply_title(ws, row_start, 1, f"  {subject_label}", merge_end_col=1, size=10)
    row = row_start + 1

    # ── 올해 / 작년 / 평년 ──
    compare_defs = [
        ('="올해("&설정!$B$4&")"',    "설정!$B$4"),
        ('="작년("&설정!$B$5&")"',    "설정!$B$5"),
        ('="평년("&(설정!$B$4-설정!$B$6)&"~"&설정!$B$5&")"',
         '"평년"&설정!$B$4'),
    ]
    row_refs = {}   # "올해" / "작년" / "평년" → 행 번호 (DOY VLOOKUP용)
    comp_names = ["올해", "작년", "평년"]

    for i, (label_f, year_f) in enumerate(compare_defs):
        comp = comp_names[i]
        row_refs[comp] = row

        cell_a = ws.cell(row, 1)
        cell_a.value = label_f
        cell_a.font = Font(name=FONT_NAME, size=10)
        cell_a.alignment = Alignment("left", "center")
        cell_a.border = BORDER

        for ci in range(2, n_ms + 2):
            cl = get_column_letter(ci)
            formula = (
                f'=IFERROR(VLOOKUP("도달일_"&{cl}${header_row}'
                f'&"_{subject_key}_"&{year_f},'
                f"'{ref_sheet}'!$A:$B,2,FALSE),\"-\")"
            )
            cell = ws.cell(row, ci)
            cell.value = formula
            cell.border = BORDER
            cell.alignment = Alignment("center", "center")
        row += 1

    # ── 차이 2행 ──
    diff_defs = [
        ("차이(올해-작년)",  "설정!$B$4",          "설정!$B$5"),
        ("차이(올해-평년)",  "설정!$B$4",          '"평년"&설정!$B$4'),
    ]

    for diff_label, year_a_f, year_b_f in diff_defs:
        C(ws, row, 1, diff_label, left=True)
        ws.cell(row, 1).border = BORDER

        for ci in range(2, n_ms + 2):
            cl = get_column_letter(ci)

            doy_a = (
                f'IFERROR(VLOOKUP("도달DOY_"&{cl}${header_row}'
                f'&"_{subject_key}_"&{year_a_f},'
                f"'{ref_sheet}'!$A:$B,2,FALSE),\"\")"
            )
            doy_b = (
                f'IFERROR(VLOOKUP("도달DOY_"&{cl}${header_row}'
                f'&"_{subject_key}_"&{year_b_f},'
                f"'{ref_sheet}'!$A:$B,2,FALSE),\"\")"
            )
            formula = (
                f'=IF(OR({doy_a}="",{doy_b}=""),"-",'
                f'{doy_a}-{doy_b})'
            )
            cell = ws.cell(row, ci)
            cell.value = formula
            cell.number_format = "+0;-0;0"
            cell.border = BORDER
            cell.alignment = Alignment("center", "center")
        row += 1

    return row  # 다음 빈 행


def _sheet_reaching_dates_region(wb, all_parsed, ref_year, normal_years,
                                  t_base, t_upper, method, milestones):
    """적산온도 도달(권역별) 시트."""
    # 권역별 데이터 묶기
    region_data = defaultdict(list)
    for stn_id, parsed in all_parsed.items():
        region = stations.region_of(stn_id) or "기타"
        region_data[region].extend(parsed)
    region_data["전국"] = [r for rows in region_data.values() for r in rows]

    # REF 시트 먼저 생성
    _build_reaching_ref_region(wb, region_data, milestones,
                                t_base, t_upper, method, normal_years, ref_year)

    ws = wb.create_sheet("적산온도 도달(권역별)")
    n_ms = len(milestones)

    apply_title(ws, 1, 1,
                f"적산온도 도달 일자 — 권역별 (Tbase={t_base}℃)",
                merge_end_col=n_ms + 1, size=12)
    _write_readonly_period(ws, 2, 1)

    HEADER_ROW = 3
    headers = ["구분"] + [f"{ms}℃" for ms in milestones]
    write_header_row(ws, HEADER_ROW, headers)
    set_col_widths(ws, [22] + [10] * n_ms)

    row = HEADER_ROW + 1
    for region in stations.REGION_ORDER:
        if region not in region_data:
            continue
        row = _write_reaching_vlookup_block(
            ws, row, region, region, milestones, HEADER_ROW, "_지역달성REF")
        row += 1  # 권역 간 빈 행


def _sheet_reaching_dates_station(wb, all_parsed, ref_year, normal_years,
                                   t_base, t_upper, method, milestones):
    """적산온도 도달(지점별) 시트."""
    # REF 시트 생성
    n_stations, first_display = _build_reaching_ref_station(
        wb, all_parsed, milestones,
        t_base, t_upper, method, normal_years, ref_year)

    ws = wb.create_sheet("적산온도 도달(지점별)")
    n_ms = len(milestones)

    apply_title(ws, 1, 1,
                f"적산온도 도달 일자 — 지점별 (Tbase={t_base}℃)",
                merge_end_col=n_ms + 1, size=12)
    _write_readonly_period(ws, 2, 1)

    HEADER_ROW = 3
    headers = ["구분"] + [f"{ms}℃" for ms in milestones]
    write_header_row(ws, HEADER_ROW, headers)
    set_col_widths(ws, [22] + [10] * n_ms)

    row = HEADER_ROW + 1
    for stn_id in sorted(all_parsed.keys()):
        stn_name = stations.station_name(stn_id)
        label = f"{stn_name}({stn_id})"
        row = _write_reaching_vlookup_block(
            ws, row, label, str(stn_id), milestones, HEADER_ROW, "_지점달성REF")
        row += 1  # 지점 간 빈 행


# ═══════════════════════════════════════════════════════════════════════════
# 시트 4: 기간별(월별) — solar 서식 (권역=열, 구분=행, 월별 블록)
# ═══════════════════════════════════════════════════════════════════════════

def _write_period_vlookup_block(ws, row_start, title, metric_label, region_cols):
    """기간별(월별) — solar 서식 VLOOKUP 블록.

    열: 구분 | 전국 | 서울경기 | ... | 제주
    행: 연간 올해/작년/평년/차이 → 1월~12월 각각 올해/작년/평년/차이
    """
    n_regions = len(region_cols)

    apply_title(ws, row_start, 1, title,
                merge_end_col=n_regions + 1, size=11)

    header_row = row_start + 1
    headers = ["구분"] + region_cols
    write_header_row(ws, header_row, headers)

    row = header_row + 1

    compare_rows = [
        ('="올해("&설정!$B$4&")"', '설정!$B$4'),
        ('="작년("&설정!$B$5&")"', '설정!$B$5'),
        ('="평년("&(설정!$B$4-설정!$B$6)&"~"&설정!$B$5&")"', '"평년"&설정!$B$4'),
    ]

    # ── 연간 섹션 ──
    row_refs = {}
    for i, (label_formula, year_ref) in enumerate(compare_rows):
        comp_name = ["올해", "작년", "평년"][i]
        row_refs[comp_name] = row

        cell_a = ws.cell(row, 1)
        cell_a.value = label_formula
        cell_a.font = Font(name=FONT_NAME, size=10)
        cell_a.alignment = Alignment("left", "center")
        cell_a.border = BORDER

        for ci in range(2, n_regions + 2):
            cl = get_column_letter(ci)
            formula = (
                f'=IFERROR(VLOOKUP("{metric_label}_"&{cl}${header_row}'
                f'&"_"&{year_ref}&"_연",'
                f"'_권역별REF'!$A:$B,2,FALSE),\"\")"
            )
            cell = ws.cell(row, ci)
            cell.value = formula
            cell.number_format = FMT_NUM1
            cell.border = BORDER
        row += 1

    this_r = row_refs["올해"]
    last_r = row_refs["작년"]
    normal_r = row_refs["평년"]

    for diff_label, ref_r in [("차이(전년)", last_r), ("차이(평년)", normal_r)]:
        C(ws, row, 1, diff_label, left=True)
        ws.cell(row, 1).border = BORDER
        for ci in range(2, n_regions + 2):
            cl = get_column_letter(ci)
            formula = (
                f'=IF(OR({cl}{this_r}="",{cl}{ref_r}=""),"",'
                f'{cl}{this_r}-{cl}{ref_r})')
            ws.cell(row, ci).value = formula
            ws.cell(row, ci).number_format = FMT_NUM1
            ws.cell(row, ci).border = BORDER
        row += 1

    # ── 월별 섹션 ──
    for month in range(1, 13):
        H(ws, row, 1, f"── {month}월 ──", fill=LIGHT_BLUE)
        for ci in range(2, n_regions + 2):
            H(ws, row, ci, "", fill=LIGHT_BLUE)
        row += 1

        m_row_refs = {}
        for i, (label_formula, year_ref) in enumerate(compare_rows):
            comp_name = ["올해", "작년", "평년"][i]
            m_row_refs[comp_name] = row

            cell_a = ws.cell(row, 1)
            cell_a.value = label_formula
            cell_a.font = Font(name=FONT_NAME, size=10)
            cell_a.alignment = Alignment("left", "center")
            cell_a.border = BORDER

            for ci in range(2, n_regions + 2):
                cl = get_column_letter(ci)
                formula = (
                    f'=IFERROR(VLOOKUP("{metric_label}_"&{cl}${header_row}'
                    f'&"_"&{year_ref}&"_{month}",'
                    f"'_권역별REF'!$A:$B,2,FALSE),\"\")"
                )
                cell = ws.cell(row, ci)
                cell.value = formula
                cell.number_format = FMT_NUM1
                cell.border = BORDER
            row += 1

        m_this = m_row_refs["올해"]
        m_last = m_row_refs["작년"]
        m_normal = m_row_refs["평년"]

        for diff_label, ref_r in [("차이(전년)", m_last), ("차이(평년)", m_normal)]:
            C(ws, row, 1, diff_label, left=True)
            ws.cell(row, 1).border = BORDER
            for ci in range(2, n_regions + 2):
                cl = get_column_letter(ci)
                formula = (
                    f'=IF(OR({cl}{m_this}="",{cl}{ref_r}=""),"",'
                    f'{cl}{m_this}-{cl}{ref_r})')
                ws.cell(row, ci).value = formula
                ws.cell(row, ci).number_format = FMT_NUM1
                ws.cell(row, ci).border = BORDER
            row += 1

    return row


def _sheet_period_analysis(wb, all_parsed, ref_year, normal_years,
                           start_date, end_date, t_base):
    ws = wb.create_sheet("기간별(월별)")

    _write_readonly_period(ws, 1, 1)

    region_cols = stations.REGION_ORDER

    metrics = [
        ("< 평균기온 (℃) >", "평균기온"),
        ("< 최고기온 (℃) >", "최고기온"),
        ("< 최저기온 (℃) >", "최저기온"),
        ("< 최고기온 평균 (℃) >", "최고기온평균"),
        ("< 최저기온 평균 (℃) >", "최저기온평균"),
        (f"< 적산온도 (℃·일, Tbase={t_base}℃) >", "적산온도"),
    ]

    row = 3
    for title_text, metric_label in metrics:
        row = _write_period_vlookup_block(
            ws, row, title_text, metric_label, region_cols,
        )
        row += 1

    set_col_widths(ws, [18] + [10] * len(region_cols))


# ═══════════════════════════════════════════════════════════════════════════
# 시트 5: 년도별 — 슬라이딩 윈도우 (solar 서식 동일)
# ═══════════════════════════════════════════════════════════════════════════

def _sheet_yearly_analysis(wb, all_parsed, ref_year, normal_years,
                           start_date, end_date, t_base):
    ws = wb.create_sheet("년도별")

    all_years = set()
    for parsed in all_parsed.values():
        for r in parsed:
            all_years.add(r["year"])
    data_start_year = min(all_years) if all_years else 2016

    _write_readonly_period(ws, 1, 1)

    metrics = [
        ("< 평균기온 (℃) >", "평균기온", "avg"),
        ("< 최고기온 (℃) >", "최고기온", "max"),
        ("< 최저기온 (℃) >", "최저기온", "min"),
        ("< 최고기온 평균 (℃) >", "최고기온평균", "avg"),
        ("< 최저기온 평균 (℃) >", "최저기온평균", "avg"),
        (f"< 적산온도 (℃·일, Tbase={t_base}℃) >", "적산온도", "cumsum"),
    ]

    row = 3
    for title_text, metric_label, agg_type in metrics:
        row = _write_yearly_vlookup_block(
            ws, row, title_text, metric_label, agg_type,
            data_start_year, ref_year,
        )
        row += 1

    set_col_widths(ws, [10, 10] + [9] * 12 + [10])


def _write_yearly_vlookup_block(ws, row_start, title, metric_label, agg_type,
                                data_start_year, ref_year):
    """년도별 — 슬라이딩 윈도우 VLOOKUP 블록.

    열 구조: A=권역, B=연도, C~N=1월~12월, O=연간
    - 올해-평년기간 ~ 올해 범위만 표시 (11개년)
    - 미래 대응: ref_year + 5까지 행을 미리 생성
    - 조건: OR(year < 설정!$B$4-설정!$B$6, year > 설정!$B$4) 이면 빈칸
    """
    summary_fn = "SUM" if agg_type == "cumsum" else "AVERAGE"

    apply_title(ws, row_start, 1, title, merge_end_col=15, size=11)
    # A=권역, B=연도, C~N=1월~12월, O=연간
    headers = ["권역", "연도"] + [f"{m}월" for m in range(1, 13)] + ["연간"]
    write_header_row(ws, row_start + 1, headers)

    row = row_start + 2
    max_year = ref_year + 5
    hide_cond = 'OR({y}<설정!$B$4-설정!$B$6,{y}>설정!$B$4)'

    for region in stations.REGION_ORDER:
        for year in range(data_start_year, max_year + 1):
            cond = hide_cond.format(y=year)

            # A열: 권역 (윈도우 밖이면 빈칸)
            cell_a = ws.cell(row, 1)
            cell_a.value = f'=IF({cond},"","{region}")'
            cell_a.font = Font(name=FONT_NAME, size=10)
            cell_a.alignment = Alignment("left", "center")
            cell_a.border = BORDER

            # B열: 연도 (윈도우 밖이면 빈칸)
            cell_b = ws.cell(row, 2)
            cell_b.value = f'=IF({cond},"",{year})'
            cell_b.number_format = '0'
            cell_b.font = Font(name=FONT_NAME, size=10)
            cell_b.alignment = Alignment("right", "center")
            cell_b.border = BORDER

            # C~N열(3~14): 1월~12월
            for mi, m in enumerate(range(1, 13)):
                ci = mi + 3  # C=3
                formula = (
                    f'=IF({cond},"",IFERROR(VLOOKUP("{metric_label}_{region}_"'
                    f'&$B{row}&"_{m}",'
                    f"'_권역별REF'!$A:$B,2,FALSE),\"\"))"
                )
                cell = ws.cell(row, ci)
                cell.value = formula
                cell.number_format = FMT_NUM1
                cell.border = BORDER

            # O열(15): 연간
            formula_o = (
                f'=IF({cond},"",IF(COUNTA(C{row}:N{row})=0,"",'
                f'{summary_fn}(C{row}:N{row})))'
            )
            cell_o = ws.cell(row, 15)
            cell_o.value = formula_o
            cell_o.number_format = FMT_NUM1
            cell_o.border = BORDER
            row += 1

        # 평년 행
        C(ws, row, 1, region, left=True)
        ws.cell(row, 1).border = BORDER

        cell_b = ws.cell(row, 2)
        cell_b.value = '="평년("&(설정!$B$4-설정!$B$6)&"~"&설정!$B$5&")"'
        cell_b.font = Font(name=FONT_NAME, size=10)
        cell_b.border = BORDER

        # C~N열: 1월~12월
        for mi, m in enumerate(range(1, 13)):
            ci = mi + 3  # C=3
            formula = (
                f'=IFERROR(VLOOKUP("{metric_label}_{region}_"'
                f'&"평년"&설정!$B$4&"_{m}",'
                f"'_권역별REF'!$A:$B,2,FALSE),\"\")"
            )
            ws.cell(row, ci).value = formula
            ws.cell(row, ci).number_format = FMT_NUM1
            ws.cell(row, ci).border = BORDER

        # O열: 연간
        formula_o = (
            f'=IF(COUNTA(C{row}:N{row})=0,"",'
            f'{summary_fn}(C{row}:N{row}))'
        )
        ws.cell(row, 15).value = formula_o
        ws.cell(row, 15).number_format = FMT_NUM1
        ws.cell(row, 15).border = BORDER
        row += 1
        row += 1  # 권역 간 빈 행

    return row


# ═══════════════════════════════════════════════════════════════════════════
# 메인: 전체 프로세스
# ═══════════════════════════════════════════════════════════════════════════

def generate_report(
    stn_ids=None, regions=None,
    start_date=None, end_date=None,
    ref_year=None, normal_years=10,
    t_base=None, t_upper=None, method="maxmin",
    crop_id=None, milestones=None,
    output_path=None, use_cache=True, verbose=True,
):
    """온도·적산온도 종합 분석 엑셀을 생성한다."""
    if start_date is None:
        start_date = solar_core.DATA_START
    if end_date is None:
        end_date = solar_core.yesterday()
    if ref_year is None:
        ref_year = date.today().year

    # 작물 지정 시 파라미터 자동 로드
    if crop_id:
        try:
            lib = temp_core.load_crops_gdd()
            crop = lib[crop_id]
            if t_base is None:
                t_base = crop["t_base"]
            if t_upper is None:
                t_upper = crop["t_upper"]
            if milestones is None:
                milestones = crop["milestones"]
        except (FileNotFoundError, KeyError):
            print(f"⚠ 작물 '{crop_id}' 정보를 찾을 수 없습니다. 기본값 사용.")

    if t_base is None:
        t_base = temp_core.DEFAULT_T_BASE
    if t_upper is None:
        t_upper = temp_core.DEFAULT_T_UPPER
    if milestones is None:
        milestones = temp_core.DEFAULT_MILESTONES

    # 대상 지점
    if stn_ids:
        target_stations = stn_ids
    elif regions:
        target_stations = []
        for region in regions:
            target_stations.extend(stations.stations_in_region(region))
    else:
        target_stations = stations.all_station_ids()

    target_regions = regions or stations.REGION_ORDER

    if verbose:
        print(f"=== 온도·적산온도 종합 분석 ===")
        print(f"기간: {start_date} ~ {end_date}")
        print(f"기준 연도: {ref_year}")
        print(f"평년: {ref_year - normal_years} ~ {ref_year - 1}")
        print(f"기준온도: {t_base}℃  상한온도: {t_upper}℃")
        print(f"계산 방법: {method}")
        if crop_id:
            print(f"작물: {crop_id}")
        print(f"대상 지점: {len(target_stations)}개")
        print()

    # 인증키
    keys = solar_core.load_apikeys()
    service_key = keys.get("DATA_GO_KR")
    if not service_key:
        raise ValueError("apikey.txt에 DATA_GO_KR 키가 없습니다.")

    # 데이터 수집
    raw_data = solar_core.fetch_all_stations(
        service_key, target_stations, start_date, end_date,
        use_cache=use_cache, verbose=verbose,
    )

    # 파싱 + GDD 계산
    all_parsed = {}
    total_rows = 0
    for stn_id, rows in raw_data.items():
        parsed = solar_core.parse_rows(rows)
        temp_core.add_gdd_to_rows(parsed, t_base, t_upper, method)
        all_parsed[stn_id] = parsed
        total_rows += len(parsed)

    if verbose:
        print(f"\n총 {total_rows}건 파싱 완료")
        print("엑셀 생성 중...")

    # 엑셀 생성
    wb = Workbook()

    # 1. 설정
    _sheet_settings(wb, start_date, end_date, ref_year, normal_years,
                    target_regions, target_stations,
                    t_base, t_upper, method, crop_id)

    # 2. 권역별 집계
    _sheet_region_summary(wb, all_parsed, ref_year, normal_years,
                          start_date, end_date, target_regions,
                          t_base, t_upper, method)

    # 3. 지점별 집계
    _sheet_station_summary(wb, all_parsed, ref_year, normal_years,
                           start_date, end_date,
                           t_base, t_upper, method)

    # 4. 기간별(월별)
    _sheet_period_analysis(wb, all_parsed, ref_year, normal_years,
                           start_date, end_date, t_base)

    # 5. 년도별
    _sheet_yearly_analysis(wb, all_parsed, ref_year, normal_years,
                           start_date, end_date, t_base)

    # 6. 적산온도 도달(권역별)
    _sheet_reaching_dates_region(wb, all_parsed, ref_year, normal_years,
                                  t_base, t_upper, method, milestones)

    # 7. 적산온도 도달(지점별)
    _sheet_reaching_dates_station(wb, all_parsed, ref_year, normal_years,
                                   t_base, t_upper, method, milestones)

    # 8. 원데이터 (맨 뒤)
    _sheet_raw_data(wb, all_parsed, t_base, method)

    # 저장
    if output_path is None:
        output_dir = Path(__file__).parent / "output"
        output_dir.mkdir(exist_ok=True)
        crop_tag = f"_{crop_id}" if crop_id else f"_tbase{int(t_base)}"
        output_path = str(
            output_dir / f"temp_analysis_{ref_year}{crop_tag}_{end_date.strftime('%Y%m%d')}.xlsx"
        )

    wb.save(output_path)
    if verbose:
        print(f"\n✓ 저장 완료: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="온도·적산온도 종합 분석 엑셀 생성",
    )
    parser.add_argument("--region", nargs="+", default=None,
                        help="분석 대상 권역 (예: 서울경기 충남)")
    parser.add_argument("--station", nargs="+", type=int, default=None,
                        help="분석 대상 지점번호 (예: 108 131)")
    parser.add_argument("--year", type=int, default=None,
                        help="기준 연도 (기본: 올해)")
    parser.add_argument("--normal-years", type=int, default=10,
                        help="평년 산정 기간 (기본: 10년)")
    parser.add_argument("--crop", default=None,
                        help="작물 ID (crops_gdd.csv 참조, 예: apple)")
    parser.add_argument("--tbase", type=float, default=None,
                        help="기준온도 직접 지정 (--crop 보다 우선)")
    parser.add_argument("--tupper", type=float, default=None,
                        help="상한온도 직접 지정")
    parser.add_argument("--method", default="maxmin",
                        choices=["maxmin", "avg", "aquacrop"],
                        help="GDD 계산 방법 (기본: maxmin)")
    parser.add_argument("--output", "-o", default=None,
                        help="출력 파일 경로")
    parser.add_argument("--no-cache", action="store_true",
                        help="CSV 캐시 사용하지 않음")
    args = parser.parse_args()

    generate_report(
        stn_ids=args.station,
        regions=args.region,
        ref_year=args.year,
        normal_years=args.normal_years,
        crop_id=args.crop,
        t_base=args.tbase,
        t_upper=args.tupper,
        method=args.method,
        output_path=args.output,
        use_cache=not args.no_cache,
    )


if __name__ == "__main__":
    main()
