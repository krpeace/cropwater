#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fao56_core.py — ETo 산출 스크립트들의 공통 모듈

eto_singlestation.py, eto_multistation.py 가 함께 참조하는 "기준 소스".
반드시 같은 폴더에 이 파일이 있어야 두 스크립트가 정상 동작한다.

포함 내용
  ① FAO-56 Penman-Monteith / 증발접시법 물리 계산 함수
     (다지점 스크립트가 그대로 호출. 단일지점 스크립트는 같은 식을
      엑셀 라이브 수식으로 재작성해 쓰지만, 물리식의 기준은 여기.)
  ② 공통 유틸: 숫자 파싱(num), ASOS 일자료 조회(fetch_asos),
     지점정보 조회(fetch_station_table), 인증키 로딩(load_apikeys)

주의: 이 파일의 계산식을 고칠 일이 생기면(단위 오류 등) 여기 한 곳만
고치면 된다. 단, eto_singlestation.py의 엑셀 수식 문자열은 별도이므로
물리식을 바꿀 경우 그쪽 수식도 함께 확인할 것(설명 시트의 [식n] 표기 참조).
"""
import math, os, csv, datetime as dt
import requests

# ── FAO-56 물리 계산 (검증 완료: FAO-56 교과서 예제와 일치) ──
def svp(T):
    """포화수증기압 e°(T) [kPa], FAO-56 Eq.11"""
    return 0.6108 * math.exp(17.27 * T / (T + 237.3))

def slope_svp(T):
    """포화수증기압곡선 기울기 Δ [kPa/℃], FAO-56 Eq.13"""
    return 4098 * svp(T) / (T + 237.3) ** 2

def extra_radiation(lat_deg, J):
    """천문일사 Ra [MJ/m²/d], FAO-56 Eq.21"""
    Gsc = 0.0820; phi = math.radians(lat_deg)
    dr = 1 + 0.033 * math.cos(2 * math.pi * J / 365)
    dec = 0.409 * math.sin(2 * math.pi * J / 365 - 1.39)
    ws = math.acos(max(-1, min(1, -math.tan(phi) * math.tan(dec))))
    return (24 * 60 / math.pi) * Gsc * dr * (
        ws * math.sin(phi) * math.sin(dec) + math.cos(phi) * math.cos(dec) * math.sin(ws))

def daylight_hours(lat_deg, J):
    """가조시간 N [hr], FAO-56 Eq.34"""
    phi = math.radians(lat_deg); dec = 0.409 * math.sin(2 * math.pi * J / 365 - 1.39)
    ws = math.acos(max(-1, min(1, -math.tan(phi) * math.tan(dec))))
    return 24 / math.pi * ws

def wind_2m(uz, zw=10.0):
    """풍속을 z m → 2 m 높이로 환산, FAO-56 Eq.47"""
    return uz * 4.87 / math.log(67.8 * zw - 5.42)

def kp_class_a(u2, RHmean, fetch=100.0):
    """Class A 증발접시 계수 Kp (녹지 fetch, Allen-Pruitt), FAO-56 Table 8 식"""
    F = math.log(fetch)
    return (0.108 - 0.0286 * u2 + 0.0422 * F + 0.1434 * math.log(RHmean)
            - 0.000631 * (F ** 2) * math.log(RHmean))

def eto_penman_monteith(Tmax, Tmin, Rs, u2, ea, elev, lat, J, Pa=None):
    """FAO-56 Eq.6 (일 단위, 지중열류 G=0 가정)
       Pa: 현지기압(hPa). None이면 고도로부터 표준대기압 추정(Eq.7)."""
    T = (Tmax + Tmin) / 2
    P = (Pa * 0.1 if Pa else 101.3 * ((293 - 0.0065 * elev) / 293) ** 5.26)  # kPa
    g = 0.000665 * P
    D = slope_svp(T)
    es = (svp(Tmax) + svp(Tmin)) / 2
    Ra = extra_radiation(lat, J)
    Rso = (0.75 + 2e-5 * elev) * Ra
    Rns = 0.77 * Rs
    Rnl = 4.903e-9 * (((Tmax + 273.16) ** 4 + (Tmin + 273.16) ** 4) / 2) \
          * (0.34 - 0.14 * math.sqrt(max(ea, 0))) * (1.35 * min(Rs / Rso, 1.0) - 0.35)
    Rn = Rns - Rnl
    return (0.408 * D * Rn + g * (900 / (T + 273)) * u2 * (es - ea)) / (D + g * (1 + 0.34 * u2))

# ── FAO-56 작물계수(Kc) 시스템 ──
# Table 12(단일 Kc, Apples/Cherries/Pears)와 Table 17(기본 Kcb, 증산만)의 4개 시나리오.
# 축1: 청경(no active ground cover) vs 초생재배(active ground cover)
# 축2: 서리 발생지역(killing frost) vs 서리 없는 지역
KC_SCENARIOS = {                                    # (Kc_ini, Kc_mid, Kc_end, 설명)
    1: (0.45, 0.95, 0.70, "청경(초생 없음), 동해서리 발생지역"),
    2: (0.60, 0.95, 0.75, "청경, 서리 없는 지역"),
    3: (0.50, 1.20, 0.95, "초생재배(active ground cover), 동해서리 발생지역 ← 한국 사과원 일반"),
    4: (0.80, 1.20, 0.85, "초생재배, 서리 없는 지역"),
}
KCB_SCENARIOS = {                                   # Table 17: 기본작물계수(증산만, 토양증발 제외)
    1: (0.35, 0.90, 0.65, "청경, 서리, 불투수 전면멀칭 과원에 권장"),
    2: (0.50, 0.90, 0.70, "청경, 서리 없는 지역"),
    3: (0.45, 1.15, 0.90, "초생재배, 동해서리 발생지역"),
    4: (0.75, 1.15, 0.80, "초생재배, 서리 없는 지역"),
}

def kc_climate_adjust(kc_table, u2_mean, rhmin_mean, h_m, is_end=False):
    """FAO-56 식(62)/(65) 기후보정.
       kc_table: 표값(Kc_mid 또는 Kc_end).  u2/rhmin: 해당 스테이지 평균.  h_m: 생육중기 초목 수고(m).
       is_end=True면 kc_table>0.45일 때만 보정(식 65). u2는 [1,6], RHmin은 [20,80]으로 자동 제한.
       반환: 기후보정된 Kc."""
    if is_end and kc_table <= 0.45:
        return kc_table
    u2c = max(1.0, min(6.0, u2_mean))
    rhc = max(20.0, min(80.0, rhmin_mean))
    return kc_table + (0.04*(u2c-2.0) - 0.004*(rhc-45.0)) * (h_m/3.0)**0.3

def stage_of_date(day_ordinal, start_ordinal, L_ini, L_dev, L_mid, L_late):
    """일자 → 생육단계 문자열. 발아일 이전은 '발아전', 이후는 초기/발육/중기/후기/후기이후."""
    d = day_ordinal - start_ordinal   # 발아일부터의 경과일(0=발아일)
    if d < 0: return "발아전"
    if d < L_ini: return "초기"
    if d < L_ini + L_dev: return "발육"
    if d < L_ini + L_dev + L_mid: return "중기"
    if d < L_ini + L_dev + L_mid + L_late: return "후기"
    return "후기이후"

def kc_of_date(day_ordinal, start_ordinal, L_ini, L_dev, L_mid, L_late,
               kc_ini, kc_mid, kc_end):
    """FAO-56 식(66) 일별 Kc 곡선 보간.
       초기: Kc_ini 유지. 발육: Kc_ini→Kc_mid 선형. 중기: Kc_mid 유지. 후기: Kc_mid→Kc_end 선형.
       발아 전/후기 이후: 0 (관수요구량 없음)."""
    d = day_ordinal - start_ordinal
    if d < 0: return 0.0
    if d < L_ini: return kc_ini
    if d < L_ini + L_dev:
        return kc_ini + (kc_mid - kc_ini) * (d - L_ini + 1) / L_dev
    if d < L_ini + L_dev + L_mid: return kc_mid
    if d < L_ini + L_dev + L_mid + L_late:
        return kc_mid + (kc_end - kc_mid) * (d - L_ini - L_dev - L_mid + 1) / L_late
    return 0.0

# ── 공통 유틸 ──
def num(x):
    """문자열/None을 안전하게 float로. 결측 표기('-', '', None 등)는 None 반환."""
    try:
        x = str(x).strip()
        return float(x) if x not in ("", "-", "null", "None") else None
    except Exception:
        return None

def fetch_asos(key, stn, start, end):
    """기상청 ASOS 일자료 조회 (공공데이터포털, getWthrDataList). key=DATA_GO_KR 키."""
    import datetime as _dt
    # 종료일이 오늘 이후이면 어제로 자동 클리핑 (ASOS는 전날까지만 제공)
    try:
        d0 = _dt.datetime.strptime(str(start), "%Y%m%d")
        d1 = _dt.datetime.strptime(str(end),   "%Y%m%d")
        yesterday = _dt.datetime.combine(_dt.date.today() - _dt.timedelta(days=1),
                                          _dt.time.min)
        if d1 > yesterday:
            d1 = yesterday
            print(f"  [주의] 종료일을 어제({d1.strftime('%Y%m%d')})로 자동 조정 (ASOS는 전날까지만 제공)")
        end = d1.strftime("%Y%m%d")
        num_rows = max(400, (d1 - d0).days + 11)
    except Exception:
        num_rows = 400

    base = "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"
    url = (f"{base}?serviceKey={key}&pageNo=1&numOfRows={num_rows}&dataType=JSON"
           f"&dataCd=ASOS&dateCd=DAY&startDt={start}&endDt={end}&stnIds={stn}")
    r = requests.get(url, timeout=30); r.raise_for_status()
    js = r.json()
    resp = js.get("response", {})
    hdr  = resp.get("header", {})
    rc   = hdr.get("resultCode", "?")
    msg  = hdr.get("resultMsg",  "?")
    if rc != "00":
        raise RuntimeError(f"API 오류 [{rc}] {msg}\nURL: {url}")
    body = resp.get("body")
    if body is None:
        raise RuntimeError(f"API 응답에 body 없음. header={hdr}\n원문: {r.text[:300]}")
    items = body.get("items") or {}
    it = items.get("item")
    if it is None:
        total = body.get("totalCount", 0)
        if total == 0:
            return []
        raise RuntimeError(f"items.item 없음. body={body}")
    return it if isinstance(it, list) else [it]

def fetch_station_table(hubkey):
    """기상청 API허브 지점일람표(stn_inf.php) 조회. key=KMA_HUB 키.
       반환: (info: {지점번호: (위도,경도,고도,풍속계높이)}, names: {지점번호: 한글명})"""
    url = f"https://apihub.kma.go.kr/api/typ01/url/stn_inf.php?inf=SFC&stn=&tm=202211300900&authKey={hubkey}"
    r = requests.get(url, timeout=30); r.raise_for_status()
    text = r.content.decode("euc-kr", "replace")
    info, names = {}, {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        p = line.split()
        if len(p) < 8:
            continue
        try:
            stn = int(p[0]); lon = float(p[1]); lat = float(p[2]); ht = float(p[4]); htwd = float(p[7])
        except Exception:
            continue
        info[stn] = (lat, lon, ht, htwd)
        if len(p) >= 11:
            names[stn] = p[10]
    return info, names

# ── 관측지점 메타정보 백업파일 (stations_backup.csv) ──
# 위도·고도·풍속계높이는 자주 바뀌지 않으므로, API로 한 번 조회에 성공하면 로컬 CSV에
# 저장해두고, 다음에 API허브 연결이 실패해도 이 백업파일에서 읽어 쓴다.
# 파일이 없으면 검증된 8개 지점(2026-09 기준 실측)으로 자동 생성(부트스트랩)한다.
STATION_BACKUP_DEFAULT = "stations_backup.csv"
_SEED_STATIONS = {   # 최초 부트스트랩용 시드 데이터 (stn: (name, lat, lon, elev, anem_height))
    101: ("춘천", 37.90262, 127.73570, 75.82, 10.0),
    93:  ("북춘천", 37.94738, 127.75443, 95.78, 10.0),
    216: ("태백", 37.17038, 128.98929, 714.45, 16.0),   # 풍속계 16m — 표준 10m 아님
    119: ("수원", 37.25746, 126.98300, 39.81, 18.7),    # 풍속계 18.7m — 표준 10m 아님
    232: ("천안", 36.76217, 127.29282, 84.78, 10.0),
    127: ("충주", 36.97045, 127.95250, 114.85, 10.0),
    136: ("안동", 36.57293, 128.70733, 141.26, 10.0),
    192: ("진주", 35.16378, 128.04004, 29.35, 10.0),
}

def _ensure_backup_file(fname):
    if not os.path.exists(fname):
        with open(fname, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["stn", "name", "lat", "lon", "elev", "anem_height"])
            for stn, (name, lat, lon, elev, anem) in _SEED_STATIONS.items():
                w.writerow([stn, name, lat, lon, elev, anem])

def load_station_backup(fname=STATION_BACKUP_DEFAULT):
    """로컬 백업파일(csv)에서 지점 메타정보를 읽는다. 파일이 없으면 시드 데이터로 새로 만든 뒤 읽는다.
       반환: (info: {지점번호: (위도,경도,고도,풍속계높이)}, names: {지점번호: 이름})"""
    _ensure_backup_file(fname)
    info, names = {}, {}
    with open(fname, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            try:
                stn = int(row["stn"])
                info[stn] = (float(row["lat"]), float(row["lon"]), float(row["elev"]), float(row["anem_height"]))
                if row.get("name"):
                    names[stn] = row["name"]
            except Exception:
                continue
    return info, names

def save_station_backup(new_info, new_names=None, fname=STATION_BACKUP_DEFAULT):
    """API 조회 성공 결과를 백업파일에 병합 저장(기존 지점은 갱신, 새 지점은 추가).
       다지점 스크립트가 한 번 성공하면 수백 개 지점이 한꺼번에 채워질 수 있다."""
    new_names = new_names or {}
    info, names = load_station_backup(fname)   # 기존 내용(시드 포함) 로드 후 병합
    for stn, meta in new_info.items():
        info[stn] = meta
        if stn in new_names:
            names[stn] = new_names[stn]
    with open(fname, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["stn", "name", "lat", "lon", "elev", "anem_height"])
        for stn in sorted(info):
            lat, lon, elev, anem = info[stn]
            w.writerow([stn, names.get(stn, ""), lat, lon, elev, anem])

# ── 작물 라이브러리 (2계층: 기본 시드 + 사용자 오버라이드) ──
# 기본 시드(crops_library.csv): FAO-56 Table 11/12에서 뽑은 카테고리별 작물 세트.
#   사용자는 이 파일을 편집하지 않고 그대로 둘 것(업데이트 시 통째로 교체 가능).
# 사용자 오버라이드(crops_overrides.csv): 실측 검증 결과로 특정 작물의 값을 덮어쓸 때 사용.
#   같은 crop_id가 있으면 오버라이드가 이깁니다. 없는 crop_id면 새 작물 추가로 취급.
# 두 파일 모두 같은 컬럼 스키마를 씁니다(crops_library.csv 헤더 참조).
CROPS_LIBRARY_DEFAULT = "crops_library.csv"
CROPS_OVERRIDES_DEFAULT = "crops_overrides.csv"

def _parse_crop_row(row):
    """CSV 한 행을 dict로 파싱. 빈 값·오류는 None으로."""
    def _f(k):
        v = row.get(k, "").strip()
        try: return float(v) if v else None
        except: return None
    def _i(k):
        v = row.get(k, "").strip()
        try: return int(float(v)) if v else None
        except: return None
    d = {
        "crop_id": row.get("crop_id", "").strip(),
        "name_ko": row.get("name_ko", "").strip(),
        "name_en": row.get("name_en", "").strip(),
        "category": row.get("category", "").strip(),
        "category_order": _i("category_order") or 99,
        "has_scenarios": bool(_i("has_scenarios") or 0),
        "L_ini": _i("L_ini"), "L_dev": _i("L_dev"),
        "L_mid": _i("L_mid"), "L_late": _i("L_late"),
        "h_m": _f("h_m"),
        "zr": _f("zr"),       # 근권심도 (m), FAO-56 Table 22
        "p": _f("p"),          # 토양수분고갈계수, FAO-56 Table 22
        "note": row.get("note", "").strip(),
    }
    # 시나리오 1~4의 Kc 3값(있는 것만)
    scenarios = []
    for i in range(1, 5):
        suf = "" if i == 1 else f"_{i}"
        ki, km, ke = _f(f"kc_ini{suf}"), _f(f"kc_mid{suf}"), _f(f"kc_end{suf}")
        if None not in (ki, km, ke):
            scenarios.append((ki, km, ke))
    d["scenarios"] = scenarios
    return d

def load_crop_library(fname=CROPS_LIBRARY_DEFAULT, overrides_fname=CROPS_OVERRIDES_DEFAULT):
    """작물 라이브러리 로드. 기본 시드 위에 오버라이드를 얹어 병합해서 반환.
       반환: {crop_id: crop_dict}. category_order·crop_id 순으로 정렬 원하면 호출측에서."""
    lib = {}
    if os.path.exists(fname):
        with open(fname, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                d = _parse_crop_row(row)
                if d["crop_id"]:
                    lib[d["crop_id"]] = d
    if os.path.exists(overrides_fname):
        with open(overrides_fname, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                d = _parse_crop_row(row)
                if d["crop_id"]:
                    lib[d["crop_id"]] = d      # 오버라이드가 이김(같은 crop_id면 덮어씀)
    return lib

def crops_sorted(lib):
    """카테고리 우선순위 순 → 같은 카테고리 안에서는 CSV 정의 순서 유지."""
    return sorted(lib.values(), key=lambda x: (x["category_order"], x["crop_id"]))

def save_crop_override(crop_dict, fname=CROPS_OVERRIDES_DEFAULT):
    """사용자가 특정 작물 값을 실측 결과로 덮어쓸 때 호출.
       기존 오버라이드에 같은 crop_id가 있으면 갱신, 없으면 추가."""
    existing = {}
    if os.path.exists(fname):
        with open(fname, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                cid = row.get("crop_id", "").strip()
                if cid: existing[cid] = row
    # 새 행 만들기(기본 스키마 순서)
    scenarios = crop_dict.get("scenarios", [])
    new_row = {"crop_id": crop_dict["crop_id"], "name_ko": crop_dict.get("name_ko", ""),
               "name_en": crop_dict.get("name_en", ""), "category": crop_dict.get("category", ""),
               "category_order": crop_dict.get("category_order", 99),
               "has_scenarios": 1 if len(scenarios) > 1 else 0}
    for i in range(1, 5):
        suf = "" if i == 1 else f"_{i}"
        if i <= len(scenarios):
            new_row[f"kc_ini{suf}"], new_row[f"kc_mid{suf}"], new_row[f"kc_end{suf}"] = scenarios[i-1]
        else:
            new_row[f"kc_ini{suf}"] = new_row[f"kc_mid{suf}"] = new_row[f"kc_end{suf}"] = ""
    for k in ("L_ini", "L_dev", "L_mid", "L_late", "h_m", "zr", "p"):
        new_row[k] = crop_dict.get(k, "")
    new_row["note"] = crop_dict.get("note", "")
    existing[crop_dict["crop_id"]] = new_row
    # 저장
    cols = ["crop_id","name_ko","name_en","category","category_order","has_scenarios",
            "kc_ini","kc_mid","kc_end","kc_ini_2","kc_mid_2","kc_end_2",
            "kc_ini_3","kc_mid_3","kc_end_3","kc_ini_4","kc_mid_4","kc_end_4",
            "L_ini","L_dev","L_mid","L_late","h_m","zr","p","note"]
    with open(fname, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols); w.writeheader()
        for row in existing.values(): w.writerow(row)

def load_apikeys(fname="apikey.txt"):
    """apikey.txt(라벨=값 형식)에서 인증키를 읽는다. 형식 예:
         DATA_GO_KR=<공공데이터포털 키, 기상데이터용>
         KMA_HUB=<기상청 API허브 키, 지점정보용>
       파일이 없거나 라벨이 없으면 빈 dict."""
    d = {}
    if os.path.exists(fname):
        for line in open(fname, encoding="utf-8"):
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            d[k.strip().upper()] = v.strip()
    return d

def safe_save(wb, path):
    """엑셀 파일이 이미 열려 있어(Windows 권한 오류) 저장이 막히면,
       파일명 뒤에 시각을 붙여 자동으로 다른 이름으로 저장한다.
       반환값: 실제로 저장된 경로."""
    try:
        wb.save(path)
        return path
    except PermissionError:
        base, ext = os.path.splitext(path)
        alt = f"{base}_{dt.datetime.now().strftime('%H%M%S')}{ext}"
        print(f"  [경고] '{path}' 파일이 열려 있어 저장할 수 없습니다.")
        print(f"         → 대신 '{alt}' 로 저장합니다. (기존 파일을 닫고 다시 실행하면 원래 이름으로 저장됩니다)")
        wb.save(alt)
        return alt
