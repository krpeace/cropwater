"""
solar_core.py — ASOS 일사·일조 데이터 수집 및 분석 핵심 모듈

기능:
  1. 공공데이터포털 ASOS 일자료 API 호출 (10년 자동 분할)
  2. CSV 파일 캐시 (과거 확정 데이터 재조회 방지)
  3. 일사·일조 지표 계산 (적산일사, 평균일사, 일조율 등)
  4. 평년(10년 평균) 비교 분석

데이터 소스:
  - 기상청 ASOS 일자료 (공공데이터포털)
  - https://www.data.go.kr/data/15059093/openapi.do
"""

import os
import csv
import math
import time
import requests
from datetime import date, timedelta, datetime
from pathlib import Path

import stations

# ---------------------------------------------------------------------------
# 상수
# ---------------------------------------------------------------------------
DATA_START = date(2016, 1, 1)          # 수집 시작일
MAX_QUERY_YEARS = 2                     # API 1회 최대 조회 기간(년) — 999건 제한 고려
MAX_NUM_OF_ROWS = 999                   # API 1회 최대 응답 건수
ASOS_DAILY_URL = (
    "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"
)

# CSV 캐시 디렉터리
CACHE_DIR = Path(__file__).parent / "data" / "raw"

# 일사·일조 관련 API 응답 필드
SOLAR_FIELDS = [
    "tm",           # 일시 (YYYY-MM-DD)
    "stnId",        # 지점번호
    "stnNm",        # 지점명
    "sumGsr",       # 합계 일사량 (MJ/m²)
    "sumSsHr",      # 합계 일조시간 (hr)
    "ssDur",        # 가조시간 — 가능일조시간 (hr)
    "hr1MaxIcsr",   # 1시간 최다 일사량 (MJ/m²)
    "hr1MaxIcsrHrmt",  # 1시간 최다 일사량 시각
    "avgTa",        # 평균기온 (℃) — 참고용
    "maxTa",        # 최고기온 (℃)
    "minTa",        # 최저기온 (℃)
    "sumRn",        # 일강수량 (mm) — 참고용
]


# ---------------------------------------------------------------------------
# API 인증키 로딩
# ---------------------------------------------------------------------------
def load_apikeys(fname: str = "apikey.txt") -> dict[str, str]:
    """apikey.txt(라벨=값 형식)에서 인증키를 읽는다.

    파일 형식 예시:
        DATA_GO_KR=인코딩된서비스키
        KMA_HUB=API허브키
    """
    d = {}
    paths = [
        Path(fname),
        Path(__file__).parent / fname,
    ]
    for p in paths:
        if p.exists():
            for line in p.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                d[k.strip().upper()] = v.strip()
            break
    if not d:
        raise FileNotFoundError(
            f"인증키 파일을 찾을 수 없습니다: {fname}\n"
            "apikey.txt.example을 복사하여 apikey.txt를 만들고 "
            "공공데이터포털 인증키를 입력하세요."
        )
    return d


# ---------------------------------------------------------------------------
# 날짜 유틸
# ---------------------------------------------------------------------------
def yesterday() -> date:
    """D-1 (전일) 날짜 반환"""
    return date.today() - timedelta(days=1)


def get_date_chunks(
    start_date: date, end_date: date, max_years: int = MAX_QUERY_YEARS
) -> list[tuple[date, date]]:
    """날짜 범위를 max_years 단위로 자동 분할.

    예) 2016-01-01 ~ 2026-09-14, max_years=10
        → [(2016-01-01, 2025-12-31), (2026-01-01, 2026-09-14)]
    """
    chunks = []
    current = start_date
    while current <= end_date:
        chunk_end = min(
            date(current.year + max_years - 1, 12, 31),
            end_date,
        )
        chunks.append((current, chunk_end))
        current = date(current.year + max_years, 1, 1)
    return chunks


def is_finalized_year(year: int) -> bool:
    """해당 연도의 데이터가 확정(변경 불가)인지 판별.

    올해 데이터는 매일 갱신되므로 확정이 아니다.
    """
    return year < date.today().year


# ---------------------------------------------------------------------------
# CSV 캐시
# ---------------------------------------------------------------------------
def _cache_path(stn_id: int, start: date, end: date) -> Path:
    """캐시 파일 경로: data/raw/{stnId}_{startYYYY}_{endYYYY}.csv"""
    return CACHE_DIR / f"{stn_id}_{start.year}_{end.year}.csv"


def _save_cache(rows: list[dict], path: Path) -> None:
    """API 결과를 CSV로 저장"""
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        return
    fieldnames = list(rows[0].keys())
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def _load_cache(path: Path) -> list[dict] | None:
    """캐시 CSV 파일 로드. 없으면 None 반환."""
    if not path.exists():
        return None
    rows = []
    with open(path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(dict(row))
    return rows if rows else None


# ---------------------------------------------------------------------------
# ASOS API 호출
# ---------------------------------------------------------------------------
def _fetch_asos_daily(
    service_key: str,
    stn_id: int,
    start: date,
    end: date,
    max_retries: int = 3,
    retry_delay: float = 5.0,
) -> list[dict]:
    """ASOS 일자료 API 1회 호출. 타임아웃 시 재시도.

    Args:
        max_retries: 최대 재시도 횟수 (기본 3회)
        retry_delay: 재시도 대기 시간(초), 지수 증가 (기본 5초 → 10초 → 20초)

    Returns:
        일별 관측 데이터 딕셔너리 리스트
    """
    # serviceKey는 이미 URL-인코딩된 상태이므로 직접 URL 조립
    # (requests.get(params=...)를 쓰면 이중 인코딩됨)
    url = (
        f"{ASOS_DAILY_URL}"
        f"?serviceKey={service_key}"
        f"&pageNo=1"
        f"&numOfRows={MAX_NUM_OF_ROWS}"
        f"&dataType=JSON"
        f"&dataCd=ASOS"
        f"&dateCd=DAY"
        f"&startDt={start.strftime('%Y%m%d')}"
        f"&endDt={end.strftime('%Y%m%d')}"
        f"&stnIds={stn_id}"
    )

    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, timeout=60)
            resp.raise_for_status()
            break  # 성공 시 루프 탈출
        except (requests.exceptions.Timeout,
                requests.exceptions.ConnectionError) as e:
            last_exc = e
            if attempt < max_retries:
                wait = retry_delay * (2 ** (attempt - 1))
                print(f"         ⚠ 연결 오류 (시도 {attempt}/{max_retries}), "
                      f"{wait:.0f}초 후 재시도...")
                time.sleep(wait)
            else:
                raise RuntimeError(
                    f"API 호출 실패 (지점={stn_id}, {start}~{end}): {e}"
                ) from e

    data = resp.json()

    # 응답 구조 파싱
    header = data.get("response", {}).get("header", {})
    result_code = header.get("resultCode", "")
    if result_code != "00":
        msg = header.get("resultMsg", "알 수 없는 오류")
        raise RuntimeError(
            f"ASOS API 오류 [{result_code}]: {msg} "
            f"(지점={stn_id}, {start}~{end})"
        )

    body = data.get("response", {}).get("body", {})

    # 페이지네이션 데이터 누락 감지
    total_count = int(body.get("totalCount", 0))
    num_of_rows = int(body.get("numOfRows", MAX_NUM_OF_ROWS))
    if total_count > num_of_rows:
        print(f"         ⚠ 경고: 총 {total_count}건 중 {num_of_rows}건만 수신됨! "
              f"(지점={stn_id}, {start}~{end})")

    items = body.get("items", {})
    if not items:
        return []

    item_list = items.get("item", [])
    if isinstance(item_list, dict):
        item_list = [item_list]

    # 필요 필드만 추출
    result = []
    for item in item_list:
        row = {}
        for field in SOLAR_FIELDS:
            row[field] = item.get(field, "")
        result.append(row)

    return result


def fetch_station_data(
    service_key: str,
    stn_id: int,
    start_date: date | None = None,
    end_date: date | None = None,
    use_cache: bool = True,
    verbose: bool = True,
) -> list[dict]:
    """단일 지점의 전체 기간 데이터를 수집한다.

    - 10년 단위 자동 분할
    - 과거 확정 연도는 CSV 캐시 활용
    - 올해 데이터는 매 실행 시 갱신

    Args:
        service_key: 공공데이터포털 인증키
        stn_id: ASOS 지점번호
        start_date: 수집 시작일 (기본: DATA_START)
        end_date: 수집 종료일 (기본: D-1)
        use_cache: CSV 캐시 사용 여부
        verbose: 진행 상황 출력

    Returns:
        일별 관측 데이터 리스트 (날짜 오름차순)
    """
    if start_date is None:
        start_date = DATA_START
    if end_date is None:
        end_date = yesterday()

    stn_name = stations.station_name(stn_id)
    chunks = get_date_chunks(start_date, end_date)
    all_rows = []

    for chunk_start, chunk_end in chunks:
        cache_file = _cache_path(stn_id, chunk_start, chunk_end)

        # 과거 확정 데이터 → 캐시 활용
        chunk_is_finalized = is_finalized_year(chunk_end.year)

        if use_cache and chunk_is_finalized:
            cached = _load_cache(cache_file)
            if cached:
                if verbose:
                    print(
                        f"  [캐시] {stn_name}({stn_id}) "
                        f"{chunk_start}~{chunk_end}: {len(cached)}건"
                    )
                all_rows.extend(cached)
                continue

        # API 호출
        if verbose:
            print(
                f"  [API]  {stn_name}({stn_id}) "
                f"{chunk_start}~{chunk_end} 조회 중..."
            )

        rows = _fetch_asos_daily(service_key, stn_id, chunk_start, chunk_end)

        if verbose:
            print(f"         → {len(rows)}건 수신")

        # 캐시 저장 (확정 연도만)
        if use_cache and rows and chunk_is_finalized:
            _save_cache(rows, cache_file)
            if verbose:
                print(f"         → 캐시 저장: {cache_file.name}")

        # 올해 데이터도 캐시 저장 (다음 실행 시 덮어쓰기)
        if use_cache and rows and not chunk_is_finalized:
            _save_cache(rows, cache_file)

        all_rows.extend(rows)

    # 날짜 오름차순 정렬
    all_rows.sort(key=lambda r: r.get("tm", ""))

    return all_rows


def fetch_all_stations(
    service_key: str,
    stn_ids: list[int] | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
    use_cache: bool = True,
    verbose: bool = True,
    skip_on_error: bool = True,
    inter_station_delay: float = 0.5,
) -> dict[int, list[dict]]:
    """여러 지점의 데이터를 일괄 수집한다.

    Args:
        stn_ids: 지점번호 리스트 (기본: 전체 66개)
        skip_on_error: True이면 실패 지점 건너뛰고 계속 진행 (기본: True)
        inter_station_delay: 지점 간 API 호출 대기 시간(초) (기본: 0.5초)

    Returns:
        {지점번호: [일별 데이터]} 딕셔너리
    """
    if stn_ids is None:
        stn_ids = stations.all_station_ids()

    total = len(stn_ids)
    result = {}
    failed = []

    for i, stn_id in enumerate(stn_ids, 1):
        stn_name = stations.station_name(stn_id)
        if verbose:
            print(f"[{i}/{total}] {stn_name}({stn_id})")

        try:
            rows = fetch_station_data(
                service_key, stn_id, start_date, end_date,
                use_cache=use_cache, verbose=verbose,
            )
            result[stn_id] = rows
        except Exception as e:
            if skip_on_error:
                print(f"  ✗ 오류 발생, 건너뜀: {e}")
                failed.append(stn_id)
                result[stn_id] = []
            else:
                raise

        # 지점 간 대기 (연속 호출 부하 방지)
        if i < total and inter_station_delay > 0:
            time.sleep(inter_station_delay)

    if failed and verbose:
        failed_names = [f"{stations.station_name(s)}({s})" for s in failed]
        print(f"\n⚠ 수집 실패 지점 ({len(failed)}개): {', '.join(failed_names)}")
        print("  → 재실행 시 캐시된 지점은 건너뛰고 실패 지점만 재조회됩니다.")

    return result


# ---------------------------------------------------------------------------
# 데이터 변환 유틸
# ---------------------------------------------------------------------------
def safe_float(value, default=None) -> float | None:
    """문자열 → float 변환. 빈 문자열이나 변환 불가 시 default 반환."""
    if value is None or str(value).strip() == "":
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def parse_rows(rows: list[dict]) -> list[dict]:
    """API/CSV 원시 데이터를 분석용 딕셔너리로 변환.

    Returns:
        각 행에 다음 필드 추가:
        - date: date 객체
        - year, month, day: 정수
        - sum_gsr: 합계일사량 (float, MJ/m²)
        - sum_ss_hr: 합계일조시간 (float, hr)
        - ss_dur: 가조시간 (float, hr)
        - ss_rate: 일조율 (float, %) — sumSsHr / ssDur × 100
        - hr1_max_icsr: 1시간 최다 일사량 (float, MJ/m²)
        - avg_ta: 평균기온 (float, ℃)
        - sum_rn: 일강수량 (float, mm)
    """
    parsed = []
    for row in rows:
        tm = row.get("tm", "")
        if not tm:
            continue

        # 날짜 파싱 (YYYY-MM-DD 또는 YYYYMMDD)
        tm_clean = tm.replace("-", "")
        try:
            d = date(int(tm_clean[:4]), int(tm_clean[4:6]), int(tm_clean[6:8]))
        except (ValueError, IndexError):
            continue

        sum_gsr = safe_float(row.get("sumGsr"))
        sum_ss_hr = safe_float(row.get("sumSsHr"))
        ss_dur = safe_float(row.get("ssDur"))

        # 일조율 계산
        ss_rate = None
        if sum_ss_hr is not None and ss_dur is not None and ss_dur > 0:
            ss_rate = round(sum_ss_hr / ss_dur * 100, 1)

        parsed.append({
            "date": d,
            "year": d.year,
            "month": d.month,
            "day": d.day,
            "stn_id": int(row.get("stnId", 0)),
            "stn_nm": row.get("stnNm", ""),
            "sum_gsr": sum_gsr,
            "sum_ss_hr": sum_ss_hr,
            "ss_dur": ss_dur,
            "ss_rate": ss_rate,
            "hr1_max_icsr": safe_float(row.get("hr1MaxIcsr")),
            "avg_ta": safe_float(row.get("avgTa")),
            "max_ta": safe_float(row.get("maxTa")),
            "min_ta": safe_float(row.get("minTa")),
            "sum_rn": safe_float(row.get("sumRn")),
            # 원본 보존
            "_raw": row,
        })

    return parsed


# ---------------------------------------------------------------------------
# 집계 함수
# ---------------------------------------------------------------------------
def _avg(values: list[float | None]) -> float | None:
    """None을 제외한 산술평균. 유효값이 없으면 None."""
    valid = [v for v in values if v is not None]
    if not valid:
        return None
    return sum(valid) / len(valid)


def _sum(values: list[float | None]) -> float | None:
    """None을 제외한 합계. 유효값이 없으면 None."""
    valid = [v for v in values if v is not None]
    if not valid:
        return None
    return sum(valid)


def aggregate_monthly(parsed_rows: list[dict]) -> list[dict]:
    """일별 데이터 → 월별 집계.

    Returns:
        각 행: year, month, avg_gsr, sum_gsr, avg_ss_hr, sum_ss_hr,
               avg_ss_dur, avg_ss_rate, days_count
    """
    from collections import defaultdict

    buckets = defaultdict(list)
    for r in parsed_rows:
        key = (r["year"], r["month"])
        buckets[key].append(r)

    result = []
    for (year, month), rows in sorted(buckets.items()):
        gsr_vals = [r["sum_gsr"] for r in rows]
        ss_hr_vals = [r["sum_ss_hr"] for r in rows]
        ss_dur_vals = [r["ss_dur"] for r in rows]
        ss_rate_vals = [r["ss_rate"] for r in rows]

        result.append({
            "year": year,
            "month": month,
            "days_count": len(rows),
            "avg_gsr": _avg(gsr_vals),        # 일평균 일사량
            "sum_gsr": _sum(gsr_vals),         # 적산 일사량 (월합)
            "avg_ss_hr": _avg(ss_hr_vals),     # 일평균 일조시간
            "sum_ss_hr": _sum(ss_hr_vals),     # 적산 일조시간 (월합)
            "avg_ss_dur": _avg(ss_dur_vals),   # 평균 가조시간
            "avg_ss_rate": _avg(ss_rate_vals), # 평균 일조율
        })

    return result


def aggregate_yearly(parsed_rows: list[dict]) -> list[dict]:
    """일별 데이터 → 연도별 집계.

    Returns:
        각 행: year, avg_gsr, sum_gsr, avg_ss_hr, sum_ss_hr,
               avg_ss_rate, days_count
    """
    from collections import defaultdict

    buckets = defaultdict(list)
    for r in parsed_rows:
        buckets[r["year"]].append(r)

    result = []
    for year, rows in sorted(buckets.items()):
        gsr_vals = [r["sum_gsr"] for r in rows]
        ss_hr_vals = [r["sum_ss_hr"] for r in rows]
        ss_rate_vals = [r["ss_rate"] for r in rows]

        result.append({
            "year": year,
            "days_count": len(rows),
            "avg_gsr": _avg(gsr_vals),
            "sum_gsr": _sum(gsr_vals),
            "avg_ss_hr": _avg(ss_hr_vals),
            "sum_ss_hr": _sum(ss_hr_vals),
            "avg_ss_rate": _avg(ss_rate_vals),
        })

    return result


def calc_normal(
    parsed_rows: list[dict],
    ref_year: int | None = None,
    normal_years: int = 10,
) -> dict:
    """평년값 계산 (월별).

    Args:
        ref_year: 기준 연도 (기본: 올해). 평년 = ref_year-normal_years ~ ref_year-1
        normal_years: 평년 산정 기간 (기본: 10년)

    Returns:
        {month: {"avg_gsr": ..., "sum_gsr": ..., "avg_ss_hr": ..., ...}}
    """
    from collections import defaultdict

    if ref_year is None:
        ref_year = date.today().year

    normal_start = ref_year - normal_years
    normal_end = ref_year - 1

    # 평년 기간 데이터 필터링
    normal_rows = [
        r for r in parsed_rows
        if normal_start <= r["year"] <= normal_end
    ]

    # 월별로 연도별 값을 모아 평균
    month_year_gsr = defaultdict(list)      # {month: [연도별 월합]}
    month_year_ss_hr = defaultdict(list)
    month_year_ss_rate = defaultdict(list)
    month_year_avg_gsr = defaultdict(list)  # {month: [연도별 일평균]}
    month_year_avg_ss_hr = defaultdict(list)

    # 먼저 연도-월별 집계
    year_month_buckets = defaultdict(list)
    for r in normal_rows:
        year_month_buckets[(r["year"], r["month"])].append(r)

    for (year, month), rows in year_month_buckets.items():
        gsr_sum = _sum([r["sum_gsr"] for r in rows])
        ss_hr_sum = _sum([r["sum_ss_hr"] for r in rows])
        gsr_avg = _avg([r["sum_gsr"] for r in rows])
        ss_hr_avg = _avg([r["sum_ss_hr"] for r in rows])
        ss_rate_avg = _avg([r["ss_rate"] for r in rows])

        if gsr_sum is not None:
            month_year_gsr[month].append(gsr_sum)
        if ss_hr_sum is not None:
            month_year_ss_hr[month].append(ss_hr_sum)
        if gsr_avg is not None:
            month_year_avg_gsr[month].append(gsr_avg)
        if ss_hr_avg is not None:
            month_year_avg_ss_hr[month].append(ss_hr_avg)
        if ss_rate_avg is not None:
            month_year_ss_rate[month].append(ss_rate_avg)

    result = {}
    for month in range(1, 13):
        result[month] = {
            "avg_gsr": _avg(month_year_avg_gsr.get(month, [])),
            "sum_gsr": _avg(month_year_gsr.get(month, [])),
            "avg_ss_hr": _avg(month_year_avg_ss_hr.get(month, [])),
            "sum_ss_hr": _avg(month_year_ss_hr.get(month, [])),
            "avg_ss_rate": _avg(month_year_ss_rate.get(month, [])),
            "normal_start": normal_start,
            "normal_end": normal_end,
        }

    return result


# ---------------------------------------------------------------------------
# CLI 진입점 (테스트용)
# ---------------------------------------------------------------------------
def main():
    """단일 지점 테스트: 서울(108)"""
    keys = load_apikeys()
    service_key = keys.get("DATA_GO_KR")
    if not service_key:
        print("ERROR: apikey.txt에 DATA_GO_KR 키가 없습니다.")
        return

    stn_id = 108  # 서울
    end = yesterday()
    print(f"=== 서울({stn_id}) 일사·일조 데이터 수집 ===")
    print(f"기간: {DATA_START} ~ {end}")
    print()

    rows = fetch_station_data(service_key, stn_id, DATA_START, end)
    print(f"\n총 {len(rows)}건 수집 완료")

    if rows:
        # 파싱 및 월별 집계 테스트
        parsed = parse_rows(rows)
        print(f"파싱 완료: {len(parsed)}건")

        monthly = aggregate_monthly(parsed)
        print(f"\n── 월별 집계 (최근 12개월) ──")
        for m in monthly[-12:]:
            avg_gsr = f"{m['avg_gsr']:.1f}" if m['avg_gsr'] else "N/A"
            sum_gsr = f"{m['sum_gsr']:.1f}" if m['sum_gsr'] else "N/A"
            avg_ss = f"{m['avg_ss_hr']:.1f}" if m['avg_ss_hr'] else "N/A"
            ss_rate = f"{m['avg_ss_rate']:.1f}" if m['avg_ss_rate'] else "N/A"
            print(
                f"  {m['year']}-{m['month']:02d}: "
                f"평균일사 {avg_gsr} MJ/m² | "
                f"적산일사 {sum_gsr} MJ/m² | "
                f"일조 {avg_ss} hr | "
                f"일조율 {ss_rate}%"
            )

        # 평년 비교
        normal = calc_normal(parsed)
        print(f"\n── 평년값 ({normal[1]['normal_start']}~{normal[1]['normal_end']}) ──")
        for month in range(1, 13):
            n = normal[month]
            avg_gsr = f"{n['avg_gsr']:.1f}" if n['avg_gsr'] else "N/A"
            avg_ss = f"{n['avg_ss_hr']:.1f}" if n['avg_ss_hr'] else "N/A"
            print(f"  {month:2d}월: 평균일사 {avg_gsr} MJ/m² | 일조 {avg_ss} hr")


if __name__ == "__main__":
    main()
