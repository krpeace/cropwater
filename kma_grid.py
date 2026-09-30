#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kma_grid.py — 위경도 ↔ 기상청 단기예보 격자(nx, ny) 변환 (02-Cycle 5단계, G5)

기상청 단기예보 조회서비스 활용가이드의 Lambert Conformal Conic(LCC) 변환과 같은 식·상수.
  지구 반경 6371.00877 km, 격자 간격 5 km, 표준위도 30°·60°, 기준점 (126°E, 38°N) = 격자 (43, 136)
격자 번호는 반올림(int(x + 0.5)) 규칙을 따른다.

농장 예보는 농장 좌표의 격자로 받는다(THEORY 9장 ◇ 격자 예보를 지점에 적용할 때 주의점 2).
  python kma_grid.py 37.90262 127.73570     → 73 134 (ASOS 101 춘천)
"""
import math
import sys

RE = 6371.00877          # 지구 반경 (km)
GRID = 5.0               # 격자 간격 (km)
SLAT1, SLAT2 = 30.0, 60.0
OLON, OLAT = 126.0, 38.0
XO, YO = 43, 136         # 기준점 격자


def _consts():
    d = math.pi / 180.0
    re = RE / GRID
    s1, s2, olon, olat = SLAT1 * d, SLAT2 * d, OLON * d, OLAT * d
    sn = math.log(math.cos(s1) / math.cos(s2)) / math.log(math.tan(math.pi * 0.25 + s2 * 0.5) / math.tan(math.pi * 0.25 + s1 * 0.5))
    sf = math.tan(math.pi * 0.25 + s1 * 0.5) ** sn * math.cos(s1) / sn
    ro = re * sf / math.tan(math.pi * 0.25 + olat * 0.5) ** sn
    return d, re, olon, sn, sf, ro


def latlon_to_grid(lat, lon):
    """위도·경도(°) → 격자 (nx, ny)"""
    d, re, olon, sn, sf, ro = _consts()
    ra = re * sf / math.tan(math.pi * 0.25 + lat * d * 0.5) ** sn
    theta = lon * d - olon
    if theta > math.pi:
        theta -= 2.0 * math.pi
    if theta < -math.pi:
        theta += 2.0 * math.pi
    theta *= sn
    x = ra * math.sin(theta) + XO
    y = ro - ra * math.cos(theta) + YO
    return int(x + 0.5), int(y + 0.5)


def grid_to_latlon(nx, ny):
    """격자 (nx, ny) → 격자 중심 위도·경도(°)"""
    d, re, olon, sn, sf, ro = _consts()
    xn, yn = nx - XO, ro - ny + YO
    ra = math.copysign(math.sqrt(xn * xn + yn * yn), sn)
    alat = 2.0 * math.atan((re * sf / ra) ** (1.0 / sn)) - math.pi * 0.5
    if abs(xn) <= 0.0:
        theta = 0.0
    elif abs(yn) <= 0.0:
        theta = math.pi * 0.5 * (1 if xn >= 0 else -1)
    else:
        theta = math.atan2(xn, yn)
    return alat / d, (theta / sn + olon) / d


def grid_key(nx, ny):
    return f"{int(nx)}_{int(ny)}"


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit("사용법: python kma_grid.py <위도> <경도>")
    nx, ny = latlon_to_grid(float(sys.argv[1]), float(sys.argv[2]))
    print(nx, ny)
