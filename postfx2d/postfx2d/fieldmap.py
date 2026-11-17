"""모드 2: 전계 분포 rainbow map 값 읽기.

컬러맵(LUT)과 최소/최대 범위를 기준으로, 클릭한 픽셀 색 → 값으로 역변환한다.
  - 내장 컬러맵: Rainbow(파랑→빨강), Jet, Turbo
  - 범례에서 추출: 이미지 속 컬러바를 최소 끝 → 최대 끝으로 드래그해 LUT 생성 (가장 정확)
"""
from __future__ import annotations

import csv
import math
from dataclasses import dataclass

import cv2
import numpy as np

N_LUT = 512
COLORMAPS = ["Rainbow (파랑→빨강)", "Jet", "Turbo", "범례에서 추출"]
WARN_DIST = 22.0     # Lab 거리: 이보다 크면 '맵 밖/불확실'


def builtin_lut(name: str, n: int = N_LUT) -> np.ndarray:
    """(n,3) BGR uint8. index 0 = 최소, n-1 = 최대."""
    if name.startswith("Rainbow"):
        h = np.linspace(120, 0, n)                    # OpenCV hue: 120=파랑 → 0=빨강
        hsv = np.stack([h, np.full(n, 255), np.full(n, 255)], -1).astype(np.uint8)
        return cv2.cvtColor(hsv[None], cv2.COLOR_HSV2BGR)[0]
    cmap = {"Jet": cv2.COLORMAP_JET, "Turbo": cv2.COLORMAP_TURBO}[name]
    grad = np.linspace(0, 255, n).astype(np.uint8)[None]
    return cv2.applyColorMap(grad, cmap)[0]


def lut_from_line(img: np.ndarray, p0, p1, n: int = N_LUT) -> np.ndarray:
    """이미지 속 컬러바를 p0(최소 끝) → p1(최대 끝) 직선으로 샘플링해 LUT 생성."""
    (x0, y0), (x1, y1) = p0, p1
    length = max(int(math.hypot(x1 - x0, y1 - y0)), 2)
    xs = np.linspace(x0, x1, length)
    ys = np.linspace(y0, y1, length)
    h, w = img.shape[:2]
    # 선에 수직 방향으로 ±2px 평균 (경계선/글자 영향 완화)
    nx, ny = -(y1 - y0) / length, (x1 - x0) / length
    samples = []
    for x, y in zip(xs, ys):
        acc = []
        for k in (-2, -1, 0, 1, 2):
            px = int(round(min(max(x + nx * k, 0), w - 1)))
            py = int(round(min(max(y + ny * k, 0), h - 1)))
            acc.append(img[py, px].astype(float))
        samples.append(np.median(acc, 0))
    samples = np.array(samples)
    idx = np.linspace(0, len(samples) - 1, n)
    lut = np.stack([np.interp(idx, np.arange(len(samples)), samples[:, c]) for c in range(3)], -1)
    return lut.astype(np.uint8)


def _lab(bgr: np.ndarray) -> np.ndarray:
    a = np.asarray(bgr, np.uint8).reshape(-1, 1, 3)
    return cv2.cvtColor(a, cv2.COLOR_BGR2LAB).reshape(-1, 3).astype(float)


@dataclass
class Reading:
    x: int
    y: int
    value: float
    t: float           # 0~1 (컬러맵 위치)
    dist: float        # 색 오차 (Lab)
    color: tuple       # 샘플 색 (BGR)

    @property
    def reliable(self) -> bool:
        return self.dist <= WARN_DIST


class FieldMap:
    def __init__(self, lut: np.ndarray, vmin: float, vmax: float, log: bool = False):
        if log and (vmin <= 0 or vmax <= 0):
            raise ValueError("로그 스케일은 최소/최대가 0보다 커야 합니다.")
        self.lut = lut
        self.lut_lab = _lab(lut)
        self.vmin, self.vmax, self.log = vmin, vmax, log

    def t_to_value(self, t: float) -> float:
        if self.log:
            a, b = math.log10(self.vmin), math.log10(self.vmax)
            return 10 ** (a + t * (b - a))
        return self.vmin + t * (self.vmax - self.vmin)

    def sample_color(self, img: np.ndarray, x: int, y: int, r: int = 2) -> np.ndarray:
        """클릭 주변 (2r+1)² 에서 채도가 있는 픽셀의 중앙값 (메쉬선/글자 제외)."""
        h, w = img.shape[:2]
        patch = img[max(y - r, 0):min(y + r + 1, h), max(x - r, 0):min(x + r + 1, w)]
        px = patch.reshape(-1, 3)
        sat = cv2.cvtColor(patch, cv2.COLOR_BGR2HSV).reshape(-1, 3)[:, 1]
        good = px[sat > 60]
        return np.median(good if len(good) else px, 0).astype(np.uint8)

    def read(self, img: np.ndarray, x: int, y: int) -> Reading:
        c = self.sample_color(img, x, y)
        d = np.linalg.norm(self.lut_lab - _lab(c), axis=1)
        i = int(np.argmin(d))
        # 이웃 인덱스와 포물선 보간으로 소수점 위치 보정
        if 0 < i < len(d) - 1:
            den = d[i - 1] - 2 * d[i] + d[i + 1]
            off = 0.5 * (d[i - 1] - d[i + 1]) / den if den > 1e-9 else 0.0
            fi = i + float(np.clip(off, -0.5, 0.5))
        else:
            fi = float(i)
        t = fi / (len(d) - 1)
        return Reading(x, y, self.t_to_value(t), t, float(d[i]), tuple(int(v) for v in c))


def fmt(v: float) -> str:
    a = abs(v)
    if a != 0 and (a >= 1e5 or a < 1e-2):
        return f"{v:.3e}"
    return f"{v:.4g}" if a < 100 else f"{v:,.1f}"


def annotate(img: np.ndarray, readings: list[Reading], unit: str = "") -> np.ndarray:
    out = img.copy()
    font, sc, th = cv2.FONT_HERSHEY_SIMPLEX, 0.5, 1
    for k, r in enumerate(readings, 1):
        col = (30, 30, 30) if r.reliable else (40, 40, 220)
        cv2.circle(out, (r.x, r.y), 6, (255, 255, 255), 2, cv2.LINE_AA)
        cv2.circle(out, (r.x, r.y), 6, col, 1, cv2.LINE_AA)
        cv2.circle(out, (r.x, r.y), 1, col, -1)
        text = f"#{k} {fmt(r.value)} {unit}".strip() + ("" if r.reliable else " ?")
        (tw, tht), b = cv2.getTextSize(text, font, sc, th)
        x = r.x + 10 if r.x + 14 + tw < out.shape[1] else r.x - 10 - tw
        y = max(r.y - 8, tht + 4)
        cv2.rectangle(out, (x - 4, y - tht - 4), (x + tw + 4, y + b + 2), (255, 255, 255), -1)
        cv2.rectangle(out, (x - 4, y - tht - 4), (x + tw + 4, y + b + 2), col, 1)
        cv2.putText(out, text, (x, y), font, sc, col, th, cv2.LINE_AA)
    return out


def rows(readings: list[Reading], unit: str = "") -> list[dict]:
    u = f" ({unit})" if unit else ""
    return [{"#": k, "X(px)": r.x, "Y(px)": r.y, f"전계{u}": float(f"{r.value:.6g}"),
             "컬러맵 위치(%)": round(r.t * 100, 1), "색 오차": round(r.dist, 1),
             "신뢰": "OK" if r.reliable else "불확실"}
            for k, r in enumerate(readings, 1)]


def save_csv(readings, path, unit=""):
    rs = rows(readings, unit)
    if not rs:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rs[0].keys()))
        w.writeheader()
        w.writerows(rs)
