"""모드 3: 연면 경로 취합.

같은 구도·같은 크기의 이미지 여러 장에서 각각의 빨간 경로선을 뽑아 한 장에 모아 그린다.
  1) 이미지별 빨간 선 마스크 추출 (HSV)
  2) 배경 복원: 어떤 이미지에서 빨간 선 아래였던 픽셀은, 그 자리가 빨갛지 않은 다른 이미지의
     픽셀로 채운다 (같은 그림이므로 실제 배경이 그대로 복원됨). 모든 이미지에서 빨간 곳만 inpaint.
  3) 경로마다 다른 색으로 다시 그리고 번호 배지 + 범례 표시
"""
from __future__ import annotations

import csv
from dataclasses import dataclass

import cv2
import numpy as np

# 경로 색상 (BGR) - 서로 잘 구분되는 순서
PATH_COLORS = [
    ((40, 40, 225), "빨강"), ((220, 110, 20), "파랑"), ((40, 160, 40), "초록"),
    ((0, 140, 255), "주황"), ((170, 50, 150), "보라"), ((190, 170, 0), "청록"),
    ((200, 40, 220), "자홍"), ((30, 80, 140), "갈색"), ((60, 60, 60), "검정"),
    ((120, 120, 0), "올리브"),
]


@dataclass
class PathImage:
    name: str
    img: np.ndarray
    mask: np.ndarray        # bool, 빨간 선 픽셀

    @property
    def pixels(self) -> int:
        return int(self.mask.sum())

    def bbox(self):
        ys, xs = np.nonzero(self.mask)
        if not len(xs):
            return None
        return int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())


class SizeMismatch(ValueError):
    pass


def red_mask(img: np.ndarray, min_area: int = 6) -> np.ndarray:
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    m = cv2.inRange(hsv, (0, 110, 80), (10, 255, 255)) | \
        cv2.inRange(hsv, (170, 110, 80), (180, 255, 255))
    # 잡티 제거 (선은 가늘 수 있으므로 열기 연산 대신 작은 조각만 제거)
    n, lab, stats, _ = cv2.connectedComponentsWithStats(m, connectivity=8)
    keep = np.zeros(n, bool)
    keep[1:] = stats[1:, cv2.CC_STAT_AREA] >= min_area
    return keep[lab]


def make_path(name: str, img: np.ndarray, ref_shape=None) -> PathImage:
    if ref_shape is not None and img.shape[:2] != ref_shape[:2]:
        h, w = img.shape[:2]
        rh, rw = ref_shape[:2]
        raise SizeMismatch(f"이미지 크기가 다릅니다: {w}×{h} (기준 {rw}×{rh})")
    return PathImage(name, img, red_mask(img))


def restore_background(paths: list[PathImage]) -> np.ndarray:
    """빨간 선을 지운 공통 배경. 다른 이미지에서 빨갛지 않은 픽셀을 가져와 채운다."""
    k = np.ones((3, 3), np.uint8)
    # 선 주변 1px 번짐(안티앨리어싱)까지 포함
    grown = [cv2.dilate(p.mask.astype(np.uint8), k) > 0 for p in paths]
    bg = paths[0].img.copy()
    todo = grown[0].copy()
    for p, g in zip(paths[1:], grown[1:]):
        take = todo & ~g
        bg[take] = p.img[take]
        todo &= ~take
    if todo.any():
        bg = cv2.inpaint(bg, todo.astype(np.uint8) * 255, 3, cv2.INPAINT_TELEA)
    return bg


def _badge_spots(mask: np.ndarray):
    ys, xs = np.nonzero(mask)
    order = [np.argmin(ys), np.argmax(ys), np.argmin(xs), np.argmax(xs)]
    return [(int(xs[i]), int(ys[i])) for i in order]


def _draw_badge(img, center, text, color, s):
    x, y = center
    r = int(11 * s)
    cv2.circle(img, (x, y), r + max(1, int(2 * s)), (255, 255, 255), -1, cv2.LINE_AA)
    cv2.circle(img, (x, y), r, color, -1, cv2.LINE_AA)
    font, fs, th = cv2.FONT_HERSHEY_SIMPLEX, 0.5 * s, max(1, int(round(1.4 * s)))
    (tw, tht), _ = cv2.getTextSize(text, font, fs, th)
    cv2.putText(img, text, (x - tw // 2, y + tht // 2), font, fs, (255, 255, 255), th, cv2.LINE_AA)


def merge(paths: list[PathImage], thickness: int = 0, colored: bool = True,
          badges: bool = True, legend: bool = True) -> np.ndarray:
    if not paths:
        raise ValueError("경로 이미지가 없습니다.")
    bg = restore_background(paths)
    out = bg.copy()
    h, w = out.shape[:2]
    s = max(min(h, w) / 600, 1.0)
    for i, p in enumerate(paths):
        color = PATH_COLORS[i % len(PATH_COLORS)][0] if colored else PATH_COLORS[0][0]
        m = p.mask.astype(np.uint8)
        if thickness > 0:
            k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2 * thickness + 1,) * 2)
            m = cv2.dilate(m, k)
        out[m > 0] = color
    if badges:
        placed = []
        r = int(14 * s)
        for i, p in enumerate(paths):
            if not p.mask.any():
                continue
            color = PATH_COLORS[i % len(PATH_COLORS)][0] if colored else PATH_COLORS[0][0]
            cands = _badge_spots(p.mask)
            best = None
            for (x, y) in cands:   # 끝점 근처, 다른 배지와 겹치지 않는 곳 우선
                bx = min(max(x + r, r), w - r - 1)
                by = min(max(y - r, r), h - r - 1)
                if all((bx - px) ** 2 + (by - py) ** 2 > (2.2 * r) ** 2 for px, py in placed):
                    best = (bx, by)
                    break
            best = best or (min(max(cands[0][0] + r, r), w - r - 1),
                            min(max(cands[0][1] - r, r), h - r - 1))
            placed.append(best)
            _draw_badge(out, best, str(i + 1), color, s)
    if legend:
        _draw_legend(out, paths, colored, s)
    return out


def _draw_legend(img, paths, colored, s):
    font, fs, th = cv2.FONT_HERSHEY_SIMPLEX, 0.45 * s, max(1, int(round(s)))
    rows = []
    for i, p in enumerate(paths):
        color = PATH_COLORS[i % len(PATH_COLORS)][0] if colored else PATH_COLORS[0][0]
        rows.append((f"Path {i + 1}", color))
    lh = int(20 * s)
    tw = max(cv2.getTextSize(t, font, fs, th)[0][0] for t, _ in rows)
    bw, bh = tw + int(44 * s), lh * len(rows) + int(12 * s)
    h, w = img.shape[:2]
    x0, y0 = w - bw - int(10 * s), h - bh - int(10 * s)
    if x0 < 0 or y0 < 0:
        return
    roi = img[y0:y0 + bh, x0:x0 + bw]
    roi[:] = cv2.addWeighted(roi, 0.15, np.full_like(roi, 255), 0.85, 0)
    cv2.rectangle(img, (x0, y0), (x0 + bw, y0 + bh), (150, 150, 150), 1)
    for k, (t, c) in enumerate(rows):
        yy = y0 + int(6 * s) + lh * k + lh // 2
        cv2.line(img, (x0 + int(8 * s), yy), (x0 + int(30 * s), yy), c, max(2, int(3 * s)))
        cv2.putText(img, t, (x0 + int(36 * s), yy + int(5 * s)), font, fs, (30, 30, 30), th,
                    cv2.LINE_AA)


def rows(paths: list[PathImage], colored: bool = True) -> list[dict]:
    out = []
    for i, p in enumerate(paths):
        bb = p.bbox()
        out.append({
            "경로": i + 1,
            "이미지": p.name,
            "표시 색": PATH_COLORS[i % len(PATH_COLORS)][1] if colored else "빨강",
            "선 픽셀 수": p.pixels,
            "영역 (x0,y0,x1,y1)": "" if bb is None else f"{bb[0]},{bb[1]},{bb[2]},{bb[3]}",
            "상태": "OK" if p.pixels else "빨간 선 없음",
        })
    return out


def save_csv(paths, path, colored=True):
    rs = rows(paths, colored)
    if not rs:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rs[0].keys()))
        w.writeheader()
        w.writerows(rs)
