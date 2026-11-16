"""2D 해석결과 이미지 후처리 - 모드 1: 최소안전율부위.

처리 흐름
  1) 색상(HSV)으로 원형 점 검출 → W1~W4 / Global 분류
  2) OCR 로 숫자(안전율) 인식 → 가장 가까운 점과 짝짓기
  3) 점+숫자 단위를 거리 기반(단일 연결)으로 그룹화, 왼쪽부터 번호
  4) 그룹 음영(연한 파스텔) + 그룹별 최소 하이라이트 + 글로벌 최소 하이라이트
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field

import cv2
import numpy as np

# ---------------------------------------------------------------- 색상 정의
# OpenCV HSV: H 0~180, S/V 0~255
POINT_CLASSES = {
    "W1": {"label": "W1 (주황)", "ranges": [((10, 120, 120), (24, 255, 255))]},
    "W2": {"label": "W2 (초록)", "ranges": [((40, 80, 60), (85, 255, 255))]},
    "W3": {"label": "W3 (파랑)", "ranges": [((95, 100, 60), (128, 255, 255))]},
    "W4": {"label": "W4 (보라)", "ranges": [((129, 60, 50), (165, 255, 255))]},
    "GLOBAL": {"label": "Global Min (빨강)",
               "ranges": [((0, 140, 110), (8, 255, 255)),
                          ((170, 140, 110), (180, 255, 255))]},
}

# 그룹 음영용 파스텔 팔레트 (BGR)
PASTELS = [
    (250, 206, 135), (170, 230, 170), (190, 175, 245), (140, 220, 250),
    (230, 190, 230), (200, 230, 150), (170, 200, 255), (225, 225, 160),
]
LOCAL_HL = (0, 215, 255)     # 진한 노랑 (BGR)
GLOBAL_HL = (60, 20, 230)    # 진한 빨강/마젠타 계열 (BGR)

NUM_RE = re.compile(r"[-+]?\d+(?:[.,]\d+)?")


@dataclass
class Settings:
    min_radius: int = 2
    max_radius: int = 30
    min_circularity: float = 0.6
    group_gap: float = 0.0          # 0 = 자동
    group_count: int = 0            # >0 이면 그룹 수 강제
    group_alpha: float = 0.18
    highlight_alpha: float = 0.45
    pad: int = 14


@dataclass
class Point:
    cls: str
    x: float
    y: float
    r: float
    value: float | None = None
    index: int | None = None
    text: str = ""
    text_box: tuple[int, int, int, int] | None = None  # x0,y0,x1,y1
    group: int = -1

    def bbox(self) -> tuple[int, int, int, int]:
        x0, y0 = int(self.x - self.r), int(self.y - self.r)
        x1, y1 = int(self.x + self.r), int(self.y + self.r)
        if self.text_box:
            tx0, ty0, tx1, ty1 = self.text_box
            x0, y0, x1, y1 = min(x0, tx0), min(y0, ty0), max(x1, tx1), max(y1, ty1)
        return x0, y0, x1, y1


@dataclass
class Result:
    image: np.ndarray                     # 후처리 결과 (BGR)
    points: list[Point] = field(default_factory=list)
    groups: list[list[Point]] = field(default_factory=list)
    rows: list[dict] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------- 1) 점 검출
def detect_points(img: np.ndarray, s: Settings) -> list[Point]:
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
    pts: list[Point] = []
    for cls, spec in POINT_CLASSES.items():
        mask = np.zeros(hsv.shape[:2], np.uint8)
        for lo, hi in spec["ranges"]:
            mask |= cv2.inRange(hsv, np.array(lo), np.array(hi))
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))
        cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        for c in cnts:
            area = cv2.contourArea(c)
            (x, y), r = cv2.minEnclosingCircle(c)
            if r < s.min_radius or r > s.max_radius or area < 4:
                continue
            per = cv2.arcLength(c, True)
            circ = 4 * np.pi * area / (per * per) if per else 0
            fill = area / (np.pi * r * r)
            if circ < s.min_circularity or fill < 0.5:
                continue
            pts.append(Point(cls, x, y, r))
    return pts


# ---------------------------------------------------------------- 2) OCR
_ocr_engine = None


def _get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


IDX_RE = re.compile(r"[(\[{]\s*(\d+)\s*[)\]}]?|(?:(?<=\s)|^)(\d+)\s*[)\]}]")


def parse_label(text: str) -> tuple[float | None, int | None]:
    """'1.28 (12)' → (1.28, 12). 괄호 안 숫자 = 인덱스, 나머지 숫자 = 안전율."""
    t = text.strip().replace("O", "0").replace("o", "0")
    idx = None
    m = IDX_RE.search(t)
    if m:
        idx = int(m.group(1) or m.group(2))
        t = t[:m.start()] + " " + t[m.end():]
    t = re.sub(r"(?<=[\d.,])\s+(?=[\d.,])", "", t)   # '1. 28' 처럼 쪼개진 숫자 붙이기
    val = None
    for v in NUM_RE.finditer(t):
        try:
            val = float(v.group().replace(",", "."))
            break
        except ValueError:
            pass
    return val, idx


def read_numbers(img: np.ndarray) -> list[dict]:
    """이미지 안의 라벨 → [{value, index, text, box}]. 작은 글씨 대비 2배 확대 후 인식."""
    scale = 2.0 if max(img.shape[:2]) < 2500 else 1.0
    big = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    res = _get_ocr()(big)
    res = res[0] if isinstance(res, tuple) else res
    items = []
    for box, text, conf in res or []:
        val, idx = parse_label(text)
        if val is None and idx is None:
            continue
        xs = [p[0] / scale for p in box]
        ys = [p[1] / scale for p in box]
        items.append({"value": val, "index": idx, "text": text.strip(),
                      "box": (int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys)))})
    # OCR 이 '(12)' 를 따로 떼어 읽은 경우 → 같은 줄의 가장 가까운 값 라벨에 합침
    vals = [it for it in items if it["value"] is not None]
    for it in [it for it in items if it["value"] is None]:
        x0, y0, x1, y1 = it["box"]
        h = y1 - y0
        best, bd = None, 3 * max(h, 8)
        for v in vals:
            if v["index"] is not None:
                continue
            vx0, vy0, vx1, vy1 = v["box"]
            if min(y1, vy1) - max(y0, vy0) < 0.4 * min(h, vy1 - vy0):
                continue                          # 같은 줄 아님
            d = max(x0 - vx1, vx0 - x1, 0)
            if d < bd:
                best, bd = v, d
        if best:
            best["index"] = it["index"]
            best["text"] += " " + it["text"]
            bx = best["box"]
            best["box"] = (min(bx[0], x0), min(bx[1], y0), max(bx[2], x1), max(bx[3], y1))
    return vals


def pair_numbers(points: list[Point], nums: list[dict]) -> None:
    """각 라벨을 가장 가까운 점에 1:1 매칭(가까운 쌍부터 탐욕적)."""
    cands = []
    for ni, it in enumerate(nums):
        x0, y0, x1, y1 = it["box"]
        # 라벨은 점 옆에 붙으므로, 박스 네 변의 중점까지 거리로 판단
        # (숫자 위에 겹쳐 찍힌 다른 점이 잘못 매칭되는 것 방지)
        cy, cx = (y0 + y1) / 2, (x0 + x1) / 2
        anchors = [(x0, cy), (x1, cy), (cx, y0), (cx, y1)]
        for pi, p in enumerate(points):
            d = min(np.hypot(p.x - ax, p.y - ay) for ax, ay in anchors)
            cands.append((d, ni, pi))
    cands.sort()
    used_n, used_p = set(), set()
    for d, ni, pi in cands:
        if ni in used_n or pi in used_p:
            continue
        p = points[pi]
        if d > max(70, p.r * 14):
            continue
        it = nums[ni]
        p.value, p.index, p.text, p.text_box = it["value"], it["index"], it["text"], it["box"]
        used_n.add(ni)
        used_p.add(pi)


# ---------------------------------------------------------------- 3) 그룹화
def _box_dist(a, b) -> float:
    dx = max(a[0] - b[2], 0, b[0] - a[2])
    dy = max(a[1] - b[3], 0, b[1] - a[3])
    return float(np.hypot(dx, dy))


def auto_gap(points: list[Point], count: int = 0) -> float:
    """MST 간선 길이 분포에서 '큰 점프'를 찾아 그룹 간격 임계값 결정."""
    n = len(points)
    if n < 2:
        return 1.0
    boxes = [p.bbox() for p in points]
    D = np.array([[_box_dist(boxes[i], boxes[j]) for j in range(n)] for i in range(n)])
    # Prim MST
    in_t = np.zeros(n, bool); in_t[0] = True
    best = D[0].copy(); edges = []
    for _ in range(n - 1):
        best[in_t] = np.inf
        j = int(np.argmin(best)); edges.append(best[j]); in_t[j] = True
        best = np.minimum(best, D[j])
    e = np.sort(np.array(edges)) + 5.0          # 박스가 겹친(0) 간선 안정화
    if 0 < count <= n:                           # 그룹 수 지정: 가장 긴 (count-1)개 간선 절단
        if count == 1:
            return float(e[-1])
        k = n - count - 1
        return float((e[k] + e[k + 1]) / 2 - 5.0) if k >= 0 else 0.0
    base = float(np.median(e))
    # 그룹 간 간선은 MST 상위 소수 → 중앙값 이상 구간에서 가장 큰 비율 점프를 찾음
    start = len(e) // 2
    ratios = e[start + 1:] / e[start:-1]
    if len(ratios) and ratios.max() > 1.8:
        k = start + int(np.argmax(ratios))
        return float((e[k] + e[k + 1]) / 2 - 5.0)
    return float(base * 3)


def group_points(points: list[Point], gap: float) -> list[list[Point]]:
    n = len(points)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]; i = parent[i]
        return i

    boxes = [p.bbox() for p in points]
    for i in range(n):
        for j in range(i + 1, n):
            if _box_dist(boxes[i], boxes[j]) <= gap:
                parent[find(i)] = find(j)
    buckets: dict[int, list[Point]] = {}
    for i, p in enumerate(points):
        buckets.setdefault(find(i), []).append(p)
    groups = sorted(buckets.values(), key=lambda g: np.mean([p.x for p in g]))
    for gi, g in enumerate(groups, 1):
        for p in g:
            p.group = gi
    return groups


# ---------------------------------------------------------------- 4) 렌더링
def _blend_mask(img, mask, color, alpha):
    over = img.copy(); over[mask > 0] = color
    return np.where(mask[..., None] > 0,
                    cv2.addWeighted(over, alpha, img, 1 - alpha, 0), img)


def _hull_mask(shape, boxes, pad):
    pts = []
    for x0, y0, x1, y1 in boxes:
        pts += [(x0, y0), (x1, y0), (x0, y1), (x1, y1)]
    m = np.zeros(shape[:2], np.uint8)
    hull = cv2.convexHull(np.array(pts, np.int32))
    cv2.fillConvexPoly(m, hull, 255)
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2 * pad + 1, 2 * pad + 1))
    return cv2.dilate(m, k)


def _outline(img, mask, color, thick):
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cv2.drawContours(img, cnts, -1, color, thick, cv2.LINE_AA)


def _label(img, text, org, color):
    font, sc, th = cv2.FONT_HERSHEY_SIMPLEX, 0.55, 1
    (w, h), b = cv2.getTextSize(text, font, sc, th)
    x, y = int(org[0]), int(max(org[1], h + 4))
    cv2.rectangle(img, (x - 3, y - h - 4), (x + w + 3, y + b), (255, 255, 255), -1)
    cv2.rectangle(img, (x - 3, y - h - 4), (x + w + 3, y + b), color, 1)
    cv2.putText(img, text, (x, y - 1), font, sc, color, th, cv2.LINE_AA)


def render(img, groups, s: Settings):
    out = img.copy()
    # 그룹 음영
    gmasks = []
    for gi, g in enumerate(groups):
        m = _hull_mask(out.shape, [p.bbox() for p in g], s.pad * 2)
        col = PASTELS[gi % len(PASTELS)]
        out = _blend_mask(out, m, col, s.group_alpha)
        dark = tuple(int(c * 0.75) for c in col)
        _outline(out, m, dark, 1)
        gmasks.append((m, dark))
    # 하이라이트
    for g in groups:
        vals = [p.value for p in g if p.value is not None]
        has_global = any(p.cls == "GLOBAL" for p in g)
        targets = [p for p in g if p.cls == "GLOBAL"] if has_global else \
                  ([p for p in g if p.value == min(vals)] if vals else [])
        for p in targets:
            glob = p.cls == "GLOBAL"
            color = GLOBAL_HL if glob else LOCAL_HL
            m = np.zeros(out.shape[:2], np.uint8)
            x0, y0, x1, y1 = p.bbox()
            cv2.rectangle(m, (x0, y0), (x1, y1), 255, -1)
            k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (2 * s.pad + 1,) * 2)
            m = cv2.dilate(m, k)
            out = _blend_mask(out, m, color, s.highlight_alpha + (0.1 if glob else 0))
            _outline(out, m, color, 3 if glob else 2)
    # 원본 점/숫자(잉크)가 음영에 묻히지 않도록 원본 픽셀 복원
    roi = np.zeros(img.shape[:2], np.uint8)
    for g in groups:
        for p in g:
            x0, y0, x1, y1 = p.bbox()
            cv2.rectangle(roi, (x0 - 2, y0 - 2), (x1 + 2, y1 + 2), 255, -1)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    sat = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)[..., 1]
    ink = (roi > 0) & ((gray < 120) | (sat > 90))
    out[ink] = img[ink]
    # 그룹 번호
    for gi, (g, (m, dark)) in enumerate(zip(groups, gmasks), 1):
        ys, xs = np.nonzero(m)
        _label(out, f"G{gi}", (xs.min() + 2, ys.min() - 4), dark)
    return out


# ---------------------------------------------------------------- 파이프라인
def process(img: np.ndarray, s: Settings | None = None, ocr=True) -> Result:
    s = s or Settings()
    warnings = []
    pts = detect_points(img, s)
    if not pts:
        return Result(img.copy(), warnings=["색상 점을 찾지 못했습니다."])
    if ocr:
        try:
            pair_numbers(pts, read_numbers(img))
        except Exception as e:  # OCR 실패해도 그룹화는 진행
            warnings.append(f"OCR 오류: {e}")
    gap = s.group_gap or auto_gap(pts, s.group_count)
    groups = group_points(pts, gap)
    if sum(p.cls == "GLOBAL" for p in pts) == 0:
        warnings.append("빨간 점(Global Minimum)을 찾지 못했습니다.")
    noidx = sum(p.value is not None and p.index is None for p in pts)
    if noidx:
        warnings.append(f"인덱스(괄호 숫자)를 읽지 못한 점 {noidx}개")
    miss = sum(p.value is None for p in pts)
    if miss:
        warnings.append(f"숫자를 읽지 못한 점 {miss}개 (표에서 직접 수정 가능)")
    res = Result(render(img, groups, s), pts, groups, warnings=warnings)
    res.rows = summarize(groups)
    res.gap = gap
    return res


def rerender(img, groups, s: Settings) -> Result:
    res = Result(render(img, groups, s), [p for g in groups for p in g], groups)
    res.rows = summarize(groups)
    return res


def summarize(groups) -> list[dict]:
    rows = []
    for gi, g in enumerate(groups, 1):
        vals = [p.value for p in g if p.value is not None]
        mn = min(vals) if vals else None
        glob = [p for p in g if p.cls == "GLOBAL"]
        mins = glob or [p for p in g if mn is not None and p.value == mn]
        for p in mins:
            rows.append({
                "그룹": f"G{gi}",
                "점 개수": len(g),
                "최소 안전율": p.value if p.value is not None else "",
                "구분": "Global Min" if p.cls == "GLOBAL" else p.cls,
                "하이라이트": "글로벌" if p.cls == "GLOBAL" else "국소",
                "X(px)": int(p.x), "Y(px)": int(p.y),
                "인덱스": "" if p.index is None else p.index,
            })
        if not mins:
            rows.append({"그룹": f"G{gi}", "점 개수": len(g), "최소 안전율": "",
                         "구분": "", "하이라이트": "", "X(px)": "", "Y(px)": "", "인덱스": ""})
    return rows


def save_csv(rows, path):
    if not rows:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)


def imread(path) -> np.ndarray:
    data = np.fromfile(path, np.uint8)          # 한글 경로 대응
    img = cv2.imdecode(data, cv2.IMREAD_UNCHANGED)
    if img is None:
        raise ValueError(f"이미지를 열 수 없습니다: {path}")
    if img.ndim == 2:
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    elif img.shape[2] == 4:
        a = img[..., 3:4] / 255.0
        img = (img[..., :3] * a + 255 * (1 - a)).astype(np.uint8)
    return img


def imwrite(path, img):
    ok, buf = cv2.imencode(".png", img)
    if ok:
        buf.tofile(path)
