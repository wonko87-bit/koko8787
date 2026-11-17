"""모드 2 테스트용 합성 rainbow 전계 분포 이미지 (+ 컬러바) 생성.

정답 값을 알고 있으므로 읽기 정확도 검증에 사용한다.
"""
import sys
from pathlib import Path

import cv2
import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from postfx2d.fieldmap import builtin_lut  # noqa: E402

VMIN, VMAX = 0.0, 5000.0          # V/m
W, H = 900, 520
MAP = (40, 40, 700, 480)          # x0, y0, x1, y1
BAR = (760, 60, 790, 460)         # 세로 컬러바 (아래=최소, 위=최대)


def field(x, y):
    """정규화 필드 0~1."""
    f = (np.exp(-(((x - 0.35) / 0.18) ** 2 + ((y - 0.45) / 0.22) ** 2))
         + 0.7 * np.exp(-(((x - 0.75) / 0.12) ** 2 + ((y - 0.7) / 0.15) ** 2)))
    return np.clip(f / 1.15, 0, 1)


def main(out):
    lut = builtin_lut("Rainbow")
    img = np.full((H, W, 3), 255, np.uint8)
    x0, y0, x1, y1 = MAP
    ys, xs = np.mgrid[y0:y1, x0:x1]
    f = field((xs - x0) / (x1 - x0), (ys - y0) / (y1 - y0))
    img[y0:y1, x0:x1] = lut[(f * (len(lut) - 1)).round().astype(int)]
    for gx in range(x0, x1, 40):                     # 메쉬선
        cv2.line(img, (gx, y0), (gx, y1), (60, 60, 60), 1)
    for gy in range(y0, y1, 40):
        cv2.line(img, (x0, gy), (x1, gy), (60, 60, 60), 1)
    cv2.rectangle(img, (x0, y0), (x1, y1), (0, 0, 0), 1)
    bx0, by0, bx1, by1 = BAR
    t = np.linspace(1, 0, by1 - by0)
    img[by0:by1, bx0:bx1] = lut[(t * (len(lut) - 1)).round().astype(int)][:, None]
    cv2.rectangle(img, (bx0, by0), (bx1, by1), (0, 0, 0), 1)
    for k in range(6):
        v = VMIN + k * (VMAX - VMIN) / 5
        yy = int(by1 - k * (by1 - by0) / 5)
        cv2.putText(img, f"{v:.0f}", (bx1 + 8, yy + 5), cv2.FONT_HERSHEY_SIMPLEX, 0.45,
                    (0, 0, 0), 1, cv2.LINE_AA)
    cv2.putText(img, "E [V/m]", (bx0 - 10, by0 - 15), cv2.FONT_HERSHEY_SIMPLEX, 0.5,
                (0, 0, 0), 1, cv2.LINE_AA)
    cv2.imwrite(out, img)
    print("saved", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent / "field_sample.png"))
