"""모드 3 테스트용: 같은 그림에 빨간 연면 경로만 다른 이미지 3장 생성."""
import sys
from pathlib import Path

import cv2
import numpy as np

W, H = 1000, 560
PATHS = [
    [(120, 380), (120, 250), (300, 250), (300, 200), (480, 200)],
    [(130, 400), (420, 400), (420, 330), (640, 330)],
    [(500, 210), (700, 210), (700, 300), (880, 300), (880, 420)],
]


def base():
    img = np.full((H, W, 3), 245, np.uint8)
    cv2.rectangle(img, (60, 180), (940, 460), (205, 210, 215), -1)
    cv2.rectangle(img, (60, 180), (940, 460), (110, 110, 110), 1)
    for x in range(60, 940, 30):
        cv2.line(img, (x, 180), (x, 460), (185, 190, 195), 1)
    cv2.rectangle(img, (250, 260), (360, 380), (150, 160, 170), -1)       # 전극/부품 형상
    cv2.rectangle(img, (560, 230), (650, 320), (150, 160, 170), -1)
    cv2.circle(img, (800, 380), 45, (150, 160, 170), -1)
    cv2.putText(img, "Model A", (70, 170), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (60, 60, 60), 1,
                cv2.LINE_AA)
    return img


def main(out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    for i, pts in enumerate(PATHS, 1):
        img = base()
        cv2.polylines(img, [np.array(pts, np.int32)], False, (20, 20, 230), 2, cv2.LINE_AA)
        cv2.imwrite(str(out / f"path{i}.png"), img)
    # 등전위선 이미지: 같은 그림 위에 색 등고선 (빨간 경로 없음)
    img = base()
    yy, xx = np.mgrid[0:H, 0:W]
    pot = np.exp(-(((xx - 305) / 160.0) ** 2 + ((yy - 320) / 110.0) ** 2)) \
        + 0.8 * np.exp(-(((xx - 800) / 120.0) ** 2 + ((yy - 380) / 90.0) ** 2))
    inside = (xx > 60) & (xx < 940) & (yy > 180) & (yy < 460)
    colors = [(200, 80, 0), (200, 160, 0), (60, 170, 60), (0, 170, 220), (0, 90, 230)]
    for k, lv in enumerate(np.linspace(0.15, 0.85, 5)):
        band = (np.abs(pot - lv) < 0.006) & inside
        img[band] = colors[k]
    cv2.putText(img, "Equipotential", (790, 170), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (60, 60, 60), 1,
                cv2.LINE_AA)
    cv2.imwrite(str(out / "equipotential.png"), img)
    print("saved", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent / "path_samples"))
