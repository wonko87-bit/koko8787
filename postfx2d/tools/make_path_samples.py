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
    print("saved", out)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else str(Path(__file__).parent / "path_samples"))
