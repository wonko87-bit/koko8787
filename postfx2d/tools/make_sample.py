"""테스트용 합성 샘플 이미지 생성 (6개 그룹, 6번 그룹에 Global Min 빨간 점)."""
import random
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

COLORS = {"W1": (255, 140, 0), "W2": (40, 170, 60), "W3": (30, 90, 230),
          "W4": (150, 60, 200), "GLOBAL": (230, 20, 20)}


def font(size):
    for f in ["arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]:
        try:
            return ImageFont.truetype(f, size)
        except OSError:
            pass
    return ImageFont.load_default()


def main(out="sample.png", seed=3):
    rnd = random.Random(seed)
    W, H = 1640, 600
    im = Image.new("RGB", (W, H), (245, 245, 245))
    d = ImageDraw.Draw(im)
    # 해석 모델처럼 보이는 회색 형상 + 메쉬 선
    d.rectangle((40, 200, W - 40, 420), fill=(205, 210, 215), outline=(120, 120, 120))
    for x in range(40, W - 40, 30):
        d.line((x, 200, x, 420), fill=(185, 190, 195))
    f = font(15)
    idx = 0
    centers = [(140, 300), (400, 280), (660, 320), (920, 290), (1180, 310), (1440, 300)]
    for gi, (cx, cy) in enumerate(centers, 1):
        n = rnd.randint(4, 6)
        used = []
        for k in range(n):
            while True:
                x, y = cx + rnd.randint(-55, 15), cy + rnd.randint(-85, 85)
                if all(abs(x - a) > 75 or abs(y - b) > 24 for a, b in used):
                    break
            used.append((x, y))
            cls = rnd.choice(["W1", "W2", "W3", "W4"])
            val = round(rnd.uniform(1.2, 3.5), 2)
            if gi == 6 and k == 0:
                cls, val = "GLOBAL", 0.98
            r = 6
            d.ellipse((x - r, y - r, x + r, y + r), fill=COLORS[cls])
            idx += 1
            d.text((x + 10, y - 9), f"{val:.2f} ({idx})", fill=(0, 0, 0), font=f)
    im.save(out)
    print("saved", out)


if __name__ == "__main__":
    main(*(sys.argv[1:2] or [str(Path(__file__).parent / "sample.png")]))
