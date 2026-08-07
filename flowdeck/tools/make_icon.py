#!/usr/bin/env python3
"""Generate flowdeck.ico — the app, tray and window icon.

The mark is a white rounded square with a dark check stroke through it, which
keeps the app's monochrome palette and stays legible on both light and dark
taskbars. Everything is rendered at 4x and box-filtered down, so the edges are
smooth without pulling in an imaging library.

Run: python3 tools/make_icon.py src/Flowdeck.Windows/Assets/flowdeck.ico
"""

from __future__ import annotations

import struct
import sys

SIZES = (16, 24, 32, 48, 64, 128, 256)
SUPERSAMPLE = 4

PLATE = (255, 255, 255)   # rounded square
GLYPH = (17, 17, 19)      # check stroke

# Geometry in unit coordinates, so every size renders from the same drawing.
PLATE_INSET = 0.055
PLATE_RADIUS = 0.235
STROKE = 0.115
CHECK = ((0.275, 0.520), (0.435, 0.685), (0.735, 0.320))


def inside_rounded_rect(x: float, y: float) -> bool:
    lo, hi = PLATE_INSET, 1.0 - PLATE_INSET
    if not (lo <= x <= hi and lo <= y <= hi):
        return False

    r = PLATE_RADIUS
    cx = min(max(x, lo + r), hi - r)
    cy = min(max(y, lo + r), hi - r)
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r


def distance_to_segment(px: float, py: float, ax: float, ay: float, bx: float, by: float) -> float:
    dx, dy = bx - ax, by - ay
    length_sq = dx * dx + dy * dy
    if length_sq == 0:
        return ((px - ax) ** 2 + (py - ay) ** 2) ** 0.5

    t = ((px - ax) * dx + (py - ay) * dy) / length_sq
    t = min(max(t, 0.0), 1.0)
    return ((px - (ax + t * dx)) ** 2 + (py - (ay + t * dy)) ** 2) ** 0.5


def on_check(x: float, y: float) -> bool:
    half = STROKE / 2
    (ax, ay), (bx, by), (cx, cy) = CHECK
    return (
        distance_to_segment(x, y, ax, ay, bx, by) <= half
        or distance_to_segment(x, y, bx, by, cx, cy) <= half
    )


def render(size: int) -> bytes:
    """Return BGRA pixels, top row first."""
    hi_res = size * SUPERSAMPLE
    samples = SUPERSAMPLE * SUPERSAMPLE
    pixels = bytearray()

    # Sample once per hi-res pixel, then average each SUPERSAMPLE x SUPERSAMPLE block.
    mask = []
    for j in range(hi_res):
        row = []
        y = (j + 0.5) / hi_res
        for i in range(hi_res):
            x = (i + 0.5) / hi_res
            if not inside_rounded_rect(x, y):
                row.append(None)
            elif on_check(x, y):
                row.append(GLYPH)
            else:
                row.append(PLATE)
        mask.append(row)

    for j in range(size):
        for i in range(size):
            r = g = b = a = 0
            for sy in range(SUPERSAMPLE):
                for sx in range(SUPERSAMPLE):
                    sample = mask[j * SUPERSAMPLE + sy][i * SUPERSAMPLE + sx]
                    if sample is None:
                        continue
                    r += sample[0]
                    g += sample[1]
                    b += sample[2]
                    a += 255

            if a == 0:
                pixels += b"\x00\x00\x00\x00"
                continue

            covered = a // 255
            # Premultiplied-free straight alpha: colour is the average of covered samples.
            pixels += bytes((b // covered, g // covered, r // covered, a // samples))

    return bytes(pixels)


def to_dib(size: int, bgra_top_down: bytes) -> bytes:
    """Wrap pixels in the BITMAPINFOHEADER + XOR + AND layout an .ico entry expects."""
    header = struct.pack(
        "<IiiHHIIiiII",
        40,          # biSize
        size,        # biWidth
        size * 2,    # biHeight: XOR mask and AND mask stacked
        1,           # biPlanes
        32,          # biBitCount
        0,           # biCompression = BI_RGB
        len(bgra_top_down),
        0, 0, 0, 0,
    )

    # DIB rows run bottom-up.
    stride = size * 4
    xor = b"".join(
        bgra_top_down[row * stride:(row + 1) * stride]
        for row in range(size - 1, -1, -1)
    )

    # 1bpp AND mask, rows padded to 4 bytes. Fully zero: alpha already carries the shape.
    mask_stride = ((size + 31) // 32) * 4
    and_mask = b"\x00" * (mask_stride * size)

    return header + xor + and_mask


def build(path: str) -> None:
    images = [to_dib(size, render(size)) for size in SIZES]

    directory = struct.pack("<HHH", 0, 1, len(SIZES))
    offset = 6 + 16 * len(SIZES)
    entries = b""
    for size, image in zip(SIZES, images):
        entries += struct.pack(
            "<BBBBHHII",
            size if size < 256 else 0,
            size if size < 256 else 0,
            0, 0, 1, 32,
            len(image),
            offset,
        )
        offset += len(image)

    with open(path, "wb") as handle:
        handle.write(directory + entries + b"".join(images))

    print(f"wrote {path} ({offset} bytes, {len(SIZES)} sizes)")


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "flowdeck.ico")
