"""화질 개선 (선택 옵션). 보기·저장용으로만 쓰고, 분석/측정 값은 항상 원본 픽셀에서 읽는다.

- 원본 (픽셀 그대로) : 확대 시 픽셀 블록 그대로 (가공 없음)
- 부드럽게 보기      : 확대 시 bicubic 보간으로 표시 (도트 현상 완화)
- 선명 확대 ×2/×4    : Lanczos 확대 + 언샤프 마스크. 저장 시에도 해당 배율로 저장

참고: 저해상도+JPEG 합성 rainbow map 검증에서 bicubic 보간이 픽셀 블록 표시보다
색→값 오차가 10~27% 작았다. '컬러맵 위치를 보간 후 재채색' 방식도 시험했지만
bicubic 보다 나아지지 않아 넣지 않았다.
"""
from __future__ import annotations

import cv2
import numpy as np

OPTIONS = ["원본 (픽셀 그대로)", "부드럽게 보기", "선명 확대 ×2", "선명 확대 ×4"]


def parse(option: str) -> tuple[str, int]:
    """옵션 문자열 → (종류, 배율). 종류: none / smooth / sharpen."""
    if option.startswith("선명"):
        return "sharpen", int(option[-1])
    if option.startswith("부드럽게"):
        return "smooth", 1
    return "none", 1


def sharpen_upscale(img: np.ndarray, k: int, amount: float = 0.7) -> np.ndarray:
    up = cv2.resize(img, None, fx=k, fy=k, interpolation=cv2.INTER_LANCZOS4)
    blur = cv2.GaussianBlur(up, (0, 0), 0.6 * k)
    return cv2.addWeighted(up, 1 + amount, blur, -amount, 0)


def enhance(img: np.ndarray, option: str) -> tuple[np.ndarray, int]:
    kind, k = parse(option)
    if kind == "sharpen":
        return sharpen_upscale(img, k), k
    return img, 1
