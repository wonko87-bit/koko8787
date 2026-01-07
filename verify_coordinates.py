#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Maxwell 철심 좌표 계산 검증 스크립트
CSV 데이터를 읽어 각 박스의 좌표와 크기를 출력합니다.
"""

import csv

def verify_core_coordinates(csv_file_path):
    """
    CSV 파일로부터 철심 좌표를 계산하고 검증합니다.
    """

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    # E1, E2 읽기
    gap = float(rows[0][4]) if len(rows[0]) > 4 else 0
    window_height = float(rows[1][4]) if len(rows[1]) > 4 else 0

    print("=" * 80)
    print("Maxwell 변압기 철심 좌표 계산 검증")
    print("=" * 80)
    print(f"Return leg 이격거리 (E1): {gap} mm")
    print(f"철심 창 높이 (E2): {window_height} mm")
    print("=" * 80)

    # 데이터 행 처리
    for i in range(1, len(rows)):
        if len(rows[i]) < 3:
            continue

        try:
            x1 = float(rows[i][0])
            x2 = float(rows[i][1])
            y = float(rows[i][2])
        except ValueError:
            continue

        if x1 <= 0 or x2 <= 0 or y <= 0:
            continue

        print(f"\n{'='*80}")
        print(f"레이어 {i}: X1={x1}, X2={x2}, Y={y}")
        print(f"{'='*80}")

        # Main leg 계산
        print("\n[1] Main Leg (중앙)")
        main_x_pos = -x1 / 2.0
        main_y_pos = -y / 2.0
        main_z_pos = -window_height / 2.0
        print(f"  위치: ({main_x_pos:.2f}, {main_y_pos:.2f}, {main_z_pos:.2f})")
        print(f"  크기: {x1} × {y} × {window_height}")
        print(f"  범위: X=[{main_x_pos:.2f}, {main_x_pos+x1:.2f}]")
        print(f"       Y=[{main_y_pos:.2f}, {main_y_pos+y:.2f}]")
        print(f"       Z=[{main_z_pos:.2f}, {main_z_pos+window_height:.2f}]")

        # Left Return leg 계산
        print("\n[2] Left Return Leg")
        left_x_pos = -gap - x2 / 2.0
        left_y_pos = -y / 2.0
        left_z_pos = -window_height / 2.0
        print(f"  위치: ({left_x_pos:.2f}, {left_y_pos:.2f}, {left_z_pos:.2f})")
        print(f"  크기: {x2} × {y} × {window_height}")
        print(f"  범위: X=[{left_x_pos:.2f}, {left_x_pos+x2:.2f}]")
        print(f"       Y=[{left_y_pos:.2f}, {left_y_pos+y:.2f}]")
        print(f"       Z=[{left_z_pos:.2f}, {left_z_pos+window_height:.2f}]")

        # Right Return leg 계산
        print("\n[3] Right Return Leg")
        right_x_pos = gap - x2 / 2.0
        right_y_pos = -y / 2.0
        right_z_pos = -window_height / 2.0
        print(f"  위치: ({right_x_pos:.2f}, {right_y_pos:.2f}, {right_z_pos:.2f})")
        print(f"  크기: {x2} × {y} × {window_height}")
        print(f"  범위: X=[{right_x_pos:.2f}, {right_x_pos+x2:.2f}]")
        print(f"       Y=[{right_y_pos:.2f}, {right_y_pos+y:.2f}]")
        print(f"       Z=[{right_z_pos:.2f}, {right_z_pos+window_height:.2f}]")

        # Yoke 치수 계산
        yoke_x_size = 2 * gap + x2
        yoke_y_size = x2
        yoke_z_size = y

        # Top Yoke 계산
        print("\n[4] Top Yoke")
        top_yoke_x_pos = -yoke_x_size / 2.0
        top_yoke_y_pos = -yoke_y_size / 2.0
        top_yoke_z_pos = window_height / 2.0
        print(f"  위치: ({top_yoke_x_pos:.2f}, {top_yoke_y_pos:.2f}, {top_yoke_z_pos:.2f})")
        print(f"  크기: {yoke_x_size} × {yoke_y_size} × {yoke_z_size}")
        print(f"  범위: X=[{top_yoke_x_pos:.2f}, {top_yoke_x_pos+yoke_x_size:.2f}]")
        print(f"       Y=[{top_yoke_y_pos:.2f}, {top_yoke_y_pos+yoke_y_size:.2f}]")
        print(f"       Z=[{top_yoke_z_pos:.2f}, {top_yoke_z_pos+yoke_z_size:.2f}]")

        # Bottom Yoke 계산
        print("\n[5] Bottom Yoke")
        bottom_yoke_x_pos = -yoke_x_size / 2.0
        bottom_yoke_y_pos = -yoke_y_size / 2.0
        bottom_yoke_z_pos = -window_height / 2.0 - yoke_z_size
        print(f"  위치: ({bottom_yoke_x_pos:.2f}, {bottom_yoke_y_pos:.2f}, {bottom_yoke_z_pos:.2f})")
        print(f"  크기: {yoke_x_size} × {yoke_y_size} × {yoke_z_size}")
        print(f"  범위: X=[{bottom_yoke_x_pos:.2f}, {bottom_yoke_x_pos+yoke_x_size:.2f}]")
        print(f"       Y=[{bottom_yoke_y_pos:.2f}, {bottom_yoke_y_pos+yoke_y_size:.2f}]")
        print(f"       Z=[{bottom_yoke_z_pos:.2f}, {bottom_yoke_z_pos+yoke_z_size:.2f}]")

        # 전체 철심 범위 계산
        print("\n" + "-" * 80)
        print("전체 철심 범위 (Unite 후):")
        total_min_x = min(left_x_pos, top_yoke_x_pos)
        total_max_x = max(right_x_pos + x2, top_yoke_x_pos + yoke_x_size)
        total_min_y = -max(y, yoke_y_size) / 2.0
        total_max_y = max(y, yoke_y_size) / 2.0
        total_min_z = bottom_yoke_z_pos
        total_max_z = top_yoke_z_pos + yoke_z_size

        print(f"  X: [{total_min_x:.2f}, {total_max_x:.2f}] (중심: 0)")
        print(f"  Y: [{total_min_y:.2f}, {total_max_y:.2f}] (중심: 0)")
        print(f"  Z: [{total_min_z:.2f}, {total_max_z:.2f}] (중심: 0)")
        print(f"  전체 크기: {total_max_x - total_min_x:.2f} × {total_max_y - total_min_y:.2f} × {total_max_z - total_min_z:.2f}")

        # 검증: 중심이 원점인지 확인
        center_x = (total_min_x + total_max_x) / 2.0
        center_y = (total_min_y + total_max_y) / 2.0
        center_z = (total_min_z + total_max_z) / 2.0
        print(f"\n중심 좌표: ({center_x:.2f}, {center_y:.2f}, {center_z:.2f})")

        if abs(center_x) < 0.01 and abs(center_y) < 0.01 and abs(center_z) < 0.01:
            print("  ✓ 중심이 원점에 정확히 위치합니다!")
        else:
            print("  ✗ 경고: 중심이 원점에서 벗어났습니다!")

        # Yoke가 Return leg 외곽에 정확히 맞는지 확인
        left_outer = left_x_pos
        right_outer = right_x_pos + x2
        yoke_left = top_yoke_x_pos
        yoke_right = top_yoke_x_pos + yoke_x_size

        print(f"\nYoke 정렬 검증:")
        print(f"  Left Return 외곽: {left_outer:.2f}")
        print(f"  Yoke 왼쪽 끝:    {yoke_left:.2f}")
        print(f"  차이:            {abs(left_outer - yoke_left):.6f}")

        print(f"  Right Return 외곽: {right_outer:.2f}")
        print(f"  Yoke 오른쪽 끝:   {yoke_right:.2f}")
        print(f"  차이:            {abs(right_outer - yoke_right):.6f}")

        if abs(left_outer - yoke_left) < 0.01 and abs(right_outer - yoke_right) < 0.01:
            print("  ✓ Yoke가 Return leg와 정확히 정렬되었습니다!")
        else:
            print("  ✗ 경고: Yoke가 Return leg와 정렬되지 않았습니다!")

if __name__ == "__main__":
    csv_file = "/home/user/koko8787/transformer_core_sample.csv"
    verify_core_coordinates(csv_file)
