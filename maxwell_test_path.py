# -*- coding: utf-8 -*-
"""
경로 테스트 스크립트
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import os
import sys

print("=" * 60)
print("경로 테스트")
print("=" * 60)

# 방법 1: __file__ 사용
try:
    script_dir_1 = os.path.dirname(os.path.abspath(__file__))
    print("방법 1 (__file__): {}".format(script_dir_1))
except Exception as e:
    print("방법 1 실패: {}".format(e))
    script_dir_1 = None

# 방법 2: sys.argv 사용
try:
    if len(sys.argv) > 0:
        script_dir_2 = os.path.dirname(os.path.abspath(sys.argv[0]))
        print("방법 2 (sys.argv): {}".format(script_dir_2))
    else:
        print("방법 2: sys.argv가 비어있음")
        script_dir_2 = None
except Exception as e:
    print("방법 2 실패: {}".format(e))
    script_dir_2 = None

# 방법 3: getcwd 사용
try:
    script_dir_3 = os.getcwd()
    print("방법 3 (getcwd): {}".format(script_dir_3))
except Exception as e:
    print("방법 3 실패: {}".format(e))
    script_dir_3 = None

# CSV 파일 확인
print("\n" + "=" * 60)
print("CSV 파일 확인")
print("=" * 60)

csv_filename = "transformer_core_sample.csv"

for i, script_dir in enumerate([script_dir_1, script_dir_2, script_dir_3], 1):
    if script_dir:
        csv_path = os.path.join(script_dir, csv_filename)
        exists = os.path.exists(csv_path)
        print("방법 {}: {}".format(i, csv_path))
        print("  존재 여부: {}".format(exists))
        if exists:
            print("  파일 크기: {} bytes".format(os.path.getsize(csv_path)))

# 하드코딩 경로로 테스트
hardcoded_path = "/home/user/koko8787/transformer_core_sample.csv"
print("\n하드코딩 경로: {}".format(hardcoded_path))
print("  존재 여부: {}".format(os.path.exists(hardcoded_path)))

print("\n" + "=" * 60)
print("테스트 완료")
print("=" * 60)
