# -*- coding: utf-8 -*-
"""
Maxwell 3D - 앵글링 생성 V1
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_csv_data(csv_file_path):
    """CSV 파일에서 앵글링 데이터를 읽어옵니다."""
    # TODO: CSV 구조에 맞게 구현 예정
    pass


def create_anglings_from_csv(csv_file_path, name_prefix="Angling"):
    """CSV 파일에서 앵글링을 생성 (V1)"""
    print("=" * 60)
    print("Maxwell 3D - 앵글링 생성 V1")
    print("=" * 60)

    print("\nCSV 파일 읽기 시작...")
    # TODO: 구현 예정

    print("\n완료!")


# 스크립트 실행
print("\n" + "=" * 60)
print("스크립트 실행 시작")
print("=" * 60)

try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print("스크립트 디렉토리 (__file__): {}".format(script_dir))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        print("스크립트 디렉토리 (sys.argv[0]): {}".format(script_dir))
    else:
        script_dir = os.getcwd()
        print("스크립트 디렉토리 (getcwd): {}".format(script_dir))

csv_file = os.path.join(script_dir, "BarrierDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n경고: BarrierDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    print("\n앵글링 생성 함수 호출 중...")
    # 앵글링 생성
    create_anglings_from_csv(
        csv_file_path=csv_file,
        name_prefix="Angling"
    )
    print("\n스크립트 실행 완료!")
