# -*- coding: utf-8 -*-
"""
Maxwell 3D - 앵글링 생성 V2 (디버그 버전)
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os
import math

print("=" * 60)
print("디버그 시작")
print("=" * 60)

try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print("스크립트 디렉토리: {}".format(script_dir))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.getcwd()
    print("스크립트 디렉토리: {}".format(script_dir))

csv_file = os.path.join(script_dir, "AnglingDim.csv")
print("CSV 파일: {}".format(csv_file))
print("CSV 존재: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("ERROR: CSV 파일이 없습니다!")
else:
    print("\n1. CSV 파일 읽기 테스트...")
    try:
        with open(csv_file, 'r') as f:
            reader = csv.reader(f)
            rows = list(reader)
        print("  총 {}행 읽음".format(len(rows)))

        print("\n2. 4행부터 데이터 읽기...")
        for i in range(3, min(len(rows), 10)):
            row = rows[i]
            print("  {}행: {}".format(i+1, row))

            if not row or all(not cell.strip() for cell in row):
                print("  -> 빈 행, 중단")
                break

            if len(row) >= 7:
                try:
                    quadrant = int(row[0].strip())
                    inner_r = float(row[1].strip())
                    thickness = float(row[2].strip())
                    x_bar = float(row[3].strip())
                    z_bar = float(row[4].strip())
                    ref_x = float(row[5].strip())
                    ref_z = float(row[6].strip())
                    print("  -> Q:{} R:{} T:{} XBar:{} ZBar:{} Ref:({},{})".format(
                        quadrant, inner_r, thickness, x_bar, z_bar, ref_x, ref_z))
                except Exception as e:
                    print("  -> 파싱 에러: {}".format(str(e)))
            else:
                print("  -> 데이터 부족 ({}개)".format(len(row)))

        print("\n3. Maxwell 프로젝트 접근 테스트...")
        oProject = oDesktop.GetActiveProject()
        if oProject is None:
            print("  활성 프로젝트 없음")
        else:
            print("  프로젝트: {}".format(oProject.GetName()))
            oDesign = oProject.GetActiveDesign()
            if oDesign is None:
                print("  활성 디자인 없음")
            else:
                print("  디자인: {}".format(oDesign.GetName()))

        print("\n성공!")

    except Exception as e:
        print("ERROR: {}".format(str(e)))
        import traceback
        traceback.print_exc()

print("\n" + "=" * 60)
print("디버그 종료")
print("=" * 60)
