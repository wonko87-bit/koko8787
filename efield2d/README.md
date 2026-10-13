# efield2d

2D 전계 FEM 해석 (평면 / 축대칭, DC · AC). 개인·연구용.
FEMM 방식(1차 삼각요소, 스칼라 포텐셜)을 참고해 Python으로 새로 구현했습니다.

## 지원 해석

모두 `div(kappa grad V) = 0` 을 풀고 kappa만 다릅니다.

| mode | kappa | 용도 |
|---|---|---|
| `es` | eps0·eps_r | 정전계 (순시/임펄스, 정전용량) |
| `dc` | sigma | **DC 내전압** (DC 전계 분포는 유전율이 아니라 도전율이 결정) |
| `ac` | sigma + jω·eps0·eps_r | **AC 내전압** (복소 페이저, 손실 유전체) |

- `symmetry="planar"` (깊이 1 m) / `"axisym"` (x=r, y=z)
- 길이 단위는 mm (`Problem.unit`), 솔버·결과는 SI
- 기하는 DXF LWPOLYLINE과 같은 (x, y, bulge) 표현 → 직선/원호
- 겹침 규칙: 나중에 추가한 것이 우선, 도체가 항상 최우선 (큰 공기 영역 → 절연물 → 전극 순으로 그리면 됨)
- 전극/대칭 이외의 경계는 자연경계(Neumann). 축대칭의 r=0 축도 동일

## 사용 예

```python
from efield2d import *
air = Material("air", 1.0, e_allow=3e6)          # V/m
p = Problem(regions=[Region(rect(0, 0, 60, 60), air)],
            conductors=[Conductor(rect(10, 10, 20, 20), 1000.0, "hv"),
                        Conductor(rect(0, 0, 60, 3), 0.0, "gnd")],
            symmetry="axisym", mesh_size=3.0, edge_size=0.5)
mesh = build_mesh(p)
sol = solve(p, mesh, "ac", freq=50.0)           # 전압은 peak 값으로 입력
post.utilization(sol)                            # 영역별 Emax, 허용전계, 이용률
post.along_path(sol, Polygon.from_points([(0, 3), (60, 3)], closed=False))  # 연면전계
```

## 주의

- 날카로운 전극 모서리의 전계는 이론상 발산합니다. 메시를 줄일수록 Emax가 커지므로
  모서리 곡률을 반영하거나 모서리에서 떨어진 값으로 판정하세요.
- 요소당 전계는 상수(1차 요소)입니다. 재료 경계의 법선 전계는 불연속이므로
  `along_path(..., region=i)` 로 보고 싶은 쪽을 지정하세요 (접선 성분은 연속).

## 검증 (`pytest`)

해석해와 비교: 평행판(정확), 동축(축대칭), 동심구(원호 + 축대칭, 2차 수렴 확인),
2층 유전체(ES: ε비, DC: σ비, AC: 복소 κ비와 ω→0/∞ 극한), 에너지-정전용량 일치.
추가 검증은 같은 형상을 Maxwell 2D로 풀어 비교합니다.

## 로드맵

M1 솔버 코어 (완료) → M2 DXF 가져오기 → M3 후처리/내전압 판정 확장 → M4 GUI (PySide6) → M5 고도화
