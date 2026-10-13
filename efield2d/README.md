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

## DXF 가져오기

```python
m = import_dxf("model.dxf")           # 단위는 $INSUNITS로 mm 변환 (없으면 mm 가정 + 경고)
print(m.describe())                    # 도형 목록: id, 레이어, 종류, 면적, bbox
p = build_problem(m, {
    3: Assignment("region", air, name="air"),
    5: Assignment("conductor", voltage=10e3, name="hv"),
    6: Assignment("conductor", voltage=0.0, name="gnd"),
}, symmetry="axisym", mesh_size=4.0, edge_size=0.5)
```

터미널에서 바로 확인: `python -m efield2d model.dxf`

- 지원: LINE, ARC, CIRCLE, LWPOLYLINE, POLYLINE(2D), SPLINE, ELLIPSE, INSERT(블록 분해)
- **SPLINE/ELLIPSE는 짧은 직선으로 근사** (`flatten_tol`, 기본 0.01 mm = 곡선에서 최대 이탈)
- 닫힌 폴리라인/원/닫힌 스플라인은 도형 1개. 흩어진 LINE/ARC/열린 폴리라인은 끝점(`snap`, 기본 0.001 mm)으로 이어 붙이며, 3개 이상 만나는 분기점에서 끊김
- **영역/전극은 각각 닫힌 윤곽으로 그리는 것을 권장** (겹치거나 맞닿아도 됨. 영역은 면적 큰 것부터 깔고 작은 것이 덮어씀)
- 끝점이 살짝 안 맞아 닫히지 않은 윤곽은 경고로 알려 줌
- 축대칭 모델은 x(=r) ≥ 0 이어야 함

## GUI

```
pip install -r requirements.txt
python run_gui.py [model.dxf | project.efp]        # 또는  python -m efield2d.app
```

1. **파일 ▸ DXF 열기** — 도형 목록이 왼쪽에, 형상이 가운데에 나타납니다.
2. 도형을 **클릭**해 선택 (겹친 곳은 가장 작은 도형, Ctrl+클릭으로 다중 선택, 목록에서 선택해도 됨)
3. 왼쪽 **할당** 패널에서 종류(영역/도체/전위선) · 재료 · 전압(kV) · 이름을 정하고 **적용**
4. 오른쪽 **해석 설정**에서 종류(ES/DC/AC), 대칭, 요소 크기를 정하고 **F5**
   - AC 전압은 기본적으로 rms로 입력합니다 (√2배 하여 peak로 해석, 결과는 peak). 끌 수 있습니다.
5. 결과: |E| / 전위 색 분포, 등전위선, 메시, 마우스 위치 값(상태바), 영역별 Emax·허용전계·**이용률 표**,
   두 전극 사이 정전용량(C) / 컨덕턴스(G) / 어드미턴스(Y)
6. **연면전계** (오른쪽 `연면 경로` 탭)
   - **경로 지정**: `경로 선택 모드`를 켜고 캔버스에서 윤곽의 **선분/곡선 조각을 클릭**(닫힌 도형의 일부 구간도 가능, 다시 클릭하면 해제).
     **Shift+클릭**은 접선이 이어지는 구간 전체(직선+필렛+직선, 스플라인 등)를 한 번에 선택합니다.
     조각이 없으면 선택한 도형의 윤곽 전체를 경로로 씁니다. 고른 조각은 끝점으로 이어 붙여 하나의 경로로 만듭니다
     (이어지지 않으면 끊어진 위치를 알려 줌).
   - **평가**: 평가 쪽(영역: 인터페이스에서 어느 쪽 값을 볼지, `자동` = 어느 쪽이든), 점 수, 방향 뒤집기
     (기본 방향은 고전위 쪽이 시작), 임계 비교 성분(접선 |Et| / 전체 |E|), 임계 전계(0이면 평가 쪽 영역의 허용전계)
   - **결과**(`경로 (연면전계)` 탭): 접선/법선/|E|, 전위 |V|, **누적 ∫Et·ds**(전압 강하, 양 끝 전위차와 일치해야 함),
     **누적 ∫|Et|·ds**, **임계 이상인 구간의 누적 길이**, 평균 |Et|, 최대값과 위치, CSV 저장
   - 평가 쪽 영역에 접하지 않은 점은 값을 만들지 않고 NaN으로 처리하며 경고합니다.

**형상 편집**: 꼭짓점 표에서 좌표(x, y, bulge) 직접 수정·삽입·삭제, *꼭짓점 드래그 편집* 체크 후 사각 마커를 끌어 이동
(격자 스냅 지원), Δx/Δy 이동, 사각형/원/폴리라인 추가, 삭제, **Ctrl+Z / Ctrl+Y** 실행 취소·다시 실행.
**재료 편집**(편집 ▸ 재료): εr, σ, 허용전계 상수 또는 곡선(`x:kV/mm; x:kV/mm`, x는 해석 설정의 *허용전계 곡선 x* 로 평가).
프로젝트는 `.efp`(JSON)로 저장됩니다.

- 해석 입력이 바뀌면 결과 패널에 경고가 뜹니다 (경로용 보조선이나 허용전계 x 변경은 경고 대상 아님).
- DC에서 σ=0 인 재료(공기 등)는 전류가 흐르지 않는 영역으로 취급되어, 그 안의 전위는 접한 경계의 전위로 정해집니다.

## 주의

- 날카로운 전극 모서리의 전계는 이론상 발산합니다. 메시를 줄일수록 Emax가 커지므로
  모서리 곡률을 반영하거나 모서리에서 떨어진 값으로 판정하세요.
- 요소당 전계는 상수(1차 요소)입니다. 재료 경계의 법선 전계는 불연속이므로
  `along_path(..., region=i)` 로 보고 싶은 쪽을 지정하세요 (접선 성분은 연속).

## 검증 (`pytest`)

해석해와 비교: 평행판(정확), 동축(축대칭), 동심구(원호 + 축대칭, 2차 수렴 확인),
2층 유전체(ES: ε비, DC: σ비, AC: 복소 κ비와 ω→0/∞ 극한), 에너지-정전용량 일치.
추가 검증은 같은 형상을 Maxwell 2D로 풀어 비교합니다.

## 수식 계산기 (`계산기` 탭, `efield2d/calc.py`)

해석 결과에 수식을 적용해 값을 뽑는 도구입니다 (FEM 도구의 Field Calculator에 해당). 모두 SI 단위.

| 계산 종류 | 하는 일 |
|---|---|
| 한 점의 값 | 점 (x, y) mm 에서 수식 값 |
| 경로 프로파일 + 선적분 | 경로를 따라 f 를 그리고 ∫f dl, 평균, 최대. **비교식(해석해)** 을 넣으면 겹쳐 그리고 차이를 표시 |
| 영역 체적적분 | ∫f dV (전체 또는 한 영역; 평면은 깊이 1 m당, 축대칭은 전체 2πr) |
| 경로 플럭스 | 경로를 회전/연장한 면을 지나는 D, J, Jc, E 의 법선 성분 적분 (Gauss 법칙 확인 등). 법선은 진행 방향의 오른쪽 |
| 전극 전하/전류 | 전극 노드 잔차합 ΣKφ — ES: 전하 [C], DC/AC: 전류 [A]. Gauss 면적분보다 정확 |
| 일관성 검토 요약 | ∫ε\|E\|²dV ↔ ½ΣQ·V, ∫σ\|E\|²dV ↔ Re(ΣI·V*) 등 서로 다른 방법의 값이 일치하는지 |

**변수**: `x y`(=`r z`), `V`, `Ex Ey`(=`Er Ez`), `E`, `eps epsr sigma kappa w`, `D Dx Dy`, `J Jx Jy`(=(σ+jωε)E), `Jc Jcx Jcy`(=σE),
경로 계산에서는 `Et En Dt Dn Jt Jn …`. **함수**: `abs re im conj angle sqrt exp log log10 sin cos tan atan2 sinh cosh tanh mag minimum maximum where`.
**상수**: `pi eps0 mm`, 자동 상수 `kappa_<영역> eps_<영역> sigma_<영역> V_<전극> f`, 직접 입력(`a=5*mm; V0=1000`).
거듭제곱은 `**` 또는 `^`. 임의 코드는 실행되지 않습니다 (산술·비교·위 함수만 허용).

예 (모델 A): 해석해 비교식 `abs(V_inner/(log(2)*(1/kappa_L1+1/kappa_L2)))/(r*abs(where(r < 10*mm, kappa_L1, kappa_L2)))`.

**정확도**: 점·경로·플럭스는 노드 전위의 2차 다항식 패치 복원(PPR)으로 구한 연속 전계를 씁니다 (영역별이라 재료 경계의 불연속 유지).
요소별 상수 전계보다 훨씬 정확하고(모델 A에서 경계 포함 0.5 % 이내, Gauss 플럭스는 전극 전하와 0.4 % 이내),
끄면(`전계 복원` 체크 해제) 요소값 그대로입니다. 체적적분은 항상 요소값이라 에너지·손실 항등식이 기계 정밀도로 맞습니다.
복원 전계의 선적분은 전위차와 ~0.3 % 차이가 납니다 (끄면 거의 일치).

| 상용 FEM 도구에서 하던 것 | 이 계산기 |
|---|---|
| 선을 따라 E 플롯 | `경로 프로파일`, 수식 `E` / `Et` / `En` |
| ∫E·dl | `경로 프로파일`의 선적분, 수식 `Et` (또는 `Er`) |
| ∮D·dS 로 전하 | `경로 플럭스` (D), 또는 `전극 전하` |
| 에너지 / 손실 | `영역 체적적분`, `0.5*eps*E**2` / `sigma*E**2` |
| 정전용량, 어드미턴스 | `전극 전하`÷V, 또는 `일관성 검토 요약`의 참고 줄 |

## 벤치마크 (다른 FEM 도구와 교차 검증)

`benchmarks/BENCHMARKS.md` — Maxwell 2D 등과 비교할 수 있도록 형상 좌표, 재료, 경계조건, 이 프로그램의 결과값과 권장 허용 오차를
정리했습니다 (모델 A: 해석해가 있는 2층 동축, ES/DC/AC / 모델 B: 필렛 전극 + 스페이서 + 접지판, 연면 경로).
DXF와 프로젝트 파일은 `benchmarks/models/`, 재생성은 `python benchmarks/bench.py`.
**Maxwell 2D에서 모델 A를 확인하는 절차**(기대값 표, 계산기 사용 순서, 안 맞을 때 점검)는 `benchmarks/MAXWELL_GUIDE.md`,
비교용 E(r)/V(r) 표는 `benchmarks/models/modelA_profile_*.csv`.

## 로드맵

M1 솔버 코어 (완료) → M2 DXF 가져오기 (완료) → M4 GUI 1차 (완료) → M3 후처리/내전압 판정 확장 → M5 고도화
