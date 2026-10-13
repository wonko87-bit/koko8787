# Maxwell 2D 에서 모델 A 결과 확인하기

이 문서는 `python benchmarks/bench.py` 로 자동 생성됩니다 (기대값은 **해석해**).

> **주의** — 메뉴 이름·버튼 위치는 Maxwell 2D / Electronics Desktop 버전에 따라 조금 다릅니다. 아래는 일반적인 흐름과 의도입니다.
> 화면과 다르면 **기대값 표의 숫자가 맞는지**로 판정하세요 (절차보다 숫자가 기준). 막히는 단계는 화면을 보내 주시면 맞춰 드립니다.

## 0. 기대값 (해석해) — Maxwell 결과가 이 값에 수렴해야 합니다

모델: 축대칭, r 5~10 mm 층 L1 (εr 2, σ 1e-12), r 10~20 mm 층 L2 (εr 4, σ 1e-13), 높이 20 mm, 전극 r=5 → 1000 V, r=20 → 0 V (자세한 형상은 `BENCHMARKS.md`).

| 확인 항목 | ES (정전계) | DC 전도 | AC 전도 (50 Hz) |
|---|---|---|---|
| 매트릭스 (Parameters ▸ Matrix) | C = **2.1403 pF** | G = **1.648e-14 S** (R = 6.067e+13 Ω) | G = **8.259e-14 S**, C = **2.1403 pF** (\|Y\| = 6.724e-10 S) |
| 전극 전하 / 전류 | Q = **2.1403 nC** | I = **1.648e-11 A** | \|I\| = **6.724e-07 A** |
| 에너지 / 손실 | 저장 에너지 W = **1.0701 µJ** (L1 0.71343, L2 0.35671) | 손실 P = **1.648e-08 W** | 평균 손실 ½GV² = **4.129e-08 W**, 평균 저장 에너지 ¼CV² = **0.53507 µJ** |
| \|E\| at r=7.5 mm, z=10 | **0.1282 kV/mm** | **0.01749 kV/mm** | **0.1282 kV/mm** (peak) |
| \|E\| at r=15 mm, z=10 | **0.03206 kV/mm** | **0.08744 kV/mm** | **0.03206 kV/mm** (peak) |
| \|V\| at r=7.5 / 10 / 15 mm | 610 / 333.3 / 138.3 V | 946.8 / 909.1 / 377.3 V | 610 / 333.3 / 138.3 V |
| ∫E·dl (r 5→20 mm) | 1000 V | 1000 V | 1000 V |

층 사이 전계 분배: ES·AC는 **L1(εr 작음)** 이 크고, DC는 **L2(σ 작음)** 가 큽니다 — 모드를 바꿨는데 분포가 안 바뀌면 재료 입력을 먼저 의심하세요.

메시 수렴된 이 프로그램 값, 그리고 r = 5 ~ 20 mm 를 0.25 mm 간격으로 적은 E(r), V(r) 표는 `models/modelA_profile_{es,dc,ac}.csv` 에 있습니다 (엑셀에서 Maxwell 내보내기 데이터와 겹쳐 비교).

## 1. 모델 만들기 (ES / DC / AC 공통)

1. Maxwell 2D 디자인 추가. **Maxwell 2D ▸ Solution Type** — *Electric* 아래에서 `Electrostatic`, `DC Conduction`, `AC Conduction` 중 하나, **Geometry mode = Cylindrical about Z** (축대칭).
2. **Modeler ▸ Units = mm**. 그림은 XY 평면에 그립니다 — 이 모드에서 **X = r(반지름), Y = z(축 방향)**.
3. **Draw ▸ Rectangle** 두 개: L1 = (X 5, Y 0) 크기 (5, 20), L2 = (X 10, Y 0) 크기 (10, 20). 이름은 L1, L2.
4. 재료: **Assign Material ▸ Add Material** 로 새 재료를 만들고 *Relative Permittivity*, *Bulk Conductivity* 를 표의 값으로 입력.
   L1: εr 2 / σ 1e-12 S/m,  L2: εr 4 / σ 1e-13 S/m. (ES는 σ 무시, DC는 εr 무시, AC는 둘 다 사용.) 영역 바깥의 배경(Region)은 만들지 마세요 — 모델은 L1, L2 두 개가 전부입니다.
5. 전극: **Edit ▸ Select ▸ Edges** 로 L1 의 안쪽 변(r = 5)을 고르고 **Excitations ▸ Assign ▸ Voltage** = **1000 V**, 바깥 변(L2의 r = 20)에 **0 V**.
   AC Conduction 은 크기 1000 V(**peak**), 위상 0° 로 넣고 Setup 에서 주파수 **50 Hz**.
6. 다른 경계(위·아래 끝면, r 축)는 **아무것도 지정하지 않습니다** (자연경계 Neumann). Balloon 금지.
7. **Analysis ▸ Add Solution Setup** (*Percent Error* 0.1 % 정도, 패스 충분히) ▸ **Validate** ▸ **Analyze**.

## 2. 계산기 없이 바로 확인 (먼저 이것부터)

- **정전용량 / 컨덕턴스**: **Maxwell 2D ▸ Parameters ▸ Assign ▸ Matrix** 에서 두 Voltage 선택 후 Analyze, 이어서 결과 창(*Solution Data*)의 **Matrix** 탭. ES 는 Capacitance, DC 는 Conductance, AC 는 Conductance 와 Capacitance 가 나옵니다. → 위 표의 C, G 와 비교.
- **에너지**: **Results ▸ Solution Data ▸ Convergence** 탭에 *Energy* (정전계: 저장 에너지) 가 나옵니다. → W = ½CV² 와 비교. 이 값이 위 표와 같으면 축대칭 전체(360°) 모델로 계산되고 있는 것입니다 (절반이면 대칭 승수/모드 확인).
- **전계·전위 분포**: **Fields ▸ Plot Fields ▸ Mag_E / Voltage** 로 영역을 골라 그려 보고, 층 경계에서 전계가 점프하는지 확인.

## 3. Field Calculator 로 확인

**Maxwell 2D ▸ Fields ▸ Calculator** 를 엽니다. 계산기는 **스택(RPN) 방식**입니다 — 값을 쌓고(`Input` 에서 Quantity, Geometry, Number…) 연산 버튼을 누르면 위의 값들을 소비합니다. 마지막에 **Eval** 로 숫자를 보고, 필요하면 **Add** 로 이름을 붙여 보관합니다.

### 3-A. 반경 방향 선 위의 E(r), V(r) 그래프 → 해석해와 비교
1. 확인용 선을 만듭니다: **Draw ▸ Line** 으로 **(5, 10) → (20, 10)** (z = 10 mm, r 5→20) 폴리라인. 속성에서 **Model 체크 해제** 권장 (형상 해석에 포함되지 않게).
2. **Results ▸ Create Fields Report ▸ Rectangular Plot**. *Geometry* = 그 Polyline, *Category* = Calculator Expressions(또는 Fields), *Quantity* = **Mag_E**, 다시 한 번 *Quantity* = **Voltage** 로 추가. X축은 선을 따른 거리 (Distance).
3. 그래프 위에서 우클릭 ▸ **Export Data** 로 CSV 저장 → 엑셀에서 `models/modelA_profile_*.csv` 의 `E_analytic` 열과 겹쳐 그립니다. 거리 d 와 r 의 관계는 r = 5 + d (mm).
4. 합격 기준: 전계 **±1 %**, 전위 **±0.2 %** (경계 r = 10 에서 전계가 점프하는 것도 확인).

### 3-B. ∫E·dl = 인가 전압 (전위차) 확인
1. 스택에 **Quantity ▸ E**, **Geometry ▸ Line ▸ (그 Polyline)** 을 올립니다.
2. 선의 접선 단위벡터를 올리고 (**Unit Vec** 메뉴의 *Tangent* — 버전에 따라 위치가 다름) **Dot** (E·t), 이어서 **∫** (적분) 후 **Eval**.
3. 기대값 **1000 V** (ES, DC, AC 모두 크기 기준). 부호는 선의 방향에 따라 ±.
   *Tangent 를 못 찾으면*: E 의 X 성분만 쓰면 됩니다 (이 선은 r 방향) — **Quantity ▸ E ▸ Scal? ▸ ScalarX**, **Geometry ▸ Line**, **∫**, **Eval**.

### 3-C. 층별 에너지 ½∫D·E dV (ES)
1. **Quantity ▸ D**, **Quantity ▸ E**, **Dot**, **Number ▸ Scalar 0.5**, **×** (곱).
2. **Geometry ▸ Volume ▸ L1**, **∫**, **Eval** → 기대 L1 = 0.71343 µJ.
3. 같은 방법으로 L2 → 기대 0.35671 µJ. 합 = 1.0701 µJ = ½CV² (Convergence 탭의 Energy 와도 같아야 함).
   *값이 이 값의 2π 분의 1* 이면 Volume 적분이 회전 방향을 포함하지 않은 것 (모드/대칭 승수 확인), *½ 배* 이면 에너지 정의(½ 누락)부터 확인.

### 3-D. 손실 ∫J·E dV (DC / AC)
1. **Quantity ▸ J**, **Quantity ▸ E**, **Dot**, (AC 는 복소 → **Real** 또는 크기 연산 주의) **Geometry ▸ Volume ▸ L1**, **∫**, **Eval**; L2 도 같은 방법으로 두 값을 더합니다.
2. 기대: DC **1.648e-08 W** (= G·V²),  AC 평균 손실 **4.129e-08 W** (= ½GV², peak 기준 — Maxwell 이 peak/RMS 중 무엇으로 보여 주는지 확인해서 √2², 즉 2배 차이를 감안).
   (σ 가 작아 값이 매우 작습니다 — 표시 자릿수를 늘리세요.)

### 3-E. (선택) Gauss 법칙으로 전하 — 축대칭에서는 까다로움
전극 전하는 3-절의 **Matrix 의 C·V** 가 가장 쉽고 정확합니다 (기대 2.1403 nC). 계산기로 하려면 반지름 r 의 원통면을 지나는 D 의 법선 성분을 적분해야 하며, 2D 축대칭에서는 선 적분에 **2π·r** 가중을 직접 넣어야 합니다: **D·Normal·2π·X** 를 선 (r = 7.5 mm, z 0→20) 위에서 적분. 되면 기대값은 위 Q 와 같고 (두 층 어디서 잡아도 같음 — D 의 법선 성분은 연속), 이 방식이 막히면 건너뛰어도 됩니다.

### 참고 — 점 전계(|E|)는 메시에 민감합니다 (실제로 확인된 사례)
정전용량과 에너지는 기본 자동 메시에서도 4~5자리가 맞지만, **한 점의 전계**는 전위의 미분이라 훨씬 민감합니다. 모델 A(ES)에서 자동 메시로는 E(r=7.5) 가 해석해보다 **1.6 % 낮게** 나왔고, **Mesh Operations ▸ Length Based 로 최대 요소 길이 0.1 mm** 를 주자 해석해(128 240 V/m) 와 일치했습니다. 점 전계를 비교할 때는 (1) 관심 영역에 길이 제한을 주고 Percent Error 를 낮추거나, (2) 전위 차분 ((V(r−δ)−V(r+δ))/2δ) 으로 구하세요. 전위·정전용량·에너지가 먼저 맞으면 모델은 맞는 것입니다.

## 4. 비교가 안 맞을 때 점검 순서

1. **C 가 크게 다르다** → 지오메트리 모드(축대칭이 아니라 Cartesian 이면 '깊이 1 m당' 값으로 다름), 단위(mm), 재료 εr 입력, Balloon 사용 여부.
2. **C 는 맞는데 점 전계만 1~2 % 다르다** → 메시 문제일 가능성이 큽니다 (위 참고). 길이 제한(예: 0.1~0.25 mm)을 주고 다시 해석. 그래도 다르면 선/점의 위치 (r = 5 + d), 경계가 아니라 층 내부에서 읽었는지 확인.
3. **DC 만 다르다** → 재료 σ 입력 (Bulk Conductivity), 해석 타입이 DC Conduction 인지 (ES 로 풀면 εr 분배가 나옴).
4. **AC 크기가 √2 배 / ½ 배** → peak 와 RMS 해석 차이: 우리는 **peak** (1000 V = 크기), Maxwell 도 입력한 값을 크기로 쓰는지, 결과 표시가 RMS 인지 확인.
5. 에너지·손실 적분이 2π 또는 ½ 어긋나면 위 3-C 의 주석.
6. **위 순서로도 안 풀리면**: 확인 항목별 Maxwell 값, 해석 설정(Solution type, Percent Error, 패스 수)을 알려 주세요. 우리 쪽 코드는 해석해와 일치함이 테스트로 고정되어 있어, 보통 설정 차이입니다.

## 5. 알려 주실 때

```
모드: ES / DC / AC
Matrix:    C = ...  또는 G = ...
에너지(Convergence 탭):  ... J        (ES)
E at r=7.5 / 15 (z=10):  ... / ... kV/mm
∫E·dl (r 5→20):  ... V
층별 에너지 L1 / L2:  ... / ... µJ
Percent Error, 패스 수, 요소 수:  ...
```
