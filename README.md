# 전자기 시뮬레이션 유틸리티 스크립트

Ansys Maxwell과 Simcenter Magnet을 위한 스크립트 모음입니다.

## 파일 구성

### Ansys Maxwell 3D 박스 모델링 스크립트
- **maxwell_create_box_simple.py**: 간단한 Python 버전 (추천 - 빠른 시작)
- **maxwell_create_box_simple.vbs**: 간단한 VBScript 버전 (추천 - 빠른 시작)
- **maxwell_create_box.py**: 전체 기능 Python 스크립트 (함수 포함)
- **maxwell_create_box.vbs**: 전체 기능 VBScript 스크립트 (함수 포함)
- **maxwell_create_core_from_csv.py**: CSV 기반 변압기 철심 생성 (Python)
- **maxwell_create_core_from_csv.vbs**: CSV 기반 변압기 철심 생성 (VBScript)
- **transformer_core_sample.csv**: 샘플 CSV 파일

### Simcenter Magnet 3D 박스 모델링 스크립트
- **create_box.vbs**: 전체 기능을 포함한 박스 생성 스크립트
- **create_box_simple.vbs**: 빠르게 사용 가능한 간단한 박스 생성 버전
- **create_box.py**: Python 버전 (참고용)

### 컴포넌트 이름 변경 스크립트
- **rename_components.vbs**: 대화형 컴포넌트 이름 일괄 변경 스크립트
- **rename_components_simple.vbs**: 간단한 버전의 이름 변경 스크립트

## 기능

### 1. Maxwell 3D 박스 모델링
- **위치**: 원점 (0, 0, 0)에서 시작
- **기본 크기**: 20mm (X) × 10mm (Y) × 5mm (Z)
- **재질 지정**: vacuum, aluminum, copper, steel, iron 등
- **자동 설정**: 프로젝트 및 디자인 자동 생성
- **생성 방법**:
  - `create_box_simple()`: 기본 박스 빠른 생성
  - `create_maxwell_box()`: 사용자 정의 박스
  - `CreateBoxInteractive()`: 대화형 입력 (VBScript)
  - `create_multiple_boxes()`: 여러 박스 일괄 생성
  - `create_transformer_core_from_csv()`: CSV 파일 기반 변압기 철심 생성

### 1-1. CSV 기반 변압기 철심 생성 (Yoke 포함!)
- **완전한 철심 구조**: Legs (3개) + Yokes (2개) = 완성된 변압기 철심
- **1x2 구조**:
  - Main leg 1개 (중앙)
  - Return legs 2개 (양쪽)
  - Top yoke (상단 수평 연결)
  - Bottom yoke (하단 수평 연결)
- **Unite 연산**: 5개 박스가 자동으로 하나의 철심으로 통합
- **CSV 입력**: 철심 치수 데이터를 CSV 파일로 관리
- **다층 구조**: 여러 레이어의 철심을 한 번에 생성
- **중심 정렬**: 모든 박스는 원점(0,0,0)을 중심으로 자동 배치
- **CSV 구조**:
  - A열: X1 (Main leg의 X 크기)
  - B열: X2 (Return leg의 X 크기)
  - C열: Y (Y축 크기)
  - E1 셀: Return leg 이격거리 (Main leg 중점 ↔ Return leg 중점)
  - E2 셀: 철심 창 높이 (Z축 extrusion)

### 2. Simcenter Magnet 3D 박스 모델링
- 원점 (0, 0, 0)을 기준으로 양의 방향으로 확장되는 직육면체 생성
- 가로(width), 세로(depth), 높이(height)를 입력 변수로 사용
- 재질(material) 지정 가능
- 여러 생성 방법 제공:
  - `CreateBox`: 기본 박스 생성
  - `CreateBoxAdvanced`: 고급 제어 가능한 방법
  - `CreateParametricBox`: 대화형 입력으로 박스 생성

### 3. 컴포넌트 이름 일괄 변경
- 대화형으로 여러 컴포넌트의 이름을 연속으로 변경
- 접두사(prefix) 자동 추가 기능
- ESC 또는 취소로 종료할 때까지 반복
- 작업 흐름:
  1. 스크립트 실행 시 prefix 입력
  2. 컴포넌트 선택
  3. 이름 입력 (자동으로 prefix_이름 형식으로 변경)
  4. 다음 컴포넌트 선택 반복
  5. ESC 또는 취소 시 종료

## 요구사항

- Ansys Maxwell Electronics Desktop (Maxwell 스크립트용)
- Simcenter Magnet (Magnet 스크립트용)

## 사용 방법

## A. Maxwell 3D 박스 모델링 스크립트

### 방법 1: 간단한 스크립트 사용 (추천 - 가장 빠름)

**Python 버전:**
1. **Ansys Maxwell Electronics Desktop 2021 R1**을 실행합니다
2. **maxwell_create_box_simple.py** 파일을 엽니다
3. 상단의 파라미터 값을 원하는 크기로 수정합니다:
   ```python
   box_width = 20      # X 방향 너비 (mm)
   box_depth = 10      # Y 방향 깊이 (mm)
   box_height = 5      # Z 방향 높이 (mm)
   box_name = "Box1"   # 박스 이름
   material = "vacuum" # 재질
   ```
4. **Tools → Run Script** 메뉴에서 스크립트를 실행합니다

**VBScript 버전:**
1. **Ansys Maxwell Electronics Desktop**을 실행합니다
2. **maxwell_create_box_simple.vbs** 파일을 엽니다
3. 상단의 파라미터 값을 원하는 크기로 수정합니다:
   ```vbscript
   boxWidth = 20        ' X 방향 너비 (mm)
   boxDepth = 10        ' Y 방향 깊이 (mm)
   boxHeight = 5        ' Z 방향 높이 (mm)
   boxName = "Box1"     ' 박스 이름
   materialName = "vacuum"  ' 재질
   ```
4. **Tools → Run Script**에서 스크립트를 실행합니다

### 방법 2: 전체 기능 스크립트 사용 (함수 호출)

**Python 버전:**
1. **maxwell_create_box.py** 파일을 엽니다
2. 스크립트 하단의 메인 실행 부분에서 원하는 방법을 선택합니다:
   ```python
   # 방법 1: 기본 박스 생성 (20x10x5)
   create_box_simple()

   # 방법 2: 사용자 정의 박스 생성
   # create_maxwell_box(100, 50, 30, "MyCustomBox", "aluminum")
   ```
3. **Tools → Run Script** 메뉴에서 스크립트를 실행합니다

**VBScript 버전:**
1. **maxwell_create_box.vbs** 파일을 엽니다
2. 스크립트 하단에서 실행할 함수를 선택합니다:
   ```vbscript
   ' 방법 1: 기본 박스 생성 (20x10x5)
   Call CreateBoxSimple()

   ' 방법 2: 대화형 입력으로 박스 생성
   ' Call CreateBoxInteractive()
   ```
3. **Tools → Run Script**에서 스크립트를 실행합니다

### Maxwell 스크립트 특징

- **실제 Maxwell 2021 R1 API 기반**: 실제 녹화된 스크립트를 기반으로 작성
- **0,0,0 위치**에서 시작하여 양의 방향으로 확장되는 박스 생성
- 기본 크기: **20mm × 10mm × 5mm** (요청하신 사양)
- 자동으로 프로젝트와 Maxwell 3D 디자인 생성
- 다양한 재질 지원 (vacuum, aluminum, copper, steel, iron 등)
- 대화형 입력 모드 지원 (전체 기능 버전)
- 간단한 버전과 전체 기능 버전 제공

### Maxwell 스크립트 예제

**간단한 버전 (파라미터 직접 수정):**
```python
# maxwell_create_box_simple.py
box_width = 100
box_depth = 50
box_height = 30
box_name = "MyCustomBox"
material = "aluminum"
# 그냥 실행하면 됨!
```

**전체 기능 버전 (함수 호출):**
```python
# maxwell_create_box.py
# Python에서 사용자 정의 박스 생성
create_maxwell_box(50, 30, 20, "MyBox", "copper")
```

```vbscript
' maxwell_create_box.vbs
' VBScript에서 사용자 정의 박스 생성
Call CreateMaxwellBox(50, 30, 20, "MyBox", "copper")
```

**실제 녹화된 API 형식:**
```python
# Maxwell에서 Tools → Record Script to File로 확인 가능
oEditor.CreateBox(
    ["NAME:BoxParameters", "XPosition:=", "0mm", ...],
    ["NAME:Attributes", "Name:=", "Box1", ...]
)
```

### 방법 3: CSV 기반 변압기 철심 생성 (신규!)

**Python 버전:**
1. **transformer_core_sample.csv** 파일을 참고하여 CSV 파일을 작성합니다
2. CSV 구조:
   ```
   X1,X2,Y,,75          (첫 행: 헤더 + E1값)
   50,30,100,,200       (데이터 행: X1,X2,Y + E2값)
   48,28,95,,,          (이후 데이터 행)
   ...
   ```
3. **maxwell_create_core_from_csv.py** 파일을 엽니다
4. CSV 파일 경로를 수정합니다:
   ```python
   csv_file = r"C:\path\to\your\transformer_core_data.csv"
   core_material = "steel_1008"  # 재질 설정
   ```
5. **Tools → Run Script**로 실행합니다

**VBScript 버전:**
1. CSV 파일을 작성합니다 (위와 동일)
2. **maxwell_create_core_from_csv.vbs** 파일을 엽니다
3. CSV 파일 경로를 수정합니다:
   ```vbscript
   csvFile = "C:\path\to\your\transformer_core_data.csv"
   coreMaterial = "steel_1008"  ' 재질 설정
   ```
4. **Tools → Run Script**로 실행합니다

**CSV 파일 예제:**
```csv
X1,X2,Y,,75
50,30,100,,200
48,28,95,,,
46,26,90,,,
44,24,85,,,
```
- E1 (첫 행 E열): 75mm (Return leg 이격거리)
- E2 (둘째 행 E열): 200mm (철심 창 높이)
- 각 행: Main leg와 Return legs의 크기 정의

**생성되는 구조:**
각 행마다 완전한 철심 생성 (5개 박스 → Unite → 1개 철심):
- Main leg (중앙): X1 × Y × E2
- Left Return leg: X2 × Y × E2
- Right Return leg: X2 × Y × E2
- Top yoke (상단): (2*E1 + 2*X2) × X2 × Y
- Bottom yoke (하단): (2*E1 + 2*X2) × X2 × Y

**좌표 계산:**
- 모든 박스는 원점(0,0,0)을 중심으로 자동 배치
- Leg Z 범위: -E2/2 ~ +E2/2
- Top yoke Z 범위: +E2/2 ~ +E2/2+Y
- Bottom yoke Z 범위: -E2/2-Y ~ -E2/2
- 전체 철심 높이: E2 + 2*Y (완벽한 중심 정렬)

**Unite 연산:**
5개의 개별 박스가 자동으로 하나의 철심 객체로 통합되어 매끄럽게 연결됩니다.

## B. Simcenter Magnet 3D 박스 모델링 스크립트

### 방법 1: 간단한 스크립트 사용 (추천)

1. Simcenter Magnet을 실행합니다
2. **create_box_simple.vbs** 파일을 엽니다
3. 상단의 파라미터 값을 원하는 크기로 수정합니다:
   ```vbscript
   boxWidth = 100.0     ' 가로 (mm)
   boxDepth = 50.0      ' 세로 (mm)
   boxHeight = 30.0     ' 높이 (mm)
   boxMaterial = "Air"  ' 재질
   ```
4. **Tools → Run Script** 메뉴에서 스크립트를 실행합니다

### 방법 2: 대화형 입력 사용

1. **create_box.vbs** 스크립트를 실행합니다
2. 대화형 입력 버전을 사용하려면 스크립트 마지막 줄을 수정:
   ```vbscript
   ' Call Main()  ' 이 줄을 주석 처리
   Call CreateParametricBox()  ' 이 줄의 주석 해제
   ```
3. 실행하면 대화상자에서 크기와 재질을 입력할 수 있습니다

### 방법 3: Simcenter Magnet 스크립트 에디터에서 직접 실행

Simcenter Magnet의 스크립트 콘솔에서 직접 코드를 입력할 수 있습니다:

```vbscript
' 간단한 박스 생성
Dim oDoc, oView, oComp
Set oDoc = GetDocument()
Set oView = oDoc.GetView()
Set oComp = oView.NewComponent("MyBox")
Call oComp.MakeComponentInABox(0, 0, 0, 100, 50, 30)
Call oComp.SetMaterial("Air")
Call oView.ViewAll()
```

## B. 컴포넌트 이름 변경 스크립트

### 방법 1: 간단한 버전 사용 (추천)

1. Simcenter Magnet에서 모델을 엽니다
2. **Tools → Run Script** 메뉴에서 **rename_components_simple.vbs** 선택
3. 대화상자에서 접두사(prefix)를 입력합니다 (예: "Part")
4. 모델에서 이름을 변경할 컴포넌트를 클릭하여 선택합니다
5. 대화상자에서 이름을 입력합니다 (예: "Core" → "Part_Core"로 변경됨)
6. 다음 컴포넌트를 선택하고 이름을 입력하는 과정을 반복합니다
7. 작업을 마치려면 컴포넌트 선택 또는 이름 입력 시 **취소** 버튼을 누릅니다

### 방법 2: 전체 기능 버전 사용

**rename_components.vbs** 파일은 세 가지 버전을 제공합니다:

1. **Main()** - 기본 버전 (상세한 안내 메시지)
2. **RenameComponentsFast()** - 빠른 버전 (확인 메시지 최소화)
3. **RenameComponentsAdvanced()** - 고급 버전 (이름 검증, 사용자 정의 구분자)

스크립트 맨 아래에서 원하는 버전의 주석을 해제하여 사용:

```vbscript
Call Main()                      ' 기본 버전
' Call RenameComponentsFast()    ' 빠른 버전
' Call RenameComponentsAdvanced() ' 고급 버전
```

### 사용 예시

**시나리오**: 자석 모델의 여러 부품 이름을 "Magnet_" 접두사로 통일

1. 스크립트 실행 → prefix: "Magnet" 입력
2. 첫 번째 컴포넌트 선택 → 이름: "Core" 입력 → "Magnet_Core"로 변경
3. 두 번째 컴포넌트 선택 → 이름: "Coil" 입력 → "Magnet_Coil"로 변경
4. 세 번째 컴포넌트 선택 → 이름: "Housing" 입력 → "Magnet_Housing"로 변경
5. 취소 버튼 클릭 → 작업 완료 메시지 (3개 컴포넌트 변경됨)

## 함수 설명

### 박스 생성 함수

### `CreateBox(w, d, h, material)`

간단한 방법으로 직육면체를 생성합니다.

**매개변수:**
- `w`: X 방향 너비 (mm)
- `d`: Y 방향 깊이 (mm)
- `h`: Z 방향 높이 (mm)
- `material`: 재질 이름

**예제:**
```vbscript
' 100mm x 50mm x 30mm 크기의 공기 박스
Call CreateBox(100, 50, 30, "Air")

' 200mm x 100mm x 75mm 크기의 철 박스
Call CreateBox(200, 100, 75, "Steel")
```

### `CreateBoxAdvanced(w, d, h, material, componentName)`

고급 메서드로 직육면체를 생성합니다. 엣지 압출(extrusion) 방식을 사용하여 더 많은 제어가 가능합니다.

**매개변수:**
- `w`: X 방향 너비 (mm)
- `d`: Y 방향 깊이 (mm)
- `h`: Z 방향 높이 (mm)
- `material`: 재질 이름
- `componentName`: 컴포넌트 이름

**예제:**
```vbscript
' 사용자 정의 이름을 가진 박스
Call CreateBoxAdvanced(150, 80, 60, "Copper", "MyCustomBox")
```

### `CreateParametricBox()`

대화상자를 통해 사용자로부터 직접 입력을 받아 박스를 생성합니다.

**예제:**
```vbscript
' 대화형 박스 생성
Call CreateParametricBox()
```

### 컴포넌트 이름 변경 함수

#### 간단한 버전 (rename_components_simple.vbs)

스크립트 전체가 하나의 실행 흐름으로 구성되어 있습니다. 파일을 실행하면 자동으로 시작됩니다.

#### 전체 기능 버전 (rename_components.vbs)

**`Main()`** - 기본 버전

상세한 안내 메시지와 함께 컴포넌트 이름을 변경합니다.

**`RenameComponentsFast()`** - 빠른 버전

확인 메시지를 최소화하여 빠르게 작업할 수 있습니다.

**`RenameComponentsAdvanced()`** - 고급 버전

- 사용자 정의 구분자 지정 가능 (기본: "_", "-", "." 등)
- 이름 검증 기능 (특수문자 차단)
- 변경 전 미리보기 제공

**예제:**
```vbscript
' 고급 버전 예시
' prefix: "Motor"
' separator: "-"
' 결과: Motor-Stator, Motor-Rotor, Motor-Housing
```

## 좌표계

스크립트는 다음 좌표계를 사용합니다:

```
        Z (높이)
        |
        |
        |_______ Y (세로/깊이)
       /
      /
     X (가로/너비)
```

박스의 한 꼭지점은 원점 (0, 0, 0)에 위치하고,
반대편 꼭지점은 (width, depth, height)에 위치합니다.

## 예제 시나리오

### 예제 1: 간단한 자석 모델링

```vbscript
' 50mm x 50mm x 10mm 크기의 네오디뮴 자석
Call CreateBox(50, 50, 10, "NdFeB")
```

### 예제 2: 여러 박스 생성

```vbscript
' create_box.vbs의 CreateMultipleBoxes 함수 사용
' 또는 직접 여러 박스 생성:

' 공기 영역 생성
Call CreateBox(500, 500, 500, "Air")

' 자석 코어
Call CreateBox(100, 100, 50, "Steel")

' 코일 영역
Call CreateBox(120, 120, 30, "Copper")
```

### 예제 3: 복잡한 형상 만들기

```vbscript
' 중첩된 박스 구조 생성
Dim oDoc, oView, oComp1, oComp2

Set oDoc = GetDocument()
Set oView = oDoc.GetView()

' 외부 케이스
Set oComp1 = oView.NewComponent("OuterCase")
Call oComp1.MakeComponentInABox(0, 0, 0, 200, 200, 100)
Call oComp1.SetMaterial("Steel")

' 내부 공기층
Set oComp2 = oView.NewComponent("InnerAir")
Call oComp2.MakeComponentInABox(10, 10, 10, 190, 190, 90)
Call oComp2.SetMaterial("Air")

Call oView.ViewAll()
```

### 예제 4: 컴포넌트 일괄 이름 변경

```vbscript
' rename_components_simple.vbs 스크립트 실행 후:
'
' 1. Prefix 입력: "Assembly"
' 2. 첫 번째 컴포넌트 선택 → 이름: "Base" → "Assembly_Base"
' 3. 두 번째 컴포넌트 선택 → 이름: "Frame" → "Assembly_Frame"
' 4. 세 번째 컴포넌트 선택 → 이름: "Cover" → "Assembly_Cover"
' 5. 취소 → 완료
'
' 결과: 3개의 컴포넌트가 일관된 명명 규칙으로 변경됨
```

## 주의사항

1. 모든 차원 값은 양수여야 합니다
2. 재질 이름은 Simcenter Magnet의 재질 라이브러리에 정의되어 있어야 합니다
3. 스크립트는 Simcenter Magnet 환경 내에서 실행되어야 합니다

## 라이선스

이 스크립트는 교육 및 연구 목적으로 자유롭게 사용할 수 있습니다.
