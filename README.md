# Simcenter Magnet 3D Box Modeling Script

Simcenter Magnet을 사용하여 3D 직육면체(Box)를 모델링하는 VBScript 스크립트입니다.

## 파일 구성

- **create_box.vbs**: 전체 기능을 포함한 메인 스크립트
- **create_box_simple.vbs**: 빠르게 사용 가능한 간단한 버전
- **create_box.py**: Python 버전 (참고용)

## 기능

- 원점 (0, 0, 0)을 기준으로 양의 방향으로 확장되는 직육면체 생성
- 가로(width), 세로(depth), 높이(height)를 입력 변수로 사용
- 재질(material) 지정 가능
- 여러 생성 방법 제공:
  - `CreateBox`: 기본 박스 생성
  - `CreateBoxAdvanced`: 고급 제어 가능한 방법
  - `CreateParametricBox`: 대화형 입력으로 박스 생성

## 요구사항

- Simcenter Magnet 소프트웨어

## 사용 방법

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

## 함수 설명

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

## 주의사항

1. 모든 차원 값은 양수여야 합니다
2. 재질 이름은 Simcenter Magnet의 재질 라이브러리에 정의되어 있어야 합니다
3. 스크립트는 Simcenter Magnet 환경 내에서 실행되어야 합니다

## 라이선스

이 스크립트는 교육 및 연구 목적으로 자유롭게 사용할 수 있습니다.
