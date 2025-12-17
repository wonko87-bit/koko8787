# Simcenter Magnet 3D Box Modeling Script

Simcenter Magnet을 사용하여 3D 직육면체(Box)를 모델링하는 Python 스크립트입니다.

## 기능

- 원점 (0, 0, 0)을 기준으로 양의 방향으로 확장되는 직육면체 생성
- 가로(width), 세로(depth), 높이(height)를 입력 변수로 사용
- 재질(material) 지정 가능
- 두 가지 생성 방법 제공:
  - `create_box()`: 간단한 박스 생성 메서드
  - `create_box_advanced()`: 고급 메서드 (더 많은 제어 가능)

## 요구사항

- Simcenter Magnet 소프트웨어
- Python (Simcenter Magnet에 내장된 Python 환경)

## 사용 방법

### 1. Simcenter Magnet 스크립트 콘솔에서 실행

```python
# 스크립트 로드
exec(open('create_box.py').read())

# 기본 박스 생성 (100 x 50 x 30 mm)
box = create_box(100, 50, 30)

# 사용자 정의 크기로 박스 생성
box = create_box(200, 100, 75, "Steel")
```

### 2. 명령줄에서 실행

```bash
# 기본 크기 사용
python create_box.py

# 사용자 정의 크기 (가로 세로 높이)
python create_box.py 150 80 60

# 재질 지정
python create_box.py 150 80 60 Steel
```

## 함수 설명

### `create_box(width, depth, height, material_name="Air")`

간단한 방법으로 직육면체를 생성합니다.

**매개변수:**
- `width` (float): X 방향 너비 (mm)
- `depth` (float): Y 방향 깊이 (mm)
- `height` (float): Z 방향 높이 (mm)
- `material_name` (str, optional): 재질 이름 (기본값: "Air")

**반환값:**
- MagNet Component 객체

**예제:**
```python
# 100mm x 50mm x 30mm 크기의 공기 박스
box = create_box(100, 50, 30)

# 200mm x 100mm x 75mm 크기의 철 박스
steel_box = create_box(200, 100, 75, "Steel")
```

### `create_box_advanced(width, depth, height, material_name="Air", name="Box")`

고급 메서드로 직육면체를 생성합니다. 더 많은 제어가 필요한 경우 사용합니다.

**매개변수:**
- `width` (float): X 방향 너비 (mm)
- `depth` (float): Y 방향 깊이 (mm)
- `height` (float): Z 방향 높이 (mm)
- `material_name` (str, optional): 재질 이름 (기본값: "Air")
- `name` (str, optional): 컴포넌트 이름 (기본값: "Box")

**반환값:**
- MagNet Component 객체

**예제:**
```python
# 사용자 정의 이름을 가진 박스
custom_box = create_box_advanced(150, 80, 60, "Copper", "MyCustomBox")
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

```python
# 50mm x 50mm x 10mm 크기의 네오디뮴 자석
magnet = create_box(50, 50, 10, "NdFeB")
```

### 예제 2: 여러 박스 생성

```python
# 공기 영역 생성
air_region = create_box(500, 500, 500, "Air")

# 자석 코어
core = create_box(100, 100, 50, "Steel")

# 코일 영역
coil = create_box(120, 120, 30, "Copper")
```

## 주의사항

1. 모든 차원 값은 양수여야 합니다
2. 재질 이름은 Simcenter Magnet의 재질 라이브러리에 정의되어 있어야 합니다
3. 스크립트는 Simcenter Magnet 환경 내에서 실행되어야 합니다

## 라이선스

이 스크립트는 교육 및 연구 목적으로 자유롭게 사용할 수 있습니다.
