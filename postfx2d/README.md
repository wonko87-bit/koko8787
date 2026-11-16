# PostFX 2D — 2D 해석결과 이미지 후처리

해석결과 캡처 이미지(클립보드 / 파일)를 받아 **후처리된 PNG** 와 **요약 데이터 테이블(CSV)** 을 만드는 Windows 프로그램입니다.

## 모드 1 · 최소안전율부위

| 점 색상 | 의미 |
|---|---|
| 🟠 주황 | W1 |
| 🟢 초록 | W2 |
| 🔵 파랑 | W3 |
| 🟣 보라 | W4 |
| 🔴 빨강 | Global Minimum Point |

점 옆 숫자 = 해당 지점의 안전율

1. 색상(HSV)으로 원형 점 검출 → W1~W4 / Global 분류
2. OCR(RapidOCR)로 숫자 인식 → 가장 가까운 점과 짝짓기
3. 점+숫자 군집을 거리 기반으로 자동 그룹화 (왼쪽부터 G1, G2, …)
4. 그룹 전체를 감싸는 **연한 파스텔 음영** (그룹별 다른 색)
5. 그룹별 최소 안전율 위치 **노란 하이라이트**, 빨간 점(글로벌 최소)은 **빨간 하이라이트**

출력: `이름_minSF.png` + `이름_minSF_summary.csv` (Excel 호환 UTF-8)

| 원본(합성 샘플) | 결과 | 결과 (정리 모드) |
|---|---|---|
| ![](tools/sample.png) | ![](docs/sample_result.png) | ![](docs/sample_result_clean.png) |

**정리 모드** (`최소 지점 외 점·숫자 지우기` 체크 / CLI `--clean`): 하이라이트된 최소 지점만 남기고 나머지 점·숫자는 주변 배경으로 지워 음영만 남깁니다.

## 사용법 (exe)

GitHub Actions → **PostFX2D Windows Build** → Artifact `PostFX2D-windows` 에서 `PostFX2D.exe` 다운로드.

- `Ctrl+V` 클립보드 캡처 붙여넣기 (Win+Shift+S 로 캡처 후) — 붙여넣으면 바로 분석
- `Ctrl+O` 파일 열기 · `Ctrl+S` PNG+CSV 저장 · `Space` 원본/결과 전환
- 그룹이 잘못 나뉘면 **그룹 수** 를 지정 후 다시 실행
- OCR 이 숫자를 잘못 읽으면 **전체 점** 탭에서 더블클릭해 수정 → 자동 재그리기

## 개발

```bash
pip install -r requirements.txt
python run_app.py                       # GUI
python cli.py 입력.png [출력.png] [--groups 6]
python tools/make_sample.py tools/sample.png
```
