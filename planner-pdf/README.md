# 하이퍼링크 플래너 (삼성노트용 PDF)

갤럭시탭 가로 화면(16:10)에 맞춘 연간·분기·월간·주간·일간 플래너입니다.
완성본은 `2027-planner.pdf` ~ `2030-planner.pdf` 이고, 연도마다 색 테마가 다릅니다
(2027 네이비, 2028 포레스트, 2029 플럼, 2030 모카).

2027년 공휴일은 우주항공청 「2027년 월력요항」과 일치합니다.
2028년 이후는 월력요항 발표 전이라 2026년 11월 현행 규정으로 계산한 예정값입니다.

## 다시 만들기

```bash
pip install reportlab pikepdf korean_lunar_calendar
```

```bash
npm pack pretendard && mkdir -p fonts && tar -xzf pretendard-*.tgz --strip-components=5 -C fonts package/dist/public/static/alternative
```

```bash
python3 make_planner.py 2027 2028 2029 2030
```

주 시작 요일(`WEEK_START`), 테마(`THEMES`), 공휴일 규칙, 페이지 크기는 `make_planner.py` 상단에서 바꿀 수 있습니다.

## 노트 · 간지 템플릿

`notebook-templates.pdf` (37쪽): 표지 → 목차 → 간지 10장 → 템플릿 목록 → 노트 템플릿 24종(각 1장).
템플릿 페이지를 삼성노트에서 복사해 원하는 간지 뒤에 붙여 쓰는 용도입니다.

```bash
python3 make_notebook.py
```

간지 수·색(`SEC`)과 템플릿 목록(`TEMPLATES`)은 `make_notebook.py` 에서 바꿀 수 있습니다.
