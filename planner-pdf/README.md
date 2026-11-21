# 2027 하이퍼링크 플래너 (삼성노트용 PDF)

갤럭시탭 가로 화면(16:10)에 맞춘 495쪽짜리 연간·분기·월간·주간·일간 플래너입니다.
완성본은 `2027-planner.pdf` 입니다.

## 다시 만들기

```bash
pip install reportlab pikepdf korean_lunar_calendar
```

```bash
npm pack pretendard && mkdir -p fonts && tar -xzf pretendard-*.tgz --strip-components=5 -C fonts package/dist/public/static/alternative
```

```bash
python3 make_planner.py
```

공휴일(`HOLIDAYS`)·절기(`TERMS`)·색상·페이지 크기는 `make_planner.py` 상단에서 바꿀 수 있습니다.
