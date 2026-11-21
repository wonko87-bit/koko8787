#!/usr/bin/env python3
"""세로형 하이퍼링크 플래너 PDF 생성기 (삼성노트 · 갤럭시탭 세로 화면용).

    python3 make_planner_portrait.py [연도 ...]      # 생략하면 2027–2030 전부
    → <연도>-planner-세로형.pdf

공휴일·음력·테마·링크 후처리는 make_planner.py(가로형)를 그대로 쓰고,
페이지 크기와 배치만 세로(540×864pt, 16:10)로 바꾼다.
"""
import calendar
import datetime as dt
import os
import sys

import make_planner as mp

# ── 세로 규격: make_planner 의 전역값을 바꿔 Painter·크롬·미니 달력이 따라오게 한다 ──
mp.W, mp.H, mp.TABW = 540, 864, 26
mp.X0, mp.X1 = 18, 540 - 26 - 12
mp.CW = mp.X1 - mp.X0           # 484pt ≈ 171mm
mp.TOP = mp.NAVH + 10
mp.BT = mp.TOP + mp.HDR + 8     # 86
mp.BB = 864 - 12                # 852
mp.BH = mp.BB - mp.BT           # 766

MINE = set()


def _sync():
    """configure(연도) 뒤의 make_planner 전역값(날짜·공휴일·색)을 이 모듈로 가져온다"""
    globals().update({k: v for k, v in vars(mp).items() if not k.startswith("_") and k not in MINE})


_sync()


# ─────────────────────────────────────────────────────────────────────────────
def pp_cover(p):
    p.start("cover", "표지")
    p.box(0, 0, W, H, fill=WHITE)
    p.box(0, 0, W, 360, fill=ACC)
    p.text(40, 64, "PLANNER", "SB", 11, WHITE, cs=6)
    p.text(40, 82, "DAILY · WEEKLY · MONTHLY · YEARLY", "M", 6.6, ACC_M, cs=1.6)
    p.text(32, 278, str(YEAR), "B", 128, WHITE)
    p.text(40, 330, "한 해를 설계하고, 하루를 완성하다", "M", 9, ACC_M, cs=0.6)
    x, xr = 40, W - 40
    p.text(x, 410, "%d 플래너" % YEAR, "B", 26, INK)
    p.text(x, 430, "연간 · 분기 · 월간 · 주간 · 일간이 모두 연결된 하이퍼링크 플래너", "R", 8.6, SUB)
    p.hl(x, xr, 446, INK, 1.1)
    p.text(x, 474, "이름", "SB", 8, ACC)
    p.hl(x + 30, x + 210, 478, RULE, 0.7)
    p.text(x + 230, 474, "연락처", "SB", 8, ACC)
    p.hl(x + 266, xr, 478, RULE, 0.7)
    p.text(x, 504, "올해의 한 문장", "SB", 8, ACC)
    p.hl(x + 70, xr, 508, RULE, 0.7)
    p.text(x, 544, "바로가기", "B", 9, ACC)
    items = [("홈 · 사용법", "index"), ("연간 달력", "year"), ("연간 일정표", "glance"),
             ("가치 · 사명", "mission"), ("연간 목표", "goals1"), ("노트", "notesidx")]
    bw = (xr - x - 2 * 8) / 3
    for i, (label, key) in enumerate(items):
        p.pill(x + (i % 3) * (bw + 8), 554 + (i // 3) * 36, bw, 28, label, key, size=8.4)
    p.text(x, 646, "분기", "B", 9, ACC)
    qw = (xr - x - 3 * 8) / 4
    for q in range(1, 5):
        p.pill(x + (q - 1) * (qw + 8), 656, qw, 28, "Q%d · %d–%d월" % (q, q * 3 - 2, q * 3), "q%d" % q,
               fill=TINT, stroke=RULE, size=8)
    p.text(x, 714, "월", "B", 9, ACC)
    mw = (xr - x - 5 * 6) / 6
    for m in range(1, 13):
        p.pill(x + ((m - 1) % 6) * (mw + 6), 724 + ((m - 1) // 6) * 38, mw, 32, "%d" % m, "m%d" % m,
               fill=TINT, stroke=RULE, font="B", size=11)
    p.text(x, 822, "링크는 삼성노트 ‘읽기 모드’(책 아이콘)를 켠 상태에서 누르세요.", "R", 7.4, SUB)
    p.pill(xr - 90, 806, 90, 26, "시작하기", "index", fill=ACC, color=WHITE, size=9)
    p.chev(xr - 14, 819, 3, True, WHITE, 1.2)
    p.end()


def pp_index(p):
    p.start("index", "홈 · 사용법")
    p.chrome([(str(YEAR), "cover"), ("홈", "index")], menu="index")
    p.title("홈", "HOME · INDEX", "플래너의 모든 페이지로 이동하는 허브")
    y = BT
    p.section(X0, y, CW, "연간", "YEARLY")
    items = [("연간 달력", "year"), ("연간 일정표", "glance"), ("가치 · 사명 · 역할", "mission"),
             ("연간 목표 ①", "goals1"), ("연간 목표 ②", "goals2")]
    bw = (CW - 4 * 6) / 5
    for i, (label, key) in enumerate(items):
        p.pill(X0 + i * (bw + 6), y + 20, bw, 26, label, key, size=7.8)
    y += 56
    p.section(X0, y, CW, "분기", "QUARTERLY")
    qw = (CW - 3 * 6) / 4
    for q in range(1, 5):
        p.pill(X0 + (q - 1) * (qw + 6), y + 20, qw, 26, "Q%d · %d–%d월" % (q, q * 3 - 2, q * 3), "q%d" % q,
               fill=TINT, stroke=RULE, size=8)
    y += 56
    p.section(X0, y, CW, "월간 · 주간", "MONTHLY · WEEKLY")
    p.text(X1, y + 9.5, "달력 / 계획 / 첫 주", "R", 6.4, FAINT, "r")
    y += 20
    cw = (CW - 3 * 6) / 4
    ch = 64
    for m in range(1, 13):
        bx = X0 + ((m - 1) % 4) * (cw + 6)
        by = y + ((m - 1) // 4) * (ch + 6)
        p.box(bx, by, cw, ch, fill=WHITE, stroke=RULE, r=6, lw=0.6)
        p.text(bx + 9, by + 18, "%d월" % m, "B", 13, INK)
        p.text(bx + 9 + p.tw("%d월" % m, "B", 13) + 5, by + 18, MON_EN[m][:3], "M", 5.6, FAINT, cs=0.8)
        p.link(bx, by, cw, ch - 22, "m%d" % m)
        sw = (cw - 14 - 2 * 4) / 3
        first = dt.date(YEAR, m, 1)
        for i, (label, key) in enumerate([("달력", "m%d" % m), ("계획", "mp%d" % m),
                                          (wlabel(first), wkey(first))]):
            p.pill(bx + 7 + i * (sw + 4), by + ch - 22, sw, 16, label, key, size=6.6)
    y += 3 * (ch + 6) + 4
    p.section(X0, y, CW, "도구", "TOOLS")
    tools = [("기념일 · 생일", "anniv"), ("연간 기록 · 픽셀", "pixels"), ("연간 가계부", "finance"),
             ("독서 · 강의 목록", "reading"), ("버킷리스트", "bucket"), ("노트 색인", "notesidx")]
    bw = (CW - 2 * 6) / 3
    for i, (label, key) in enumerate(tools):
        p.pill(X0 + (i % 3) * (bw + 6), y + 20 + (i // 3) * 30, bw, 24, label, key, size=7.8)
    y += 20 + 60 + 4
    p.section(X0, y, CW, "프로젝트", "PROJECTS")
    pw = (CW - 5 * 6) / 6
    for i in range(6):
        bx = X0 + i * (pw + 6)
        p.box(bx, y + 20, pw, 32, fill=WHITE, stroke=RULE, r=6, lw=0.6)
        p.text(bx + 7, y + 32, "P%d" % (i + 1), "B", 8, ACC)
        p.hl(bx + 7, bx + pw - 7, y + 45, RULE_L, 0.6)
        p.link(bx, y + 20, pw, 32, "proj%d" % (i + 1))
    y += 20 + 32 + 10
    p.section(X0, y, CW, "노트", "NOTES")
    p.text(X1, y + 9.5, "줄 · 도트 · 모눈 각 10장", "R", 6.4, FAINT, "r")
    nw = (CW - 14 * 4) / 15
    for i in range(30):
        bx = X0 + (i % 15) * (nw + 4)
        by = y + 20 + (i // 15) * 26
        p.pill(bx, by, nw, 22, "%02d" % (i + 1), "n%d" % (i + 1), fill=TINT, stroke=RULE, font="M", size=7, )
    y += 20 + 52 + 10
    # 사용법 / 표기·연휴
    lw = 270
    p.box(X0, y, lw, BB - y, fill=TINT, r=6)
    p.text(X0 + 10, y + 16, "사용법", "B", 8, ACC)
    tips = ["링크는 삼성노트 ‘읽기 모드’(책 아이콘)에서 누릅니다.",
            "상단 왼쪽은 현재 위치, 오른쪽은 주요 메뉴입니다.",
            "오른쪽 1–12월 탭 → 월간 달력",
            "날짜 → 일간 · W번호 → 주간",
            "가치·사명 → 목표 → 분기 → 월 → 주 → 오늘",
            "A 필수 → B 중요 → C 선택 순서로 실행합니다."]
    for i, t in enumerate(tips):
        p.text(X0 + 10, y + 32 + i * 13.5, "%d. %s" % (i + 1, t), "R", 6.9, SUB)
    sx = X0 + 10
    sy = y + 32 + 6 * 13.5 + 6
    for s, body in [("✓", "완료"), ("→", "내일로"), ("×", "취소"), ("●", "진행"), ("Ⓓ", "위임")]:
        if s == "Ⓓ":
            p.circle(sx + 4, sy - 2.5, 4.2, stroke=INK, lw=0.6)
            p.text(sx + 4, sy - 0.4, "D", "B", 5.2, INK, "c")
        else:
            p.text(sx + 4, sy, s, "B", 8, INK, "c")
        p.text(sx + 11, sy, body, "R", 6.9, SUB)
        sx += 11 + p.tw(body, "R", 6.9) + 10
    if YEAR not in OFFICIAL_HOLIDAYS:
        p.text(X0 + 10, BB - 8, "※ 공휴일은 2026년 11월 현행 법령으로 계산한 예정값입니다.", "R", 6, FAINT)
    rx = X0 + lw + 10
    rw = X1 - rx
    p.box(rx, y, rw, BB - y, fill=TINT, r=6)
    p.text(rx + 10, y + 16, "달력 표기", "B", 8, ACC)
    for i, (k, col, body) in enumerate([("빨강", RED, "일요일 · 공휴일"), ("파랑", BLUE, "토요일"),
                                        ("음", SUB, "음력 · 절기 · 기념일")]):
        p.text(rx + 10, y + 31 + i * 12.5, k, "SB", 6.9, col)
        p.text(rx + 34, y + 31 + i * 12.5, body, "R", 6.9, SUB)
    ty = y + 31 + 3 * 12.5 + 8
    p.text(rx + 10, ty, "%d 연휴 (3일 이상)" % YEAR, "B", 8, ACC)
    for i, (name, a, b) in enumerate(long_breaks()[:10]):
        bx = rx + 10 + (i % 2) * ((rw - 20) / 2)
        by = ty + 6 + (i // 2) * 12.5
        p.text(bx, by + 9, name, "SB", 6.4, RED)
        rng = "%d.%d–%d" % (a.month, a.day, b.day) if a.month == b.month else \
            "%d.%d–%d.%d" % (a.month, a.day, b.month, b.day)
        p.text(bx + 30, by + 9, rng, "R", 6.4, SUB)
        p.link(bx - 2, by, (rw - 20) / 2, 12.5, dkey(a))
    p.end()


def pp_year(p):
    p.start("year", "연간 달력", 0)
    p.chrome([(str(YEAR), "cover"), ("연간 달력", "year")], menu="year")
    p.title(str(YEAR), "YEAR CALENDAR", "날짜 → 일간 · W번호 → 주간")
    for q in range(1, 5):
        p.pill(X1 - (5 - q) * 44 + 4, TOP + 10, 38, 22, "Q%d" % q, "q%d" % q, size=8.2)
    gx, gy = 18, 12
    mw = (CW - 2 * gx) / 3
    rowh = (BH - 3 * gy) / 4
    for m in range(1, 13):
        p.mini_month(X0 + ((m - 1) % 3) * (mw + gx), BT + ((m - 1) // 3) * (rowh + gy), mw, m,
                     rh=(rowh - 28) / 6)
    p.end()


def pp_glance(p):
    p.start("glance", "연간 일정표", 0)
    p.chrome([(str(YEAR), "cover"), ("연간 일정표", "glance")], menu="glance")
    p.title("연간 일정표", "YEAR AT A GLANCE", "한 해의 큰 일정을 한눈에")
    hh = 15
    cw = CW / 12
    rh = (BH - hh) / 31
    for m in range(1, 13):
        x = X0 + (m - 1) * cw
        p.box(x, BT, cw, hh, fill=ACC_L)
        p.ctext(x, BT, cw, hh, "%d월" % m, "B", 7.6, ACC)
        p.link(x, BT, cw, hh, "m%d" % m)
        nd = calendar.monthrange(YEAR, m)[1]
        for dd in range(1, 32):
            y = BT + hh + (dd - 1) * rh
            if dd > nd:
                p.box(x, y, cw, rh, fill=TINT)
                continue
            d = dt.date(YEAR, m, dd)
            col = day_color(d)
            if col != INK:
                p.box(x, y, cw, rh, fill=HexColor("#FBF3F2") if col == RED else HexColor("#F2F6FB"))
            p.text(x + 8, y + 9, str(dd), "SB", 6.6, col, "c")
            p.text(x + 8, y + 16.5, WD_KO[d.weekday()], "R", 5, col if col != INK else FAINT, "c")
            p.link(x, y, cw, rh, dkey(d))
    for i in range(32):
        p.hl(X0, X1, BT + hh + i * rh, RULE if i in (0, 31) else RULE_L, 0.45)
    for m in range(13):
        p.vl(X0 + m * cw, BT, BB, RULE, 0.5)
    for m in range(12):
        p.vl(X0 + m * cw + 16, BT + hh, BB, RULE_L, 0.4)
    p.hl(X0, X1, BT, RULE, 0.5)
    p.end()


def pp_mission(p):
    p.start("mission", "가치 · 사명 · 역할", 0)
    p.chrome([(str(YEAR), "cover"), ("가치 · 사명", "mission")], menu="mission")
    p.title("가치 · 사명", "VALUES · MISSION · ROLES", "모든 계획의 출발점")
    p.pill(X1 - 84, TOP + 10, 84, 22, "연간 목표로", "goals1", size=8)
    p.chev(X1 - 11, TOP + 21, 2.6, True, ACC, 1.0)
    hw = (CW - 14) / 2
    x = X0
    p.section(x, BT, hw, "핵심 가치", "GOVERNING VALUES")
    y = BT + 14
    for i in range(6):
        p.box(x, y + 4, 15, 15, fill=ACC_L, r=3)
        p.ctext(x, y + 4, 15, 15, str(i + 1), "B", 7.6, ACC)
        p.hl(x + 21, x + hw, y + 20, RULE, 0.6)
        p.hl(x + 21, x + hw, y + 37, RULE_L, 0.6)
        y += 38
    x = X0 + hw + 14
    p.section(x, BT, hw, "나의 역할", "ROLES")
    y = BT + 14
    rr = (6 * 38) / 7
    for i in range(7):
        p.text(x, y + 13, "역할 %d" % (i + 1), "SB", 6.6, ACC)
        p.hl(x + 30, x + hw, y + 16, RULE, 0.6)
        p.hl(x, x + hw, y + rr - 2, RULE_L, 0.6)
        y += rr
    y = BT + 14 + 6 * 38 + 12
    p.section(X0, y, CW, "사명 선언서", "PERSONAL MISSION STATEMENT")
    p.text(X0, y + 26, "나는 어떤 사람이 되고 싶은가 · 무엇을 이루고 싶은가 · 무엇에 기여하고 싶은가", "R", 6.8, FAINT)
    n = 12
    p.lines(X0, y + 30, CW, n, 22, RULE_L)
    y += 30 + n * 22 + 14
    # 피라미드
    p.section(X0, y, hw, "계획의 피라미드", "PRODUCTIVITY PYRAMID")
    py0 = y + 20
    ph = BB - py0
    tiers = [("일일 과업", "일간"), ("중간 목표", "분기 · 월 · 주"), ("장기 목표", "연간 목표"), ("핵심 가치", "가치 · 사명")]
    c = p.c
    cx = X0 + 82
    half = 80
    th = ph / 4.0
    for i in range(4):
        ya, yb = py0 + i * th, py0 + (i + 1) * th - 2
        wa, wb = half * (i / 4.0), half * ((i + 1) / 4.0) - 1
        path = c.beginPath()
        path.moveTo(cx - wa, H - ya)
        path.lineTo(cx + wa, H - ya)
        path.lineTo(cx + wb, H - yb)
        path.lineTo(cx - wb, H - yb)
        path.close()
        c.setFillColor(PYR[i])
        c.drawPath(path, stroke=0, fill=1)
        ym = (ya + yb) / 2
        p.hl(cx + (wa + wb) / 2 + 4, X0 + 176, ym, RULE, 0.5)
        p.text(X0 + 180, ym - 1, tiers[i][0], "B", 7.4, INK)
        p.text(X0 + 180, ym + 9, tiers[i][1], "R", 6.2, SUB)
    x = X0 + hw + 14
    p.box(x, y, hw, BB - y, fill=TINT, r=6)
    p.text(x + 10, y + 18, "올해의 한 단어", "B", 8, ACC)
    p.text(x + 72, y + 18, "WORD OF THE YEAR", "M", 5.6, FAINT, cs=0.7)
    p.hl(x + 10, x + hw - 10, y + 62, RULE, 0.7)
    p.text(x + 10, y + 88, "올해의 다짐", "B", 8, ACC)
    for k in range(4):
        p.hl(x + 10, x + hw - 10, y + 112 + k * 22, RULE_L, 0.6)
    p.end()


def pp_goals(p, n):
    p.start("goals%d" % n, "연간 목표 %s" % "①②"[n - 1], 0)
    p.chrome([(str(YEAR), "cover"), ("연간 목표", "goals1")], menu="goals1")
    p.title("연간 목표", "ANNUAL GOALS · %d / 2" % n, "가치에서 출발해 측정 가능한 목표로")
    other = 3 - n
    p.pill(X1 - 80, TOP + 10, 80, 22, "목표 %s 보기" % "①②"[other - 1], "goals%d" % other, size=8)
    gap = 10
    ch = (BH - 2 * gap) / 3
    for i in range(3):
        y = BT + i * (ch + gap)
        gno = (n - 1) * 3 + i + 1
        p.box(X0, y, CW, ch, stroke=RULE, r=7, lw=0.7)
        p.box(X0 + 8, y + 8, 46, 17, fill=ACC, r=8.5)
        p.ctext(X0 + 8, y + 8, 46, 17, "GOAL %d" % gno, "B", 7.4, WHITE)
        p.text(X1 - 10, y + 20, "영역", "M", 6.6, FAINT, "r")
        p.hl(X0 + 62, X1 - 30, y + 23, RULE_L, 0.6)
        ix = X0 + 10
        iw = (CW - 20 - 14) / 2
        yy = y + 32
        for label in ["목표 — 무엇을 이룰 것인가", "왜 중요한가 — 연결된 가치", "측정 기준 — 완료의 모습"]:
            p.text(ix, yy + 9, label, "SB", 6.6, ACC)
            p.lines(ix, yy + 11, iw, 2, 17, RULE_L)
            yy += 11 + 34 + 5
        p.text(ix, yy + 12, "시작", "SB", 6.6, ACC)
        p.hl(ix + 22, ix + iw / 2 - 8, yy + 15, RULE_L, 0.6)
        p.text(ix + iw / 2, yy + 12, "기한", "SB", 6.6, ACC)
        p.hl(ix + iw / 2 + 22, ix + iw, yy + 15, RULE_L, 0.6)
        rx = ix + iw + 14
        yy = y + 32
        p.text(rx, yy + 9, "실행 단계", "SB", 6.6, ACC)
        p.text(rx + iw, yy + 9, "기한", "M", 6, FAINT, "r")
        p.hl(rx, rx + iw, yy + 12, RULE, 0.6)
        yy += 12
        sh = 17.5
        for s in range(7):
            p.check(rx + 1, yy + s * sh + 4.3, 9)
            p.hl(rx, rx + iw, yy + (s + 1) * sh, RULE_L, 0.6)
        p.vl(rx + iw - 36, yy, yy + 7 * sh, RULE_L, 0.6)
        yy += 7 * sh + 8
        p.text(rx, yy + 8, "분기 점검", "SB", 6.6, ACC)
        yy += 12
        qw = (iw - 3 * 5) / 4
        qh = y + ch - 9 - yy
        for q in range(1, 5):
            qx = rx + (q - 1) * (qw + 5)
            p.box(qx, yy, qw, qh, fill=TINT, r=5)
            p.text(qx + 6, yy + 11, "Q%d" % q, "B", 7.2, ACC)
            p.link(qx, yy, qw, qh, "q%d" % q)
    p.end()


def pp_quarter(p, q):
    key = "q%d" % q
    months = [q * 3 - 2, q * 3 - 1, q * 3]
    p.start(key, "Q%d · %d–%d월" % (q, months[0], months[2]), 0)
    p.chrome([(str(YEAR), "year"), ("Q%d" % q, key)])
    p.title("Q%d" % q, "QUARTERLY PLAN", "%d월 – %d월 · 90일의 집중" % (months[0], months[2]))
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "q%d" % (q - 1) if q > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "q%d" % (q + 1) if q < 4 else None, right=True)
    p.pill(X1 - 132, TOP + 10, 70, 22, "연간 목표", "goals1", size=8)
    gx = 16
    mw = (CW - 2 * gx) / 3
    mh = 172
    for i, m in enumerate(months):
        p.mini_month(X0 + i * (mw + gx), BT, mw, m, rh=(mh - 28) / 6)
    y0 = BT + mh + 10
    rev_h = 92
    mid = BB - rev_h - 12 - y0
    lw = 288
    x = X0
    p.section(x, y0, lw, "분기 핵심 목표", "TOP 3 GOALS")
    blk = (mid - 14) / 3
    y = y0 + 18
    for g in range(3):
        p.box(x, y, 34, 14, fill=ACC, r=7)
        p.ctext(x, y, 34, 14, "목표 %d" % (g + 1), "B", 6.8, WHITE)
        p.hl(x + 40, x + lw, y + 15, RULE, 0.6)
        p.hl(x, x + lw, y + 33, RULE_L, 0.6)
        p.text(x, y + 46, "핵심 결과", "SB", 6.4, ACC)
        p.text(x + lw, y + 46, "달성도", "M", 6, FAINT, "r")
        nkr = int((blk - 60) / 18)
        for r in range(nkr):
            ry = y + 49 + r * 18
            p.check(x + 1, ry + 4.5, 9)
            p.hl(x, x + lw, ry + 18, RULE_L, 0.6)
        p.vl(x + lw - 34, y + 49, y + 49 + nkr * 18, RULE_L, 0.6)
        y += blk
    x = X0 + lw + 14
    rw = X1 - x
    seg = (mid - 2 * 8) / 3
    for i, m in enumerate(months):
        y = y0 + i * (seg + 8)
        p.section(x, y, rw, "%d월 마일스톤" % m, MON_EN[m][:3])
        p.link(x, y - 4, rw - 44, 20, "m%d" % m)
        p.pill(x + rw - 38, y - 1, 38, 12, "계획", "mp%d" % m, size=6.4)
        p.link(x + rw - 44, y - 5, 44, 20, "mp%d" % m)
        n = int((seg - 14) / 19)
        for r in range(n):
            p.check(x + 1, y + 14 + r * 19 + 5, 9)
            p.hl(x, x + rw, y + 14 + (r + 1) * 19, RULE_L, 0.6)
    y = BB - rev_h
    w3 = (CW - 2 * 14) / 3
    for i, lab in enumerate(["잘된 것", "배운 것", "바꿀 것"]):
        rx = X0 + i * (w3 + 14)
        p.section(rx, y, w3, lab, ["KEEP", "LEARN", "CHANGE"][i])
        for k in range(3):
            p.hl(rx, rx + w3, y + 14 + (k + 1) * 26, RULE_L, 0.6)
    p.end()


def _month_side(p):
    y = BT + 480 + 16
    gap = 14
    w3 = (CW - 2 * gap) / 3
    h = BB - y
    x = X0
    p.section(x, y, w3, "이달의 목표", "FOCUS")
    n = int((h - 14) / 22)
    for r in range(n):
        p.check(x + 1, y + 14 + r * 22 + 7, 9)
        p.hl(x, x + w3, y + 14 + (r + 1) * 22, RULE_L, 0.6)
    x += w3 + gap
    p.section(x, y, w3, "주요 일정 · 마감", "KEY DATES")
    for r in range(n):
        p.hl(x, x + w3, y + 14 + (r + 1) * 22, RULE_L, 0.6)
    p.vl(x + 28, y + 14, y + 14 + n * 22, RULE_L, 0.6)
    x += w3 + gap
    p.section(x, y, w3, "메모", "NOTES")
    p.dots(x, y + 16, w3, h - 16, 14)


def pp_month(p, m):
    key = "m%d" % m
    p.start(key, "%d월" % m, 0)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(m), "q%d" % quarter(m)), ("%d월" % m, key)], month=m)
    p.title("%d월" % m, MON_EN[m] + " %d" % YEAR, "날짜 → 일간 · W → 주간")
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "m%d" % (m - 1) if m > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "m%d" % (m + 1) if m < 12 else None, right=True)
    p.pill(X1 - 146, TOP + 10, 82, 22, "월간 계획 · 습관", "mp%d" % m, size=7.4)
    gh = 480
    wkw, hh = 22, 16
    cw = (CW - wkw) / 7
    first = dt.date(YEAR, m, 1)
    start = monday(first)
    nd = calendar.monthrange(YEAR, m)[1]
    nrows = ((first - start).days + nd + 6) // 7
    rh = (gh - hh) / nrows
    p.box(X0, BT, CW, hh, fill=ACC_L)
    p.ctext(X0, BT, wkw, hh, "주", "M", 6.2, ACC_M)
    for i in range(7):
        p.ctext(X0 + wkw + i * cw, BT, cw, hh, WD_KO[COLS[i]], "B", 8, col_color(i, ACC))
    for r in range(nrows):
        mon = start + dt.timedelta(7 * r)
        y = BT + hh + r * rh
        p.box(X0, y, wkw, rh, fill=TINT)
        p.text(X0 + wkw / 2, y + 13, "W", "M", 5.4, ACC_M, "c")
        p.text(X0 + wkw / 2, y + 23, wlabel(mon)[1:], "B", 8, ACC, "c")
        p.link(X0, y, wkw, rh, wkey(mon))
        for i in range(7):
            d = mon + dt.timedelta(i)
            x = X0 + wkw + i * cw
            if d.month != m:
                p.box(x, y, cw, rh, fill=TINT)
                p.text(x + 5, y + 13, "%d.%d" % (d.month, d.day), "M", 7, FAINT)
                if in_year(d):
                    p.link(x, y, cw, rh, dkey(d))
                continue
            p.text(x + 5, y + 15, str(d.day), "B", 12, day_color(d))
            md = (d.month, d.day)
            if md in HOLIDAYS:
                p.text(x + 5, y + 26, HOLIDAYS[md], "SB", 6, RED)
            elif md in MEMOS:
                p.text(x + 5, y + 26, MEMOS[md], "M", 6, SUB)
            sub = []
            if lunar(d)[1] in (1, 15) or d.day == 1:
                sub.append(lunar_str(d))
            if md in TERMS:
                sub.append(TERMS[md])
            if md in HOLIDAYS and md in MEMOS:
                sub.append(MEMOS[md])
            if sub:
                p.text(x + 5, y + rh - 5, " · ".join(sub), "R", 5.4, FAINT)
            p.link(x, y, cw, rh, dkey(d))
    for r in range(nrows + 1):
        p.hl(X0, X1, BT + hh + r * rh, RULE, 0.5)
    for i in range(8):
        p.vl(X0 + wkw + i * cw, BT, BT + gh, RULE, 0.5)
    p.vl(X0, BT, BT + gh, RULE, 0.5)
    p.hl(X0, X1, BT, RULE, 0.5)
    p.form("p_month_side", lambda: _month_side(p))
    p.end()


def _monthplan_static(p):
    y = BT
    p.section(X0, y, CW, "영역별 월간 목표", "GOALS BY AREA")
    areas = ["일 · 커리어", "건강", "관계", "성장 · 학습", "재정"]
    for i, a in enumerate(areas):
        yy = y + 14 + i * 25
        p.box(X0, yy + 3, 56, 20, fill=ACC_L, r=4)
        p.ctext(X0, yy + 3, 56, 20, a, "SB", 6.8, ACC)
        p.hl(X0 + 62, X1, yy + 23, RULE_L, 0.6)
    y += 14 + 5 * 25 + 10
    lw = 288
    p.section(X0, y, lw, "마스터 태스크 리스트", "MASTER TASKS")
    ty = y + 14
    p.box(X0, ty, lw, 12, fill=TINT)
    p.text(X0 + 10, ty + 8.6, "✓", "M", 6.6, SUB, "c")
    p.text(X0 + 32, ty + 8.6, "순위", "M", 6, SUB, "c")
    p.text(X0 + 50, ty + 8.6, "이번 달에 할 일 → 주간 · 일간으로", "M", 6, SUB)
    p.text(X0 + lw - 18, ty + 8.6, "기한", "M", 6, SUB, "c")
    ty += 12
    n, rh = 14, 20
    for r in range(n):
        p.check(X0 + 5.5, ty + r * rh + 5.5, 9)
        p.hl(X0, X0 + lw, ty + (r + 1) * rh, RULE_L, 0.6)
    for vx in (X0 + 20, X0 + 44, X0 + lw - 36):
        p.vl(vx, ty - 12, ty + n * rh, RULE_L, 0.6)
    x = X0 + lw + 14
    rw = X1 - x
    p.section(x, y, rw, "월간 인덱스", "INDEX")
    p.box(x, y + 14, rw, 12, fill=TINT)
    p.text(x + 15, y + 22.6, "날짜", "M", 6, SUB, "c")
    p.text(x + 36, y + 22.6, "일간에 적어 둔 중요한 기록", "M", 6, SUB)
    for r in range(n):
        p.hl(x, X1, ty + (r + 1) * rh, RULE_L, 0.6)
    p.vl(x + 30, ty - 12, ty + n * rh, RULE_L, 0.6)
    ry = BB - 104
    w3 = (CW - 2 * 14) / 3
    for i, (ko, en) in enumerate([("이번 달의 성과", "WINS"), ("배운 점", "LESSONS"), ("다음 달 집중", "NEXT")]):
        rx = X0 + i * (w3 + 14)
        p.section(rx, ry, w3, ko, en)
        for k in range(4):
            p.hl(rx, rx + w3, ry + 14 + (k + 1) * 22, RULE_L, 0.6)


def pp_monthplan(p, m):
    key = "mp%d" % m
    p.start(key, "%d월 계획 · 습관" % m, 1)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(m), "q%d" % quarter(m)), ("%d월" % m, "m%d" % m),
              ("계획", key)], month=m)
    p.title("%d월 계획" % m, "MONTHLY PLAN", "목표 · 태스크 · 습관 · 회고")
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "mp%d" % (m - 1) if m > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "mp%d" % (m + 1) if m < 12 else None, right=True)
    p.pill(X1 - 128, TOP + 10, 66, 22, "%d월 달력" % m, "m%d" % m, size=8)
    p.form("p_monthplan", lambda: _monthplan_static(p))
    hy = BT + 14 + 5 * 25 + 10 + 14 + 12 + 14 * 20 + 14
    p.section(X0, hy, CW, "습관 트래커", "HABIT TRACKER")
    p.text(X1, hy + 9.5, "날짜 → 일간", "R", 6.2, FAINT, "r")
    hy += 14
    lw, tw_ = 68, 22
    nd = calendar.monthrange(YEAR, m)[1]
    cw = (CW - lw - tw_) / 31
    hh = 22
    nh = 8
    rh = (BB - 104 - 10 - hy - hh) / nh
    for dd in range(1, 32):
        x = X0 + lw + (dd - 1) * cw
        if dd > nd:
            p.box(x, hy, cw, hh + nh * rh, fill=TINT)
            continue
        d = dt.date(YEAR, m, dd)
        col = day_color(d)
        if col != INK:
            p.box(x, hy, cw, hh + nh * rh, fill=HexColor("#FBF3F2") if col == RED else HexColor("#F2F6FB"))
        p.text(x + cw / 2, hy + 9, WD_KO[d.weekday()], "R", 5, col if col != INK else FAINT, "c")
        p.text(x + cw / 2, hy + 18.5, str(dd), "SB", 6.6, col, "c")
        p.link(x, hy, cw, hh, dkey(d))
        for r in range(nh):
            p.circle(x + cw / 2, hy + hh + r * rh + rh / 2, 4.2, stroke=RULE, lw=0.55)
    p.text(X0 + 4, hy + 14, "습관", "M", 6.4, SUB)
    p.text(X1 - tw_ / 2, hy + 14, "합계", "M", 6, SUB, "c")
    for r in range(nh + 1):
        p.hl(X0, X1, hy + hh + r * rh, RULE if r in (0, nh) else RULE_L, 0.55)
    p.hl(X0, X1, hy, RULE, 0.55)
    p.vl(X0 + lw, hy, hy + hh + nh * rh, RULE, 0.55)
    p.vl(X1 - tw_, hy, hy + hh + nh * rh, RULE, 0.55)
    p.end()


# 주간 레이아웃 치수
WK_TH = 100
WK_REV = 58


def _week_static(p):
    gap = 14
    w3 = (CW - 2 * gap) / 3
    y = BT
    p.section(X0, y, w3, "핵심 목표", "BIG 3")
    for r in range(3):
        p.circle(X0 + 6, y + 14 + r * 24 + 13, 5.4, fill=ACC_L)
        p.text(X0 + 6, y + 14 + r * 24 + 15.5, str(r + 1), "B", 6.8, ACC, "c")
        p.hl(X0 + 15, X0 + w3, y + 14 + (r + 1) * 24, RULE_L, 0.6)
    x = X0 + w3 + gap
    p.section(x, y, w3, "역할과 목표", "ROLES")
    for r in range(4):
        p.text(x, y + 14 + r * 21 + 14, "역할", "SB", 5.8, ACC)
        p.hl(x + 18, x + w3, y + 14 + (r + 1) * 21, RULE_L, 0.6)
    x += w3 + gap
    p.section(x, y, w3, "주간 습관", "HABITS")
    lw = 46
    cw = (w3 - lw) / 7
    for i in range(7):
        p.text(x + lw + i * cw + cw / 2, y + 24, WD_KO[COLS[i]], "M", 5.6, col_color(i, SUB), "c")
    for r in range(4):
        yy = y + 28 + r * 17
        p.hl(x, x + w3, yy + 17, RULE_L, 0.6)
        for i in range(7):
            p.circle(x + lw + i * cw + cw / 2, yy + 8.5, 4.4, stroke=RULE, lw=0.6)
    y0 = BT + WK_TH + 6
    rh = (BB - WK_REV - 8 - y0) / 7
    dw = 54
    kx = X0 + dw + 8
    kw = (CW - dw - 8) * 0.44
    sx = kx + kw + 10
    p.hl(X0, X1, y0, INK, 0.9)
    p.text(kx, y0 - 3, "핵심 할 일", "M", 5.6, FAINT)
    p.text(sx, y0 - 3, "일정 · 메모", "M", 5.6, FAINT)
    for i in range(7):
        y = y0 + i * rh
        for r in range(3):
            ry = y + 6 + r * (rh - 12) / 3
            p.check(kx, ry + ((rh - 12) / 3 - 9) / 2, 9)
            p.hl(kx + 14, kx + kw, ry + (rh - 12) / 3, RULE_L, 0.6)
        for r in range(1, 4):
            p.hl(sx, X1, y + r * rh / 4, RULE_L, 0.5, dash=[1, 2])
        p.hl(X0, X1, y + rh, RULE, 0.6)
    p.vl(sx - 5, y0, y0 + 7 * rh, RULE_L, 0.6)
    y = BB - WK_REV
    for i, (ko, en) in enumerate([("잘한 일", "WINS"), ("배운 점", "LESSONS"), ("다음 주로", "CARRY OVER")]):
        rx = X0 + i * (w3 + gap)
        p.section(rx, y, w3, ko, en)
        p.hl(rx, rx + w3, y + 36, RULE_L, 0.6)
        p.hl(rx, rx + w3, y + 56, RULE_L, 0.6)


def pp_week(p, mon):
    key = wkey(mon)
    sun = mon + dt.timedelta(6)
    rep = mon if in_year(mon) else sun
    m = rep.month
    lab = wlabel(mon)
    p.start(key, "%s · %s" % (lab, week_range_str(mon)), 1)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(m), "q%d" % quarter(m)), ("%d월" % m, "m%d" % m),
              (lab, key)], month=m)
    p.text(X0, TOP + 26, lab, "B", 22, INK)
    x = X0 + p.tw(lab, "B", 22) + 10
    p.text(x, TOP + 17, "WEEKLY PLAN · %d" % YEAR, "SB", 7.2, ACC, cs=1.4)
    p.text(x, TOP + 30, week_range_str(mon), "SB", 9.4, INK)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)
    i = WEEKS.index(mon)
    p.navbtn(X1 - 56, TOP + 9, 26, 24, wkey(WEEKS[i - 1]) if i > 0 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, wkey(WEEKS[i + 1]) if i < len(WEEKS) - 1 else None, right=True)
    months = []
    for d in (mon, sun):
        if in_year(d) and d.month not in months:
            months.append(d.month)
    bx = X1 - 64
    for mm in reversed(months):
        bx -= 60
        p.pill(bx, TOP + 10, 54, 22, "%d월 달력" % mm, "m%d" % mm, size=7.4)
    p.form("p_week", lambda: _week_static(p))
    y0 = BT + WK_TH + 6
    rh = (BB - WK_REV - 8 - y0) / 7
    dw = 54
    for k in range(7):
        d = mon + dt.timedelta(k)
        y = y0 + k * rh
        ok = in_year(d)
        col = day_color(d) if ok else FAINT
        p.box(X0, y + 3, dw, rh - 6, fill=ACC_L if ok else TINT, r=5)
        p.text(X0 + dw / 2, y + 30, str(d.day), "B", 20, col, "c")
        p.text(X0 + dw / 2, y + 44, WD_KO[d.weekday()] + "요일", "SB", 7, col, "c")
        md = (d.month, d.day)
        tag = (HOLIDAYS.get(md) or TERMS.get(md) or MEMOS.get(md)) if ok else "%d.%d" % (d.month, d.day)
        if tag:
            tag = SHORT.get(tag, tag)
            p.text(X0 + dw / 2, y + 57, tag[:7], "M", 5.6, RED if ok and md in HOLIDAYS else SUB, "c")
        if ok:
            p.link(X0, y + 3, dw, rh - 6, dkey(d))
    p.end()


def _day_static(p):
    y = BT + 40
    p.section(X0, y, CW, "오늘 꼭 끝낼 3가지", "BIG 3")
    p.box(X0, y + 17, CW, 62, fill=ACC_L, r=5)
    for r in range(3):
        p.circle(X0 + 12, y + 17 + r * 20 + 11, 6, fill=ACC)
        p.text(X0 + 12, y + 17 + r * 20 + 13.8, str(r + 1), "B", 7.4, WHITE, "c")
        if r < 2:
            p.hl(X0 + 24, X1 - 8, y + 17 + (r + 1) * 20, ACC_M, 0.4)
    y += 90
    bot = BB - 214
    lw = 262
    p.section(X0, y, lw, "우선순위 업무", "ABC TASKS")
    ty = y + 14
    p.box(X0, ty, lw, 12, fill=TINT)
    p.text(X0 + 10, ty + 8.6, "✓", "M", 6.6, SUB, "c")
    p.text(X0 + 32, ty + 8.6, "ABC", "M", 5.8, SUB, "c")
    p.text(X0 + 50, ty + 8.6, "할 일", "M", 6, SUB)
    ty += 12
    n = 15
    rh = (bot - ty) / n
    for r in range(n):
        p.check(X0 + 5.5, ty + r * rh + (rh - 9) / 2, 9)
        p.hl(X0, X0 + lw, ty + (r + 1) * rh, RULE_L, 0.6)
    p.vl(X0 + 20, ty - 12, bot, RULE_L, 0.6)
    p.vl(X0 + 44, ty - 12, bot, RULE_L, 0.6)
    sx = X0 + lw + 14
    sw = X1 - sx
    p.section(sx, y, sw, "시간 계획", "SCHEDULE")
    ty = y + 14
    hw = 20
    half = (sw - hw) / 2
    p.box(sx, ty, sw, 12, fill=TINT)
    p.text(sx + hw + half / 2, ty + 8.6, "계획", "M", 6, SUB, "c")
    p.text(sx + hw + half * 1.5, ty + 8.6, "실제", "M", 6, SUB, "c")
    ty += 12
    hrs = list(range(6, 24))
    rh = (bot - ty) / len(hrs)
    for i, hr in enumerate(hrs):
        yy = ty + i * rh
        p.text(sx + hw - 4, yy + 9, str(hr), "SB", 7, ACC if hr in (9, 12, 18) else SUB, "r")
        p.hl(sx + hw, X1, yy + rh / 2, RULE_L, 0.45, dash=[1, 2.2])
        p.hl(sx, X1, yy + rh, RULE_L, 0.6)
    p.vl(sx + hw, ty - 12, bot, RULE_L, 0.6)
    p.vl(sx + hw + half, ty - 12, bot, RULE_L, 0.6)
    y = bot + 12
    nh = 112
    p.section(X0, y, CW, "오늘의 기록", "NOTES")
    p.dots(X0, y + 16, CW, nh - 16, 14)
    y += nh + 8
    ew = 150
    p.section(X0, y, ew, "지출", "EXPENSES")
    erh = (BB - y - 14) / 4
    for r in range(4):
        p.hl(X0, X0 + ew, y + 14 + (r + 1) * erh, RULE_L if r < 3 else RULE, 0.6)
    p.vl(X0 + ew - 50, y + 14, BB, RULE_L, 0.6)
    p.text(X0 + 2, BB - 5.5, "합계", "SB", 6.4, SUB)
    gx = X0 + ew + 14
    p.section(gx, y, X1 - gx, "감사 · 한 줄 회고", "GRATITUDE")
    for r in range(2):
        p.hl(gx, X1, y + 14 + (r + 1) * 19, RULE_L, 0.6)
    ty = y + 14 + 38 + 8
    p.text(gx, ty + 9, "물", "SB", 6.6, SUB)
    for i in range(8):
        p.circle(gx + 16 + i * 11.5, ty + 6.5, 4.2, stroke=RULE, lw=0.6)
    p.text(gx + 112, ty + 9, "컨디션", "SB", 6.6, SUB)
    for i in range(5):
        p.circle(gx + 140 + i * 12.5, ty + 6.5, 4.8, stroke=RULE, lw=0.6)
        p.text(gx + 140 + i * 12.5, ty + 8.6, str(i + 1), "M", 5.4, FAINT, "c")
    p.text(gx + 210, ty + 9, "수면", "SB", 6.6, SUB)
    p.hl(gx + 228, X1, ty + 11, RULE_L, 0.6)
    p.text(gx, ty + 27, "운동", "SB", 6.6, SUB)
    p.hl(gx + 20, gx + 140, ty + 29, RULE_L, 0.6)
    p.text(gx + 150, ty + 27, "날씨", "SB", 6.6, SUB)
    p.hl(gx + 170, X1, ty + 29, RULE_L, 0.6)


def pp_day(p, d):
    key = dkey(d)
    idx = (d - DAYS[0]).days
    p.start(key)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(d.month), "q%d" % quarter(d.month)),
              ("%d월" % d.month, "m%d" % d.month), (wlabel(d), wkey(d)),
              ("%d일 %s" % (d.day, WD_KO[d.weekday()]), key)], month=d.month)
    col = day_color(d)
    p.text(X0, TOP + 31, str(d.day), "B", 34, col)
    x = X0 + p.tw(str(d.day), "B", 34) + 9
    p.text(x, TOP + 15, WD_KO[d.weekday()] + "요일", "B", 12.5, col)
    p.text(x + p.tw(WD_KO[d.weekday()] + "요일", "B", 12.5) + 7, TOP + 15, "%d년 %d월" % (YEAR, d.month),
           "M", 8.6, SUB)
    info = "%s · %s %d  ·  %s" % (WD_EN[d.weekday()][:3], MON_EN[d.month][:3], d.day, lunar_str(d, "음력 "))
    p.text(x, TOP + 29, info, "M", 7, SUB)
    ix = x + p.tw(info, "M", 7) + 8
    md = (d.month, d.day)
    tags = []
    if md in HOLIDAYS:
        tags.append((HOLIDAYS[md], RED, HexColor("#FBECEA")))
    if md in TERMS:
        tags.append((TERMS[md], ACC, ACC_L))
    if md in MEMOS:
        tags.append((MEMOS[md], ACC, ACC_L))
    for label, fg, bg in tags:
        w = p.tw(label, "SB", 7) + 12
        if ix + w > X1 - 62:
            break
        p.box(ix, TOP + 19.5, w, 13, fill=bg, r=6.5)
        p.ctext(ix, TOP + 19.5, w, 13, label, "SB", 7, fg)
        ix += w + 5
    p.text(x, TOP + 40, "%d일째 · 남은 날 %d일" % (idx + 1, NDAYS - idx - 1), "R", 6.4, FAINT)
    p.navbtn(X1 - 56, TOP + 6, 26, 30, dkey(DAYS[idx - 1]) if idx > 0 else None, right=False)
    p.navbtn(X1 - 26, TOP + 6, 26, 30, dkey(DAYS[idx + 1]) if idx < NDAYS - 1 else None, right=True)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)
    y = BT - 2
    p.pill(X0, y, 40, 30, wlabel(d), wkey(d), font="B", size=8.6)
    mon = monday(d)
    cw = (CW - 46) / 7
    for i in range(7):
        dd = mon + dt.timedelta(i)
        cx = X0 + 46 + i * cw
        ok = in_year(dd)
        if dd == d:
            p.box(cx + 2, y, cw - 4, 30, fill=ACC, r=6)
            p.text(cx + cw / 2, y + 11, WD_KO[COLS[i]], "M", 6.2, WHITE, "c")
            p.text(cx + cw / 2, y + 25, str(dd.day), "B", 11, WHITE, "c")
        else:
            c2 = day_color(dd, SUB) if ok else RULE
            p.box(cx + 2, y, cw - 4, 30, fill=TINT, r=6)
            p.text(cx + cw / 2, y + 11, WD_KO[COLS[i]], "M", 6.2, c2, "c")
            p.text(cx + cw / 2, y + 25, str(dd.day), "SB", 11, c2, "c")
        if ok:
            p.link(cx, y, cw, 30, dkey(dd))
    p.form("p_day", lambda: _day_static(p))
    p.end()


# ── 도구 ──
def pp_anniv(p):
    p.start("anniv", "기념일 · 생일", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("기념일 · 생일", "anniv")])
    p.title("기념일 · 생일", "SPECIAL DATES", "매년 반복되는 소중한 날")
    gx, gy = 12, 8
    bw = (CW - 2 * gx) / 3
    bh = (BH - 3 * gy) / 4
    for m in range(1, 13):
        x = X0 + ((m - 1) % 3) * (bw + gx)
        y = BT + ((m - 1) // 3) * (bh + gy)
        p.section(x, y, bw, "%d월" % m, MON_EN[m])
        p.link(x, y, bw, 14, "m%d" % m)
        n = 7
        rh = (bh - 14) / n
        for r in range(n):
            p.hl(x, x + bw, y + 14 + (r + 1) * rh, RULE_L, 0.6)
        p.vl(x + 26, y + 14, y + bh, RULE_L, 0.6)
    p.end()


def pp_pixels(p):
    p.start("pixels", "연간 기록 · 픽셀", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("연간 기록", "pixels")])
    p.title("연간 기록", "YEAR IN PIXELS", "하루 한 칸 — 기분·컨디션을 색으로")
    lw = 18
    cw = (CW - lw) / 12
    hh = 14
    rh = 14.4
    for m in range(1, 13):
        x = X0 + lw + (m - 1) * cw
        p.ctext(x, BT, cw, hh, "%d월" % m, "B", 7, ACC)
        p.link(x, BT, cw, hh, "m%d" % m)
        nd = calendar.monthrange(YEAR, m)[1]
        for dd in range(1, 32):
            y = BT + hh + (dd - 1) * rh
            if dd > nd:
                p.box(x, y, cw, rh, fill=TINT)
            elif dt.date(YEAR, m, dd).weekday() == 6:
                p.box(x + cw - 3.4, y + 1.2, 2.2, 2.2, fill=HexColor("#E3A9A3"))
    for dd in range(1, 32):
        p.text(X0 + lw - 4, BT + hh + (dd - 1) * rh + rh / 2 + 2.4, str(dd), "M", 6.2, SUB, "r")
    gb = BT + hh + 31 * rh
    for r in range(32):
        p.hl(X0 + lw, X1, BT + hh + r * rh, RULE if r in (0, 31) else RULE_L, 0.5)
    for i in range(13):
        p.vl(X0 + lw + i * cw, BT + hh, gb, RULE, 0.5)
    y = gb + 14
    p.section(X0, y, CW, "색 범례", "COLOR KEY")
    kw = (CW - 5 * 10) / 6
    for i in range(6):
        kx = X0 + i * (kw + 10)
        p.box(kx, y + 22, 16, 16, stroke=RULE, r=3)
        p.hl(kx + 22, kx + kw, y + 37, RULE_L, 0.6)
    p.text(X0, y + 54, "예) 기분: 최고 · 좋음 · 보통 · 지침 · 나쁨  /  빨간 점 = 일요일", "R", 6.6, FAINT)
    y += 68
    p.section(X0, y, CW, "월별 요약", "MONTHLY SCORECARD")
    y += 16
    mlw = 50
    ccw = (CW - mlw) / 13
    hh2 = 18
    nr = 6
    rh2 = (BB - y - hh2) / nr
    p.box(X0, y, CW, hh2, fill=TINT)
    for k in range(13):
        cx = X0 + mlw + k * ccw
        if k < 12:
            p.ctext(cx, y, ccw, hh2, "%d월" % (k + 1), "SB", 6.6, ACC)
            p.link(cx, y, ccw, hh2, "mp%d" % (k + 1))
        else:
            p.box(cx, y, ccw, hh2 + nr * rh2, fill=ACC_L)
            p.ctext(cx, y, ccw, hh2, "연간", "B", 6.6, ACC)
    for r in range(nr):
        yy = y + hh2 + r * rh2
        p.text(X0 + 4, yy + rh2 / 2 + 2.4, "지표 %d" % (r + 1), "M", 6.2, FAINT)
        p.hl(X0, X1, yy + rh2, RULE_L if r < nr - 1 else RULE, 0.6)
    p.hl(X0, X1, y + hh2, RULE, 0.6)
    for k in range(14):
        p.vl(X0 + mlw + k * ccw, y, BB, RULE_L, 0.6)
    p.end()


def pp_finance(p, half):
    key = "finance" if half == 1 else "finance2"
    months = list(range(1, 7)) if half == 1 else list(range(7, 13))
    p.start(key, "연간 가계부 (%d–%d월)" % (months[0], months[-1]), 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("연간 가계부", "finance")])
    p.title("연간 가계부", "ANNUAL BUDGET · %d / 2" % half, "%s반기 · %d월 – %d월" % ("상하"[half - 1], months[0], months[-1]))
    other = "finance2" if half == 1 else "finance"
    p.pill(X1 - 76, TOP + 10, 76, 22, "%s반기 보기" % "하상"[half - 1], other, size=8)
    tcols = ["상반기 합계"] if half == 1 else ["하반기 합계", "연간 합계"]
    lw, tw_ = 86, 54
    cw = (CW - lw - tw_ * len(tcols)) / 6
    hh = 20
    sections = [("수입", 4), ("고정 지출", 7), ("변동 지출", 7), ("저축 · 투자", 4)]
    nrows = sum(n for _, n in sections) + len(sections) + 2
    rh = (BH - hh) / nrows
    p.box(X0, BT, CW, hh, fill=ACC)
    p.text(X0 + 8, BT + 13, "항목", "B", 7.6, WHITE)
    for k, m in enumerate(months):
        p.ctext(X0 + lw + k * cw, BT, cw, hh, "%d월" % m, "B", 7.6, WHITE)
        p.link(X0 + lw + k * cw, BT, cw, hh, "mp%d" % m)
    for k, t in enumerate(tcols):
        p.ctext(X0 + lw + 6 * cw + k * tw_, BT, tw_, hh, t, "B", 6.8, WHITE)
    y = BT + hh
    for name, n in sections:
        p.box(X0, y, CW, rh, fill=ACC_L)
        p.text(X0 + 8, y + rh / 2 + 2.8, name, "B", 7.4, ACC)
        p.text(X1 - 6, y + rh / 2 + 2.6, "소계", "M", 6, ACC_M, "r")
        p.hl(X0, X1, y + rh, RULE, 0.5)
        y += rh
        for r in range(n):
            p.hl(X0, X1, y + rh, RULE_L, 0.5)
            y += rh
    for name in ["총지출", "잔액 (수입−지출−저축)"]:
        p.box(X0, y, CW, rh, fill=TINT)
        p.text(X0 + 8, y + rh / 2 + 2.8, name, "B", 7, INK)
        p.hl(X0, X1, y + rh, RULE, 0.5)
        y += rh
    for k in range(7 + len(tcols) - 1):
        xx = X0 + lw + k * cw if k <= 6 else X0 + lw + 6 * cw + (k - 6) * tw_
        p.vl(xx, BT + hh, BB, RULE_L if k not in (0, 6) else RULE, 0.5)
    p.vl(X0, BT, BB, RULE, 0.5)
    p.vl(X1, BT, BB, RULE, 0.5)
    p.end()


def pp_reading(p):
    p.start("reading", "독서 · 강의 목록", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("독서 · 강의", "reading")])
    p.title("독서 · 강의 목록", "READING LOG", "읽은 책, 들은 강의, 본 영화")
    n, hh = 34, 14
    rh = (BH - hh) / n
    cols = [(0, 22, "No"), (22, 214, "제목"), (236, 96, "저자 · 강사"), (332, 56, "완료일"), (388, CW - 388, "평점")]
    p.box(X0, BT, CW, hh, fill=ACC_L)
    for cx, cw, label in cols:
        p.ctext(X0 + cx, BT, cw, hh, label, "SB", 6.6, ACC)
    for r in range(n):
        y = BT + hh + r * rh
        p.ctext(X0, y, 22, rh, "%d" % (r + 1), "M", 6.8, FAINT)
        sx = X0 + 388 + (CW - 388 - 5 * 13) / 2 + 6.5
        for s in range(5):
            p.text(sx + s * 13, y + rh / 2 + 3.2, "☆", "R", 9, RULE, "c")
        p.hl(X0, X1, y + rh, RULE_L, 0.6)
    for cx, cw, label in cols[1:]:
        p.vl(X0 + cx, BT + hh, BB, RULE_L, 0.6)
    p.end()


def pp_bucket(p):
    p.start("bucket", "버킷리스트", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("버킷리스트", "bucket")])
    p.title("%d 버킷리스트" % YEAR, "BUCKET LIST", "올해 꼭 해 보고 싶은 것들")
    cats = [("경험 · 여행", "EXPERIENCE"), ("성장 · 도전", "GROWTH"), ("사람 · 나눔", "PEOPLE")]
    gap = 16
    sh = (BH - 2 * gap) / 3
    cw = (CW - 16) / 2
    for t, (ko, en) in enumerate(cats):
        y = BT + t * (sh + gap)
        p.section(X0, y, CW, ko, en)
        n = int((sh - 16) / 22)
        for c in range(2):
            x = X0 + c * (cw + 16)
            for r in range(n):
                ry = y + 16 + r * 22
                p.check(x + 1, ry + 6, 10)
                p.hl(x + 16, x + cw, ry + 22, RULE_L, 0.6)
    p.end()


def _project_static(p):
    gap = 14
    w3 = (CW - 2 * gap) / 3
    y = BT
    for i, (ko, en) in enumerate([("목적", "WHY"), ("완료의 모습", "DONE"), ("자원 · 사람", "RESOURCES")]):
        x = X0 + i * (w3 + gap)
        p.section(x, y, w3, ko, en)
        p.lines(x, y + 14, w3, 4, 19, RULE_L)
    y += 14 + 76 + 12
    p.section(X0, y, CW, "작업 목록", "ACTION ITEMS")
    ty = y + 14
    p.box(X0, ty, CW, 12, fill=TINT)
    p.text(X0 + 10, ty + 8.6, "✓", "M", 6.6, SUB, "c")
    p.text(X0 + 26, ty + 8.6, "작업", "M", 6, SUB)
    p.text(X1 - 104, ty + 8.6, "담당", "M", 6, SUB, "c")
    p.text(X1 - 56, ty + 8.6, "기한", "M", 6, SUB, "c")
    p.text(X1 - 15, ty + 8.6, "상태", "M", 6, SUB, "c")
    ty += 12
    n, rh = 18, 20
    for r in range(n):
        p.check(X0 + 5.5, ty + r * rh + 5.5, 9)
        p.hl(X0, X1, ty + (r + 1) * rh, RULE_L, 0.6)
    for vx in (X0 + 20, X1 - 128, X1 - 80, X1 - 30):
        p.vl(vx, ty - 12, ty + n * rh, RULE_L, 0.6)
    y = ty + n * rh + 14
    hw = (CW - gap) / 2
    p.section(X0, y, hw, "마일스톤", "MILESTONES")
    nn = int((BB - y - 14) / 19)
    for r in range(nn):
        p.check(X0 + 1, y + 14 + r * 19 + 5, 9)
        p.hl(X0, X0 + hw, y + 14 + (r + 1) * 19, RULE_L, 0.6)
    p.vl(X0 + hw - 40, y + 14, y + 14 + nn * 19, RULE_L, 0.6)
    x = X0 + hw + gap
    p.section(x, y, hw, "메모 · 아이디어", "NOTES")
    p.dots(x, y + 16, hw, BB - y - 16, 14)


def pp_project(p, n):
    key = "proj%d" % n
    p.start(key, "프로젝트 %d" % n, 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("프로젝트 %d" % n, key)])
    p.text(X0, TOP + 26, "P%d" % n, "B", 22, ACC)
    x = X0 + p.tw("P%d" % n, "B", 22) + 10
    p.text(x, TOP + 13, "PROJECT", "SB", 6.6, ACC, cs=1.4)
    p.hl(x, X1 - 230, TOP + 32, RULE, 0.7)
    p.text(X1 - 222, TOP + 30, "시작", "SB", 7, SUB)
    p.hl(X1 - 202, X1 - 148, TOP + 32, RULE, 0.7)
    p.text(X1 - 140, TOP + 30, "기한", "SB", 7, SUB)
    p.hl(X1 - 120, X1 - 66, TOP + 32, RULE, 0.7)
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "proj%d" % (n - 1) if n > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "proj%d" % (n + 1) if n < 6 else None, right=True)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)
    p.form("p_project", lambda: _project_static(p))
    p.end()


def pp_notes_index(p):
    p.start("notesidx", "노트 색인", 1)
    p.chrome([(str(YEAR), "cover"), ("노트", "notesidx")], menu="notesidx")
    p.title("노트 색인", "NOTES INDEX", "번호를 누르면 해당 노트로 이동")
    gap = 14
    tw_ = (CW - 2 * gap) / 3
    n = 10
    rh = (BH - 14) / n
    for t in range(3):
        x = X0 + t * (tw_ + gap)
        p.section(x, BT, tw_, "%s 노트" % NOTE_KINDS[t][0], NOTE_KINDS[t][1].upper())
        for r in range(n):
            no = t * 10 + r + 1
            y = BT + 14 + r * rh
            p.pill(x, y + (rh - 22) / 2, 30, 22, "%02d" % no, "n%d" % no, font="B", size=8.4)
            p.hl(x + 36, x + tw_ - 42, y + rh - 14, RULE_L, 0.6)
            p.hl(x + tw_ - 36, x + tw_, y + rh - 14, RULE_L, 0.6)
            if r == 0:
                p.text(x + 36, y + 12, "제목", "M", 5.6, FAINT)
                p.text(x + tw_ - 36, y + 12, "날짜", "M", 5.6, FAINT)
            p.link(x, y, tw_, rh, "n%d" % no)
    p.end()


def pp_note(p, no):
    key = "n%d" % no
    kind = NOTE_KINDS[(no - 1) // 10]
    p.start(key, "노트 %02d (%s)" % (no, kind[0]), 2)
    p.chrome([(str(YEAR), "cover"), ("노트", "notesidx"), ("%02d" % no, key)], menu="notesidx")
    p.text(X0, TOP + 26, "%02d" % no, "B", 22, ACC)
    x = X0 + p.tw("%02d" % no, "B", 22) + 10
    p.text(x, TOP + 13, "NOTE · " + kind[1].upper(), "SB", 6.6, ACC, cs=1.4)
    p.hl(x, X1 - 170, TOP + 32, RULE, 0.7)
    p.text(X1 - 160, TOP + 30, "날짜", "SB", 7, SUB)
    p.hl(X1 - 140, X1 - 66, TOP + 32, RULE, 0.7)
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "n%d" % (no - 1) if no > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "n%d" % (no + 1) if no < 30 else None, right=True)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)

    def static():
        if kind[1] == "lined":
            n = int(BH / 22.5)
            p.lines(X0, BT, CW, n, BH / float(n), RULE_L, 0.6)
        elif kind[1] == "dot":
            p.dots(X0, BT, CW, BH, 14)
        else:
            p.grid(X0, BT, CW, BH, 14)

    p.form("p_note_" + kind[1], static)
    p.end()


MINE.update(k for k in list(globals()) if k.startswith(("pp_", "_")) or k in ("MINE", "WK_TH", "WK_REV"))


def compose_portrait(p):
    _sync()
    pp_cover(p)
    pp_index(p)
    pp_year(p)
    pp_glance(p)
    pp_mission(p)
    pp_goals(p, 1)
    pp_goals(p, 2)
    for d in DAYS:
        if d.day == 1:
            if d.month % 3 == 1:
                pp_quarter(p, quarter(d.month))
            pp_month(p, d.month)
            pp_monthplan(p, d.month)
        if d == monday(d) or d == DAYS[0]:
            pp_week(p, monday(d))
        pp_day(p, d)
    p.outline.append((0, "도구 · 프로젝트", "anniv"))
    pp_anniv(p)
    pp_pixels(p)
    pp_finance(p, 1)
    pp_finance(p, 2)
    pp_reading(p)
    pp_bucket(p)
    for n in range(1, 7):
        pp_project(p, n)
    p.outline.append((0, "노트", "notesidx"))
    pp_notes_index(p)
    for no in range(1, 31):
        pp_note(p, no)


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    fonts = os.environ.get("PLANNER_FONT_DIR", os.path.join(here, "fonts"))
    years = [int(a) for a in sys.argv[1:]] or sorted(mp.THEMES)
    for year in years:
        mp.build(year, os.path.join(here, "%d-planner-세로형.pdf" % year), fonts, compose_portrait)
