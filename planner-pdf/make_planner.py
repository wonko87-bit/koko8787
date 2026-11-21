#!/usr/bin/env python3
"""하이퍼링크 플래너 PDF 생성기 (삼성노트 · 갤럭시탭 가로 화면용).

    python3 make_planner.py [연도 ...]      # 생략하면 2027–2030 전부
    → <연도>-planner.pdf

필요: reportlab, pikepdf, korean_lunar_calendar
폰트: Pretendard TTF 4종(Regular/Medium/SemiBold/Bold)을 ./fonts 또는
      환경변수 PLANNER_FONT_DIR 경로에 둔다. (npm pack pretendard 로 받을 수 있음)
"""
import calendar
import datetime as dt
import os
import sys

import pikepdf
from korean_lunar_calendar import KoreanLunarCalendar
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


# ── 페이지 규격 (갤럭시탭 울트라 가로 16:10, 화면 폭에 맞추면 1pt ≈ 실제 1pt) ──
W, H = 900, 540
TABW = 26                 # 오른쪽 월 탭 폭
NAVH = 26                 # 상단 내비게이션 높이
X0, X1 = 18, W - TABW - 12
CW = X1 - X0              # 본문 폭 844
TOP = NAVH + 10           # 본문 시작(위에서부터)
HDR = 42                  # 페이지 제목 영역 높이
BT = TOP + HDR + 8        # 본문 body 시작 = 86
BB = H - 12               # 본문 body 끝 = 528
BH = BB - BT              # 442

# ── 색 ──
INK = HexColor("#262A33")
SUB = HexColor("#6B7280")
FAINT = HexColor("#A3A9B3")
RULE = HexColor("#C3C8D1")
RULE_L = HexColor("#E1E4EA")
# 연도별 테마: (강조색, 중간, 옅은 배경, 면 배경, 피라미드 위 3단)
THEMES = {
    2027: ("네이비", "#29476B", "#8FA5C0", "#EAEFF6", "#F5F6F9", ("#DCE5F0", "#B9C9DD", "#6E8BAE")),
    2028: ("포레스트", "#2C5B4B", "#93B4A7", "#E7F1EC", "#F4F7F5", ("#DAE9E2", "#B3CFC4", "#6B9987")),
    2029: ("플럼", "#58407A", "#AD9DC5", "#F0ECF6", "#F7F6F9", ("#E4DDEF", "#C7B9DD", "#8F79B1")),
    2030: ("모카", "#6A4A36", "#BFA895", "#F4EEE8", "#F8F6F3", ("#EADFD5", "#D6C2B1", "#A68569")),
}
RED = HexColor("#C0392B")
BLUE = HexColor("#2F6DB5")
WHITE = HexColor("#FFFFFF")
DOT = HexColor("#B4BAC4")

WD_KO = "월화수목금토일"
WD_EN = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
MON_EN = ["", "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY",
          "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]

# ── 공휴일 ──
# 2027년은 우주항공청 「2027년 월력요항」(대체공휴일 7일)과 대조해 검증한다.
# 2028년 이후는 월력요항 발표 전이므로 2026년 11월 현행 규정으로 계산한 예정값이다.
OFFICIAL_HOLIDAYS = {2027: {
    (1, 1): "신정",
    (2, 6): "설날 연휴", (2, 7): "설날", (2, 8): "설날 연휴", (2, 9): "대체공휴일",
    (3, 1): "삼일절",
    (5, 1): "노동절", (5, 3): "대체공휴일", (5, 5): "어린이날", (5, 13): "부처님오신날",
    (6, 6): "현충일",
    (7, 17): "제헌절", (7, 19): "대체공휴일",
    (8, 15): "광복절", (8, 16): "대체공휴일",
    (9, 14): "추석 연휴", (9, 15): "추석", (9, 16): "추석 연휴",
    (10, 3): "개천절", (10, 4): "대체공휴일", (10, 9): "한글날", (10, 11): "대체공휴일",
    (12, 25): "성탄절", (12, 27): "대체공휴일",
}}
FIXED_HOLIDAYS = {(1, 1): "신정", (3, 1): "삼일절", (5, 1): "노동절", (5, 5): "어린이날", (6, 6): "현충일",
                  (7, 17): "제헌절", (8, 15): "광복절", (10, 3): "개천절", (10, 9): "한글날", (12, 25): "성탄절"}
# 임기만료 선거일(공직선거법 제34조로 계산한 법정 공휴일)
ELECTIONS = {2028: {(4, 12): "국회의원선거"}, 2030: {(3, 27): "대통령선거", (6, 12): "지방선거"}}
# 토·일·다른 공휴일과 겹치면 대체공휴일이 붙는 날 (설·추석 연휴는 일요일·다른 공휴일과 겹칠 때만)
SUB_WEEKEND = {"삼일절", "노동절", "어린이날", "부처님오신날", "제헌절", "광복절", "개천절", "한글날", "성탄절"}


def lunar_to_solar(y, m, d):
    k = KoreanLunarCalendar()
    k.setLunarDate(y, m, d, False)
    return dt.date.fromisoformat(k.SolarIsoFormat())


def compute_holidays(year):
    one = dt.timedelta(1)
    names = {}

    def add(d, n):
        if d.year == year:
            names.setdefault(d, []).append(n)

    for (m, d), n in FIXED_HOLIDAYS.items():
        add(dt.date(year, m, d), n)
    for lm, ld, n in ((1, 1, "설날"), (8, 15, "추석")):
        base = lunar_to_solar(year, lm, ld)
        add(base - one, n + " 연휴")
        add(base, n)
        add(base + one, n + " 연휴")
    add(lunar_to_solar(year, 4, 8), "부처님오신날")
    for (m, d), n in ELECTIONS.get(year, {}).items():
        add(dt.date(year, m, d), n)
    subs = set()
    for d in sorted(names):
        ns = names[d]
        if any(n in SUB_WEEKEND for n in ns):
            need = d.weekday() >= 5 or len(ns) > 1
        elif any(n.startswith(("설날", "추석")) for n in ns):
            need = d.weekday() == 6 or len(ns) > 1
        else:
            need = False
        if need:
            x = d + one
            while x.weekday() >= 5 or x in names or x in subs:
                x += one
            subs.add(x)
    out = {(d.month, d.day): " · ".join(ns) for d, ns in names.items()}
    out.update({(d.month, d.day): "대체공휴일" for d in subs if d.year == year})
    return out


# 24절기 (태양 황경 계산값, 한국 표준시 기준)
TERMS_BY_YEAR = {
    2027: {(1, 5): '소한', (1, 20): '대한', (2, 4): '입춘', (2, 19): '우수', (3, 6): '경칩', (3, 21): '춘분', (4, 5): '청명', (4, 20): '곡우', (5, 6): '입하', (5, 21): '소만', (6, 6): '망종', (6, 21): '하지', (7, 7): '소서', (7, 23): '대서', (8, 8): '입추', (8, 23): '처서', (9, 8): '백로', (9, 23): '추분', (10, 8): '한로', (10, 24): '상강', (11, 8): '입동', (11, 22): '소설', (12, 7): '대설', (12, 22): '동지'},
    2028: {(1, 6): '소한', (1, 20): '대한', (2, 4): '입춘', (2, 19): '우수', (3, 5): '경칩', (3, 20): '춘분', (4, 4): '청명', (4, 19): '곡우', (5, 5): '입하', (5, 20): '소만', (6, 5): '망종', (6, 21): '하지', (7, 6): '소서', (7, 22): '대서', (8, 7): '입추', (8, 22): '처서', (9, 7): '백로', (9, 22): '추분', (10, 8): '한로', (10, 23): '상강', (11, 7): '입동', (11, 22): '소설', (12, 6): '대설', (12, 21): '동지'},
    2029: {(1, 5): '소한', (1, 20): '대한', (2, 3): '입춘', (2, 18): '우수', (3, 5): '경칩', (3, 20): '춘분', (4, 4): '청명', (4, 20): '곡우', (5, 5): '입하', (5, 21): '소만', (6, 5): '망종', (6, 21): '하지', (7, 7): '소서', (7, 22): '대서', (8, 7): '입추', (8, 23): '처서', (9, 7): '백로', (9, 23): '추분', (10, 8): '한로', (10, 23): '상강', (11, 7): '입동', (11, 22): '소설', (12, 7): '대설', (12, 21): '동지'},
    # 2030년 우수는 계산상 2월 18일 23:59:40 무렵이라 공식 발표에서 19일로 나올 수 있다.
    2030: {(1, 5): '소한', (1, 20): '대한', (2, 4): '입춘', (2, 18): '우수', (3, 5): '경칩', (3, 20): '춘분', (4, 5): '청명', (4, 20): '곡우', (5, 5): '입하', (5, 21): '소만', (6, 5): '망종', (6, 21): '하지', (7, 7): '소서', (7, 23): '대서', (8, 7): '입추', (8, 23): '처서', (9, 7): '백로', (9, 23): '추분', (10, 8): '한로', (10, 23): '상강', (11, 7): '입동', (11, 22): '소설', (12, 7): '대설', (12, 22): '동지'},
}
# 기념일(쉬는 날 아님)
FIXED_MEMOS = {(4, 5): "식목일", (5, 8): "어버이날", (5, 15): "스승의날", (10, 1): "국군의날"}


def configure(year):
    """연도에 따라 달라지는 전역값(날짜·공휴일·테마)을 설정"""
    global YEAR, DAYS, NDAYS, WEEKS, HOLIDAYS, TERMS, MEMOS, ACC, ACC_M, ACC_L, TINT, PYR
    YEAR = year
    DAYS = [dt.date(year, 1, 1) + dt.timedelta(n) for n in range(366 if calendar.isleap(year) else 365)]
    NDAYS = len(DAYS)
    WEEKS = sorted({monday(d) for d in DAYS})
    HOLIDAYS = compute_holidays(year)
    if year in OFFICIAL_HOLIDAYS and HOLIDAYS != OFFICIAL_HOLIDAYS[year]:
        raise SystemExit("%d년 공휴일 계산이 공식 자료와 다릅니다" % year)
    TERMS = TERMS_BY_YEAR.get(year, {})
    MEMOS = dict(FIXED_MEMOS)
    for lm, ld, n in ((1, 15, "정월대보름"), (5, 5, "단오")):
        d = lunar_to_solar(year, lm, ld)
        MEMOS[(d.month, d.day)] = n
    name, acc, acc_m, acc_l, tint, pyr = THEMES.get(year) or THEMES[2027 + (year - 2027) % 4]
    ACC, ACC_M, ACC_L, TINT = HexColor(acc), HexColor(acc_m), HexColor(acc_l), HexColor(tint)
    PYR = [HexColor(c) for c in pyr] + [ACC]


WEEK_START = 6            # 주 시작 요일: 6 = 일요일, 0 = 월요일
COLS = [(WEEK_START + i) % 7 for i in range(7)]   # 달력 열 순서(요일 번호)


def monday(d):
    """d가 속한 주의 시작일 (WEEK_START 기준)"""
    return d - dt.timedelta((d.weekday() - WEEK_START) % 7)


def col_color(i, base):
    return RED if COLS[i] == 6 else BLUE if COLS[i] == 5 else base


def dkey(d):
    return "d" + d.isoformat()


def wkey(d):
    return "w" + monday(d).isoformat()


def wlabel(d):
    if WEEK_START == 0:                      # 월요일 시작: ISO 주차
        return "W%02d" % monday(d).isocalendar()[1]
    return "W%02d" % ((monday(d) - WEEKS[0]).days // 7 + 1)   # 1월 1일이 든 주가 W01


def in_year(d):
    return d.year == YEAR


def is_hol(d):
    return in_year(d) and (d.month, d.day) in HOLIDAYS


def day_color(d, base=INK):
    if d.weekday() == 6 or is_hol(d):
        return RED
    if d.weekday() == 5:
        return BLUE
    return base


_lunar = KoreanLunarCalendar()
_lcache = {}


def lunar(d):
    """(월, 일, 윤달여부)"""
    if d not in _lcache:
        _lunar.setSolarDate(d.year, d.month, d.day)
        s = _lunar.LunarIsoFormat()
        leap = "Intercalation" in s
        y, m, dd = s.split(" ")[0].split("-")
        _lcache[d] = (int(m), int(dd), leap)
    return _lcache[d]


def lunar_str(d, prefix="음 "):
    m, dd, leap = lunar(d)
    return "%s%s%d.%d" % (prefix, "윤" if leap else "", m, dd)


def quarter(m):
    return (m - 1) // 3 + 1


SHORT = {"부처님오신날": "석가탄신", "국회의원선거": "총선", "대통령선거": "대선", "지방선거": "지선"}


def long_breaks():
    """공휴일이 낀 3일 이상 연휴: [(이름, 시작일, 끝일)]"""
    out, run = [], []
    for d in DAYS + [None]:
        if d is not None and (d.weekday() >= 5 or is_hol(d)):
            run.append(d)
            continue
        hol = [HOLIDAYS[(x.month, x.day)] for x in run if is_hol(x)]
        if len(run) >= 3 and hol:
            real = [h for h in hol if h != "대체공휴일"] or hol
            name = real[0].split(" · ")[0].replace(" 연휴", "")
            out.append((SHORT.get(name, name), run[0], run[-1]))
        run = []
    return out


# ─────────────────────────────────────────────────────────────────────────────
class Painter:
    def __init__(self, path, font_dir):
        for name, fn in [("R", "Regular"), ("M", "Medium"), ("SB", "SemiBold"), ("B", "Bold")]:
            pdfmetrics.registerFont(TTFont(name, os.path.join(font_dir, "Pretendard-%s.ttf" % fn)))
        self.c = canvas.Canvas(path, pagesize=(W, H), pageCompression=1)
        self.c.setTitle("%d 플래너" % YEAR)
        self.c.setAuthor("Planner")
        self.c.setSubject("일간 · 주간 · 월간 · 연간 하이퍼링크 플래너")
        self.page = 0
        self.links = []     # (page, (x0,y0,x1,y1), key)
        self.dests = {}     # key -> page
        self.outline = []   # (level, title, key)
        self.forms = set()

    # 페이지
    def start(self, key, outline=None, level=0):
        self.dests[key] = self.page
        if outline:
            self.outline.append((level, outline, key))

    def end(self):
        self.c.showPage()
        self.page += 1

    def link(self, x, top, w, h, key):
        if key:
            self.links.append((self.page, (x, H - top - h, x + w, H - top), key))

    # 기본 도형 (좌표는 모두 '위에서부터')
    def text(self, x, y, s, font="R", size=9, color=INK, align="l", cs=0):
        c = self.c
        c.setFillColor(color)
        c.setFont(font, size)
        if cs:
            wd = pdfmetrics.stringWidth(s, font, size) + cs * max(len(s) - 1, 0)
            if align == "c":
                x -= wd / 2
            elif align == "r":
                x -= wd
            c.drawString(x, H - y, s, charSpace=cs)
        elif align == "c":
            c.drawCentredString(x, H - y, s)
        elif align == "r":
            c.drawRightString(x, H - y, s)
        else:
            c.drawString(x, H - y, s)

    def tw(self, s, font="R", size=9):
        return pdfmetrics.stringWidth(s, font, size)

    def ctext(self, x, top, w, h, s, font="R", size=9, color=INK):
        """상자 한가운데 정렬"""
        self.text(x + w / 2, top + h / 2 + size * 0.36, s, font, size, color, "c")

    def hl(self, x0, x1, y, color=RULE, lw=0.6, dash=None):
        c = self.c
        c.setStrokeColor(color)
        c.setLineWidth(lw)
        c.setDash(dash or [])
        c.line(x0, H - y, x1, H - y)
        c.setDash([])

    def vl(self, x, y0, y1, color=RULE, lw=0.6, dash=None):
        c = self.c
        c.setStrokeColor(color)
        c.setLineWidth(lw)
        c.setDash(dash or [])
        c.line(x, H - y0, x, H - y1)
        c.setDash([])

    def box(self, x, top, w, h, fill=None, stroke=None, r=0, lw=0.6):
        c = self.c
        if fill is not None:
            c.setFillColor(fill)
        if stroke is not None:
            c.setStrokeColor(stroke)
            c.setLineWidth(lw)
        if r:
            c.roundRect(x, H - top - h, w, h, r, stroke=int(stroke is not None), fill=int(fill is not None))
        else:
            c.rect(x, H - top - h, w, h, stroke=int(stroke is not None), fill=int(fill is not None))

    def circle(self, cx, cy, r, fill=None, stroke=None, lw=0.6):
        c = self.c
        if fill is not None:
            c.setFillColor(fill)
        if stroke is not None:
            c.setStrokeColor(stroke)
            c.setLineWidth(lw)
        c.circle(cx, H - cy, r, stroke=int(stroke is not None), fill=int(fill is not None))

    def check(self, x, top, s=9, color=RULE):
        self.box(x, top, s, s, stroke=color, r=1.6, lw=0.7)

    def chev(self, cx, cy, s=3.2, right=True, color=SUB, lw=1.1):
        c = self.c
        c.setStrokeColor(color)
        c.setLineWidth(lw)
        c.setLineCap(1)
        c.setLineJoin(1)
        d = s if right else -s
        p = c.beginPath()
        p.moveTo(cx - d / 2, H - cy + s)
        p.lineTo(cx + d / 2, H - cy)
        p.lineTo(cx - d / 2, H - cy - s)
        c.drawPath(p, stroke=1, fill=0)
        c.setLineCap(0)
        c.setLineJoin(0)

    def dots(self, x, top, w, h, step=14, color=DOT):
        """도트 그리드: 둥근 끝의 점선 한 줄 = 점 한 행 (용량 최소화)"""
        c = self.c
        c.setStrokeColor(color)
        c.setLineWidth(1.1)
        c.setLineCap(1)
        c.setDash([0.05, step - 0.05])
        nx = int(w // step)
        ny = int(h // step)
        ox = x + (w - nx * step) / 2 + step / 2
        oy = top + (h - ny * step) / 2 + step / 2
        for j in range(ny):
            yy = H - (oy + j * step)
            c.line(ox, yy, ox + (nx - 1) * step + 0.06, yy)
        c.setDash([])
        c.setLineCap(0)

    def grid(self, x, top, w, h, step=14, color=RULE_L):
        nx = int(w // step)
        ny = int(h // step)
        gw, gh = nx * step, ny * step
        ox = x + (w - gw) / 2
        oy = top + (h - gh) / 2
        for i in range(nx + 1):
            self.vl(ox + i * step, oy, oy + gh, color, 0.4)
        for j in range(ny + 1):
            self.hl(ox, ox + gw, oy + j * step, color, 0.4)

    def lines(self, x, top, w, n, rh, color=RULE_L, lw=0.6, first=True):
        """가로 필기선 n행 (각 행의 아랫선)"""
        for i in range(n):
            self.hl(x, x + w, top + (i + 1) * rh, color, lw)

    def section(self, x, top, w, ko, en=None, h=14):
        """섹션 제목: 굵은 한글 + 작은 영문, 아래 강조선"""
        self.text(x, top + 9.5, ko, "B", 8.6, ACC)
        if en:
            self.text(x + self.tw(ko, "B", 8.6) + 6, top + 9.5, en, "M", 5.6, FAINT, cs=0.7)
        self.hl(x, x + w, top + h - 0.5, ACC, 1.0)

    # 정적 배경(Form XObject) – 같은 그림을 수백 페이지가 공유
    def form(self, name, draw):
        if name not in self.forms:
            self.c.beginForm(name, 0, 0, W, H)
            draw()
            self.c.endForm()
            self.forms.add(name)
        self.c.doForm(name)

    # ── 공통 크롬: 상단 내비 + 오른쪽 월 탭 ──
    MENU = [("홈", "index"), ("연간달력", "year"), ("연간일정", "glance"), ("사명", "mission"),
            ("목표", "goals1"), ("노트", "notesidx")]

    def _menu_geom(self):
        out = []
        x = W - TABW - 8
        for label, key in reversed(self.MENU):
            w = self.tw(label, "M", 8) + 16
            x -= w
            out.append((label, key, x, w))
        return list(reversed(out))

    def _chrome_static(self):
        self.box(0, 0, W - TABW, NAVH, fill=TINT)
        self.hl(0, W - TABW, NAVH, RULE, 0.6)
        for label, key, x, w in self._menu_geom():
            self.ctext(x, 0, w, NAVH, label, "M", 8, SUB)
        th = H / 12.0
        self.box(W - TABW, 0, TABW, H, fill=TINT)
        self.vl(W - TABW, 0, H, RULE, 0.6)
        for m in range(1, 13):
            if m > 1:
                self.hl(W - TABW + 5, W - 5, (m - 1) * th, RULE, 0.5)
            self.ctext(W - TABW, (m - 1) * th, TABW, th, "%d월" % m, "M", 7.6, SUB)

    def chrome(self, crumbs, month=None, menu=None):
        """crumbs: [(라벨, key)] 왼쪽 경로 / month: 강조할 월 탭 / menu: 강조할 메뉴 key"""
        self.form("chrome", self._chrome_static)
        for label, key, x, w in self._menu_geom():
            self.link(x, 0, w, NAVH, key)
            if key == menu:
                self.box(x + 2, 4.5, w - 4, NAVH - 9, fill=ACC, r=4)
                self.ctext(x, 0, w, NAVH, label, "SB", 8, WHITE)
        th = H / 12.0
        for m in range(1, 13):
            self.link(W - TABW, (m - 1) * th, TABW, th, "m%d" % m)
            if m == month:
                self.box(W - TABW, (m - 1) * th, TABW, th, fill=ACC)
                self.ctext(W - TABW, (m - 1) * th, TABW, th, "%d월" % m, "B", 7.8, WHITE)
        x = 12
        for i, (label, key) in enumerate(crumbs):
            last = i == len(crumbs) - 1
            font, col = ("B", ACC) if last else ("M", SUB)
            w = self.tw(label, font, 8.6)
            self.text(x + 6, NAVH / 2 + 3.1, label, font, 8.6, col)
            self.link(x, 0, w + 12, NAVH, key)
            x += w + 12
            if not last:
                self.chev(x + 2, NAVH / 2, 2.4, True, FAINT, 0.9)
                x += 6

    def title(self, ko, en=None, sub=None):
        """페이지 제목 (왼쪽)"""
        self.text(X0, TOP + 26, ko, "B", 22, INK)
        x = X0 + self.tw(ko, "B", 22) + 10
        if en:
            self.text(x, TOP + 17, en, "SB", 7.2, ACC, cs=1.4)
        if sub:
            self.text(x, TOP + 29, sub, "R", 7.6, SUB)
        self.hl(X0, X1, TOP + HDR, INK, 1.1)

    def navbtn(self, x, top, w, h, key, right=True):
        """이전/다음 버튼"""
        if not key:
            return
        self.box(x, top, w, h, fill=TINT, stroke=RULE, r=5, lw=0.5)
        self.chev(x + w / 2, top + h / 2, 3.6, right, ACC, 1.3)
        self.link(x, top, w, h, key)

    def pill(self, x, top, w, h, label, key, fill=None, color=None, font="SB", size=8, stroke=None):
        fill = ACC_L if fill is None else fill
        color = ACC if color is None else color
        self.box(x, top, w, h, fill=fill, stroke=stroke, r=h / 2 if h <= 20 else 6, lw=0.5)
        self.ctext(x, top, w, h, label, font, size, color)
        self.link(x, top, w, h, key)

    # ── 미니 달력 ──
    def mini_month(self, x, top, w, m, rh=18, title=True, wk=True):
        """작은 월 달력. 날짜→일간, 주차→주간, 제목→월간 링크. 높이를 반환."""
        y = top
        if title:
            self.text(x + 2, y + 11, "%d월" % m, "B", 10.5, ACC)
            self.text(x + 2 + self.tw("%d월" % m, "B", 10.5) + 5, y + 11, MON_EN[m][:3], "M", 6, FAINT, cs=0.8)
            self.link(x, y, w, 15, "m%d" % m)
            y += 16
        wkw = 17 if wk else 0
        cw = (w - wkw) / 7.0
        for i in range(7):
            col = col_color(i, SUB)
            self.text(x + wkw + cw * i + cw / 2, y + 7.5, WD_KO[COLS[i]], "M", 6.2, col, "c")
        y += 10
        self.hl(x, x + w, y, RULE, 0.5)
        first = dt.date(YEAR, m, 1)
        start = monday(first)
        nd = calendar.monthrange(YEAR, m)[1]
        nrows = ((first - start).days + nd + 6) // 7
        for r in range(nrows):
            mon = start + dt.timedelta(7 * r)
            if wk:
                self.ctext(x, y + r * rh, wkw, rh, wlabel(mon)[1:], "M", 5.6, ACC_M)
                self.link(x, y + r * rh, wkw, rh, wkey(mon))
            for i in range(7):
                d = mon + dt.timedelta(i)
                if d.month != m:
                    continue
                self.ctext(x + wkw + cw * i, y + r * rh, cw, rh, str(d.day),
                           "SB" if is_hol(d) else "M", 8.2, day_color(d))
                self.link(x + wkw + cw * i, y + r * rh, cw, rh, dkey(d))
        if wk:
            self.vl(x + wkw, y + 1, y + nrows * rh, RULE_L, 0.5)
        return (y + nrows * rh) - top


# ─────────────────────────────────────────────────────────────────────────────
# 페이지별 그리기
# ─────────────────────────────────────────────────────────────────────────────
def page_cover(p):
    p.start("cover", "표지")
    p.box(0, 0, W, H, fill=WHITE)
    p.box(0, 0, 300, H, fill=ACC)
    p.text(40, 70, "PLANNER", "SB", 11, WHITE, cs=6)
    p.text(40, 88, "DAILY · WEEKLY · MONTHLY · YEARLY", "M", 6.6, ACC_M, cs=1.6)
    p.text(36, 300, str(YEAR)[:2], "B", 150, WHITE)
    p.text(36, 430, str(YEAR)[2:], "B", 150, WHITE)
    p.text(40, 500, "한 해를 설계하고, 하루를 완성하다", "M", 9, ACC_M, cs=0.6)

    x = 350
    p.text(x, 92, "%d 플래너" % YEAR, "B", 30, INK)
    p.text(x, 114, "연간 · 분기 · 월간 · 주간 · 일간이 모두 연결된 하이퍼링크 플래너", "R", 10, SUB)
    p.hl(x, W - 50, 132, INK, 1.1)
    # 소유자
    p.text(x, 160, "이름", "SB", 8, ACC)
    p.hl(x + 30, x + 230, 164, RULE, 0.7)
    p.text(x + 260, 160, "연락처", "SB", 8, ACC)
    p.hl(x + 298, W - 50, 164, RULE, 0.7)
    p.text(x, 190, "올해의 한 문장", "SB", 8, ACC)
    p.hl(x + 70, W - 50, 194, RULE, 0.7)

    # 바로가기
    p.text(x, 236, "바로가기", "B", 9, ACC)
    p.text(x + 44, 236, "QUICK LINKS", "M", 5.8, FAINT, cs=0.8)
    items = [("홈 · 사용법", "index"), ("연간 달력", "year"), ("연간 일정표", "glance"),
             ("가치 · 사명", "mission"), ("연간 목표", "goals1"), ("노트", "notesidx")]
    bw = (W - 50 - x - 5 * 8) / 6.0
    for i, (label, key) in enumerate(items):
        p.pill(x + i * (bw + 8), 246, bw, 30, label, key, fill=ACC_L, size=8.4)
    p.text(x, 312, "분기", "B", 9, ACC)
    qw = (W - 50 - x - 3 * 8) / 4.0
    for q in range(1, 5):
        p.pill(x + (q - 1) * (qw + 8), 322, qw, 28, "Q%d  ·  %d–%d월" % (q, q * 3 - 2, q * 3), "q%d" % q,
               fill=TINT, stroke=RULE, size=8.4)
    p.text(x, 386, "월", "B", 9, ACC)
    mw = (W - 50 - x - 11 * 6) / 12.0
    for m in range(1, 13):
        p.pill(x + (m - 1) * (mw + 6), 396, mw, 34, "%d" % m, "m%d" % m, fill=TINT, stroke=RULE,
               font="B", size=11)
    p.text(x, 470, "삼성노트에서는 상단의 ‘읽기 모드’(책 아이콘)를 켠 상태에서 링크를 누르면 해당 페이지로 이동합니다.",
           "R", 7.8, SUB)
    p.text(x, 484, "필기할 때는 읽기 모드를 끄고 S펜으로 자유롭게 작성하세요.", "R", 7.8, SUB)
    p.pill(W - 50 - 96, 494, 96, 26, "시작하기", "index", fill=ACC, color=WHITE, size=9)
    p.chev(W - 50 - 16, 507, 3, True, WHITE, 1.2)
    p.end()


def page_index(p):
    p.start("index", "홈 · 사용법")
    p.chrome([(str(YEAR), "cover"), ("홈", "index")], menu="index")
    p.title("홈", "HOME · INDEX", "플래너의 모든 페이지로 이동하는 허브")
    # 왼쪽: 사용법
    lw = 262
    p.section(X0, BT, lw, "사용법", "HOW TO USE")
    tips = [
        ("링크 이동", "삼성노트 상단 ‘읽기 모드’(책 아이콘)를 켜고 탭합니다."),
        ("상단 바", "왼쪽은 현재 위치(연›분기›월›주›일), 오른쪽은 주요 메뉴입니다."),
        ("오른쪽 탭", "1월–12월 탭으로 어느 페이지에서든 월간 달력으로 갑니다."),
        ("달력 날짜", "모든 달력의 날짜는 일간, W번호는 주간 페이지로 연결됩니다."),
        ("계획 흐름", "가치·사명 → 연간 목표 → 분기 → 월 → 주간 나침반 → 오늘의 ABC."),
        ("매일 10분", "아침에 할 일을 적고 A·B·C 순위를 매긴 뒤 A1부터 실행합니다."),
    ]
    y = BT + 22
    for i, (h, body) in enumerate(tips):
        p.circle(X0 + 8, y + 5, 7, fill=ACC)
        p.text(X0 + 8, y + 8, str(i + 1), "B", 7.6, WHITE, "c")
        p.text(X0 + 22, y + 5, h, "B", 8.6, INK)
        p.text(X0 + 22, y + 17, body, "R", 7.2, SUB)
        y += 31
    y += 4
    p.section(X0, y, lw, "우선순위 · 진행 기호", "LEGEND")
    y += 22
    legend_abc = [("A", "반드시 해야 하는 일 (오늘 안 하면 문제가 생김)"),
                  ("B", "해야 하는 일 (중요하지만 미뤄도 큰일은 아님)"),
                  ("C", "하면 좋은 일 (여유가 있을 때)")]
    for k, body in legend_abc:
        p.box(X0, y - 4, 14, 14, fill=ACC_L, r=3)
        p.ctext(X0, y - 4, 14, 14, k, "B", 8, ACC)
        p.text(X0 + 22, y + 6.5, body, "R", 7.6, INK)
        y += 19
    y += 4
    sym = [("✓", "완료"), ("→", "다음 날로"), ("×", "취소"), ("●", "진행 중"), ("Ⓓ", "위임")]
    sx = X0
    for s, body in sym:
        if s == "Ⓓ":
            p.circle(sx + 5, y + 2.5, 5.2, stroke=INK, lw=0.7)
            p.text(sx + 5, y + 5.2, "D", "B", 6.4, INK, "c")
        else:
            p.text(sx + 5, y + 6, s, "B", 9.5, INK, "c")
        p.text(sx + 14, y + 6, body, "R", 7.6, SUB)
        sx += 20 + p.tw(body, "R", 7.6) + 8
    y += 24
    p.box(X0, y, lw, BB - y, fill=TINT, r=6)
    p.text(X0 + 10, y + 16, "달력 표기", "B", 8, ACC)
    p.text(X0 + 10, y + 31, "빨강", "SB", 7.6, RED)
    p.text(X0 + 32, y + 31, "일요일 · 공휴일", "R", 7.6, SUB)
    p.text(X0 + 10, y + 44, "파랑", "SB", 7.6, BLUE)
    p.text(X0 + 32, y + 44, "토요일", "R", 7.6, SUB)
    p.text(X0 + 10, y + 57, "음", "SB", 7.6, SUB)
    p.text(X0 + 32, y + 57, "음력 · 절기 · 기념일", "R", 7.6, SUB)
    bx0 = X0 + 128
    p.text(bx0, y + 16, "%d 연휴 (3일 이상)" % YEAR, "B", 8, ACC)
    for i, (name, a, b) in enumerate(long_breaks()[:12]):
        bx = bx0 + (i % 2) * 66
        by = y + 22 + (i // 2) * 12.5
        p.text(bx, by + 9, name, "SB", 6.4, RED)
        rng = "%d.%d–%d" % (a.month, a.day, b.day) if a.month == b.month else \
            "%d.%d–%d.%d" % (a.month, a.day, b.month, b.day)
        p.text(bx + 28, by + 9, rng, "R", 6.4, SUB)
        p.link(bx - 2, by, 64, 12.5, dkey(a))
    if YEAR not in OFFICIAL_HOLIDAYS:
        p.text(X0 + 10, BB - 8, "※ 공휴일은 2026년 11월 현행 법령으로 계산한 예정값입니다.", "R", 6.2, FAINT)

    # 가운데: 연간/분기/월
    mx = X0 + lw + 20
    mw_ = 330
    p.section(mx, BT, mw_, "연간", "YEARLY")
    items = [("연간 달력", "year"), ("연간 일정표", "glance"), ("가치 · 사명 · 역할", "mission"),
             ("연간 목표 ①", "goals1"), ("연간 목표 ②", "goals2")]
    bw = (mw_ - 2 * 8) / 3.0
    for i, (label, key) in enumerate(items):
        p.pill(mx + (i % 3) * (bw + 8), BT + 22 + (i // 3) * 34, bw, 27, label, key, fill=ACC_L, size=8.2)
    y = BT + 22 + 2 * 34 + 6
    p.section(mx, y, mw_, "분기", "QUARTERLY")
    qw = (mw_ - 3 * 8) / 4.0
    for q in range(1, 5):
        p.pill(mx + (q - 1) * (qw + 8), y + 22, qw, 27, "Q%d · %d–%d월" % (q, q * 3 - 2, q * 3), "q%d" % q,
               fill=TINT, stroke=RULE, size=8)
    y += 22 + 27 + 12
    p.section(mx, y, mw_, "월간 · 주간", "MONTHLY · WEEKLY")
    p.text(mx + mw_, y + 9.5, "달력 / 계획 / 첫 주", "R", 6.6, FAINT, "r")
    y += 21
    cw3 = (mw_ - 2 * 8) / 3.0
    rh = (BB - y - 3 * 6) / 4.0
    for m in range(1, 13):
        bx = mx + ((m - 1) % 3) * (cw3 + 8)
        by = y + ((m - 1) // 3) * (rh + 6)
        p.box(bx, by, cw3, rh, fill=WHITE, stroke=RULE, r=6, lw=0.6)
        p.text(bx + 10, by + 20, "%d월" % m, "B", 14, INK)
        p.text(bx + 10, by + 31, MON_EN[m], "M", 5.4, FAINT, cs=0.8)
        p.link(bx, by, cw3, rh - 22, "m%d" % m)
        sw = (cw3 - 16 - 2 * 4) / 3.0
        first = dt.date(YEAR, m, 1)
        subs = [("달력", "m%d" % m), ("계획", "mp%d" % m), (wlabel(first), wkey(first))]
        for i, (label, key) in enumerate(subs):
            p.pill(bx + 8 + i * (sw + 4), by + rh - 23, sw, 17, label, key, fill=ACC_L, size=6.8)

    # 오른쪽: 도구/프로젝트/노트
    rx = mx + mw_ + 20
    rw = X1 - rx
    p.section(rx, BT, rw, "도구", "TOOLS")
    tools = [("기념일 · 생일", "anniv"), ("연간 기록 · 픽셀", "pixels"), ("연간 가계부", "finance"),
             ("독서 · 강의 목록", "reading"), ("버킷리스트", "bucket"), ("노트 색인", "notesidx")]
    bw = (rw - 8) / 2.0
    for i, (label, key) in enumerate(tools):
        p.pill(rx + (i % 2) * (bw + 8), BT + 22 + (i // 2) * 33, bw, 27, label, key, fill=ACC_L, size=8)
    y = BT + 22 + 3 * 33 + 6
    p.section(rx, y, rw, "프로젝트", "PROJECTS")
    bw = (rw - 2 * 8) / 3.0
    for i in range(6):
        bx = rx + (i % 3) * (bw + 8)
        by = y + 22 + (i // 3) * 42
        p.box(bx, by, bw, 36, fill=WHITE, stroke=RULE, r=6, lw=0.6)
        p.text(bx + 7, by + 12, "P%d" % (i + 1), "B", 8, ACC)
        p.hl(bx + 7, bx + bw - 7, by + 28, RULE_L, 0.6)
        p.link(bx, by, bw, 36, "proj%d" % (i + 1))
    y += 22 + 2 * 42 + 6
    p.section(rx, y, rw, "노트", "NOTES")
    p.text(rx + rw, y + 9.5, "줄 · 도트 · 모눈 각 10장", "R", 6.6, FAINT, "r")
    y += 21
    nw = (rw - 9 * 4) / 10.0
    nh = (BB - y - 2 * 5) / 3.0
    for i in range(30):
        bx = rx + (i % 10) * (nw + 4)
        by = y + (i // 10) * (nh + 5)
        p.pill(bx, by, nw, nh, "%02d" % (i + 1), "n%d" % (i + 1), fill=TINT, stroke=RULE, font="M", size=7.4)
    p.end()


def page_year(p):
    p.start("year", "연간 달력", 0)
    p.chrome([(str(YEAR), "cover"), ("연간 달력", "year")], menu="year")
    p.title("%d" % YEAR, "YEAR CALENDAR", "날짜 → 일간 · W번호 → 주간 · 월 이름 → 월간")
    # 분기 버튼
    for q in range(1, 5):
        p.pill(X1 - (5 - q) * 46 + 6, TOP + 10, 40, 22, "Q%d" % q, "q%d" % q, fill=ACC_L, size=8.4)
    gx, gy = 22, 10
    mw_ = (CW - 3 * gx) / 4.0
    rowh = (BH - 2 * gy) / 3.0
    for m in range(1, 13):
        x = X0 + ((m - 1) % 4) * (mw_ + gx)
        y = BT + ((m - 1) // 4) * (rowh + gy)
        p.mini_month(x, y, mw_, m, rh=(rowh - 28) / 6.0)
    p.end()


def page_glance(p):
    p.start("glance", "연간 일정표", 0)
    p.chrome([(str(YEAR), "cover"), ("연간 일정표", "glance")], menu="glance")
    p.title("연간 일정표", "YEAR AT A GLANCE", "기념일 · 여행 · 마감 등 한 해의 큰 일정을 한눈에")
    hh = 15
    cw = CW / 12.0
    rh = (BH - hh) / 31.0
    for m in range(1, 13):
        x = X0 + (m - 1) * cw
        p.box(x, BT, cw, hh, fill=ACC_L)
        p.ctext(x, BT, cw, hh, "%d월" % m, "B", 8, ACC)
        p.link(x, BT, cw, hh, "m%d" % m)
        nd = calendar.monthrange(YEAR, m)[1]
        for dd in range(1, 32):
            y = BT + hh + (dd - 1) * rh
            if dd > nd:
                p.box(x, y, cw, rh, fill=TINT)
                continue
            d = dt.date(YEAR, m, dd)
            col = day_color(d)
            if d.weekday() >= 5 or is_hol(d):
                p.box(x, y, cw, rh, fill=HexColor("#FBF3F2") if col == RED else HexColor("#F2F6FB"))
            p.text(x + 9, y + rh / 2 + 2.6, str(dd), "SB", 7, col, "c")
            p.text(x + 17, y + rh / 2 + 2.3, WD_KO[d.weekday()], "R", 5.6, col if col != INK else FAINT)
            if is_hol(d):
                p.text(x + cw - 3, y + rh / 2 + 2.2, HOLIDAYS[(m, dd)], "M", 5.2, RED, "r")
            p.link(x, y, cw, rh, dkey(d))
    for i in range(32):
        p.hl(X0, X1, BT + hh + i * rh, RULE_L if i not in (0, 31) else RULE, 0.45)
    for m in range(13):
        p.vl(X0 + m * cw, BT, BB, RULE, 0.5)
    p.hl(X0, X1, BT, RULE, 0.5)
    p.end()


def page_mission(p):
    p.start("mission", "가치 · 사명 · 역할", 0)
    p.chrome([(str(YEAR), "cover"), ("가치 · 사명", "mission")], menu="mission")
    p.title("가치 · 사명 · 역할", "VALUES · MISSION · ROLES", "모든 계획의 출발점 — 무엇이 가장 소중한가")
    p.pill(X1 - 90, TOP + 10, 90, 22, "연간 목표로", "goals1", fill=ACC_L, size=8)
    p.chev(X1 - 12, TOP + 21, 2.6, True, ACC, 1.0)
    aw, mw_, rw = 262, 300, CW - 262 - 300 - 28
    # 가치
    x = X0
    p.section(x, BT, aw, "핵심 가치", "GOVERNING VALUES")
    y = BT + 14
    for i in range(6):
        p.box(x, y + 4, 15, 15, fill=ACC_L, r=3)
        p.ctext(x, y + 4, 15, 15, str(i + 1), "B", 7.6, ACC)
        p.hl(x + 21, x + aw, y + 20, RULE, 0.6)
        p.hl(x + 21, x + aw, y + 38, RULE_L, 0.6)
        y += 39
    # 생산성 피라미드
    y += 8
    p.section(x, y, aw, "계획의 피라미드", "PRODUCTIVITY PYRAMID")
    py0 = y + 22
    ph = BB - py0
    tiers = [("일일 과업", "일간"), ("중간 목표", "분기 · 월 · 주"), ("장기 목표", "연간 목표"), ("핵심 가치", "가치 · 사명")]
    fills = PYR
    c = p.c
    cx = x + 96
    half = 92
    th = ph / 4.0
    for i in range(4):
        ya, yb = py0 + i * th, py0 + (i + 1) * th - 2
        wa = half * (i / 4.0)
        wb = half * ((i + 1) / 4.0) - 1
        path = c.beginPath()
        path.moveTo(cx - wa, H - ya)
        path.lineTo(cx + wa, H - ya)
        path.lineTo(cx + wb, H - yb)
        path.lineTo(cx - wb, H - yb)
        path.close()
        c.setFillColor(fills[i])
        c.drawPath(path, stroke=0, fill=1)
        ym = (ya + yb) / 2
        p.hl(cx + (wa + wb) / 2 + 4, x + 200, ym, RULE, 0.5)
        p.text(x + 204, ym - 1, tiers[i][0], "B", 7.6, INK)
        p.text(x + 204, ym + 9, tiers[i][1], "R", 6.4, SUB)
    # 사명
    x = X0 + aw + 14
    p.section(x, BT, mw_, "사명 선언서", "PERSONAL MISSION STATEMENT")
    p.text(x, BT + 26, "나는 어떤 사람이 되고 싶은가 · 무엇을 이루고 싶은가 · 무엇에 기여하고 싶은가", "R", 6.8, FAINT)
    n = int((BB - BT - 32) // 22)
    p.lines(x, BT + 32, mw_, n, (BB - BT - 32) / n, RULE_L)
    # 역할
    x = x + mw_ + 14
    p.section(x, BT, rw, "나의 역할", "ROLES")
    y = BT + 14
    rr = (BB - y - 72) / 7.0
    for i in range(7):
        p.text(x, y + 13, "역할 %d" % (i + 1), "SB", 6.6, ACC)
        p.hl(x + 30, x + rw, y + 16, RULE, 0.6)
        p.hl(x, x + rw, y + rr - 4 - (rr - 20) / 2, RULE_L, 0.6)
        p.hl(x, x + rw, y + rr - 4, RULE_L, 0.6)
        y += rr
    p.box(x, BB - 62, rw, 62, fill=TINT, r=6)
    p.text(x + 10, BB - 46, "올해의 한 단어", "B", 8, ACC)
    p.text(x + 72, BB - 46, "WORD OF THE YEAR", "M", 5.6, FAINT, cs=0.7)
    p.hl(x + 10, x + rw - 10, BB - 14, RULE, 0.7)
    p.end()


def page_goals(p, n):
    key = "goals%d" % n
    p.start(key, "연간 목표 %s" % ("①②"[n - 1]), 0)
    p.chrome([(str(YEAR), "cover"), ("연간 목표", "goals1")], menu="goals1")
    p.title("연간 목표", "ANNUAL GOALS · %d / 2" % n, "가치에서 출발해 측정 가능한 목표와 실행 단계로")
    other = 3 - n
    p.pill(X1 - 84, TOP + 10, 84, 22, "목표 %s 보기" % "①②"[other - 1], "goals%d" % other, fill=ACC_L, size=8)
    gap = 14
    cw = (CW - 2 * gap) / 3.0
    for i in range(3):
        x = X0 + i * (cw + gap)
        gno = (n - 1) * 3 + i + 1
        p.box(x, BT, cw, BH, fill=None, stroke=RULE, r=7, lw=0.7)
        p.box(x + 8, BT + 8, 46, 17, fill=ACC, r=8.5)
        p.ctext(x + 8, BT + 8, 46, 17, "GOAL %d" % gno, "B", 7.4, WHITE)
        p.text(x + cw - 10, BT + 20, "영역", "M", 6.6, FAINT, "r")
        p.hl(x + 62, x + cw - 30, BT + 23, RULE_L, 0.6)
        y = BT + 30
        ix, iw = x + 10, cw - 20
        for label, rows in [("목표 — 무엇을 이룰 것인가", 2), ("왜 중요한가 — 연결된 가치", 2), ("측정 기준 — 완료의 모습", 2)]:
            p.text(ix, y + 9, label, "SB", 6.8, ACC)
            p.lines(ix, y + 11, iw, rows, 18, RULE_L)
            y += 11 + rows * 18 + 4
        p.text(ix, y + 12, "시작", "SB", 6.8, ACC)
        p.hl(ix + 22, ix + iw / 2 - 8, y + 15, RULE_L, 0.6)
        p.text(ix + iw / 2, y + 12, "기한", "SB", 6.8, ACC)
        p.hl(ix + iw / 2 + 22, ix + iw, y + 15, RULE_L, 0.6)
        y += 24
        p.text(ix, y + 9, "실행 단계", "SB", 6.8, ACC)
        p.text(ix + iw, y + 9, "기한", "M", 6.2, FAINT, "r")
        p.hl(ix, ix + iw, y + 12, RULE, 0.6)
        y += 12
        steps = 7
        sh = 18.5
        for s in range(steps):
            p.check(ix + 1, y + s * sh + 4.5, 9)
            p.hl(ix, ix + iw, y + (s + 1) * sh, RULE_L, 0.6)
        p.vl(ix + iw - 38, y, y + steps * sh, RULE_L, 0.6)
        y += steps * sh + 8
        p.text(ix, y + 9, "분기 점검", "SB", 6.8, ACC)
        y += 13
        qw = (iw - 3 * 5) / 4.0
        qh = BT + BH - 10 - y
        for q in range(1, 5):
            qx = ix + (q - 1) * (qw + 5)
            p.box(qx, y, qw, qh, fill=TINT, r=5)
            p.text(qx + 6, y + 12, "Q%d" % q, "B", 7.6, ACC)
            p.link(qx, y, qw, qh, "q%d" % q)
    p.end()


def page_quarter(p, q):
    key = "q%d" % q
    months = [q * 3 - 2, q * 3 - 1, q * 3]
    p.start(key, "Q%d · %d–%d월" % (q, months[0], months[2]), 0)
    p.chrome([(str(YEAR), "year"), ("Q%d" % q, key)], month=None)
    p.title("Q%d" % q, "QUARTERLY PLAN", "%d월 – %d월 · 90일의 집중" % (months[0], months[2]))
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "q%d" % (q - 1) if q > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "q%d" % (q + 1) if q < 4 else None, right=True)
    p.pill(X1 - 150, TOP + 10, 84, 22, "연간 목표", "goals1", fill=ACC_L, size=8)
    lw, mw_ = 176, 340
    rw = CW - lw - mw_ - 28
    # 미니 달력 3개
    seg = (BH - 2 * 6) / 3.0
    for i, m in enumerate(months):
        p.mini_month(X0, BT + i * (seg + 6), lw, m, rh=(seg - 28) / 6.0)
    # 분기 목표
    x = X0 + lw + 14
    p.section(x, BT, mw_, "분기 핵심 목표", "TOP 3 GOALS & KEY RESULTS")
    y = BT + 20
    blk = 108
    for g in range(3):
        p.box(x, y, 34, 14, fill=ACC, r=7)
        p.ctext(x, y, 34, 14, "목표 %d" % (g + 1), "B", 6.8, WHITE)
        p.hl(x + 40, x + mw_, y + 15, RULE, 0.6)
        p.hl(x, x + mw_, y + 33, RULE_L, 0.6)
        p.text(x, y + 45, "핵심 결과", "SB", 6.4, ACC)
        p.text(x + mw_, y + 45, "달성도", "M", 6, FAINT, "r")
        for r in range(3):
            ry = y + 48 + r * 18
            p.check(x + 1, ry + 4.5, 9)
            p.hl(x, x + mw_, ry + 18, RULE_L, 0.6)
        p.vl(x + mw_ - 34, y + 48, y + 48 + 54, RULE_L, 0.6)
        y += blk + 8
    p.section(x, y, mw_, "분기 회고", "REVIEW")
    rest = BB - (y + 14)
    n = 3
    labels = ["잘된 것", "배운 것", "바꿀 것"]
    for r in range(n):
        p.text(x, y + 14 + (r + 0.5) * rest / n + 3, labels[r], "SB", 6.6, SUB)
        p.hl(x + 34, x + mw_, y + 14 + (r + 1) * rest / n - 2, RULE_L, 0.6)
    # 월별 마일스톤
    x = x + mw_ + 14
    for i, m in enumerate(months):
        y = BT + i * (seg + 6)
        p.section(x, y, rw, "%d월 마일스톤" % m, MON_EN[m])
        p.link(x, y - 4, rw - 50, 20, "m%d" % m)
        p.pill(x + rw - 40, y - 1, 40, 12, "계획", "mp%d" % m, fill=ACC_L, size=6.4)
        p.link(x + rw - 46, y - 5, 46, 20, "mp%d" % m)
        n = 6
        rh = (seg - 14) / n
        for r in range(n):
            p.check(x + 1, y + 14 + r * rh + (rh - 9) / 2, 9)
            p.hl(x, x + rw, y + 14 + (r + 1) * rh, RULE_L, 0.6)
    p.end()


def page_month(p, m):
    key = "m%d" % m
    p.start(key, "%d월" % m, 0)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(m), "q%d" % quarter(m)), ("%d월" % m, key)], month=m)
    p.title("%d월" % m, MON_EN[m] + " %d" % YEAR, "날짜를 누르면 일간, W번호를 누르면 주간 페이지")
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "m%d" % (m - 1) if m > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "m%d" % (m + 1) if m < 12 else None, right=True)
    p.pill(X1 - 150, TOP + 10, 84, 22, "월간 계획 · 습관", "mp%d" % m, fill=ACC_L, size=7.6)

    sw = 170
    gw = CW - sw - 12
    wkw = 24
    cw = (gw - wkw) / 7.0
    hh = 16
    first = dt.date(YEAR, m, 1)
    start = monday(first)
    nd = calendar.monthrange(YEAR, m)[1]
    nrows = ((first - start).days + nd + 6) // 7
    rh = (BH - hh) / float(nrows)
    p.box(X0, BT, gw, hh, fill=ACC_L)
    p.ctext(X0, BT, wkw, hh, "주", "M", 6.4, ACC_M)
    for i in range(7):
        col = col_color(i, ACC)
        p.ctext(X0 + wkw + i * cw, BT, cw, hh, WD_KO[COLS[i]], "B", 8, col)
    for r in range(nrows):
        mon = start + dt.timedelta(7 * r)
        y = BT + hh + r * rh
        p.box(X0, y, wkw, rh, fill=TINT)
        p.text(X0 + wkw / 2, y + 13, "W", "M", 5.6, ACC_M, "c")
        p.text(X0 + wkw / 2, y + 23, wlabel(mon)[1:], "B", 8.4, ACC, "c")
        p.link(X0, y, wkw, rh, wkey(mon))
        for i in range(7):
            d = mon + dt.timedelta(i)
            x = X0 + wkw + i * cw
            if d.month != m:
                p.box(x, y, cw, rh, fill=TINT)
                p.text(x + 6, y + 14, "%d.%d" % (d.month, d.day), "M", 7.6, FAINT)
                if in_year(d):
                    p.link(x, y, cw, rh, dkey(d))
                continue
            col = day_color(d)
            p.text(x + 6, y + 15.5, str(d.day), "B", 12, col)
            nx = x + 6 + p.tw(str(d.day), "B", 12) + 4
            md = (d.month, d.day)
            if md in HOLIDAYS:
                p.text(nx, y + 14.5, HOLIDAYS[md], "SB", 6.6, RED)
            elif md in MEMOS:
                p.text(nx, y + 14.5, MEMOS[md], "M", 6.4, SUB)
            lm, ld, leap = lunar(d)
            sub = []
            if ld in (1, 15) or d.day == 1:
                sub.append(lunar_str(d))
            if md in TERMS:
                sub.append(TERMS[md])
            if md in HOLIDAYS and md in MEMOS:
                sub.append(MEMOS[md])
            if sub:
                p.text(x + cw - 5, y + 11, " · ".join(sub), "R", 5.8, FAINT, "r")
            p.link(x, y, cw, rh, dkey(d))
    for r in range(nrows + 1):
        p.hl(X0, X0 + gw, BT + hh + r * rh, RULE, 0.5)
    for i in range(8):
        p.vl(X0 + wkw + i * cw, BT, BB, RULE, 0.5)
    p.vl(X0, BT, BB, RULE, 0.5)
    p.hl(X0, X0 + gw, BT, RULE, 0.5)

    # 사이드바
    x = X1 - sw
    p.section(x, BT, sw, "이달의 목표", "FOCUS")
    y = BT + 14
    for r in range(4):
        p.check(x + 1, y + r * 20 + 6, 9)
        p.hl(x, x + sw, y + (r + 1) * 20, RULE_L, 0.6)
    y += 4 * 20 + 10
    p.section(x, y, sw, "주요 일정 · 마감", "KEY DATES")
    y += 14
    for r in range(7):
        p.hl(x, x + sw, y + (r + 1) * 20, RULE_L, 0.6)
    p.vl(x + 30, y, y + 7 * 20, RULE_L, 0.6)
    y += 7 * 20 + 10
    p.section(x, y, sw, "메모", "NOTES")
    p.dots(x, y + 16, sw, BB - y - 16, 14)
    p.end()


def page_monthplan(p, m):
    key = "mp%d" % m
    p.start(key, "%d월 계획 · 습관" % m, 1)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(m), "q%d" % quarter(m)), ("%d월" % m, "m%d" % m),
              ("계획", key)], month=m)
    p.title("%d월 계획" % m, "MONTHLY PLAN · " + MON_EN[m], "목표 · 마스터 태스크 · 습관 · 회고")
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "mp%d" % (m - 1) if m > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "mp%d" % (m + 1) if m < 12 else None, right=True)
    p.pill(X1 - 150, TOP + 10, 84, 22, "%d월 달력" % m, "m%d" % m, fill=ACC_L, size=8)

    def static():
        aw, bw_, cw_ = 250, 330, CW - 250 - 330 - 24
        th = 216
        # A 월간 목표
        x = X0
        p.section(x, BT, aw, "영역별 월간 목표", "GOALS BY AREA")
        areas = ["일 · 커리어", "건강", "관계", "성장 · 학습", "재정"]
        ah = (th - 14) / 5.0
        for i, a in enumerate(areas):
            y = BT + 14 + i * ah
            p.box(x, y + 3, 50, ah - 5, fill=ACC_L, r=4)
            p.ctext(x, y + 3, 50, ah - 5, a, "SB", 7, ACC)
            p.hl(x + 56, x + aw, y + ah / 2, RULE_L, 0.6)
            p.hl(x + 56, x + aw, y + ah - 1, RULE_L, 0.6)
        # B 마스터 태스크
        x = X0 + aw + 12
        p.section(x, BT, bw_, "마스터 태스크 리스트", "MASTER TASK LIST")
        y = BT + 14
        p.box(x, y, bw_, 12, fill=TINT)
        p.text(x + 10, y + 8.6, "✓", "M", 6.6, SUB, "c")
        p.text(x + 32, y + 8.6, "순위", "M", 6, SUB, "c")
        p.text(x + 50, y + 8.6, "이번 달에 할 일 (날짜 미정) → 주간 · 일간으로 옮기기", "M", 6, SUB)
        p.text(x + bw_ - 20, y + 8.6, "기한", "M", 6, SUB, "c")
        y += 12
        rh = (th - 26) / 10.0
        for r in range(10):
            p.check(x + 5.5, y + r * rh + (rh - 9) / 2, 9)
            p.hl(x, x + bw_, y + (r + 1) * rh, RULE_L, 0.6)
        for vx in (x + 20, x + 44, x + bw_ - 40):
            p.vl(vx, y - 12, y + 10 * rh, RULE_L, 0.6)
        # C 월간 인덱스
        x = x + bw_ + 12
        p.section(x, BT, cw_, "월간 인덱스", "MONTHLY INDEX")
        y = BT + 14
        p.box(x, y, cw_, 12, fill=TINT)
        p.text(x + 17, y + 8.6, "날짜", "M", 6, SUB, "c")
        p.text(x + 40, y + 8.6, "일간 페이지에 적어 둔 중요한 기록", "M", 6, SUB)
        y += 12
        for r in range(10):
            p.hl(x, x + cw_, y + (r + 1) * rh, RULE_L, 0.6)
        p.vl(x + 34, y - 12, y + 10 * rh, RULE_L, 0.6)
        # 회고
        ry = BB - 58
        labels = [("이번 달의 성과", "WINS"), ("배운 점 · 아쉬운 점", "LESSONS"), ("다음 달에 집중할 것", "NEXT")]
        rw = (CW - 2 * 12) / 3.0
        for i, (ko, en) in enumerate(labels):
            rx = X0 + i * (rw + 12)
            p.section(rx, ry, rw, ko, en)
            p.hl(rx, rx + rw, ry + 36, RULE_L, 0.6)
            p.hl(rx, rx + rw, ry + 57, RULE_L, 0.6)

    p.form("monthplan", static)
    # 습관 트래커 (월마다 요일이 달라 동적으로 그림)
    hy = BT + 216 + 10
    p.section(X0, hy, CW, "습관 트래커", "HABIT TRACKER")
    p.text(X1, hy + 9.5, "날짜를 누르면 그날의 일간 페이지", "R", 6.4, FAINT, "r")
    lw = 104
    tw_ = 30
    nd = calendar.monthrange(YEAR, m)[1]
    cw = (CW - lw - tw_) / 31.0
    hy += 14
    hh = 22
    nh = 7
    rh = (BB - 58 - 8 - hy - hh) / float(nh)
    for dd in range(1, 32):
        x = X0 + lw + (dd - 1) * cw
        if dd > nd:
            p.box(x, hy, cw, hh + nh * rh, fill=TINT)
            continue
        d = dt.date(YEAR, m, dd)
        col = day_color(d)
        if col != INK:
            p.box(x, hy, cw, hh + nh * rh, fill=HexColor("#FBF3F2") if col == RED else HexColor("#F2F6FB"))
        p.text(x + cw / 2, hy + 9, WD_KO[d.weekday()], "R", 5.6, col if col != INK else FAINT, "c")
        p.text(x + cw / 2, hy + 18.5, str(dd), "SB", 7.4, col, "c")
        p.link(x, hy, cw, hh, dkey(d))
        for r in range(nh):
            p.circle(x + cw / 2, hy + hh + r * rh + rh / 2, 4.6, stroke=RULE, lw=0.6)
    p.text(X0 + 4, hy + 14, "습관", "M", 6.4, SUB)
    p.text(X1 - tw_ / 2, hy + 14, "합계", "M", 6.4, SUB, "c")
    for r in range(nh + 1):
        p.hl(X0, X1, hy + hh + r * rh, RULE if r in (0, nh) else RULE_L, 0.55)
    p.hl(X0, X1, hy, RULE, 0.55)
    p.vl(X0 + lw, hy, hy + hh + nh * rh, RULE, 0.55)
    p.vl(X1 - tw_, hy, hy + hh + nh * rh, RULE, 0.55)
    p.end()


def week_range_str(mon):
    sun = mon + dt.timedelta(6)
    return "%d월 %d일 – %d월 %d일" % (mon.month, mon.day, sun.month, sun.day)


def page_week(p, mon):
    key = wkey(mon)
    sun = mon + dt.timedelta(6)
    rep = mon if in_year(mon) else sun          # 이 주를 대표하는 2027년 날짜
    m = rep.month
    lab = wlabel(mon)
    p.start(key, "%s · %s" % (lab, week_range_str(mon)), 1)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(m), "q%d" % quarter(m)), ("%d월" % m, "m%d" % m),
              (lab, key)], month=m)
    p.text(X0, TOP + 26, lab, "B", 22, INK)
    x = X0 + p.tw(lab, "B", 22) + 10
    p.text(x, TOP + 17, "WEEKLY PLAN · %d" % (YEAR if WEEK_START else (mon + dt.timedelta(3)).year), "SB", 7.2, ACC, cs=1.4)
    p.text(x, TOP + 30, week_range_str(mon), "SB", 9.6, INK)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)
    i = WEEKS.index(mon)
    p.navbtn(X1 - 56, TOP + 9, 26, 24, wkey(WEEKS[i - 1]) if i > 0 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, wkey(WEEKS[i + 1]) if i < len(WEEKS) - 1 else None, right=True)
    bx = X1 - 66
    months = []
    for d in (mon, sun):
        if in_year(d) and d.month not in months:
            months.append(d.month)
    for mm in reversed(months):
        bx -= 62
        p.pill(bx, TOP + 10, 56, 22, "%d월 달력" % mm, "m%d" % mm, fill=ACC_L, size=7.6)

    lw = 178
    dx0 = X0 + lw + 10
    dw = (X1 - dx0) / 7.0
    rev_h = 44
    hab_h = 14 + 4 * 16
    up_h = BH - rev_h - 8 - hab_h - 8      # 위쪽 영역 높이
    hab_y = BT + up_h + 8
    rev_y = BB - rev_h
    dhh = 30

    def static():
        # 주간 나침반
        x = X0
        y = BT
        p.section(x, y, lw, "이번 주 핵심 목표", "BIG 3")
        y += 14
        for r in range(3):
            p.circle(x + 6.5, y + r * 19 + 10, 5.6, fill=ACC_L)
            p.text(x + 6.5, y + r * 19 + 12.6, str(r + 1), "B", 7, ACC, "c")
            p.hl(x + 16, x + lw, y + (r + 1) * 19, RULE_L, 0.6)
        y += 3 * 19 + 8
        p.section(x, y, lw, "역할과 목표", "ROLES & GOALS")
        y += 14
        roles_h = up_h - (14 + 57 + 8) - 14 - 8 - (14 + 52)
        rr = roles_h / 4.0
        for r in range(4):
            p.text(x, y + 11.5, "역할", "SB", 6.2, ACC)
            p.hl(x + 20, x + lw * 0.62, y + 14, RULE, 0.55)
            p.hl(x, x + lw, y + rr - 2, RULE_L, 0.6)
            y += rr
        y += 8
        p.section(x, y, lw, "톱날 갈기", "SHARPEN THE SAW")
        y += 14
        saw = ["신체", "지성", "관계", "정신"]
        for r in range(4):
            sx = x + (r % 2) * (lw / 2 + 3)
            sy = y + (r // 2) * 26
            p.text(sx, sy + 12, saw[r], "SB", 6.6, SUB)
            p.hl(sx + 20, sx + lw / 2 - 3, sy + 22, RULE_L, 0.6)
        # 요일 칸
        for i in range(7):
            cx = dx0 + i * dw
            y = BT + dhh
            p.text(cx + 4, y + 9, "핵심", "M", 5.6, FAINT)
            for r in range(3):
                p.check(cx + 4, y + 12 + r * 18 + 4.5, 8.5)
                p.hl(cx + 3, cx + dw - 3, y + 12 + (r + 1) * 18, RULE_L, 0.6)
            y += 12 + 54 + 4
            p.text(cx + 4, y + 8, "일정 · 메모", "M", 5.6, FAINT)
            nn = int((BT + up_h - y - 10) // 18)
            lh = (BT + up_h - y - 10) / nn
            for r in range(nn - 1):
                p.hl(cx + 3, cx + dw - 3, y + 10 + (r + 1) * lh, RULE_L, 0.5, dash=[1, 2])
        for i in range(8):
            p.vl(dx0 + i * dw, BT, BT + up_h, RULE, 0.5)
        p.hl(dx0, X1, BT + up_h, RULE, 0.5)
        p.hl(dx0, X1, BT + dhh + 12 + 54 + 2, RULE, 0.5)
        # 습관
        p.section(X0, hab_y, CW, "주간 습관", "HABITS")
        for r in range(4):
            yy = hab_y + 14 + r * 16
            p.hl(X0, X1, yy + 16, RULE_L, 0.6)
            for i in range(7):
                p.circle(dx0 + i * dw + dw / 2, yy + 8, 4.8, stroke=RULE, lw=0.6)
        p.vl(dx0, hab_y + 14, hab_y + hab_h, RULE_L, 0.6)
        # 회고
        labels = [("잘한 일", "WINS"), ("아쉬운 점 · 배운 점", "LESSONS"), ("다음 주로 넘길 일", "CARRY OVER")]
        rw = (CW - 2 * 12) / 3.0
        for i, (ko, en) in enumerate(labels):
            rx = X0 + i * (rw + 12)
            p.section(rx, rev_y, rw, ko, en)
            p.hl(rx, rx + rw, rev_y + 30, RULE_L, 0.6)
            p.hl(rx, rx + rw, rev_y + 44, RULE_L, 0.6)

    p.form("weekly", static)
    for i in range(7):
        d = mon + dt.timedelta(i)
        cx = dx0 + i * dw
        ok = in_year(d)
        col = day_color(d) if ok else FAINT
        p.box(cx, BT, dw, dhh, fill=ACC_L if ok else TINT)
        p.text(cx + 6, BT + 21, str(d.day), "B", 16, col)
        nx = cx + 6 + p.tw(str(d.day), "B", 16) + 4
        p.text(nx, BT + 12, WD_KO[d.weekday()] + "요일", "SB", 7, col)
        md = (d.month, d.day)
        if ok and md in HOLIDAYS:
            p.text(nx, BT + 22.5, HOLIDAYS[md], "SB", 6.2, RED)
        elif ok and md in TERMS:
            p.text(nx, BT + 22.5, TERMS[md], "R", 6.2, SUB)
        elif ok and md in MEMOS:
            p.text(nx, BT + 22.5, MEMOS[md], "R", 6.2, SUB)
        else:
            p.text(nx, BT + 22.5, "%d.%d" % (d.month, d.day), "R", 6.2, FAINT)
        if ok:
            p.chev(cx + dw - 8, BT + dhh / 2, 2.6, True, ACC_M, 1.0)
            p.link(cx, BT, dw, dhh, dkey(d))
    p.vl(dx0, BT, BT + dhh, RULE, 0.5)
    p.hl(dx0, X1, BT, RULE, 0.5)
    p.hl(dx0, X1, BT + dhh, RULE, 0.5)
    p.end()


def daily_static(p):
    aw, bw_ = 268, 196
    cw_ = CW - aw - bw_ - 24
    # A 우선순위 업무
    x = X0
    p.section(x, BT, aw, "우선순위 업무", "PRIORITIZED DAILY TASK LIST")
    y = BT + 14
    p.box(x, y + 3, aw, 60, fill=ACC_L, r=5)
    p.text(x + aw - 8, y + 12, "오늘 꼭 끝낼 3가지", "M", 5.8, ACC_M, "r")
    for r in range(3):
        p.circle(x + 12, y + 3 + r * 20 + 11, 6, fill=ACC)
        p.text(x + 12, y + 3 + r * 20 + 13.8, str(r + 1), "B", 7.4, WHITE, "c")
        if r < 2:
            p.hl(x + 24, x + aw - 8, y + 3 + (r + 1) * 20, ACC_M, 0.4)
    y += 69
    p.box(x, y, aw, 12, fill=TINT)
    p.text(x + 10, y + 8.6, "✓", "M", 6.6, SUB, "c")
    p.text(x + 32, y + 8.6, "ABC", "M", 5.8, SUB, "c")
    p.text(x + 50, y + 8.6, "할 일", "M", 6, SUB)
    y += 12
    leg_h = 15
    n = 16
    rh = (BB - leg_h - y) / float(n)
    for r in range(n):
        p.check(x + 5.5, y + r * rh + (rh - 9) / 2, 9)
        p.hl(x, x + aw, y + (r + 1) * rh, RULE_L if r < n - 1 else RULE, 0.6)
    p.vl(x + 20, y - 12, y + n * rh, RULE_L, 0.6)
    p.vl(x + 44, y - 12, y + n * rh, RULE_L, 0.6)
    ly = BB - 3.5
    lx = x
    for s, body in [("✓", "완료"), ("→", "내일로"), ("×", "취소"), ("●", "진행")]:
        p.text(lx + 3, ly, s, "B", 7, SUB, "c")
        p.text(lx + 9, ly, body, "R", 6.2, FAINT)
        lx += 9 + p.tw(body, "R", 6.2) + 8
    p.circle(lx + 3, ly - 2.3, 3.6, stroke=SUB, lw=0.5)
    p.text(lx + 3, ly - 0.4, "D", "B", 4.6, SUB, "c")
    p.text(lx + 9, ly, "위임", "R", 6.2, FAINT)
    p.text(x + aw, ly, "A 필수 · B 중요 · C 선택", "M", 6.2, FAINT, "r")

    # B 시간 계획
    x = X0 + aw + 12
    p.section(x, BT, bw_, "시간 계획", "SCHEDULE")
    y = BT + 14
    hw = 22
    sw = (bw_ - hw) / 2.0
    p.box(x, y, bw_, 12, fill=TINT)
    p.text(x + hw + sw / 2, y + 8.6, "계획 PLAN", "M", 6, SUB, "c")
    p.text(x + hw + sw + sw / 2, y + 8.6, "실제 ACTUAL", "M", 6, SUB, "c")
    y += 12
    hrs = list(range(6, 24))
    rh = (BB - y) / float(len(hrs))
    for i, hr in enumerate(hrs):
        yy = y + i * rh
        p.text(x + hw - 5, yy + 9.5, "%d" % hr, "SB", 7.4, ACC if hr in (9, 12, 18) else SUB, "r")
        p.hl(x + hw, x + bw_, yy + rh / 2, RULE_L, 0.45, dash=[1, 2.2])
        p.hl(x, x + bw_, yy + rh, RULE_L if i < len(hrs) - 1 else RULE, 0.6)
    p.vl(x + hw, y - 12, BB, RULE_L, 0.6)
    p.vl(x + hw + sw, y - 12, BB, RULE_L, 0.6)

    # C 기록
    x = x + bw_ + 12
    p.section(x, BT, cw_, "오늘의 기록", "DAILY NOTES")
    bot_h = 106
    p.dots(x, BT + 18, cw_, BH - 18 - bot_h - 8, 14)
    by = BB - bot_h
    ew = 150
    # 지출
    p.section(x, by, ew, "지출", "EXPENSES")
    erh = (bot_h - 14) / 5.0
    for r in range(5):
        p.hl(x, x + ew, by + 14 + (r + 1) * erh, RULE_L if r < 4 else RULE, 0.6)
    p.vl(x + ew - 52, by + 14, BB, RULE_L, 0.6)
    p.text(x + 2, BB - 5.5, "합계", "SB", 6.6, SUB)
    # 감사·회고
    gx = x + ew + 12
    gw = cw_ - ew - 12
    p.section(gx, by, gw, "감사 · 한 줄 회고", "GRATITUDE")
    for r in range(3):
        p.hl(gx, gx + gw, by + 14 + (r + 1) * 19, RULE_L, 0.6)
    ty = by + 14 + 57 + 8
    p.text(gx, ty + 9, "물", "SB", 6.6, SUB)
    for i in range(8):
        p.circle(gx + 16 + i * 11.5, ty + 6.5, 4.2, stroke=RULE, lw=0.6)
    p.text(gx + 112, ty + 9, "컨디션", "SB", 6.6, SUB)
    for i in range(5):
        cx = gx + 140 + i * 12.5
        p.circle(cx, ty + 6.5, 4.8, stroke=RULE, lw=0.6)
        p.text(cx, ty + 8.6, str(i + 1), "M", 5.4, FAINT, "c")
    ty += 19
    p.text(gx, ty + 8, "수면", "SB", 6.6, SUB)
    p.hl(gx + 18, gx + 58, ty + 10, RULE_L, 0.6)
    p.text(gx + 68, ty + 8, "운동", "SB", 6.6, SUB)
    p.hl(gx + 86, gx + 126, ty + 10, RULE_L, 0.6)
    p.text(gx + 136, ty + 8, "날씨", "SB", 6.6, SUB)
    p.hl(gx + 154, gx + gw, ty + 10, RULE_L, 0.6)


def page_day(p, d):
    key = dkey(d)
    idx = (d - DAYS[0]).days
    md = (d.month, d.day)
    p.start(key)
    p.chrome([(str(YEAR), "year"), ("Q%d" % quarter(d.month), "q%d" % quarter(d.month)),
              ("%d월" % d.month, "m%d" % d.month), (wlabel(d), wkey(d)),
              ("%d일 %s" % (d.day, WD_KO[d.weekday()]), key)], month=d.month)
    col = day_color(d)
    # 날짜 제목
    num = str(d.day)
    p.text(X0, TOP + 31, num, "B", 34, col)
    x = X0 + p.tw(num, "B", 34) + 9
    p.text(x, TOP + 15, WD_KO[d.weekday()] + "요일", "B", 12.5, col)
    wx = x + p.tw(WD_KO[d.weekday()] + "요일", "B", 12.5) + 7
    p.text(wx, TOP + 15, "%d년 %d월" % (YEAR, d.month), "M", 8.6, SUB)
    info = [WD_EN[d.weekday()][:3] + " · " + MON_EN[d.month][:3] + " %d" % d.day, lunar_str(d, "음력 ")]
    p.text(x, TOP + 29, "  ·  ".join(info), "M", 7, SUB, cs=0.2)
    ix = x + p.tw("  ·  ".join(info), "M", 7) + 0.2 * len("  ·  ".join(info)) + 8
    tags = []
    if md in HOLIDAYS:
        tags.append((HOLIDAYS[md], RED, HexColor("#FBECEA")))
    if md in TERMS:
        tags.append((TERMS[md], ACC, ACC_L))
    if md in MEMOS:
        tags.append((MEMOS[md], ACC, ACC_L))
    for label, fg, bg in tags:
        w = p.tw(label, "SB", 7) + 12
        p.box(ix, TOP + 19.5, w, 13, fill=bg, r=6.5)
        p.ctext(ix, TOP + 19.5, w, 13, label, "SB", 7, fg)
        ix += w + 5
    p.text(x, TOP + 40, "%d일째 · 남은 날 %d일" % (idx + 1, NDAYS - idx - 1), "R", 6.4, FAINT)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)

    # 오른쪽: 주간 스트립 + 이전/다음
    bw = 26
    cw = 37
    sx1 = X1 - 2 * bw - 8
    sx0 = sx1 - 7 * cw
    p.navbtn(X1 - 2 * bw - 2, TOP + 5, bw, 30, dkey(DAYS[idx - 1]) if idx > 0 else None, right=False)
    p.navbtn(X1 - bw, TOP + 5, bw, 30, dkey(DAYS[idx + 1]) if idx < NDAYS - 1 else None, right=True)
    mon = monday(d)
    for i in range(7):
        dd = mon + dt.timedelta(i)
        cx = sx0 + i * cw
        ok = in_year(dd)
        if dd == d:
            p.box(cx + 2, TOP + 4, cw - 4, 32, fill=ACC, r=6)
            p.text(cx + cw / 2, TOP + 15, WD_KO[COLS[i]], "M", 6.4, WHITE, "c")
            p.text(cx + cw / 2, TOP + 29.5, str(dd.day), "B", 11.5, WHITE, "c")
        else:
            c2 = day_color(dd, SUB) if ok else RULE
            p.box(cx + 2, TOP + 4, cw - 4, 32, fill=TINT, r=6)
            p.text(cx + cw / 2, TOP + 15, WD_KO[COLS[i]], "M", 6.4, c2 if ok else RULE, "c")
            p.text(cx + cw / 2, TOP + 29.5, str(dd.day), "SB", 11.5, c2, "c")
        if ok:
            p.link(cx, TOP + 2, cw, 36, dkey(dd))
    p.pill(sx0 - 44, TOP + 4, 40, 32, wlabel(d), wkey(d), fill=ACC_L, font="B", size=9)

    p.form("daily", lambda: daily_static(p))
    p.end()


# ── 도구 ──
def page_anniv(p):
    p.start("anniv", "기념일 · 생일", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("기념일 · 생일", "anniv")])
    p.title("기념일 · 생일", "SPECIAL DATES", "매년 반복되는 소중한 날을 월별로 정리")
    gx, gy = 12, 8
    bw = (CW - 3 * gx) / 4.0
    bh = (BH - 2 * gy) / 3.0
    for m in range(1, 13):
        x = X0 + ((m - 1) % 4) * (bw + gx)
        y = BT + ((m - 1) // 4) * (bh + gy)
        p.section(x, y, bw, "%d월" % m, MON_EN[m])
        p.link(x, y, bw, 14, "m%d" % m)
        n = 6
        rh = (bh - 14) / n
        for r in range(n):
            p.hl(x, x + bw, y + 14 + (r + 1) * rh, RULE_L, 0.6)
        p.vl(x + 26, y + 14, y + bh, RULE_L, 0.6)
    p.end()


def page_pixels(p):
    p.start("pixels", "연간 기록 · 픽셀", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("연간 기록", "pixels")])
    p.title("연간 기록", "YEAR IN PIXELS", "하루 한 칸 — 기분·컨디션·습관을 색으로 남기기")
    lw = 18
    cw = 22
    hh = 14
    rh = (BH - hh) / 31.0
    gx = X0
    for m in range(1, 13):
        x = gx + lw + (m - 1) * cw
        p.ctext(x, BT, cw, hh, str(m), "B", 7.4, ACC)
        p.link(x, BT, cw, hh, "m%d" % m)
        nd = calendar.monthrange(YEAR, m)[1]
        for dd in range(1, 32):
            y = BT + hh + (dd - 1) * rh
            if dd > nd:
                p.box(x, y, cw, rh, fill=TINT)
            elif dt.date(YEAR, m, dd).weekday() == 6:
                p.box(x + cw - 3.2, y + 1.2, 2, 2, fill=HexColor("#E3A9A3"))
    for dd in range(1, 32):
        p.text(gx + lw - 4, BT + hh + (dd - 1) * rh + rh / 2 + 2.4, str(dd), "M", 6.4, SUB, "r")
    for r in range(32):
        p.hl(gx + lw, gx + lw + 12 * cw, BT + hh + r * rh, RULE if r in (0, 31) else RULE_L, 0.5)
    for i in range(13):
        p.vl(gx + lw + i * cw, BT + hh, BB, RULE, 0.5)
    # 범례
    x = gx + lw + 12 * cw + 24
    w = X1 - x
    p.section(x, BT, w, "색 범례", "COLOR KEY")
    kw = (w - 5 * 10) / 6.0
    for i in range(6):
        kx = x + i * (kw + 10)
        p.box(kx, BT + 24, 16, 16, stroke=RULE, r=3)
        p.hl(kx + 22, kx + kw, BT + 39, RULE_L, 0.6)
    p.text(x, BT + 58, "예) 기분: 최고 · 좋음 · 보통 · 지침 · 나쁨  /  빨간 점 = 일요일", "R", 6.8, FAINT)
    # 월별 요약표
    y = BT + 74
    p.section(x, y, w, "월별 요약", "MONTHLY SCORECARD")
    y += 14
    ncol = 6
    mw_ = 44
    cw2 = (w - mw_) / ncol
    hh2 = 22
    rh2 = (BB - y - hh2) / 13.0
    p.box(x, y, w, hh2, fill=TINT)
    p.text(x + mw_ / 2, y + 14, "월", "SB", 7, SUB, "c")
    for i in range(ncol):
        p.text(x + mw_ + i * cw2 + 5, y + 9, "지표 %d" % (i + 1), "M", 5.6, FAINT)
    for m in range(1, 14):
        yy = y + hh2 + (m - 1) * rh2
        if m <= 12:
            p.ctext(x, yy, mw_, rh2, "%d월" % m, "SB", 7.6, ACC)
            p.link(x, yy, mw_, rh2, "mp%d" % m)
        else:
            p.box(x, yy, w, rh2, fill=ACC_L)
            p.ctext(x, yy, mw_, rh2, "연간", "B", 7.6, ACC)
        p.hl(x, x + w, yy + rh2, RULE_L if m < 13 else RULE, 0.6)
    p.hl(x, x + w, y + hh2, RULE, 0.6)
    for i in range(ncol):
        p.vl(x + mw_ + i * cw2, y, BB, RULE_L, 0.6)
    p.end()


def page_finance(p):
    p.start("finance", "연간 가계부", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("연간 가계부", "finance")])
    p.title("연간 가계부", "ANNUAL BUDGET", "월별 수입 · 지출 · 저축을 한 장에")
    lw, tw_ = 104, 80
    cw = (CW - lw - tw_) / 12.0
    hh = 18
    sections = [("수입", 3), ("고정 지출", 5), ("변동 지출", 5), ("저축 · 투자", 3)]
    nrows = sum(n for _, n in sections) + len(sections) + 2
    rh = (BH - hh) / float(nrows)
    p.box(X0, BT, CW, hh, fill=ACC)
    p.text(X0 + 8, BT + 12, "항목", "B", 7.6, WHITE)
    for m in range(1, 13):
        p.ctext(X0 + lw + (m - 1) * cw, BT, cw, hh, "%d월" % m, "B", 7.6, WHITE)
        p.link(X0 + lw + (m - 1) * cw, BT, cw, hh, "mp%d" % m)
    p.ctext(X1 - tw_, BT, tw_, hh, "합계", "B", 7.6, WHITE)
    y = BT + hh
    for name, n in sections:
        p.box(X0, y, CW, rh, fill=ACC_L)
        p.text(X0 + 8, y + rh / 2 + 2.8, name, "B", 7.6, ACC)
        p.text(X1 - 6, y + rh / 2 + 2.6, "소계", "M", 6, ACC_M, "r")
        p.hl(X0, X1, y + rh, RULE, 0.5)
        y += rh
        for r in range(n):
            p.hl(X0, X1, y + rh, RULE_L, 0.5)
            y += rh
    for name in ["총지출", "잔액 (수입 − 지출 − 저축)"]:
        p.box(X0, y, CW, rh, fill=TINT)
        p.text(X0 + 8, y + rh / 2 + 2.8, name, "B", 7.2, INK)
        p.hl(X0, X1, y + rh, RULE, 0.5)
        y += rh
    for m in range(13):
        p.vl(X0 + lw + m * cw, BT + hh, BB, RULE_L, 0.5)
    p.vl(X0, BT, BB, RULE, 0.5)
    p.vl(X1, BT, BB, RULE, 0.5)
    p.end()


def page_reading(p):
    p.start("reading", "독서 · 강의 목록", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("독서 · 강의", "reading")])
    p.title("독서 · 강의 목록", "READING & LEARNING LOG", "읽은 책, 들은 강의, 본 영화를 기록")
    gap = 16
    tw_ = (CW - gap) / 2.0
    n = 20
    hh = 14
    rh = (BH - hh) / float(n)
    for t in range(2):
        x = X0 + t * (tw_ + gap)
        p.box(x, BT, tw_, hh, fill=ACC_L)
        cols = [(0, 22, "No"), (22, 190, "제목"), (212, 80, "저자 · 강사"), (292, 50, "완료일"), (342, tw_ - 342, "평점")]
        for cx, cw, label in cols:
            p.ctext(x + cx, BT, cw, hh, label, "SB", 6.6, ACC)
        for r in range(n):
            y = BT + hh + r * rh
            p.ctext(x, y, 22, rh, "%d" % (t * n + r + 1), "M", 6.8, FAINT)
            sx = x + 342 + (tw_ - 342 - 5 * 12) / 2 + 6
            for s in range(5):
                p.text(sx + s * 12, y + rh / 2 + 3.2, "☆", "R", 9, RULE, "c")
            p.hl(x, x + tw_, y + rh, RULE_L, 0.6)
        for cx, cw, label in cols[1:]:
            p.vl(x + cx, BT + hh, BB, RULE_L, 0.6)
    p.end()


def page_bucket(p):
    p.start("bucket", "버킷리스트", 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("버킷리스트", "bucket")])
    p.title("%d 버킷리스트" % YEAR, "BUCKET LIST", "올해 꼭 해 보고 싶은 것들")
    gap = 20
    tw_ = (CW - 2 * gap) / 3.0
    cats = [("경험 · 여행", "EXPERIENCE"), ("성장 · 도전", "GROWTH"), ("사람 · 나눔", "PEOPLE")]
    n = 18
    rh = (BH - 14) / float(n)
    for t in range(3):
        x = X0 + t * (tw_ + gap)
        p.section(x, BT, tw_, cats[t][0], cats[t][1])
        for r in range(n):
            y = BT + 14 + r * rh
            p.check(x + 1, y + (rh - 10) / 2, 10)
            p.hl(x + 16, x + tw_, y + rh, RULE_L, 0.6)
    p.end()


def page_project(p, n):
    key = "proj%d" % n
    p.start(key, "프로젝트 %d" % n, 1)
    p.chrome([(str(YEAR), "cover"), ("홈", "index"), ("프로젝트 %d" % n, key)])
    p.text(X0, TOP + 26, "P%d" % n, "B", 22, ACC)
    x = X0 + p.tw("P%d" % n, "B", 22) + 10
    p.text(x, TOP + 13, "PROJECT PLANNER", "SB", 6.6, ACC, cs=1.4)
    p.hl(x, x + 330, TOP + 32, RULE, 0.7)
    p.text(x + 350, TOP + 30, "시작", "SB", 7.4, SUB)
    p.hl(x + 372, x + 440, TOP + 32, RULE, 0.7)
    p.text(x + 452, TOP + 30, "기한", "SB", 7.4, SUB)
    p.hl(x + 474, x + 542, TOP + 32, RULE, 0.7)
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "proj%d" % (n - 1) if n > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "proj%d" % (n + 1) if n < 6 else None, right=True)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)

    def static():
        lw, mw_ = 250, 370
        rw = CW - lw - mw_ - 24
        x = X0
        y = BT
        for ko, en, rows in [("목적", "WHY", 3), ("완료의 모습 · 성공 기준", "DONE MEANS", 3), ("자원 · 관련된 사람", "RESOURCES", 3)]:
            p.section(x, y, lw, ko, en)
            p.lines(x, y + 14, lw, rows, 19, RULE_L)
            y += 14 + rows * 19 + 9
        p.section(x, y, lw, "마일스톤", "MILESTONES")
        nn = int((BB - y - 14) // 19)
        rh = (BB - y - 14) / nn
        for r in range(nn):
            p.check(x + 1, y + 14 + r * rh + (rh - 9) / 2, 9)
            p.hl(x, x + lw, y + 14 + (r + 1) * rh, RULE_L, 0.6)
        p.vl(x + lw - 44, y + 14, BB, RULE_L, 0.6)
        x = X0 + lw + 12
        p.section(x, BT, mw_, "작업 목록", "ACTION ITEMS")
        y = BT + 14
        p.box(x, y, mw_, 12, fill=TINT)
        p.text(x + 10, y + 8.6, "✓", "M", 6.6, SUB, "c")
        p.text(x + 26, y + 8.6, "작업", "M", 6, SUB)
        p.text(x + mw_ - 84, y + 8.6, "담당", "M", 6, SUB, "c")
        p.text(x + mw_ - 44, y + 8.6, "기한", "M", 6, SUB, "c")
        p.text(x + mw_ - 12, y + 8.6, "상태", "M", 6, SUB, "c")
        y += 12
        nn = 20
        rh = (BB - y) / float(nn)
        for r in range(nn):
            p.check(x + 5.5, y + r * rh + (rh - 9) / 2, 9)
            p.hl(x, x + mw_, y + (r + 1) * rh, RULE_L, 0.6)
        for vx in (x + 20, x + mw_ - 104, x + mw_ - 64, x + mw_ - 24):
            p.vl(vx, y - 12, BB, RULE_L, 0.6)
        x = x + mw_ + 12
        p.section(x, BT, rw, "메모 · 아이디어", "NOTES")
        p.dots(x, BT + 18, rw, BH - 18, 14)

    p.form("project", static)
    p.end()


NOTE_KINDS = [("줄", "lined"), ("도트", "dot"), ("모눈", "grid")]


def page_notes_index(p):
    p.start("notesidx", "노트 색인", 1)
    p.chrome([(str(YEAR), "cover"), ("노트", "notesidx")], menu="notesidx")
    p.title("노트 색인", "NOTES INDEX", "번호를 누르면 해당 노트로 이동 · 제목을 적어 두면 찾기 쉬워요")
    gap = 16
    tw_ = (CW - 2 * gap) / 3.0
    n = 10
    rh = (BH - 14) / float(n)
    for t in range(3):
        x = X0 + t * (tw_ + gap)
        p.section(x, BT, tw_, "%s 노트" % NOTE_KINDS[t][0], NOTE_KINDS[t][1].upper())
        for r in range(n):
            no = t * 10 + r + 1
            y = BT + 14 + r * rh
            p.pill(x, y + (rh - 22) / 2, 34, 22, "%02d" % no, "n%d" % no, fill=ACC_L, font="B", size=8.6)
            p.hl(x + 42, x + tw_ - 50, y + rh - 8, RULE_L, 0.6)
            p.hl(x + tw_ - 42, x + tw_, y + rh - 8, RULE_L, 0.6)
            if r == 0:
                p.text(x + 42, y + 9, "제목", "M", 5.8, FAINT)
                p.text(x + tw_ - 42, y + 9, "날짜", "M", 5.8, FAINT)
            p.link(x, y, tw_, rh, "n%d" % no)
    p.end()


def page_note(p, no):
    key = "n%d" % no
    kind = NOTE_KINDS[(no - 1) // 10]
    p.start(key, "노트 %02d (%s)" % (no, kind[0]), 2)
    p.chrome([(str(YEAR), "cover"), ("노트", "notesidx"), ("%02d" % no, key)], menu="notesidx")
    p.text(X0, TOP + 26, "%02d" % no, "B", 22, ACC)
    x = X0 + p.tw("%02d" % no, "B", 22) + 10
    p.text(x, TOP + 13, "NOTE · " + kind[1].upper(), "SB", 6.6, ACC, cs=1.4)
    p.hl(x, x + 470, TOP + 32, RULE, 0.7)
    p.text(x + 490, TOP + 30, "날짜", "SB", 7.4, SUB)
    p.hl(x + 512, x + 600, TOP + 32, RULE, 0.7)
    p.navbtn(X1 - 56, TOP + 9, 26, 24, "n%d" % (no - 1) if no > 1 else None, right=False)
    p.navbtn(X1 - 26, TOP + 9, 26, 24, "n%d" % (no + 1) if no < 30 else None, right=True)
    p.hl(X0, X1, TOP + HDR, INK, 1.1)

    def static():
        if kind[1] == "lined":
            n = 20
            p.lines(X0, BT, CW, n, BH / float(n), RULE_L, 0.6)
        elif kind[1] == "dot":
            p.dots(X0, BT, CW, BH, 14)
        else:
            p.grid(X0, BT, CW, BH, 14)

    p.form("note_" + kind[1], static)
    p.end()


# ─────────────────────────────────────────────────────────────────────────────
def compose(p):
    """가로형 페이지 순서"""
    page_cover(p)
    page_index(p)
    page_year(p)
    page_glance(p)
    page_mission(p)
    page_goals(p, 1)
    page_goals(p, 2)
    for d in DAYS:
        if d.day == 1:
            if d.month % 3 == 1:
                page_quarter(p, quarter(d.month))
            page_month(p, d.month)
            page_monthplan(p, d.month)
        if d == monday(d) or d == DAYS[0]:
            page_week(p, monday(d))
        page_day(p, d)
    p.outline.append((0, "도구 · 프로젝트", "anniv"))
    page_anniv(p)
    page_pixels(p)
    page_finance(p)
    page_reading(p)
    page_bucket(p)
    for n in range(1, 7):
        page_project(p, n)
    p.outline.append((0, "노트", "notesidx"))
    page_notes_index(p)
    for no in range(1, 31):
        page_note(p, no)


def build(year, out, font_dir, compose_fn=None):
    configure(year)
    tmp = out + ".tmp"
    p = Painter(tmp, font_dir)
    (compose_fn or compose)(p)
    p.c.save()

    # ── 링크·목차 후처리: 뷰어 호환성이 가장 넓은 '명시적 페이지 목적지'로 기록 ──
    pdf = pikepdf.open(tmp)
    pages = pdf.pages
    per_page = {}
    missing = set()
    for pg, rect, key in p.links:
        if key not in p.dests:
            missing.add(key)
            continue
        per_page.setdefault(pg, []).append((rect, p.dests[key]))
    if missing:
        raise SystemExit("링크 대상 없음: %s" % sorted(missing))
    nlinks = 0
    for pg, items in per_page.items():
        annots = pikepdf.Array()
        for rect, target in items:
            annots.append(pdf.make_indirect(pikepdf.Dictionary(
                Type=pikepdf.Name.Annot, Subtype=pikepdf.Name.Link,
                Rect=[round(v, 2) for v in rect], Border=[0, 0, 0],
                Dest=pikepdf.Array([pages[target].obj, pikepdf.Name.XYZ, 0, H, 0]),
            )))
            nlinks += 1
        pages[pg].obj.Annots = annots
    with pdf.open_outline() as ol:
        stack = []
        for level, title, key in p.outline:
            item = pikepdf.OutlineItem(title, p.dests[key])
            del stack[level:]
            if level == 0 or not stack:
                ol.root.append(item)
            else:
                stack[-1].children.append(item)
            stack.append(item)
    pdf.Root.PageLayout = pikepdf.Name.SinglePage
    pdf.save(out, compress_streams=True, object_stream_mode=pikepdf.ObjectStreamMode.disable,
             linearize=False)
    pdf.close()
    os.remove(tmp)
    print("%s: %d쪽, 링크 %d개, %.1f MB" % (out, len(p.dests) and p.page, nlinks, os.path.getsize(out) / 1e6))


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    fonts = os.environ.get("PLANNER_FONT_DIR", os.path.join(here, "fonts"))
    years = [int(a) for a in sys.argv[1:]] or sorted(THEMES)
    for year in years:
        build(year, os.path.join(here, "%d-planner.pdf" % year), fonts)
