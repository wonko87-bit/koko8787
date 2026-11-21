#!/usr/bin/env python3
"""간지 + 노트 템플릿 PDF 생성기 (삼성노트 · 갤럭시탭 세로 화면용).

    python3 make_notebook.py [출력.pdf]      # 기본: notebook-templates.pdf

구성: 표지 → 목차 → 간지 10장 → 템플릿 목록 → 노트 템플릿 24종(각 1장)
템플릿 페이지를 삼성노트에서 복사해 원하는 간지 뒤에 붙여 쓰는 용도.
화면 폭(약 19cm)에 맞춰 보면 1pt ≈ 실제 1pt 이므로 줄 간격은 실제 mm 로 잡았다.

필요: reportlab, pikepdf / 폰트: Pretendard (make_planner.py 와 같은 위치)
"""
import math
import os
import sys

import pikepdf
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

MM = 72 / 25.4
W, H = 540, 864           # 세로 16:10
TABW = 28                 # 오른쪽 간지 탭 폭
NAVH = 22                 # 상단 바 높이
X0, X1 = 14, W - TABW - 10
CW = X1 - X0              # 488pt ≈ 172mm
TOP = NAVH + 8            # 필기 영역 시작
BB = H - 10               # 필기 영역 끝
HEAD = 24                 # TITLE/DATE 머리줄 높이
BODY = TOP + HEAD         # 본문 시작 (높이 800pt ≈ 282mm)
BT = TOP + 50             # 목차·간지 등 '제목 있는 페이지'의 본문 시작

INK = HexColor("#262A33")
SUB = HexColor("#6B7280")
FAINT = HexColor("#A3A9B3")
RULE = HexColor("#C3C8D1")
RULE_L = HexColor("#E1E4EA")
TINT = HexColor("#F5F6F8")
WHITE = HexColor("#FFFFFF")
UI = HexColor("#34475C")
UI_L = HexColor("#E8EDF2")
PAPER = HexColor("#F6F3EC")
LINE = HexColor("#A9BCD0")      # 노트 줄 (연한 파랑 회색)
LINE_L = HexColor("#D3DEE9")
GRID = HexColor("#C4D2E0")
GRID_L = HexColor("#E2E9F0")
GRID_D = HexColor("#97ADC4")
DOT = HexColor("#9DAAB8")
RED = HexColor("#E0645A")       # 빨간 마진 줄
WONGO = HexColor("#E07A5F")     # 원고지
GREEN = HexColor("#5FA27E")     # 칸노트
HANJA = HexColor("#C9776A")
STAFF = HexColor("#8C96A4")

# 간지 10장 색 (진한색, 옅은색)
SEC = [("#D9645A", "#FBE9E7"), ("#E08A4E", "#FCEEE3"), ("#C99A2E", "#FAF2DC"), ("#86A847", "#EEF5E1"),
       ("#3E9C74", "#E2F2EA"), ("#3A949C", "#E1F1F2"), ("#4C84C8", "#E5EEF8"), ("#6670C4", "#ECEDF8"),
       ("#8E66B8", "#F1EBF7"), ("#C9638F", "#F8E8EF")]
SEC = [(HexColor(a), HexColor(b)) for a, b in SEC]
NSEC = len(SEC)


# ─────────────────────────────────────────────────────────────────────────────
class Painter:
    def __init__(self, path, font_dir):
        for name, fn in [("R", "Regular"), ("M", "Medium"), ("SB", "SemiBold"), ("B", "Bold")]:
            pdfmetrics.registerFont(TTFont(name, os.path.join(font_dir, "Pretendard-%s.ttf" % fn)))
        self.c = canvas.Canvas(path, pagesize=(W, H), pageCompression=1)
        self.c.setTitle("노트 · 간지 템플릿")
        self.c.setSubject("간지 10장과 노트 템플릿 24종")
        self.page = 0
        self.links, self.dests, self.outline, self.forms = [], {}, [], set()

    def start(self, key, outline=None, level=0):
        self.dests[key] = self.page
        if outline:
            self.outline.append((level, outline, key))

    def end(self):
        self.c.showPage()
        self.page += 1

    def link(self, x, top, w, h, key):
        self.links.append((self.page, (x, H - top - h, x + w, H - top), key))

    # 기본 도형 (좌표는 '위에서부터')
    def text(self, x, y, s, font="R", size=9, color=INK, align="l", cs=0):
        c = self.c
        c.setFillColor(color)
        c.setFont(font, size)
        if cs:
            wd = pdfmetrics.stringWidth(s, font, size) + cs * max(len(s) - 1, 0)
            x -= wd / 2 if align == "c" else wd if align == "r" else 0
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
        args = (x, H - top - h, w, h)
        if r:
            c.roundRect(*args, r, stroke=int(stroke is not None), fill=int(fill is not None))
        else:
            c.rect(*args, stroke=int(stroke is not None), fill=int(fill is not None))

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

    def clip(self, x, top, w, h):
        c = self.c
        c.saveState()
        p = c.beginPath()
        p.rect(x, H - top - h, w, h)
        c.clipPath(p, stroke=0, fill=0)

    def unclip(self):
        self.c.restoreState()

    def dotrow(self, x0, x1, y, step, color=DOT, size=1.1):
        """둥근 끝 점선 = 점 한 줄"""
        c = self.c
        c.setStrokeColor(color)
        c.setLineWidth(size)
        c.setLineCap(1)
        c.setDash([0.05, step - 0.05])
        c.line(x0, H - y, x1 + 0.06, H - y)
        c.setDash([])
        c.setLineCap(0)

    def grid(self, x, top, w, h, step, color=GRID, lw=0.45):
        nx, ny = int(w / step + 1e-6), int(h / step + 1e-6)
        ox, oy = x + (w - nx * step) / 2, top + (h - ny * step) / 2
        for i in range(nx + 1):
            self.vl(ox + i * step, oy, oy + ny * step, color, lw)
        for j in range(ny + 1):
            self.hl(ox, ox + nx * step, oy + j * step, color, lw)

    def section(self, x, top, w, ko, en=None, color=UI, h=14):
        self.text(x, top + 9.5, ko, "B", 8.4, color)
        if en:
            self.text(x + self.tw(ko, "B", 8.4) + 6, top + 9.5, en, "M", 5.6, FAINT, cs=0.7)
        self.hl(x, x + w, top + h - 0.5, color, 1.0)

    def pill(self, x, top, w, h, label, key, fill=UI_L, color=UI, font="SB", size=8, stroke=None, r=None):
        self.box(x, top, w, h, fill=fill, stroke=stroke, r=h / 2 if r is None else r, lw=0.5)
        self.ctext(x, top, w, h, label, font, size, color)
        if key:
            self.link(x, top, w, h, key)

    def form(self, name, draw, bbox=(0, 0, W, H)):
        if name not in self.forms:
            self.c.beginForm(name, *bbox)
            draw()
            self.c.endForm()
            self.forms.add(name)

    # ── 공통 크롬: 상단 바(목차·템플릿) + 오른쪽 간지 탭 ──
    def _chrome_static(self):
        self.box(0, 0, W - TABW, NAVH, fill=TINT)
        self.hl(0, W - TABW, NAVH, RULE, 0.6)
        for label, key, x, w in NAV:
            self.ctext(x, 0, w, NAVH, label, "M", 7.8, SUB)
        th = H / float(NSEC)
        for i, (main, light) in enumerate(SEC):
            self.box(W - TABW, i * th, TABW, th, fill=light)
            self.box(W - TABW, i * th, 3.5, th, fill=main)
            self.ctext(W - TABW + 2, i * th, TABW - 2, th, "%d" % (i + 1), "B", 9.5, main)
            if i:
                self.hl(W - TABW, W, i * th, WHITE, 1.2)

    def chrome(self, label=None, active=None, menu=None):
        self.form("chrome", self._chrome_static)
        self.c.doForm("chrome")
        for lab, key, x, w in NAV:
            self.link(x, 0, w, NAVH, key)
            if key == menu:
                self.box(x + 2, 4, w - 4, NAVH - 8, fill=UI, r=4)
                self.ctext(x, 0, w, NAVH, lab, "SB", 7.8, WHITE)
        th = H / float(NSEC)
        for i in range(NSEC):
            self.link(W - TABW, i * th, TABW, th, "s%d" % (i + 1))
        if active:
            main = SEC[active - 1][0]
            self.box(W - TABW, (active - 1) * th, TABW, th, fill=main)
            self.ctext(W - TABW, (active - 1) * th, TABW, th, "%d" % active, "B", 10, WHITE)
        if label:
            self.text(NAV[-1][2] + NAV[-1][3] + 10, NAVH / 2 + 3, label, "M", 7.4, FAINT)

    def title(self, ko, en=None, sub=None):
        self.text(X0, TOP + 26, ko, "B", 22, INK)
        x = X0 + self.tw(ko, "B", 22) + 10
        if en:
            self.text(x, TOP + 17, en, "SB", 7.2, UI, cs=1.4)
        if sub:
            self.text(x, TOP + 29, sub, "R", 7.6, SUB)
        self.hl(X0, X1, TOP + 42, INK, 1.1)


NAV = [("목차", "index", 6, 40), ("템플릿", "tmpl", 46, 48)]


# ─────────────────────────────────────────────────────────────────────────────
# 노트 템플릿 — 모두 (p, x, y, w, h) 본문 영역에 그린다
# ─────────────────────────────────────────────────────────────────────────────
def note_header(p):
    """모닝글로리식 머리줄: TITLE ______  DATE  .  .   No. ___"""
    y = TOP + 15
    p.text(X0, y, "TITLE", "SB", 5.6, FAINT, cs=0.8)
    p.hl(X0 + 28, X1 - 196, y + 2, RULE, 0.6)
    dx = X1 - 182
    p.text(dx, y, "DATE", "SB", 5.6, FAINT, cs=0.8)
    p.hl(dx + 26, dx + 116, y + 2, RULE, 0.6)
    for k in (1, 2):
        p.text(dx + 26 + k * 30, y, ".", "B", 8, FAINT, "c")
    p.text(X1 - 52, y, "No.", "SB", 6.2, FAINT)
    p.hl(X1 - 36, X1, y + 2, RULE, 0.6)


def lines(p, x, y, w, h, pitch, color=LINE, lw=0.5):
    n = int(h / pitch + 1e-6)
    for i in range(1, n + 1):
        p.hl(x, x + w, y + i * pitch, color, lw)


def t_lined(pitch):
    return lambda p, x, y, w, h: lines(p, x, y, w, h, pitch)


def t_margin(p, x, y, w, h):
    lines(p, x, y, w, h, 7 * MM)
    p.vl(x + 30 * MM, y, y + h, RED, 0.8)


def t_blank(p, x, y, w, h):
    pass


def t_grid(step, color, lw):
    return lambda p, x, y, w, h: p.grid(x, y, w, h, step, color, lw)


def t_graph(p, x, y, w, h):
    nx, ny = int(w / (10 * MM)) * 10, int(h / (10 * MM)) * 10
    ox, oy = x + (w - nx * MM) / 2, y + (h - ny * MM) / 2
    for pass_ in range(3):
        for i in range(nx + 1):
            k = 2 if i % 10 == 0 else 1 if i % 5 == 0 else 0
            if k == pass_:
                p.vl(ox + i * MM, oy, oy + ny * MM, [GRID_L, GRID, GRID_D][k], [0.25, 0.45, 0.65][k])
        for j in range(ny + 1):
            k = 2 if j % 10 == 0 else 1 if j % 5 == 0 else 0
            if k == pass_:
                p.hl(ox, ox + nx * MM, oy + j * MM, [GRID_L, GRID, GRID_D][k], [0.25, 0.45, 0.65][k])


def t_dot(p, x, y, w, h, step=5 * MM):
    nx, ny = int(w / step + 1e-6), int(h / step + 1e-6)
    ox, oy = x + (w - nx * step) / 2, y + (h - ny * step) / 2
    for j in range(ny + 1):
        p.dotrow(ox, ox + nx * step, oy + j * step, step)


def t_cornell(p, x, y, w, h):
    cue, summ = 55 * MM, 50 * MM
    lines(p, x, y, w, h - summ, 8 * MM)
    lines(p, x, y + h - summ, w, summ, 8 * MM)
    p.vl(x + cue, y, y + h - summ, UI, 0.9)
    p.hl(x, x + w, y + h - summ, UI, 0.9)
    p.text(x + 2, y + 10, "단서 · 키워드 · 질문", "SB", 6.2, FAINT)
    p.text(x + cue + 6, y + 10, "필기", "SB", 6.2, FAINT)
    p.text(x + 2, y + h - summ + 11, "요약", "SB", 6.2, FAINT)


def t_twocol(p, x, y, w, h):
    gap = 10 * MM
    cw = (w - gap) / 2
    for c in range(2):
        lines(p, x + c * (cw + gap), y, cw, h, 8 * MM)
    p.vl(x + w / 2, y, y + h, RULE, 0.5, dash=[2, 2])


def t_english(p, x, y, w, h):
    s, gap = 3 * MM, 4.5 * MM
    pitch = 3 * s + gap
    n = int((h + gap) / pitch)
    oy = y + (h - (n * pitch - gap)) / 2
    for i in range(n):
        yy = oy + i * pitch
        p.hl(x, x + w, yy, LINE, 0.5)
        p.hl(x, x + w, yy + s, LINE_L, 0.5, dash=[2, 2])
        p.hl(x, x + w, yy + 2 * s, RED, 0.6)
        p.hl(x, x + w, yy + 3 * s, LINE, 0.5)


def t_vocab(p, x, y, w, h):
    gap = 10 * MM
    cw = (w - gap) / 2
    hh, rh = 16, 8 * MM
    n = int((h - hh) / rh)
    for c in range(2):
        cx = x + c * (cw + gap)
        p.box(cx, y, cw, hh, fill=TINT)
        p.text(cx + 9, y + 11, "✓", "M", 7, SUB, "c")
        p.text(cx + 24, y + 11, "단어  WORD", "SB", 6.6, SUB)
        p.text(cx + 24 + cw * 0.42, y + 11, "뜻  MEANING", "SB", 6.6, SUB)
        for r in range(n):
            p.check(cx + 4.5, y + hh + r * rh + (rh - 9) / 2, 9)
            p.hl(cx, cx + cw, y + hh + (r + 1) * rh, LINE, 0.5)
        p.vl(cx + 18, y, y + hh + n * rh, LINE_L, 0.5)
        p.vl(cx + 18 + cw * 0.42, y, y + hh + n * rh, LINE, 0.5, dash=[2, 2])


def t_wongo(p, x, y, w, h):
    cols = 20
    cell = w / cols
    gap = cell * 0.3
    rows = int((h - 8 + gap) / (cell + gap))
    th = rows * cell + (rows - 1) * gap
    oy = y + (h - th) / 2
    hx = x
    p.box(hx - 3, oy - 4, cols * cell + 6, th + 8, stroke=WONGO, lw=1.0)
    for r in range(rows):
        ry = oy + r * (cell + gap)
        p.hl(hx, hx + cols * cell, ry, WONGO, 0.5)
        p.hl(hx, hx + cols * cell, ry + cell, WONGO, 0.5)
        for k in range(cols + 1):
            p.vl(hx + k * cell, ry, ry + cell, WONGO, 0.5)
        if (r + 1) % 5 == 0:
            p.text(hx + cols * cell + 5, ry + cell - 1, "%d" % ((r + 1) * cols), "M", 4.6, WONGO)


def t_music(p, x, y, w, h):
    ls, gap = 2.5 * MM, 9 * MM
    staff = 4 * ls
    pitch = staff + gap
    n = int((h + gap) / pitch)
    oy = y + (h - (n * pitch - gap)) / 2
    for i in range(n):
        yy = oy + i * pitch
        for k in range(5):
            p.hl(x, x + w, yy + k * ls, STAFF, 0.55)
        p.vl(x, yy, yy + staff, STAFF, 0.9)
        p.vl(x + w, yy, yy + staff, STAFF, 0.9)


def cell_rows(p, x, y, w, h, cols, color, strip=0, model=False, halves=1, gap_ratio=0.25):
    """십자 가이드가 있는 칸 노트 (한자·받아쓰기)"""
    gutter = 16 if halves == 2 else 0
    cell = (w - gutter) / (cols * halves)
    gap = cell * gap_ratio
    pitch = strip + cell + gap
    rows = int((h + gap) / pitch)
    oy = y + (h - (rows * pitch - gap)) / 2
    for half in range(halves):
        hx = x + half * (cols * cell + gutter)
        hw = cols * cell
        for r in range(rows):
            ry = oy + r * pitch
            if strip:
                p.hl(hx, hx + hw, ry, color, 0.5)
                p.text(hx + 3, ry + strip - 4, "훈 · 음", "M", 5.4, FAINT)
                p.vl(hx + 2 * cell, ry, ry + strip, LINE_L, 0.5)
            cy = ry + strip
            if model:
                p.box(hx, cy, cell, cell, fill=HexColor("#FBF0EE"))
            p.hl(hx, hx + hw, cy + cell / 2, LINE_L, 0.5, dash=[1.5, 2])
            for k in range(cols):
                p.vl(hx + k * cell + cell / 2, cy, cy + cell, LINE_L, 0.5, dash=[1.5, 2])
            p.box(hx, cy, hw, cell, stroke=color, lw=0.8)
            for k in range(1, cols):
                p.vl(hx + k * cell, cy, cy + cell, color, 0.55)


def t_hanja(p, x, y, w, h):
    cell_rows(p, x, y, w, h, cols=int(w / (15 * MM)), color=HANJA, strip=5 * MM, model=True, gap_ratio=0.12)


def t_kids(p, x, y, w, h):
    cell_rows(p, x, y, w, h, cols=10, color=GREEN, halves=1)


def t_math(p, x, y, w, h):
    gap = 12
    cw = (w - gap) / 2
    for c in range(2):
        cx = x + c * (cw + gap)
        p.grid(cx, y, cw, h, 5 * MM, GRID_L, 0.4)
    p.vl(x + w / 2, y, y + h, GRID_D, 0.7)


def t_iso(p, x, y, w, h):
    s = 5 * MM
    rp = s * math.sqrt(3) / 2
    ny = int(h / rp)
    nx = int(w / s) - 1
    ox = x + (w - nx * s - s / 2) / 2
    oy = y + (h - ny * rp) / 2
    for j in range(ny + 1):
        off = s / 2 if j % 2 else 0
        p.dotrow(ox + off, ox + off + nx * s, oy + j * rp, s, DOT, 1.2)


def t_calli(p, x, y, w, h):
    asc, xh, desc, gap = 5 * MM, 7 * MM, 5 * MM, 5 * MM
    gh = asc + xh + desc
    pitch = gh + gap
    n = int((h + gap) / pitch)
    oy = y + (h - (n * pitch - gap)) / 2
    dx = gh / math.tan(math.radians(55))
    sp = 8 * MM
    for i in range(n):
        yy = oy + i * pitch
        p.clip(x, yy, w, gh)
        k = -int(dx / sp) - 1
        while x + k * sp <= x + w:
            bx = x + k * sp
            p.c.setStrokeColor(GRID_L)
            p.c.setLineWidth(0.45)
            p.c.line(bx, H - (yy + gh), bx + dx, H - yy)
            k += 1
        p.unclip()
        p.hl(x, x + w, yy, LINE_L, 0.5)
        p.hl(x, x + w, yy + asc, LINE, 0.5, dash=[2, 2])
        p.hl(x, x + w, yy + asc + xh, RED, 0.6)
        p.hl(x, x + w, yy + gh, LINE_L, 0.5)


def t_story(p, x, y, w, h, cols=2, rows=4):
    gx, gy = 16, 16
    cw = (w - (cols - 1) * gx) / cols
    fh = cw * 9 / 16
    lh = 16
    ch = 12 + fh + 2 * lh
    oy = y + (h - (rows * ch + (rows - 1) * gy)) / 2
    for i in range(cols * rows):
        cx = x + (i % cols) * (cw + gx)
        cy = oy + (i // cols) * (ch + gy)
        p.text(cx, cy + 8, "#%d" % (i + 1), "B", 7, UI)
        p.text(cx + cw, cy + 8, "장면 · 시간", "M", 5.6, FAINT, "r")
        p.box(cx, cy + 12, cw, fh, stroke=GRID_D, r=3, lw=0.8)
        for k in range(2):
            p.hl(cx, cx + cw, cy + 12 + fh + (k + 1) * lh, LINE, 0.5)


def t_diary(p, x, y, w, h):
    y0 = TOP + 2
    cx = x
    for lab, lw in [("년", 40), ("월", 26), ("일", 26), ("요일", 26)]:
        p.hl(cx, cx + lw, y0 + 14, RULE, 0.6)
        p.text(cx + lw + 3, y0 + 12, lab, "M", 7.6, SUB)
        cx += lw + 3 + p.tw(lab, "M", 7.6) + 8
    wx = X1 - 226
    p.text(wx, y0 + 12, "날씨", "SB", 7.6, SUB)
    for i, lab in enumerate(["맑음", "구름", "흐림", "비", "눈"]):
        bx = wx + 26 + i * 40
        p.circle(bx + 4, y0 + 9.5, 4, stroke=RULE, lw=0.7)
        p.text(bx + 11, y0 + 12, lab, "R", 7.2, SUB)
    ty = y0 + 34
    p.text(x, ty, "제목", "SB", 7.6, SUB)
    p.hl(x + 26, X1, ty + 2, RULE, 0.6)
    top = ty + 12
    dh = (BB - top) * 0.46
    p.box(x, top, w, dh, stroke=GRID_D, r=6, lw=0.8)
    p.text(x + 8, top + 12, "그림", "M", 6.2, FAINT)
    gy = top + dh + 12
    cols = 12
    cell = w / cols
    rows = int((BB - gy) / cell)
    for r in range(rows + 1):
        p.hl(x, x + w, gy + r * cell, GREEN if r in (0, rows) else HexColor("#A9CDB8"), 0.55)
    for k in range(cols + 1):
        p.vl(x + k * cell, gy, gy + rows * cell, GREEN if k in (0, cols) else HexColor("#A9CDB8"), 0.55)


def t_wrong(p, x, y, w, h):
    y0 = TOP + 2
    rows = [[("과목", 0.34), ("단원", 0.66)], [("출처 · 쪽", 0.62), ("날짜", 0.38)]]
    for r, row in enumerate(rows):
        cx = x
        ry = y0 + r * 20
        for lab, frac in row:
            fw = w * frac
            p.text(cx, ry + 12, lab, "SB", 7.2, SUB)
            p.hl(cx + p.tw(lab, "SB", 7.2) + 6, cx + fw - 14, ry + 14, RULE, 0.6)
            cx += fw
    top = y0 + 46
    ph = (BB - top) * 0.3
    p.box(x, top, w, ph, stroke=GRID_D, r=5, lw=0.8)
    p.box(x, top, 40, 15, fill=UI, r=4)
    p.ctext(x, top, 40, 15, "문제", "B", 7.2, WHITE)
    top += ph + 10
    lw = w * 0.58
    p.text(x, top + 10, "풀이", "B", 7.6, UI)
    p.hl(x, x + lw, top + 13.5, UI, 0.9)
    p.grid(x, top + 16, lw, BB - top - 16, 5 * MM, GRID_L, 0.4)
    rx = x + lw + 14
    rw = x + w - rx
    yy = top
    p.section(rx, yy, rw, "틀린 이유", "WHY")
    yy += 18
    reasons = ["계산 실수", "개념 부족", "문제 이해", "시간 부족", "실수 · 착각", "기타"]
    for i, r in enumerate(reasons):
        bx = rx + (i % 2) * (rw / 2)
        by = yy + (i // 2) * 17
        p.check(bx, by + 2, 9)
        p.text(bx + 14, by + 10, r, "R", 7.2, INK)
    yy += 3 * 17 + 8
    p.section(rx, yy, rw, "핵심 개념 · 공식", "KEY")
    lines(p, rx, yy + 14, rw, 7 * 20, 20, LINE)
    yy += 14 + 7 * 20 + 10
    p.section(rx, yy, rw, "다시 풀기", "REVIEW")
    yy += 20
    bw = (rw - 2 * 6) / 3
    for i in range(3):
        bx = rx + i * (bw + 6)
        p.box(bx, yy, bw, 34, fill=TINT, r=5)
        p.check(bx + 6, yy + 7, 9)
        p.text(bx + 19, yy + 15, "%d차" % (i + 1), "SB", 7.2, UI)
        p.text(bx + 6, yy + 28, "   /   ", "M", 7, FAINT)
    yy += 34 + 12
    p.section(rx, yy, rw, "메모", "NOTE")
    t_dot(p, rx, yy + 16, rw, BB - yy - 16)


def t_meeting(p, x, y, w, h):
    y0 = TOP + 2
    rows = [[("회의명", 0.62), ("일시", 0.38)], [("장소", 0.38), ("참석자", 0.62)], [("안건", 1.0)]]
    for r, row in enumerate(rows):
        cx = x
        ry = y0 + r * 20
        for lab, frac in row:
            fw = w * frac
            p.text(cx, ry + 13, lab, "SB", 7.4, SUB)
            p.hl(cx + p.tw(lab, "SB", 7.4) + 6, cx + fw - 14, ry + 15, RULE, 0.6)
            cx += fw
    top = y0 + 66
    avail = BB - top
    dh = avail * 0.46
    p.section(x, top, w, "논의 내용", "DISCUSSION")
    lines(p, x, top + 14, w, dh - 14, 8 * MM)
    top += dh + 8
    eh = avail * 0.2
    p.section(x, top, w, "결정 사항", "DECISIONS")
    lines(p, x, top + 14, w, eh - 14, 8 * MM)
    top += eh + 8
    p.section(x, top, w, "할 일", "ACTION ITEMS")
    ty = top + 16
    p.box(x, ty, w, 13, fill=TINT)
    p.text(x + 9, ty + 9, "✓", "M", 6.6, SUB, "c")
    p.text(x + 24, ty + 9, "할 일", "M", 6.2, SUB)
    p.text(x + w - 96, ty + 9, "담당", "M", 6.2, SUB, "c")
    p.text(x + w - 32, ty + 9, "기한", "M", 6.2, SUB, "c")
    ty += 13
    rh = 8 * MM
    n = int((BB - ty) / rh)
    for r in range(n):
        p.check(x + 4.5, ty + r * rh + (rh - 9) / 2, 9)
        p.hl(x, x + w, ty + (r + 1) * rh, LINE, 0.5)
    for vx in (x + 18, x + w - 128, x + w - 64):
        p.vl(vx, ty - 13, ty + n * rh, LINE_L, 0.5)


def t_todo(p, x, y, w, h, ncol=2):
    gap = 14
    cw = (w - (ncol - 1) * gap) / ncol
    rh = 8 * MM
    for c in range(ncol):
        cx = x + c * (cw + gap)
        p.box(cx, y + 2, cw, 18, fill=TINT, r=4)
        p.text(cx + 6, y + 14, "목록", "SB", 6.6, FAINT)
        n = int((h - 24) / rh)
        for r in range(n):
            ry = y + 24 + r * rh
            p.check(cx + 2, ry + (rh - 10) / 2, 10)
            p.hl(cx + 17, cx + cw, ry + rh, LINE, 0.5)


# (키, 이름, 설명, 그리기, 머리줄 사용)
TEMPLATES = [
    ("lined8", "줄노트 8mm", "기본 유선", t_lined(8 * MM), True),
    ("lined6", "줄노트 6mm", "좁은 줄", t_lined(6 * MM), True),
    ("margin", "대학노트", "7mm 줄 · 왼쪽 3cm 빨간 줄", t_margin, True),
    ("blank", "무지노트", "빈 페이지", t_blank, True),
    ("grid10", "큰 격자", "10mm 모눈", t_grid(10 * MM, GRID, 0.5), True),
    ("grid5", "작은 격자", "5mm 모눈", t_grid(5 * MM, GRID, 0.4), True),
    ("graph", "방안지", "1mm 공학용 모눈", t_graph, True),
    ("dot", "도트 노트", "5mm 점", t_dot, True),
    ("cornell", "코넬노트", "단서 · 필기 · 요약", t_cornell, True),
    ("twocol", "2단 줄노트", "좌우 2단 · 8mm", t_twocol, True),
    ("english", "영어노트", "4선 · 빨간 기준선", t_english, True),
    ("vocab", "단어장", "단어 · 뜻 2단", t_vocab, True),
    ("wongo", "원고지", "20자 × 25줄", t_wongo, True),
    ("music", "오선지", "악보 · 15단", t_music, True),
    ("hanja", "한자노트", "훈음 칸 · 십자 가이드", t_hanja, True),
    ("kids", "칸노트", "10칸 · 받아쓰기", t_kids, True),
    ("math", "수학노트", "2단 · 연한 5mm 모눈", t_math, True),
    ("wrong", "오답노트", "문제 · 풀이 · 틀린 이유", t_wrong, False),
    ("iso", "아이소메트릭", "삼각 도트 5mm", t_iso, True),
    ("calli", "레터링 연습지", "4선 · 55° 사선", t_calli, True),
    ("story", "스토리보드", "16:9 · 8컷", t_story, True),
    ("diary", "그림일기", "그림 + 글 칸", t_diary, False),
    ("meeting", "회의록", "안건 · 결정 · 할 일", t_meeting, False),
    ("todo", "체크리스트", "2단 할 일 목록", t_todo, True),
]
FORM_BBOX = (X0 - 2, H - BB - 2, X1 + 2, H - TOP + 6)


def template_form(p, t):
    key, name, sub, fn, head = t

    def draw():
        if head:
            note_header(p)
        fn(p, X0, BODY, CW, BB - BODY)

    p.form("tpl_" + key, draw, FORM_BBOX)
    return "tpl_" + key


# ─────────────────────────────────────────────────────────────────────────────
# 페이지
# ─────────────────────────────────────────────────────────────────────────────
def page_cover(p):
    p.start("cover", "표지")
    p.box(0, 0, W, H, fill=PAPER)
    sw = 11
    for i, (main, light) in enumerate(SEC):
        p.box(W - (NSEC - i) * sw, 0, sw, H, fill=main)
    lx, ly, lw, lh = 44, 210, 320, 300
    p.box(lx + 4, ly + 5, lw, lh, fill=HexColor("#E7E1D4"), r=10)
    p.box(lx, ly, lw, lh, fill=WHITE, stroke=HexColor("#D9D2C3"), r=10, lw=0.8)
    p.box(lx, ly, lw, 70, fill=UI, r=10)
    p.box(lx, ly + 56, lw, 14, fill=UI)
    p.text(lx + 24, ly + 40, "NOTEBOOK", "B", 22, WHITE, cs=3.5)
    p.text(lx + 24, ly + 58, "간지 %d · 템플릿 %d" % (NSEC, len(TEMPLATES)), "M", 8, HexColor("#B8C5D3"))
    for i, lab in enumerate(["과목 · 주제", "이름", "기간"]):
        yy = ly + 118 + i * 58
        p.text(lx + 24, yy, lab, "SB", 8.6, UI)
        p.hl(lx + 24, lx + lw - 24, yy + 30, RULE, 0.8)
    bx = lx
    for lab, key, w in [("목차", "index", 100), ("노트 템플릿", "tmpl", 130)]:
        p.pill(bx, ly + lh + 44, w, 34, lab, key, fill=UI, color=WHITE, size=9.4, r=8)
        bx += w + 12
    p.text(lx, ly + lh + 110, "링크는 삼성노트 ‘읽기 모드’(책 아이콘)를 켠 상태에서", "R", 7.8, SUB)
    p.text(lx, ly + lh + 124, "눌러야 이동합니다.", "R", 7.8, SUB)
    p.end()


def page_index(p):
    p.start("index", "목차")
    p.chrome(menu="index")
    p.title("목차", "INDEX", "간지를 누르면 해당 섹션으로 이동")
    gx, gy = 12, 8
    cw = (CW - gx) / 2
    ch = 62
    for i, (main, light) in enumerate(SEC):
        cx = X0 + (i // 5) * (cw + gx)
        cy = BT + (i % 5) * (ch + gy)
        p.box(cx, cy, cw, ch, fill=WHITE, stroke=RULE, r=7, lw=0.6)
        p.box(cx, cy, 44, ch, fill=main, r=7)
        p.box(cx + 34, cy, 10, ch, fill=main)
        p.ctext(cx, cy, 44, ch, "%02d" % (i + 1), "B", 14, WHITE)
        p.text(cx + 54, cy + 15, "제목", "SB", 6.2, FAINT)
        p.hl(cx + 54, cx + cw - 24, cy + 36, RULE, 0.7)
        p.text(cx + 54, cy + ch - 9, "기간", "SB", 6.2, FAINT)
        p.hl(cx + 74, cx + cw * 0.7, cy + ch - 7, RULE_L, 0.6)
        p.chev(cx + cw - 12, cy + ch / 2, 3.2, True, main, 1.3)
        p.link(cx, cy, cw, ch, "s%d" % (i + 1))
    y = BT + 5 * (ch + gy) + 6
    p.pill(X0, y, CW, 38, "노트 템플릿 %d종" % len(TEMPLATES), "tmpl", fill=UI, color=WHITE, size=10.5, r=8)
    p.chev(X1 - 20, y + 19, 3.8, True, WHITE, 1.4)
    y += 50
    p.section(X0, y, CW, "바로가기", "TEMPLATES")
    y += 18
    ncol, nrow = 3, 8
    bw = (CW - (ncol - 1) * 6) / ncol
    rh = 20
    for i, (key, name, sub, fn, head) in enumerate(TEMPLATES):
        bx = X0 + (i // nrow) * (bw + 6)
        by = y + (i % nrow) * rh
        p.box(bx, by + 1, bw, rh - 3, fill=TINT, r=4)
        p.text(bx + 7, by + 13, "%02d" % (i + 1), "B", 6.6, FAINT)
        p.text(bx + 22, by + 13, name, "M", 7.4, INK)
        p.link(bx, by, bw, rh, "t_" + key)
    y += nrow * rh + 10
    p.box(X0, y, CW, BB - y, fill=UI_L, r=7)
    tips = ["템플릿에서 원하는 노트를 고릅니다.",
            "삼성노트에서 그 페이지를 복사해 간지 뒤에 붙입니다.",
            "간지의 섹션 목차 칸에 내용과 날짜를 적어 둡니다.",
            "오른쪽 색 탭을 누르면 간지로, 위쪽 ‘목차’·‘템플릿’은 해당 페이지로 갑니다.",
            "링크는 읽기 모드(책 아이콘)에서 눌러야 이동합니다."]
    p.text(X0 + 12, y + 18, "사용법", "B", 8.4, UI)
    for i, t in enumerate(tips):
        p.text(X0 + 12, y + 36 + i * 15, "%d. %s" % (i + 1, t), "R", 7.4, SUB)
    p.end()


def page_divider(p, i):
    main, light = SEC[i - 1]
    p.start("s%d" % i, "간지 %02d" % i, 1)
    p.chrome(label="간지 %02d" % i, active=i)
    ph = 300
    p.box(X0, TOP, CW, ph, fill=main, r=10)
    p.text(X0 + 24, TOP + 36, "SECTION", "SB", 10, WHITE, cs=5)
    p.text(X0 + 16, TOP + 210, "%02d" % i, "B", 130, WHITE)
    rx = X0 + 220
    p.text(rx, TOP + 120, "제목", "SB", 7.6, light)
    p.hl(rx, X1 - 24, TOP + 160, WHITE, 1.0)
    p.hl(rx, X1 - 24, TOP + 196, light, 0.8)
    p.text(rx, TOP + 228, "기간", "SB", 7.6, light)
    p.hl(rx + 28, X1 - 24, TOP + 230, light, 0.8)
    p.text(X0 + 24, TOP + ph - 18, "이 간지 뒤에 노트 페이지를 붙여 쓰세요", "R", 7.2, light)

    y = TOP + ph + 14
    p.text(X0, y + 20, "섹션 목차", "B", 16, INK)
    p.text(X0 + p.tw("섹션 목차", "B", 16) + 10, y + 19, "CONTENTS", "SB", 7.2, main, cs=1.4)
    bx = X1
    for lab, key, w, right in [("다음 간지", "s%d" % (i + 1) if i < NSEC else None, 84, True),
                               ("이전 간지", "s%d" % (i - 1) if i > 1 else None, 84, False)]:
        bx -= w
        if key:
            p.pill(bx, y + 4, w, 24, ("   " if not right else "") + lab + ("   " if right else ""), key,
                   fill=light, color=main, size=8, r=6)
            p.chev(bx + (w - 12 if right else 12), y + 16, 2.8, right, main, 1.2)
        bx -= 8
    ty = y + 38
    p.hl(X0, X1, ty - 4, INK, 1.1)
    hh = 16
    p.box(X0, ty, CW, hh, fill=light)
    p.ctext(X0, ty, 30, hh, "No.", "SB", 6.8, main)
    p.text(X0 + 40, ty + 11, "내용", "SB", 6.8, main)
    p.ctext(X1 - 64, ty, 64, hh, "날짜", "SB", 6.8, main)
    ty += hh
    nav_y = BB - 26
    rh = 22
    n = int((nav_y - 8 - ty) / rh)
    for r in range(n):
        p.text(X0 + 15, ty + r * rh + 14.5, "%d" % (r + 1), "M", 7, FAINT, "c")
        p.hl(X0, X1, ty + (r + 1) * rh, RULE_L, 0.6)
    p.vl(X0 + 30, ty - hh, ty + n * rh, RULE_L, 0.6)
    p.vl(X1 - 64, ty - hh, ty + n * rh, RULE_L, 0.6)
    p.pill(X0, nav_y, 70, 24, "목차", "index", fill=TINT, stroke=RULE, color=UI, size=8, r=6)
    p.pill(X0 + 78, nav_y, 100, 24, "노트 템플릿", "tmpl", fill=TINT, stroke=RULE, color=UI, size=8, r=6)
    p.end()


def page_tmpl_list(p):
    p.start("tmpl", "노트 템플릿")
    p.chrome(menu="tmpl")
    p.title("노트 템플릿", "NOTE TEMPLATES", "누르면 해당 노트로 이동")
    cols, gx, gy = 6, 8, 16
    cw = (CW - (cols - 1) * gx) / cols
    fx0, fy0, fx1, fy1 = FORM_BBOX
    s = cw / (fx1 - fx0)
    ph = (fy1 - fy0) * s
    ch = ph + 24
    for i, t in enumerate(TEMPLATES):
        key, name, sub, fn, head = t
        cx = X0 + (i % cols) * (cw + gx)
        cy = BT + (i // cols) * (ch + gy)
        p.box(cx, cy, cw, ph, fill=WHITE, stroke=RULE, r=3, lw=0.6)
        c = p.c
        c.saveState()
        c.translate(cx - fx0 * s, H - (cy + ph) - fy0 * s)
        c.scale(s, s)
        c.doForm(template_form(p, t))
        c.restoreState()
        p.text(cx + 1, cy + ph + 11, "%02d" % (i + 1), "B", 6.6, FAINT)
        p.text(cx + 13, cy + ph + 11, name, "SB", 7.4, INK)
        p.text(cx + 1, cy + ph + 20.5, sub, "R", 5.6, SUB)
        p.link(cx, cy, cw, ch, "t_" + key)
    y = BT + 4 * (ch + gy) + 4
    p.box(X0, y, CW, BB - y, fill=UI_L, r=7)
    p.text(X0 + 12, y + 18, "쓰는 법", "B", 8.4, UI)
    for i, t in enumerate(["원하는 템플릿을 눌러 이동한 뒤, 삼성노트에서 그 페이지를 복사합니다.",
                           "쓰려는 간지 뒤에 붙여 넣으면 됩니다. 원본 템플릿은 그대로 두세요.",
                           "줄 간격은 화면 폭에 맞춰 볼 때 실제 mm 크기입니다."]):
        p.text(X0 + 12, y + 36 + i * 15, "%d. %s" % (i + 1, t), "R", 7.4, SUB)
    p.end()


def page_template(p, i, t):
    key, name, sub, fn, head = t
    p.start("t_" + key, "%02d %s" % (i + 1, name), 1)
    p.chrome(label="%s · %s" % (name, sub))
    p.c.doForm(template_form(p, t))
    p.end()


# ─────────────────────────────────────────────────────────────────────────────
def build(out, font_dir):
    tmp = out + ".tmp"
    p = Painter(tmp, font_dir)
    page_cover(p)
    page_index(p)
    p.outline.append((0, "간지", "s1"))
    for i in range(1, NSEC + 1):
        page_divider(p, i)
    page_tmpl_list(p)
    for i, t in enumerate(TEMPLATES):
        page_template(p, i, t)
    p.c.save()

    pdf = pikepdf.open(tmp)
    pages = pdf.pages
    per_page = {}
    missing = sorted({k for _, _, k in p.links if k not in p.dests})
    if missing:
        raise SystemExit("링크 대상 없음: %s" % missing)
    for pg, rect, key in p.links:
        per_page.setdefault(pg, []).append((rect, p.dests[key]))
    for pg, items in per_page.items():
        pages[pg].obj.Annots = pikepdf.Array([pdf.make_indirect(pikepdf.Dictionary(
            Type=pikepdf.Name.Annot, Subtype=pikepdf.Name.Link,
            Rect=[round(v, 2) for v in rect], Border=[0, 0, 0],
            Dest=pikepdf.Array([pages[target].obj, pikepdf.Name.XYZ, 0, H, 0]),
        )) for rect, target in items])
    with pdf.open_outline() as ol:
        stack = []
        for level, title, key in p.outline:
            item = pikepdf.OutlineItem(title, p.dests[key])
            del stack[level:]
            (stack[-1].children if level and stack else ol.root).append(item)
            stack.append(item)
    pdf.Root.PageLayout = pikepdf.Name.SinglePage
    pdf.save(out, compress_streams=True)
    pdf.close()
    os.remove(tmp)
    print("%s: %d쪽, 링크 %d개, %.1f MB" % (out, p.page, len(p.links), os.path.getsize(out) / 1e6))


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "notebook-templates.pdf")
    build(out, os.environ.get("PLANNER_FONT_DIR", os.path.join(here, "fonts")))
