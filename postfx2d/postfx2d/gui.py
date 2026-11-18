"""PostFX 2D - GUI (Tkinter)."""
from __future__ import annotations

import dataclasses
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cv2
import numpy as np
from PIL import Image, ImageGrab, ImageTk

from . import core
from . import enhance
from . import fieldmap as fmap

APP = "PostFX 2D"
VERSION = "0.1.0"

BG = "#14171c"
PANEL = "#1d2128"
CARD = "#252a33"
FG = "#e6e9ef"
MUTED = "#8b93a1"
ACCENT = "#ffb020"
ACCENT2 = "#e8364f"
LINE = "#323845"

CLASS_COLORS = {"W1": "#ff8c00", "W2": "#28aa3c", "W3": "#1e5ae6",
                "W4": "#963cc8", "GLOBAL": "#e61414"}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP}  ·  2D 해석결과 후처리")
        self.geometry("1360x820")
        self.minsize(1000, 640)
        self.configure(bg=BG)
        self.src: np.ndarray | None = None
        self.result: core.Result | None = None
        self.src_name = "capture"
        self.show_original = False
        self._tk_img = None
        self.readings: list[fmap.Reading] = []
        self._fm: fmap.FieldMap | None = None
        self._legend_lut = None
        self._legend_pick = False
        self._drag = None
        self._view = (1.0, 0.0, 0.0)          # scale, offset x, offset y
        self._zoom = 1.0                       # 1.0 = 화면 맞춤
        self._pan = [0.0, 0.0]
        self._pan_start = None                 # (ex, ey, panx, pany)
        self._panned = False
        self._style()
        self._build()
        self.bind_all("<Control-v>", lambda e: self.paste())
        self.bind_all("<Control-o>", lambda e: self.open_file())
        self.bind_all("<Control-s>", lambda e: self.save())
        self.bind_all("<space>", lambda e: self.toggle_view())
        self.bind_all("<Control-z>", lambda e: self.undo_reading())
        for k in ("<Control-plus>", "<Control-equal>", "<Control-KP_Add>"):
            self.bind_all(k, lambda e: self.zoom_by(1.25))
        for k in ("<Control-minus>", "<Control-KP_Subtract>"):
            self.bind_all(k, lambda e: self.zoom_by(0.8))
        self.bind_all("<Control-0>", lambda e: self.zoom_fit())
        self.set_status("Ctrl+V 로 클립보드 캡처를 붙여넣거나, 이미지 파일을 여세요.")

    # ------------------------------------------------------------ style
    def _style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        font = ("Segoe UI", 10)
        st.configure(".", background=PANEL, foreground=FG, font=font,
                     fieldbackground=CARD, bordercolor=LINE)
        st.configure("TFrame", background=PANEL)
        st.configure("Bg.TFrame", background=BG)
        st.configure("Card.TFrame", background=CARD)
        st.configure("TLabel", background=PANEL, foreground=FG)
        st.configure("Card.TLabel", background=CARD, foreground=FG)
        st.configure("Muted.TLabel", background=PANEL, foreground=MUTED, font=("Segoe UI", 9))
        st.configure("H.TLabel", background=PANEL, foreground=FG, font=("Segoe UI Semibold", 11))
        st.configure("Title.TLabel", background=BG, foreground=FG, font=("Segoe UI Semibold", 15))
        st.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))
        st.configure("TButton", background=CARD, foreground=FG, borderwidth=0,
                     padding=(12, 7), focusthickness=0)
        st.map("TButton", background=[("active", "#323946"), ("disabled", PANEL)],
               foreground=[("disabled", MUTED)])
        st.configure("Accent.TButton", background=ACCENT, foreground="#111",
                     font=("Segoe UI Semibold", 10))
        st.map("Accent.TButton", background=[("active", "#ffc24d"), ("disabled", "#5a4a22")])
        st.configure("TCombobox", fieldbackground=CARD, background=CARD, foreground=FG,
                     arrowcolor=FG, borderwidth=0)
        st.map("TCombobox", fieldbackground=[("readonly", CARD)], foreground=[("readonly", FG)])
        st.configure("TSpinbox", fieldbackground=CARD, foreground=FG, arrowcolor=FG,
                     background=CARD, borderwidth=0)
        st.configure("Horizontal.TScale", background=PANEL, troughcolor=CARD)
        st.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=FG,
                     rowheight=26, borderwidth=0)
        st.configure("Treeview.Heading", background=PANEL, foreground=MUTED,
                     font=("Segoe UI Semibold", 9), borderwidth=0)
        st.map("Treeview", background=[("selected", "#3a4252")])
        st.configure("TNotebook", background=PANEL, borderwidth=0)
        st.configure("TNotebook.Tab", background=PANEL, foreground=MUTED, padding=(14, 6))
        st.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", FG)])
        st.configure("TEntry", fieldbackground=CARD, foreground=FG, insertcolor=FG, borderwidth=0)
        st.configure("TCheckbutton", background=PANEL, foreground=FG, indicatorcolor=CARD)
        st.map("TCheckbutton", background=[("active", PANEL)],
               indicatorcolor=[("selected", ACCENT)])
        st.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor=CARD, borderwidth=0)

    # ------------------------------------------------------------ layout
    def _build(self):
        top = ttk.Frame(self, style="Bg.TFrame", padding=(18, 12, 18, 8))
        top.pack(fill="x")
        dot = tk.Canvas(top, width=22, height=22, bg=BG, highlightthickness=0)
        dot.create_oval(2, 2, 20, 20, fill=ACCENT, outline="")
        dot.create_oval(8, 8, 14, 14, fill=ACCENT2, outline="")
        dot.pack(side="left", padx=(0, 10))
        ttk.Label(top, text=APP, style="Title.TLabel").pack(side="left")
        ttk.Label(top, text=f"  v{VERSION}  ·  2D 해석결과 이미지 후처리",
                  style="Sub.TLabel").pack(side="left", pady=(5, 0))

        body = ttk.Frame(self, style="Bg.TFrame", padding=(18, 0, 18, 0))
        body.pack(fill="both", expand=True)

        # --- 왼쪽 사이드바
        side = ttk.Frame(body, padding=16, width=290)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)

        ttk.Label(side, text="입력", style="H.TLabel").pack(anchor="w")
        ttk.Button(side, text="📋  클립보드 붙여넣기  (Ctrl+V)",
                   command=self.paste).pack(fill="x", pady=(8, 4))
        ttk.Button(side, text="📂  이미지 파일 열기  (Ctrl+O)",
                   command=self.open_file).pack(fill="x")

        ttk.Label(side, text="모드", style="H.TLabel").pack(anchor="w", pady=(20, 6))
        self.mode = ttk.Combobox(side, state="readonly",
                                 values=["1. 최소안전율부위", "2. 전계 값 읽기 (Rainbow map)"])
        self.mode.current(0)
        self.mode.pack(fill="x")
        self.mode.bind("<<ComboboxSelected>>", lambda e: self._on_mode())

        ttk.Label(side, text="옵션", style="H.TLabel").pack(anchor="w", pady=(20, 6))
        self.opt_host = ttk.Frame(side)
        self.opt_host.pack(fill="x")
        self.opt1 = ttk.Frame(self.opt_host)
        self.opt2 = ttk.Frame(self.opt_host)
        self.opt1.pack(fill="x")
        self._build_opt2(self.opt2)
        side_main = side
        side = self.opt1
        self.v_count = tk.IntVar(value=0)
        self._spin(side, "그룹 수 (0 = 자동)", self.v_count, 0, 50)
        self.v_galpha = tk.DoubleVar(value=0.18)
        self._slider(side, "그룹 음영 농도", self.v_galpha, 0.05, 0.5)
        self.v_halpha = tk.DoubleVar(value=0.45)
        self._slider(side, "하이라이트 농도", self.v_halpha, 0.2, 0.8)
        self.v_pad = tk.IntVar(value=14)
        self._spin(side, "음영 여백 (px)", self.v_pad, 2, 60)

        self.btn_run = ttk.Button(side, text="▶  분석 실행", style="Accent.TButton",
                                  command=self.run)
        self.btn_run.pack(fill="x", pady=(22, 6))
        self.btn_redraw = ttk.Button(side, text="↻  표 수정값으로 다시 그리기",
                                     command=self.redraw)
        self.btn_redraw.pack(fill="x")
        self.progress = ttk.Progressbar(side, mode="indeterminate")
        self.progress.pack(fill="x", pady=(10, 0))
        side = side_main

        ttk.Label(side, text="출력", style="H.TLabel").pack(anchor="w", pady=(20, 6))
        ttk.Button(side, text="💾  PNG + CSV 저장  (Ctrl+S)",
                   command=self.save).pack(fill="x", pady=(0, 4))
        ttk.Button(side, text="📄  결과 이미지 클립보드 복사",
                   command=self.copy_result).pack(fill="x")

        legend = ttk.Frame(side)
        legend.pack(side="bottom", fill="x")
        self.legend1 = legend
        ttk.Label(legend, text="점 색상 범례", style="Muted.TLabel").pack(anchor="w")
        for k, c in CLASS_COLORS.items():
            row = ttk.Frame(legend); row.pack(anchor="w", pady=1)
            cv = tk.Canvas(row, width=12, height=12, bg=PANEL, highlightthickness=0)
            cv.create_oval(1, 1, 11, 11, fill=c, outline="")
            cv.pack(side="left", padx=(0, 6))
            ttk.Label(row, text=core.POINT_CLASSES[k]["label"],
                      style="Muted.TLabel").pack(side="left")

        # --- 가운데: 뷰어 + 테이블
        main = ttk.Frame(body, style="Bg.TFrame")
        main.pack(side="left", fill="both", expand=True, padx=(14, 0))

        viewer = ttk.Frame(main, padding=1)
        viewer.pack(fill="both", expand=True)
        bar = ttk.Frame(viewer, padding=(12, 8))
        bar.pack(fill="x")
        self.lbl_view = ttk.Label(bar, text="미리보기", style="H.TLabel")
        self.lbl_view.pack(side="left")
        zbar = ttk.Frame(bar)
        zbar.pack(side="left", padx=(18, 0))
        ttk.Button(zbar, text="－", width=3, command=lambda: self.zoom_by(0.8)).pack(side="left")
        self.lbl_zoom = ttk.Label(zbar, text="맞춤", width=6, anchor="center",
                                  style="Muted.TLabel")
        self.lbl_zoom.pack(side="left", padx=2)
        ttk.Button(zbar, text="＋", width=3, command=lambda: self.zoom_by(1.25)).pack(side="left")
        ttk.Button(zbar, text="맞춤 (Ctrl+0)", command=self.zoom_fit).pack(side="left", padx=(4, 0))
        ttk.Label(zbar, text="화질", style="Muted.TLabel").pack(side="left", padx=(14, 4))
        self.quality = ttk.Combobox(zbar, state="readonly", width=16, values=enhance.OPTIONS)
        self.quality.current(0)
        self.quality.pack(side="left")
        self.quality.bind("<<ComboboxSelected>>", lambda e: self._on_quality())
        self._enh_cache = (None, None, None, 1)     # (img id, option, enhanced, k)
        self.btn_toggle = ttk.Button(bar, text="원본 / 결과 전환 (Space)",
                                     command=self.toggle_view)
        self.btn_toggle.pack(side="right")
        self.canvas = tk.Canvas(viewer, bg="#0f1115", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self._show())
        self.canvas.bind("<Button-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Button-3>", self._on_right)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<MouseWheel>", self._on_wheel)                 # Windows / macOS
        self.canvas.bind("<Button-4>", lambda e: self._on_wheel(e, 120))  # Linux
        self.canvas.bind("<Button-5>", lambda e: self._on_wheel(e, -120))
        self.canvas.bind("<Button-2>", self._pan_begin)                   # 휠 클릭 드래그 = 이동
        self.canvas.bind("<B2-Motion>", self._pan_move)
        self.canvas.bind("<ButtonRelease-2>", lambda e: self._pan_end())

        tables = self.tables = ttk.Notebook(main, height=210)
        tables.pack(fill="x", pady=(12, 0))
        self.tree_sum = self._tree(tables, ["그룹", "점 개수", "최소 안전율", "구분",
                                            "하이라이트", "X(px)", "Y(px)", "인덱스"])
        tables.add(self.tree_sum.master, text="  요약 (그룹별 최소)  ")
        self.tree_pts = self._tree(tables, ["#", "그룹", "구분", "안전율", "인덱스", "X(px)", "Y(px)"])
        tables.add(self.tree_pts.master, text="  전체 점  (더블클릭: 값 수정)  ")
        self.tree_pts.bind("<Double-1>", self._edit_value)
        self.tree_fld = self._tree(tables, ["#", "X(px)", "Y(px)", "전계", "컬러맵 위치(%)",
                                            "색 오차", "신뢰"])
        tables.add(self.tree_fld.master, text="  전계 측정점  (우클릭: 삭제)  ")
        tables.tab(2, state="hidden")

        self.status = ttk.Label(self, style="Sub.TLabel", padding=(20, 6))
        self.status.pack(fill="x")

    def _build_opt2(self, f):
        self.v_vmin = tk.StringVar(value="0")
        self.v_vmax = tk.StringVar(value="1")
        self.v_unit = tk.StringVar(value="V/m")
        for text, var in (("최소값 (컬러맵 아래 끝)", self.v_vmin),
                          ("최대값 (컬러맵 위 끝)", self.v_vmax), ("단위", self.v_unit)):
            row = ttk.Frame(f); row.pack(fill="x", pady=3)
            ttk.Label(row, text=text).pack(side="left")
            ttk.Entry(row, textvariable=var, width=10, justify="right").pack(side="right")
            var.trace_add("write", lambda *_: self._refresh_field())
        ttk.Label(f, text="컬러맵", style="Muted.TLabel").pack(anchor="w", pady=(8, 2))
        self.cmap = ttk.Combobox(f, state="readonly", values=fmap.COLORMAPS)
        self.cmap.current(0)
        self.cmap.pack(fill="x")
        self.cmap.bind("<<ComboboxSelected>>", lambda e: self._on_cmap())
        self.lut_preview = tk.Canvas(f, height=14, bg=PANEL, highlightthickness=0)
        self.lut_preview.pack(fill="x", pady=(6, 0))
        self.lut_preview.bind("<Configure>", lambda e: self._draw_lut_preview())
        self.btn_legend = ttk.Button(f, text="🎯  범례(컬러바) 지정하기",
                                     command=self.start_legend_pick)
        self.btn_legend.pack(fill="x", pady=(8, 0))
        self.v_log = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="로그 스케일", variable=self.v_log,
                        command=self._refresh_field).pack(anchor="w", pady=(8, 0))
        row = ttk.Frame(f); row.pack(fill="x", pady=(10, 0))
        ttk.Button(row, text="↶ 되돌리기", command=self.undo_reading).pack(
            side="left", fill="x", expand=True, padx=(0, 3))
        ttk.Button(row, text="✕ 모두 지우기", command=self.clear_readings).pack(
            side="left", fill="x", expand=True, padx=(3, 0))
        ttk.Label(f, style="Muted.TLabel", justify="left", wraplength=250, text=(
            "클릭: 측정 · 우클릭: 삭제 · 휠: 확대/축소 · 드래그: 이동\n"
            "정확도↑: 범례 지정 → 컬러바 최소 끝→최대 끝 드래그"
        )).pack(anchor="w", pady=(10, 0))

    # ------------------------------------------------------------ mode 2
    def is_field(self):
        return self.mode.current() == 1

    def _on_mode(self):
        if self.is_field():
            self.opt1.pack_forget(); self.legend1.pack_forget()
            self.opt2.pack(fill="x")
            self.btn_toggle.pack_forget()
            for i in (0, 1):
                self.tables.tab(i, state="hidden")
            self.tables.tab(2, state="normal"); self.tables.select(2)
            self.show_original = True
            self._refresh_field()
            self.set_status("최소/최대값을 입력하고 맵 위를 클릭하세요.")
        else:
            self.opt2.pack_forget()
            self.opt1.pack(fill="x"); self.legend1.pack(side="bottom", fill="x")
            self.btn_toggle.pack(side="right")
            for i in (0, 1):
                self.tables.tab(i, state="normal")
            self.tables.tab(2, state="hidden"); self.tables.select(0)
            if self.src is not None and self.result is None:
                self.set_status("모드 1: '분석 실행' 을 누르면 점/숫자 분석을 시작합니다.")
        self._show()

    def _on_cmap(self):
        if self.cmap.get() == "범례에서 추출" and self._legend_lut is None:
            self.start_legend_pick()
        self._refresh_field()

    def start_legend_pick(self):
        if self.src is None:
            self.set_status("먼저 이미지를 붙여넣거나 여세요.")
            return
        self._legend_pick = True
        self.canvas.configure(cursor="crosshair")
        self.set_status("컬러바의 최소값 끝에서 최대값 끝까지 드래그하세요. (Esc 취소)")
        self.bind("<Escape>", lambda e: self._end_legend_pick())

    def _end_legend_pick(self):
        self._legend_pick = False
        self._drag = None
        self.canvas.configure(cursor="")
        self.canvas.delete("drag")

    def _current_lut(self):
        name = self.cmap.get()
        if name == "범례에서 추출":
            return self._legend_lut
        return fmap.builtin_lut(name)

    def _refresh_field(self):
        """범위/컬러맵이 바뀌면 FieldMap 재생성 + 기존 측정점 다시 계산."""
        self._fm = None
        try:
            vmin = float(self.v_vmin.get().replace(",", ""))
            vmax = float(self.v_vmax.get().replace(",", ""))
            lut = self._current_lut()
            if lut is not None and vmin != vmax:
                self._fm = fmap.FieldMap(lut, vmin, vmax, bool(self.v_log.get()))
        except (ValueError, tk.TclError) as e:
            if "로그" in str(e):
                self.set_status(f"⚠ {e}")
        if self._fm and self.src is not None:
            self.readings = [self._fm.read(self.src, r.x, r.y) for r in self.readings]
        self._draw_lut_preview()
        self._fill_field_table()
        if self.is_field():
            self._show()

    def _draw_lut_preview(self):
        c = self.lut_preview
        c.delete("all")
        lut = self._current_lut()
        w = max(c.winfo_width(), 10)
        if lut is None:
            c.create_text(w / 2, 7, text="범례 미지정", fill=MUTED, font=("Segoe UI", 8))
            return
        for i in range(w):
            b, g, r = lut[int(i / (w - 1) * (len(lut) - 1))]
            c.create_line(i, 0, i, 14, fill=f"#{r:02x}{g:02x}{b:02x}")

    def _img_xy(self, ex, ey):
        sc, ox, oy = self._view
        x, y = int((ex - ox) / sc), int((ey - oy) / sc)
        if self.src is None:
            return None
        h, w = self.src.shape[:2]
        return (x, y) if 0 <= x < w and 0 <= y < h else None

    # ------------------------------------------------------------ zoom / pan
    def _base_scale(self):
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        h, w = self.src.shape[:2]
        return min((cw - 20) / w, (ch - 20) / h, 4.0)

    def zoom_by(self, f, ex=None, ey=None):
        if self.src is None:
            return
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        ex = cw / 2 if ex is None else ex
        ey = ch / 2 if ey is None else ey
        sc, ox, oy = self._view
        base = self._base_scale()
        new_zoom = min(max(self._zoom * f, 1.0), max(32.0 / base, 1.0))
        if abs(new_zoom - self._zoom) < 1e-9:
            return
        nsc = base * new_zoom
        h, w = self.src.shape[:2]
        # 커서 아래 이미지 지점이 그대로 머물도록 이동량 보정
        px, py = (ex - ox) / sc, (ey - oy) / sc
        self._pan = [ex - px * nsc - (cw / 2 - w * nsc / 2),
                     ey - py * nsc - (ch / 2 - h * nsc / 2)]
        self._zoom = new_zoom
        if self._zoom == 1.0:
            self._pan = [0.0, 0.0]
        self._show()

    def zoom_fit(self):
        self._zoom, self._pan = 1.0, [0.0, 0.0]
        self._show()

    def _on_wheel(self, e, delta=None):
        d = e.delta if delta is None else delta
        self.zoom_by(1.25 if d > 0 else 0.8, e.x, e.y)

    def _pan_begin(self, e):
        self._pan_start = (e.x, e.y, *self._pan)
        self._panned = False

    def _pan_move(self, e):
        if not self._pan_start:
            return
        sx, sy, px, py = self._pan_start
        if not self._panned and abs(e.x - sx) + abs(e.y - sy) < 4:
            return
        self._panned = True
        if self._zoom > 1.0:
            self._pan = [px + e.x - sx, py + e.y - sy]
            self.canvas.configure(cursor="fleur")
            self._show()

    def _pan_end(self):
        self._pan_start = None
        self.canvas.configure(cursor="crosshair" if self._legend_pick else "")

    # ------------------------------------------------------------ mouse
    def _on_press(self, e):
        if self.src is None:
            return
        if self._legend_pick:
            self._drag = (e.x, e.y)
            return
        self._pan_begin(e)

    def _measure(self, e):
        xy = self._img_xy(e.x, e.y)
        if not xy:
            return
        if not self._fm:
            self.set_status("⚠ 최소/최대값과 컬러맵(범례)을 먼저 지정하세요.")
            return
        r = self._fm.read(self.src, *xy)
        self.readings.append(r)
        self._fill_field_table()
        self._show()
        self.set_status(self._reading_text(r, len(self.readings)))

    def _on_drag(self, e):
        if not self._legend_pick:
            self._pan_move(e)
            return
        if self._drag:
            self.canvas.delete("drag")
            self.canvas.create_line(*self._drag, e.x, e.y, fill="#ffffff", width=3, tags="drag")
            self.canvas.create_line(*self._drag, e.x, e.y, fill=ACCENT, width=1,
                                    dash=(4, 2), tags="drag")

    def _on_release(self, e):
        if not self._legend_pick:
            moved = self._panned
            self._pan_end()
            if not moved and self.is_field():
                self._measure(e)
            return
        if not self._drag:
            return
        p0, p1 = self._img_xy(*self._drag), self._img_xy(e.x, e.y)
        self._end_legend_pick()
        if not p0 or not p1 or abs(p0[0] - p1[0]) + abs(p0[1] - p1[1]) < 10:
            self.set_status("⚠ 범례 지정 실패: 컬러바 위에서 더 길게 드래그하세요.")
            return
        self._legend_lut = fmap.lut_from_line(self.src, p0, p1)
        self.cmap.set("범례에서 추출")
        self._refresh_field()
        self.set_status(f"범례 지정 완료: {p0} → {p1}. 이제 맵을 클릭하세요.")

    def _on_right(self, e):
        if not self.is_field() or not self.readings:
            return
        sc, _, _ = self._view
        xy = self._img_xy(e.x, e.y)
        if not xy:
            return
        k = min(range(len(self.readings)),
                key=lambda i: (self.readings[i].x - xy[0]) ** 2 + (self.readings[i].y - xy[1]) ** 2)
        r = self.readings[k]
        if ((r.x - xy[0]) ** 2 + (r.y - xy[1]) ** 2) ** 0.5 * sc < 20:
            self.readings.pop(k)
            self._fill_field_table()
            self._show()

    def _on_motion(self, e):
        if not self.is_field() or self.src is None or self._legend_pick or not self._fm:
            return
        xy = self._img_xy(e.x, e.y)
        if xy:
            r = self._fm.read(self.src, *xy)
            self.set_status("커서  " + self._reading_text(r))

    def _reading_text(self, r, k=None):
        unit = self.v_unit.get().strip()
        head = f"#{k}  " if k else ""
        warn = "" if r.reliable else "   ⚠ 컬러맵에 없는 색 (맵 밖/경계선일 수 있음)"
        return (f"{head}({r.x}, {r.y})  →  {fmap.fmt(r.value)} {unit}"
                f"   [컬러맵 {r.t * 100:.1f}%, 색오차 {r.dist:.1f}]{warn}")

    def undo_reading(self):
        if self.is_field() and self.readings:
            self.readings.pop()
            self._fill_field_table()
            self._show()

    def clear_readings(self):
        self.readings = []
        self._fill_field_table()
        self._show()

    def _fill_field_table(self):
        tv = self.tree_fld
        tv.delete(*tv.get_children())
        unit = self.v_unit.get().strip()
        tv.heading("전계", text=f"전계 ({unit})" if unit else "전계")
        for k, r in enumerate(self.readings, 1):
            tv.insert("", "end", values=[k, r.x, r.y, fmap.fmt(r.value), f"{r.t * 100:.1f}",
                                         f"{r.dist:.1f}", "OK" if r.reliable else "불확실"],
                      tags=() if r.reliable else ("global",))

    def _on_quality(self):
        kind, k = enhance.parse(self.quality.get())
        self._show()
        if kind == "sharpen":
            self.set_status(f"화질: 선명 확대 ×{k} 로 표시합니다. 저장할 때도 ×{k} 크기로 저장됩니다. "
                            "(측정·분석 값은 항상 원본 픽셀 기준)")
        elif kind == "smooth":
            self.set_status("화질: 확대 시 부드럽게(bicubic) 표시합니다. 저장 이미지는 원본 크기.")
        else:
            self.set_status("화질: 원본 픽셀 그대로 표시합니다.")

    def _enhanced(self, img):
        """화질 옵션 적용 이미지 (캐시)."""
        opt = self.quality.get()
        cid, copt, cimg, ck = self._enh_cache
        if cid == id(img) and copt == opt:
            return cimg, ck
        if enhance.parse(opt)[0] == "sharpen":
            self.set_status("화질 개선 처리 중…")
            self.update_idletasks()
        eimg, k = enhance.enhance(img, opt)
        self._enh_cache = (id(img), opt, eimg, k)
        return eimg, k

    def _output_image(self):
        unit = self.v_unit.get().strip()
        if self.is_field():
            if self.src is None:
                return None
            base, k = self._enhanced(self.src)
            if k == 1:
                return fmap.annotate(self.src, self.readings, unit)
            scaled = [dataclasses.replace(r, x=r.x * k + k // 2, y=r.y * k + k // 2)
                      for r in self.readings]
            return fmap.annotate(base, scaled, unit, scale=k)
        if not self.result:
            return None
        return self._enhanced(self.result.image)[0]

    def _spin(self, parent, text, var, lo, hi):
        f = ttk.Frame(parent); f.pack(fill="x", pady=3)
        ttk.Label(f, text=text).pack(side="left")
        ttk.Spinbox(f, from_=lo, to=hi, textvariable=var, width=6).pack(side="right")

    def _slider(self, parent, text, var, lo, hi):
        f = ttk.Frame(parent); f.pack(fill="x", pady=3)
        head = ttk.Frame(f); head.pack(fill="x")
        ttk.Label(head, text=text).pack(side="left")
        val = ttk.Label(head, text=f"{var.get():.2f}", style="Muted.TLabel")
        val.pack(side="right")
        ttk.Scale(f, from_=lo, to=hi, variable=var,
                  command=lambda v: val.configure(text=f"{float(v):.2f}")).pack(fill="x")

    def _tree(self, nb, cols):
        frame = ttk.Frame(nb, style="Card.TFrame")
        tv = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=90, anchor="center")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        tv.tag_configure("global", foreground="#ff6b7d")
        tv.tag_configure("local", foreground=ACCENT)
        return tv

    # ------------------------------------------------------------ input
    def paste(self):
        try:
            data = ImageGrab.grabclipboard()
        except Exception as e:
            messagebox.showerror(APP, f"클립보드를 읽을 수 없습니다.\n{e}")
            return
        if isinstance(data, list) and data and os.path.isfile(data[0]):
            return self._load_path(data[0])
        if not isinstance(data, Image.Image):
            self.set_status("클립보드에 이미지가 없습니다. 화면 캡처(Win+Shift+S) 후 다시 시도하세요.")
            return
        rgb = np.array(data.convert("RGB"))
        self._set_source(cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR), "capture")

    def open_file(self):
        p = filedialog.askopenfilename(
            title="해석결과 이미지 열기",
            filetypes=[("이미지", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"), ("모든 파일", "*.*")])
        if p:
            self._load_path(p)

    def _load_path(self, p):
        try:
            img = core.imread(p)
        except Exception as e:
            messagebox.showerror(APP, str(e))
            return
        self._set_source(img, os.path.splitext(os.path.basename(p))[0])

    def _set_source(self, img, name):
        self.src, self.src_name, self.result = img, name, None
        self._zoom, self._pan = 1.0, [0.0, 0.0]
        self.show_original = True
        self.readings, self._legend_lut = [], None
        self._fill_tables()
        h, w = img.shape[:2]
        if self.is_field():
            if self.cmap.get() == "범례에서 추출":
                self.cmap.current(0)
            self._refresh_field()
            self._show()
            self.set_status(f"입력 이미지: {name}  ({w}×{h})  · 최소/최대값 입력 후 맵을 클릭하세요.")
            return
        self._show()
        self.set_status(f"입력 이미지: {name}  ({w}×{h})  → 분석 실행 중…")
        self.run()

    # ------------------------------------------------------------ run
    def _settings(self):
        return core.Settings(group_count=int(self.v_count.get() or 0),
                             group_alpha=float(self.v_galpha.get()),
                             highlight_alpha=float(self.v_halpha.get()),
                             pad=int(self.v_pad.get()))

    def run(self):
        if self.is_field():
            return
        if self.src is None:
            self.set_status("먼저 이미지를 붙여넣거나 여세요.")
            return
        self.btn_run.state(["disabled"])
        self.progress.start(12)
        s = self._settings()
        threading.Thread(target=self._worker, args=(self.src.copy(), s), daemon=True).start()

    def _worker(self, img, s):
        try:
            res = core.process(img, s)
            err = None
        except Exception as e:
            res, err = None, e
        self.after(0, self._done, res, err)

    def _done(self, res, err):
        self.progress.stop()
        self.btn_run.state(["!disabled"])
        if err:
            messagebox.showerror(APP, f"분석 실패:\n{err}")
            return
        self.result = res
        self.show_original = False
        self._fill_tables()
        self._show()
        n_loc = sum(r["하이라이트"] == "국소" for r in res.rows)
        n_glb = sum(r["하이라이트"] == "글로벌" for r in res.rows)
        msg = (f"점 {len(res.points)}개 · 그룹 {len(res.groups)}개 · "
               f"국소 최소 {n_loc} + 글로벌 최소 {n_glb}")
        if res.warnings:
            msg += "   ⚠ " + " / ".join(res.warnings)
        self.set_status(msg)

    def redraw(self):
        if not self.result:
            return
        self.result = core.rerender(self.src, self.result.groups, self._settings())
        self.show_original = False
        self._fill_tables()
        self._show()
        self.set_status("수정된 값으로 다시 그렸습니다.")

    # ------------------------------------------------------------ tables
    def _fill_tables(self):
        for tv in (self.tree_sum, self.tree_pts):
            tv.delete(*tv.get_children())
        if not self.result:
            return
        for r in self.result.rows:
            tag = {"글로벌": "global", "국소": "local"}.get(r["하이라이트"], "")
            self.tree_sum.insert("", "end", values=list(r.values()), tags=(tag,))
        self._pt_index = []
        i = 0
        for g in self.result.groups:
            for p in sorted(g, key=lambda p: (p.value is None, p.value or 0)):
                i += 1
                self._pt_index.append(p)
                self.tree_pts.insert("", "end", iid=str(i - 1), values=[
                    i, f"G{p.group}", "Global" if p.cls == "GLOBAL" else p.cls,
                    "" if p.value is None else p.value,
                    "" if p.index is None else p.index, int(p.x), int(p.y)],
                    tags=("global",) if p.cls == "GLOBAL" else ())

    def _edit_value(self, event):
        iid = self.tree_pts.identify_row(event.y)
        if not iid:
            return
        p = self._pt_index[int(iid)]
        x, y, w, h = self.tree_pts.bbox(iid, "#4")
        e = ttk.Entry(self.tree_pts)
        e.insert(0, "" if p.value is None else str(p.value))
        e.place(x=x, y=y, width=w, height=h)
        e.focus_set(); e.select_range(0, "end")

        def commit(_=None):
            t = e.get().strip().replace(",", ".")
            try:
                p.value = float(t) if t else None
            except ValueError:
                pass
            e.destroy()
            self.redraw()
        e.bind("<Return>", commit)
        e.bind("<FocusOut>", commit)
        e.bind("<Escape>", lambda _: e.destroy())

    # ------------------------------------------------------------ view
    def toggle_view(self):
        if self.src is None or self.is_field():
            return
        self.show_original = not self.show_original
        self._show()

    def _show(self):
        self.canvas.delete("all")
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        if self.src is None:
            self.canvas.create_text(cw / 2, ch / 2 - 12, fill=MUTED, font=("Segoe UI", 13),
                                    text="Ctrl+V 로 해석결과 캡처를 붙여넣으세요")
            self.canvas.create_text(cw / 2, ch / 2 + 16, fill="#5d6574", font=("Segoe UI", 10),
                                    text="또는 Ctrl+O 로 이미지 파일 열기")
            return
        if self.is_field():
            img = self.src
            self.lbl_view.configure(text="전계 값 읽기 · 클릭하여 측정")
        else:
            img = self.src if (self.show_original or not self.result) else self.result.image
            self.lbl_view.configure(text="원본" if img is self.src else "결과 · 최소안전율부위")
        h, w = img.shape[:2]
        base = min((cw - 20) / w, (ch - 20) / h, 4.0)
        if base <= 0:
            return
        sc = base * self._zoom
        ox = cw / 2 - w * sc / 2 + self._pan[0]
        oy = ch / 2 - h * sc / 2 + self._pan[1]
        # 이미지가 화면 밖으로 완전히 벗어나지 않도록 이동 범위 제한
        if self._zoom > 1.0:
            ax, ay = cw - w * sc - 40, ch - h * sc - 40
            ox = min(max(ox, min(ax, 40)), max(ax, 40))
            oy = min(max(oy, min(ay, 40)), max(ay, 40))
            self._pan = [ox - (cw / 2 - w * sc / 2), oy - (ch / 2 - h * sc / 2)]
        # 보이는 부분만 잘라서 확대 (큰 배율에서도 빠르게)
        ix0, iy0 = max(int(-ox / sc), 0), max(int(-oy / sc), 0)
        ix1, iy1 = min(int((cw - ox) / sc) + 2, w), min(int((ch - oy) / sc) + 2, h)
        self._view = (sc, ox, oy)
        self.lbl_zoom.configure(text="맞춤" if self._zoom == 1.0 else f"{sc * 100:.0f}%")
        if ix1 > ix0 and iy1 > iy0:
            kind = enhance.parse(self.quality.get())[0]
            eimg, k = self._enhanced(img)
            crop = cv2.cvtColor(eimg[iy0 * k:iy1 * k, ix0 * k:ix1 * k], cv2.COLOR_BGR2RGB)
            size = (max(1, round((ix1 - ix0) * sc)), max(1, round((iy1 - iy0) * sc)))
            if kind == "none":
                resample = Image.NEAREST if sc >= 3 else Image.LANCZOS
            elif sc / k >= 1:
                resample = Image.BICUBIC
            else:
                resample = Image.LANCZOS
            pil = Image.fromarray(crop).resize(size, resample)
            self._tk_img = ImageTk.PhotoImage(pil)
            self.canvas.create_image(ox + ix0 * sc, oy + iy0 * sc, image=self._tk_img,
                                     anchor="nw")
        if self.is_field():
            self._draw_markers()

    def _draw_markers(self):
        sc, ox, oy = self._view
        unit = self.v_unit.get().strip()
        for k, r in enumerate(self.readings, 1):
            x, y = ox + r.x * sc, oy + r.y * sc
            col = "#111111" if r.reliable else "#e8364f"
            self.canvas.create_oval(x - 6, y - 6, x + 6, y + 6, outline="#ffffff", width=3)
            self.canvas.create_oval(x - 6, y - 6, x + 6, y + 6, outline=col, width=1)
            text = f"#{k}  {fmap.fmt(r.value)} {unit}" + ("" if r.reliable else " ?")
            t = self.canvas.create_text(x + 12, y - 12, text=text, anchor="w",
                                        fill=col, font=("Segoe UI Semibold", 9))
            bx = self.canvas.bbox(t)
            bg = self.canvas.create_rectangle(bx[0] - 4, bx[1] - 2, bx[2] + 4, bx[3] + 2,
                                              fill="#ffffff", outline=col)
            self.canvas.tag_raise(t, bg)

    # ------------------------------------------------------------ output
    def save(self):
        img = self._output_image()
        if img is None or (self.is_field() and not self.readings):
            self.set_status("저장할 결과가 없습니다.")
            return
        suffix = "_field" if self.is_field() else "_minSF"
        p = filedialog.asksaveasfilename(
            title="결과 저장", defaultextension=".png",
            initialfile=f"{self.src_name}{suffix}.png", filetypes=[("PNG", "*.png")])
        if not p:
            return
        core.imwrite(p, img)
        csv_path = os.path.splitext(p)[0] + "_summary.csv"
        if self.is_field():
            fmap.save_csv(self.readings, csv_path, self.v_unit.get().strip())
        else:
            core.save_csv(self.result.rows, csv_path)
        self.set_status(f"저장 완료: {p}  +  {os.path.basename(csv_path)}")

    def copy_result(self):
        img = self._output_image()
        if img is None:
            return
        try:
            import io
            import win32clipboard  # type: ignore
            buf = io.BytesIO()
            Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB)).convert(
                "RGB").save(buf, "BMP")
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, buf.getvalue()[14:])
            win32clipboard.CloseClipboard()
            self.set_status("결과 이미지를 클립보드에 복사했습니다.")
        except Exception as e:
            self.set_status(f"클립보드 복사 실패 (Windows 전용): {e}")

    def set_status(self, text):
        self.status.configure(text=text)


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()
