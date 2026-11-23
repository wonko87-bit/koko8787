"""PostFX 2D - GUI (Tkinter)."""
from __future__ import annotations

import dataclasses
import datetime
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
from . import paths as pmod
from . import settings

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

# ---------------------------------------------------------------- 해상도 대응
# UI 배율: Windows 디스플레이 배율(DPI)을 자동으로 따르고, 설정에서 직접 지정할 수도 있다.
# 1080p 100% → 1.0,  1440p 125% → 1.25,  4K 150% → 1.5,  4K 200% → 2.0
UI = 1.0
UI_CHOICES = ["자동", "100%", "125%", "150%", "175%", "200%"]


def S(v):
    """설계 기준(96 DPI) 픽셀 → 현재 화면 픽셀."""
    return int(round(v * UI))


def detect_scale(root: tk.Tk) -> float:
    pref = settings.load().get("ui_scale", "자동")
    if pref != "자동":
        try:
            return int(pref.rstrip("%")) / 100
        except ValueError:
            pass
    try:
        dpi = root.winfo_fpixels("1i")
    except tk.TclError:
        dpi = 96.0
    return min(max(round(dpi / 96 * 4) / 4, 1.0), 3.0)     # 0.25 단위로 반올림


CLASS_COLORS = {"W1": "#ff8c00", "W2": "#28aa3c", "W3": "#1e5ae6",
                "W4": "#963cc8", "GLOBAL": "#e61414"}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        global UI
        UI = detect_scale(self)
        self.tk.call("tk", "scaling", UI * 96 / 72)      # 글꼴(pt)도 같은 배율로
        self.title(f"{APP}  ·  2D 해석결과 후처리")
        # 창 크기: 기준 1360x820 을 배율만큼 키우되, 화면의 90% 를 넘지 않게 + 가운데 배치
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(S(1360), int(sw * 0.9)), min(S(820), int(sh * 0.88))
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{max((sh - h) // 2 - S(20), 0)}")
        self.minsize(min(S(1000), w), min(S(640), h))
        self.configure(bg=BG)
        self.src: np.ndarray | None = None
        self.result: core.Result | None = None
        self.src_name = "capture"
        self.path_list: list[pmod.PathImage] = []      # 모드 3
        self.path_merged: np.ndarray | None = None
        self.path_sel: int | None = None
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
        self._load_legend()
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
        for i in range(3):                                 # Ctrl+1/2/3 = 모드 전환
            self.bind_all(f"<Control-Key-{i + 1}>", lambda e, i=i: self.set_mode(i))
        self.bind_all("<Control-Return>", lambda e: self.merge_paths())
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
                     padding=(S(12), S(7)), focusthickness=0)
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
                     rowheight=S(26), borderwidth=0)
        st.configure("Treeview.Heading", background=PANEL, foreground=MUTED,
                     font=("Segoe UI Semibold", 9), borderwidth=0)
        st.map("Treeview", background=[("selected", "#3a4252")])
        st.configure("TNotebook", background=PANEL, borderwidth=0)
        st.configure("TNotebook.Tab", background=PANEL, foreground=MUTED, padding=(S(14), S(6)))
        st.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", FG)])
        st.configure("TEntry", fieldbackground=CARD, foreground=FG, insertcolor=FG, borderwidth=0)
        st.configure("TCheckbutton", background=PANEL, foreground=FG, indicatorcolor=CARD)
        st.map("TCheckbutton", background=[("active", PANEL)],
               indicatorcolor=[("selected", ACCENT)])
        # 고해상도에서 작게 남는 고정 크기 요소들도 배율 적용
        st.configure("TCombobox", arrowsize=S(14), padding=S(3))
        st.configure("TSpinbox", arrowsize=S(12), padding=S(2))
        st.configure("TCheckbutton", indicatorsize=S(13), indicatormargin=S(3))
        st.configure("Vertical.TScrollbar", arrowsize=S(14), width=S(14))
        st.configure("Horizontal.TScale", sliderlength=S(30), sliderthickness=S(15))
        st.configure("Horizontal.TProgressbar", thickness=S(14))
        st.configure("Horizontal.TProgressbar", background=ACCENT, troughcolor=CARD, borderwidth=0)

    # ------------------------------------------------------------ layout
    def _build(self):
        top = ttk.Frame(self, style="Bg.TFrame", padding=(S(18), S(12), S(18), S(8)))
        top.pack(fill="x")
        dot = tk.Canvas(top, width=S(22), height=S(22), bg=BG, highlightthickness=0)
        dot.create_oval(S(2), S(2), S(20), S(20), fill=ACCENT, outline="")
        dot.create_oval(S(8), S(8), S(14), S(14), fill=ACCENT2, outline="")
        dot.pack(side="left", padx=(S(0), S(10)))
        ttk.Label(top, text=APP, style="Title.TLabel").pack(side="left")
        ttk.Label(top, text=f"  v{VERSION}  ·  2D 해석결과 이미지 후처리",
                  style="Sub.TLabel").pack(side="left", pady=(S(5), S(0)))
        self.ui_choice = ttk.Combobox(top, state="readonly", width=6, values=UI_CHOICES)
        self.ui_choice.set(settings.load().get("ui_scale", "자동"))
        self.ui_choice.pack(side="right")
        self.ui_choice.bind("<<ComboboxSelected>>", lambda e: self._on_ui_scale())
        ttk.Label(top, text=f"UI 크기 (현재 {UI * 100:.0f}%)", style="Sub.TLabel").pack(
            side="right", padx=(0, S(6)))

        body = ttk.Frame(self, style="Bg.TFrame", padding=(S(18), S(0), S(18), S(0)))
        body.pack(fill="both", expand=True)

        # --- 왼쪽 사이드바
        side = ttk.Frame(body, padding=S(16), width=S(290))
        side.pack(side="left", fill="y")
        side.pack_propagate(False)

        ttk.Label(side, text="입력", style="H.TLabel").pack(anchor="w")
        ttk.Button(side, text="📋  클립보드 붙여넣기  (Ctrl+V)",
                   command=self.paste).pack(fill="x", pady=(S(8), S(4)))
        ttk.Button(side, text="📂  이미지 파일 열기  (Ctrl+O)",
                   command=self.open_file).pack(fill="x")

        ttk.Label(side, text="모드", style="H.TLabel").pack(anchor="w", pady=(S(20), S(6)))
        self.mode = ttk.Combobox(side, state="readonly",
                                 values=["1. 최소안전율부위  (Ctrl+1)",
                                         "2. 전계 값 읽기  (Ctrl+2)",
                                         "3. 연면 경로 취합  (Ctrl+3)"])
        self.mode.current(0)
        self.mode.pack(fill="x")
        self.mode.bind("<<ComboboxSelected>>", lambda e: self._on_mode())

        ttk.Label(side, text="옵션", style="H.TLabel").pack(anchor="w", pady=(S(20), S(6)))
        self.opt_host = ttk.Frame(side)
        self.opt_host.pack(fill="x")
        self.opt1 = ttk.Frame(self.opt_host)
        self.opt2 = ttk.Frame(self.opt_host)
        self.opt3 = ttk.Frame(self.opt_host)
        self.opt1.pack(fill="x")
        self._build_opt2(self.opt2)
        self._build_opt3(self.opt3)
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
        self.btn_run.pack(fill="x", pady=(S(22), S(6)))
        self.btn_redraw = ttk.Button(side, text="↻  표 수정값으로 다시 그리기",
                                     command=self.redraw)
        self.btn_redraw.pack(fill="x")
        self.progress = ttk.Progressbar(side, mode="indeterminate")
        self.progress.pack(fill="x", pady=(S(10), S(0)))
        side = side_main

        ttk.Label(side, text="출력", style="H.TLabel").pack(anchor="w", pady=(S(20), S(6)))
        ttk.Button(side, text="💾  PNG + CSV 저장  (Ctrl+S)",
                   command=self.save).pack(fill="x", pady=(S(0), S(4)))
        ttk.Button(side, text="📄  결과 이미지 클립보드 복사",
                   command=self.copy_result).pack(fill="x")

        legend = ttk.Frame(side)
        legend.pack(side="bottom", fill="x")
        self.legend1 = legend
        ttk.Label(legend, text="점 색상 범례", style="Muted.TLabel").pack(anchor="w")
        for k, c in CLASS_COLORS.items():
            row = ttk.Frame(legend); row.pack(anchor="w", pady=S(1))
            cv = tk.Canvas(row, width=S(12), height=S(12), bg=PANEL, highlightthickness=0)
            cv.create_oval(S(1), S(1), S(11), S(11), fill=c, outline="")
            cv.pack(side="left", padx=(S(0), S(6)))
            ttk.Label(row, text=core.POINT_CLASSES[k]["label"],
                      style="Muted.TLabel").pack(side="left")

        # --- 가운데: 뷰어 + 테이블
        main = ttk.Frame(body, style="Bg.TFrame")
        main.pack(side="left", fill="both", expand=True, padx=(S(14), S(0)))

        viewer = ttk.Frame(main, padding=S(1))
        viewer.pack(fill="both", expand=True)
        bar = ttk.Frame(viewer, padding=(S(12), S(8)))
        bar.pack(fill="x")
        self.lbl_view = ttk.Label(bar, text="미리보기", style="H.TLabel")
        self.lbl_view.pack(side="left")
        zbar = ttk.Frame(bar)
        zbar.pack(side="left", padx=(S(18), S(0)))
        ttk.Button(zbar, text="－", width=3, command=lambda: self.zoom_by(0.8)).pack(side="left")
        self.lbl_zoom = ttk.Label(zbar, text="맞춤", width=6, anchor="center",
                                  style="Muted.TLabel")
        self.lbl_zoom.pack(side="left", padx=S(2))
        ttk.Button(zbar, text="＋", width=3, command=lambda: self.zoom_by(1.25)).pack(side="left")
        ttk.Button(zbar, text="맞춤 (Ctrl+0)", command=self.zoom_fit).pack(side="left", padx=(S(4), S(0)))
        ttk.Label(zbar, text="화질", style="Muted.TLabel").pack(side="left", padx=(S(14), S(4)))
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

        tables = self.tables = ttk.Notebook(main, height=S(210))
        tables.pack(fill="x", pady=(S(12), S(0)))
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
        self.tree_path = self._tree(tables, ["경로", "이미지", "크기", "선 픽셀 수", "상태"])
        tables.add(self.tree_path.master, text="  경로 목록  (선택: 미리보기 · Delete: 삭제)  ")
        tables.tab(3, state="hidden")
        self.tree_path.bind("<<TreeviewSelect>>", lambda e: self._on_path_select())
        self.tree_path.bind("<Delete>", lambda e: self.delete_path())

        self.status = ttk.Label(self, style="Sub.TLabel", padding=(S(20), S(6)))
        self.status.pack(fill="x")

    def _build_opt3(self, f):
        ttk.Label(f, style="Muted.TLabel", justify="left", wraplength=S(250), text=(
            "같은 구도·같은 크기 이미지를 경로별로\nCtrl+V (또는 Ctrl+O 여러 장) 로 추가"
        )).pack(anchor="w")
        self.lbl_pathcount = ttk.Label(f, text="추가된 경로: 0장", style="H.TLabel")
        self.lbl_pathcount.pack(anchor="w", pady=(S(8), S(4)))
        row = ttk.Frame(f); row.pack(fill="x")
        ttk.Button(row, text="✕ 선택 삭제", command=self.delete_path).pack(
            side="left", fill="x", expand=True, padx=(0, S(3)))
        ttk.Button(row, text="⟲ 모두 지우기", command=self.clear_paths).pack(
            side="left", fill="x", expand=True, padx=(S(3), 0))
        self.v_pcolor = tk.BooleanVar(value=True)
        self.v_pbadge = tk.BooleanVar(value=True)
        self.v_plegend = tk.BooleanVar(value=True)
        for text, var in (("경로별 색 구분", self.v_pcolor), ("번호 표시", self.v_pbadge),
                          ("범례 표시", self.v_plegend)):
            ttk.Checkbutton(f, text=text, variable=var, command=self._remerge).pack(
                anchor="w", pady=(S(6), 0))
        self.v_pthick = tk.IntVar(value=0)
        row = ttk.Frame(f); row.pack(fill="x", pady=(S(8), 0))
        ttk.Label(row, text="선 굵게 (px)").pack(side="left")
        ttk.Spinbox(row, from_=0, to=6, textvariable=self.v_pthick, width=6,
                    command=self._remerge).pack(side="right")
        ttk.Button(f, text="🔗  경로 취합  (Ctrl+Enter)", style="Accent.TButton",
                   command=self.merge_paths).pack(fill="x", pady=(S(16), 0))

    # ------------------------------------------------------------ mode 3
    def _add_path(self, img, name):
        ref = self.path_list[0].img.shape if self.path_list else None
        try:
            p = pmod.make_path(f"{len(self.path_list) + 1}_{name}", img, ref)
        except pmod.SizeMismatch as e:
            messagebox.showerror(APP, f"{e}\n\n모든 경로 이미지는 같은 크기여야 합니다. "
                                      "이 이미지는 추가하지 않았습니다.")
            self.set_status(f"⚠ {e} → 추가 안 함")
            return
        self.path_list.append(p)
        self.path_merged = None
        self.path_sel = len(self.path_list) - 1
        self._zoom, self._pan = 1.0, [0.0, 0.0]
        self._fill_path_table()
        self._show()
        n = len(self.path_list)
        msg = f"경로 {n} 추가 ({img.shape[1]}×{img.shape[0]}, 빨간 선 {p.pixels} px)"
        if not p.pixels:
            msg += "   ⚠ 빨간 선을 찾지 못했습니다"
        if n > len(pmod.PATH_COLORS):
            msg += f"   ⚠ {len(pmod.PATH_COLORS)}개를 넘어 색이 반복됩니다 (번호로 구분)"
        self.set_status(msg)

    def _fill_path_table(self):
        tv = self.tree_path
        tv.delete(*tv.get_children())
        colored = bool(self.v_pcolor.get())
        for i, p in enumerate(self.path_list):
            h, w = p.img.shape[:2]
            cname = pmod.PATH_COLORS[i % len(pmod.PATH_COLORS)][1] if colored else "빨강"
            tv.insert("", "end", iid=str(i), values=[f"{i + 1} ({cname})", p.name, f"{w}×{h}",
                                                     p.pixels, "OK" if p.pixels else "선 없음"],
                      tags=() if p.pixels else ("global",))
        self.lbl_pathcount.configure(text=f"추가된 경로: {len(self.path_list)}장")

    def _on_path_select(self):
        sel = self.tree_path.selection()
        if sel:
            self.path_sel = int(sel[0])
            self.show_original = True
            self._show()

    def delete_path(self):
        sel = self.tree_path.selection()
        if not sel or not self.path_list:
            return
        self.path_list.pop(int(sel[0]))
        self.path_merged = None
        self.path_sel = len(self.path_list) - 1 if self.path_list else None
        self._fill_path_table()
        self._show()

    def clear_paths(self):
        if self.path_list and not messagebox.askyesno(APP, "추가한 경로 이미지를 모두 지울까요?"):
            return
        self.path_list, self.path_merged, self.path_sel = [], None, None
        self._fill_path_table()
        self._show()
        self.set_status("경로 목록을 비웠습니다.")

    def merge_paths(self):
        if not self.is_paths():
            return
        if not self.path_list:
            self.set_status("먼저 경로 이미지를 붙여넣으세요.")
            return
        self.path_merged = pmod.merge(self.path_list, int(self.v_pthick.get() or 0),
                                      bool(self.v_pcolor.get()), bool(self.v_pbadge.get()),
                                      bool(self.v_plegend.get()))
        self.show_original = False
        self._fill_path_table()
        self._show()
        empty = sum(not p.pixels for p in self.path_list)
        self.set_status(f"경로 {len(self.path_list)}개 취합 완료"
                        + (f"   ⚠ 빨간 선 없는 이미지 {empty}장" if empty else "")
                        + "   · Space: 개별 이미지 / 결과 전환")

    def _remerge(self):
        self._fill_path_table()
        if self.path_merged is not None:
            self.merge_paths()

    def _view_img(self):
        """지금 화면에 보이는 이미지 (좌표 계산 기준)."""
        if self.is_paths():
            if self.path_merged is not None and not self.show_original:
                return self.path_merged
            if self.path_list:
                i = self.path_sel if self.path_sel is not None else len(self.path_list) - 1
                return self.path_list[min(i, len(self.path_list) - 1)].img
            return None
        if self.src is None:
            return None
        if self.is_field() or self.show_original or not self.result:
            return self.src
        return self.result.image

    def _build_opt2(self, f):
        self.v_vmin = tk.StringVar(value="0")
        self.v_vmax = tk.StringVar(value="1")
        self.v_unit = tk.StringVar(value="V/m")
        for text, var in (("최소값 (컬러맵 아래 끝)", self.v_vmin),
                          ("최대값 (컬러맵 위 끝)", self.v_vmax), ("단위", self.v_unit)):
            row = ttk.Frame(f); row.pack(fill="x", pady=S(3))
            ttk.Label(row, text=text).pack(side="left")
            ttk.Entry(row, textvariable=var, width=10, justify="right").pack(side="right")
            var.trace_add("write", lambda *_: self._refresh_field())
        ttk.Label(f, text="컬러맵", style="Muted.TLabel").pack(anchor="w", pady=(S(8), S(2)))
        self.cmap = ttk.Combobox(f, state="readonly", values=fmap.COLORMAPS)
        self.cmap.current(0)
        self.cmap.pack(fill="x")
        self.cmap.bind("<<ComboboxSelected>>", lambda e: self._on_cmap())
        self.lut_preview = tk.Canvas(f, height=S(14), bg=PANEL, highlightthickness=0)
        self.lut_preview.pack(fill="x", pady=(S(6), S(0)))
        self.lut_preview.bind("<Configure>", lambda e: self._draw_lut_preview())
        self.lbl_legend = ttk.Label(f, text="", style="Muted.TLabel")
        self.lbl_legend.pack(anchor="w")
        row = ttk.Frame(f); row.pack(fill="x", pady=(S(4), S(0)))
        self.btn_legend = ttk.Button(row, text="🎯 범례 지정", command=self.start_legend_pick)
        self.btn_legend.pack(side="left", fill="x", expand=True, padx=(S(0), S(3)))
        ttk.Button(row, text="⟲ 범례 초기화", command=self.reset_legend).pack(
            side="left", fill="x", expand=True, padx=(S(3), S(0)))
        self.v_log = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="로그 스케일", variable=self.v_log,
                        command=self._refresh_field).pack(anchor="w", pady=(S(8), S(0)))
        row = ttk.Frame(f); row.pack(fill="x", pady=(S(10), S(0)))
        ttk.Button(row, text="↶ 되돌리기", command=self.undo_reading).pack(
            side="left", fill="x", expand=True, padx=(S(0), S(3)))
        ttk.Button(row, text="✕ 모두 지우기", command=self.clear_readings).pack(
            side="left", fill="x", expand=True, padx=(S(3), S(0)))
        ttk.Label(f, style="Muted.TLabel", justify="left", wraplength=S(250), text=(
            "클릭: 측정 · 우클릭: 삭제 · 휠: 확대/축소 · 드래그: 이동\n"
            "정확도↑: 범례 지정 → 컬러바 최소 끝→최대 끝 드래그"
        )).pack(anchor="w", pady=(S(10), S(0)))

    # ------------------------------------------------------------ mode 2
    def is_field(self):
        return self.mode.current() == 1

    def is_paths(self):
        return self.mode.current() == 2

    def set_mode(self, i):
        if self.mode.current() != i:
            self.mode.current(i)
            self._on_mode()

    def _on_mode(self):
        m = self.mode.current()
        for f in (self.opt1, self.opt2, self.opt3):
            f.pack_forget()
        self.legend1.pack_forget()
        self.btn_toggle.pack_forget()
        tabs = {0: (0, 1), 1: (2,), 2: (3,)}[m]
        for i in range(4):
            self.tables.tab(i, state="normal" if i in tabs else "hidden")
        self.tables.select(tabs[0])
        self._zoom, self._pan = 1.0, [0.0, 0.0]
        if m == 0:
            self.opt1.pack(fill="x"); self.legend1.pack(side="bottom", fill="x")
            self.btn_toggle.pack(side="right")
            if self.src is not None and self.result is None:
                self.set_status("모드 1: '분석 실행' 을 누르면 점/숫자 분석을 시작합니다.")
            else:
                self.set_status("모드 1: 최소안전율부위")
        elif m == 1:
            self.opt2.pack(fill="x")
            self.show_original = True
            self._refresh_field()
            self.set_status("모드 2: 최소/최대값을 입력하고 맵 위를 클릭하세요.")
        else:
            self.opt3.pack(fill="x")
            self.btn_toggle.pack(side="right")
            self.set_status(f"모드 3: 경로 이미지를 Ctrl+V 로 차례로 붙여넣고 '경로 취합'. "
                            f"(현재 {len(self.path_list)}장)")
        self._show()

    def _on_cmap(self):
        if self.cmap.get() == "범례에서 추출" and self._legend_lut is None:
            self.start_legend_pick()
        else:
            settings.update(cmap=self.cmap.get())
        self._refresh_field()

    def _load_legend(self):
        """지난 실행에서 저장한 범례/컬러맵 복원."""
        d = settings.load()
        lut = d.get("legend_lut")
        if lut:
            try:
                self._legend_lut = np.array(lut, np.uint8).reshape(-1, 3)
                self._legend_saved = d.get("legend_saved", "")
            except ValueError:
                self._legend_lut = None
        name = d.get("cmap")
        if name in fmap.COLORMAPS and (name != "범례에서 추출" or self._legend_lut is not None):
            self.cmap.set(name)
        self._update_legend_label()

    def _update_legend_label(self):
        if self._legend_lut is not None and self.cmap.get() == "범례에서 추출":
            when = getattr(self, "_legend_saved", "")
            self.lbl_legend.configure(text=f"저장된 범례 사용 중 · {when} 지정" if when
                                      else "저장된 범례 사용 중")
        elif self._legend_lut is not None:
            self.lbl_legend.configure(text="저장된 범례 있음 (컬러맵에서 '범례에서 추출' 선택)")
        else:
            self.lbl_legend.configure(text="")

    def reset_legend(self):
        if self._legend_lut is None and self.cmap.current() == 0:
            return
        if not messagebox.askyesno(APP, "저장된 범례 색상을 지우고 기본 컬러맵(Rainbow)으로 "
                                        "되돌릴까요?"):
            return
        self._legend_lut = None
        self._legend_saved = ""
        self.cmap.current(0)
        settings.update(legend_lut=None, legend_saved=None, cmap=None)
        self._refresh_field()
        self.set_status("범례를 초기화했습니다. 기본 Rainbow 컬러맵을 사용합니다.")

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
        self._update_legend_label()
        self._fill_field_table()
        if self.is_field():
            self._show()

    def _draw_lut_preview(self):
        c = self.lut_preview
        c.delete("all")
        lut = self._current_lut()
        w = max(c.winfo_width(), 10)
        if lut is None:
            c.create_text(w / 2, S(7), text="범례 미지정", fill=MUTED, font=("Segoe UI", 8))
            return
        for i in range(w):
            b, g, r = lut[int(i / (w - 1) * (len(lut) - 1))]
            c.create_line(i, 0, i, S(14), fill=f"#{r:02x}{g:02x}{b:02x}")

    def _img_xy(self, ex, ey):
        sc, ox, oy = self._view
        x, y = int((ex - ox) / sc), int((ey - oy) / sc)
        img = self._view_img()
        if img is None:
            return None
        h, w = img.shape[:2]
        return (x, y) if 0 <= x < w and 0 <= y < h else None

    # ------------------------------------------------------------ zoom / pan
    def _base_scale(self):
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        h, w = self._view_img().shape[:2]
        return min((cw - 20) / w, (ch - 20) / h, 4.0)

    def zoom_by(self, f, ex=None, ey=None):
        if self._view_img() is None:
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
        h, w = self._view_img().shape[:2]
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
        if self._view_img() is None:
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
            self.canvas.create_line(*self._drag, e.x, e.y, fill="#ffffff", width=S(3), tags="drag")
            self.canvas.create_line(*self._drag, e.x, e.y, fill=ACCENT, width=S(1),
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
        self._legend_saved = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        self.cmap.set("범례에서 추출")
        settings.update(legend_lut=self._legend_lut.tolist(), legend_saved=self._legend_saved,
                        cmap="범례에서 추출")
        self._refresh_field()
        self.set_status("범례 지정 완료 · 저장됨 (초기화 전까지 유지). 이제 맵을 클릭하세요.")

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
        if ((r.x - xy[0]) ** 2 + (r.y - xy[1]) ** 2) ** 0.5 * sc < S(20):
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
        if self.is_paths():
            return None if self.path_merged is None else self._enhanced(self.path_merged)[0]
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
        f = ttk.Frame(parent); f.pack(fill="x", pady=S(3))
        ttk.Label(f, text=text).pack(side="left")
        ttk.Spinbox(f, from_=lo, to=hi, textvariable=var, width=6).pack(side="right")

    def _slider(self, parent, text, var, lo, hi):
        f = ttk.Frame(parent); f.pack(fill="x", pady=S(3))
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
        if self.is_paths():
            ps = filedialog.askopenfilenames(
                title="경로 이미지 열기 (여러 장 선택 가능)",
                filetypes=[("이미지", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"), ("모든 파일", "*.*")])
            for p in ps:
                self._load_path(p)
            return
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
        if self.is_paths():
            self._add_path(img, name)
            return
        self.src, self.src_name, self.result = img, name, None
        self._zoom, self._pan = 1.0, [0.0, 0.0]
        self.show_original = True
        self.readings = []          # 범례(컬러맵)와 범위는 초기화 전까지 유지
        self._fill_tables()
        h, w = img.shape[:2]
        if self.is_field():
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
        if self.is_field() or self.is_paths():
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
        if self.is_paths():
            if self.path_merged is not None:
                self.show_original = not self.show_original
                self._show()
            return
        if self.src is None or self.is_field():
            return
        self.show_original = not self.show_original
        self._show()

    def _show(self):
        self.canvas.delete("all")
        cw, ch = self.canvas.winfo_width(), self.canvas.winfo_height()
        img = self._view_img()
        if img is None:
            self.lbl_view.configure(text=("연면 경로 취합" if self.is_paths() else
                                          "전계 값 읽기" if self.is_field() else "미리보기"))
            first = ("Ctrl+V 로 경로 이미지를 차례로 붙여넣으세요" if self.is_paths()
                     else "Ctrl+V 로 해석결과 캡처를 붙여넣으세요")
            self.canvas.create_text(cw / 2, ch / 2 - S(12), fill=MUTED, font=("Segoe UI", 13),
                                    text=first)
            self.canvas.create_text(cw / 2, ch / 2 + S(16), fill="#5d6574", font=("Segoe UI", 10),
                                    text="또는 Ctrl+O 로 이미지 파일 열기")
            return
        if self.is_paths():
            if img is self.path_merged:
                self.lbl_view.configure(text=f"결과 · 경로 {len(self.path_list)}개 취합")
            else:
                i = self.path_list.index(next(p for p in self.path_list if p.img is img)) + 1
                self.lbl_view.configure(text=f"경로 {i} / {len(self.path_list)}")
        elif self.is_field():
            self.lbl_view.configure(text="전계 값 읽기 · 클릭하여 측정")
        else:
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
            m = S(40)
            ax, ay = cw - w * sc - m, ch - h * sc - m
            ox = min(max(ox, min(ax, m)), max(ax, m))
            oy = min(max(oy, min(ay, m)), max(ay, m))
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
            r6 = S(6)
            self.canvas.create_oval(x - r6, y - r6, x + r6, y + r6, outline="#ffffff", width=S(3))
            self.canvas.create_oval(x - r6, y - r6, x + r6, y + r6, outline=col, width=S(1))
            text = f"#{k}  {fmap.fmt(r.value)} {unit}" + ("" if r.reliable else " ?")
            t = self.canvas.create_text(x + S(12), y - S(12), text=text, anchor="w",
                                        fill=col, font=("Segoe UI Semibold", 9))
            bx = self.canvas.bbox(t)
            bg = self.canvas.create_rectangle(bx[0] - S(4), bx[1] - S(2), bx[2] + S(4), bx[3] + S(2),
                                              fill="#ffffff", outline=col)
            self.canvas.tag_raise(t, bg)

    # ------------------------------------------------------------ output
    def save(self):
        img = self._output_image()
        if img is None or (self.is_field() and not self.readings):
            self.set_status("저장할 결과가 없습니다." + (" '경로 취합' 을 먼저 누르세요."
                                                       if self.is_paths() else ""))
            return
        if self.is_paths():
            initial = "paths_merged.png"
        else:
            initial = f"{self.src_name}{'_field' if self.is_field() else '_minSF'}.png"
        p = filedialog.asksaveasfilename(
            title="결과 저장", defaultextension=".png",
            initialfile=initial, filetypes=[("PNG", "*.png")])
        if not p:
            return
        core.imwrite(p, img)
        csv_path = os.path.splitext(p)[0] + "_summary.csv"
        if self.is_paths():
            pmod.save_csv(self.path_list, csv_path, bool(self.v_pcolor.get()))
        elif self.is_field():
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

    def _on_ui_scale(self):
        choice = self.ui_choice.get()
        settings.update(ui_scale=choice)
        messagebox.showinfo(APP, f"UI 크기를 '{choice}' 로 저장했습니다.\n"
                                 "프로그램을 다시 시작하면 적용됩니다.")


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()
