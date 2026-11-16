"""PostFX 2D - GUI (Tkinter)."""
from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import cv2
import numpy as np
from PIL import Image, ImageGrab, ImageTk

from . import core

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
        self._style()
        self._build()
        self.bind_all("<Control-v>", lambda e: self.paste())
        self.bind_all("<Control-o>", lambda e: self.open_file())
        self.bind_all("<Control-s>", lambda e: self.save())
        self.bind_all("<space>", lambda e: self.toggle_view())
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
                                 values=["1. 최소안전율부위"])
        self.mode.current(0)
        self.mode.pack(fill="x")

        ttk.Label(side, text="옵션", style="H.TLabel").pack(anchor="w", pady=(20, 6))
        self.v_count = tk.IntVar(value=0)
        self._spin(side, "그룹 수 (0 = 자동)", self.v_count, 0, 50)
        self.v_galpha = tk.DoubleVar(value=0.18)
        self._slider(side, "그룹 음영 농도", self.v_galpha, 0.05, 0.5)
        self.v_halpha = tk.DoubleVar(value=0.45)
        self._slider(side, "하이라이트 농도", self.v_halpha, 0.2, 0.8)
        self.v_pad = tk.IntVar(value=14)
        self._spin(side, "음영 여백 (px)", self.v_pad, 2, 60)
        self.v_clean = tk.BooleanVar(value=False)
        ttk.Checkbutton(side, text="최소 지점 외 점·숫자 지우기", variable=self.v_clean,
                        command=self.redraw).pack(anchor="w", pady=(6, 0))

        self.btn_run = ttk.Button(side, text="▶  분석 실행", style="Accent.TButton",
                                  command=self.run)
        self.btn_run.pack(fill="x", pady=(22, 6))
        self.btn_redraw = ttk.Button(side, text="↻  표 수정값으로 다시 그리기",
                                     command=self.redraw)
        self.btn_redraw.pack(fill="x")
        self.progress = ttk.Progressbar(side, mode="indeterminate")
        self.progress.pack(fill="x", pady=(10, 0))

        ttk.Label(side, text="출력", style="H.TLabel").pack(anchor="w", pady=(20, 6))
        ttk.Button(side, text="💾  PNG + CSV 저장  (Ctrl+S)",
                   command=self.save).pack(fill="x", pady=(0, 4))
        ttk.Button(side, text="📄  결과 이미지 클립보드 복사",
                   command=self.copy_result).pack(fill="x")

        legend = ttk.Frame(side)
        legend.pack(side="bottom", fill="x")
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
        ttk.Button(bar, text="원본 / 결과 전환 (Space)",
                   command=self.toggle_view).pack(side="right")
        self.canvas = tk.Canvas(viewer, bg="#0f1115", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self._show())

        tables = ttk.Notebook(main, height=210)
        tables.pack(fill="x", pady=(12, 0))
        self.tree_sum = self._tree(tables, ["그룹", "점 개수", "최소 안전율", "구분",
                                            "하이라이트", "X(px)", "Y(px)", "그룹 최대"])
        tables.add(self.tree_sum.master, text="  요약 (그룹별 최소)  ")
        self.tree_pts = self._tree(tables, ["#", "그룹", "구분", "안전율", "X(px)", "Y(px)"])
        tables.add(self.tree_pts.master, text="  전체 점  (더블클릭: 값 수정)  ")
        self.tree_pts.bind("<Double-1>", self._edit_value)

        self.status = ttk.Label(self, style="Sub.TLabel", padding=(20, 6))
        self.status.pack(fill="x")

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
        self.show_original = True
        self._fill_tables()
        self._show()
        h, w = img.shape[:2]
        self.set_status(f"입력 이미지: {name}  ({w}×{h})  → 분석 실행 중…")
        self.run()

    # ------------------------------------------------------------ run
    def _settings(self):
        return core.Settings(group_count=int(self.v_count.get() or 0),
                             group_alpha=float(self.v_galpha.get()),
                             highlight_alpha=float(self.v_halpha.get()),
                             pad=int(self.v_pad.get()),
                             clean=bool(self.v_clean.get()))

    def run(self):
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
                    "" if p.value is None else p.value, int(p.x), int(p.y)],
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
        if self.src is None:
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
        img = self.src if (self.show_original or not self.result) else self.result.image
        self.lbl_view.configure(text="원본" if img is self.src else "결과 · 최소안전율부위")
        h, w = img.shape[:2]
        sc = min((cw - 20) / w, (ch - 20) / h, 4.0)
        if sc <= 0:
            return
        rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        pil = Image.fromarray(rgb).resize((max(1, int(w * sc)), max(1, int(h * sc))),
                                          Image.LANCZOS)
        self._tk_img = ImageTk.PhotoImage(pil)
        self.canvas.create_image(cw / 2, ch / 2, image=self._tk_img)

    # ------------------------------------------------------------ output
    def save(self):
        if not self.result:
            self.set_status("저장할 결과가 없습니다.")
            return
        p = filedialog.asksaveasfilename(
            title="결과 저장", defaultextension=".png",
            initialfile=f"{self.src_name}_minSF.png", filetypes=[("PNG", "*.png")])
        if not p:
            return
        core.imwrite(p, self.result.image)
        csv_path = os.path.splitext(p)[0] + "_summary.csv"
        core.save_csv(self.result.rows, csv_path)
        self.set_status(f"저장 완료: {p}  +  {os.path.basename(csv_path)}")

    def copy_result(self):
        if not self.result:
            return
        try:
            import io
            import win32clipboard  # type: ignore
            buf = io.BytesIO()
            Image.fromarray(cv2.cvtColor(self.result.image, cv2.COLOR_BGR2RGB)).convert(
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
