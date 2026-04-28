"""
ui_utils.py — 共用顏色、字型、元件
"""
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import calendar

# ── 顏色 & 字型 ───────────────────────────────────────────────────────────────
# 淺色模式
_LIGHT = {
    "primary":    "#2C6FAC",
    "primary_dk": "#1A4F7A",
    "bg":         "#F0F4F8",
    "white":      "#FFFFFF",
    "text":       "#2D3748",
    "text_lt":    "#718096",
    "border":     "#CBD5E0",
    "danger":     "#E53E3E",
    "sidebar":    "#1A4F7A",
    "sidebar_sel":"#2C6FAC",
    "ok_fg":      "#3B6D11",
    "ok_bg":      "#EAF3DE",
    "ng_fg":      "#854F0B",
    "ng_bg":      "#FAEEDA",
    "row_alt":    "#F7FAFC",
    "header_bg":  "#EBF5FF",
}

# 深色模式
_DARK = {
    "primary":    "#4A90D9",
    "primary_dk": "#2C6FAC",
    "bg":         "#1A1D23",
    "white":      "#252930",
    "text":       "#E2E8F0",
    "text_lt":    "#A0AEC0",
    "border":     "#3D4451",
    "danger":     "#FC8181",
    "sidebar":    "#111317",
    "sidebar_sel":"#2D3748",
    "ok_fg":      "#68D391",
    "ok_bg":      "#1C3D2A",
    "ng_fg":      "#F6AD55",
    "ng_bg":      "#3D2B10",
    "row_alt":    "#1E2128",
    "header_bg":  "#1E2736",
}

C = dict(_LIGHT)  # 預設淺色

def apply_theme(dark=False):
    """切換深色/淺色模式，更新全域 C 字典"""
    src = _DARK if dark else _LIGHT
    C.update(src)
F  = "Microsoft JhengHei"
FS = (F, 9);  FM = (F, 10);  FL = (F, 11);  FX = (F, 13)
FMB = (F, 10, "bold");  FLB = (F, 11, "bold");  FXB = (F, 13, "bold")

# ── 通用工廠 ──────────────────────────────────────────────────────────────────
def lbl(parent, text, size=10, bold=False, fg=None, bg=None, **kw):
    bg = bg or C["bg"]; fg = fg or C["text"]
    font = (F, size, "bold" if bold else "normal")
    return tk.Label(parent, text=text, font=font, bg=bg, fg=fg, **kw)

def btn(parent, text, cmd, bg=None, fg="white", size=10, width=None, **kw):
    bg = bg or C["primary"]
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                  font=(F, size, "bold"), relief="flat", cursor="hand2",
                  activebackground=C["primary_dk"], activeforeground="white",
                  padx=8, pady=5, **kw)
    if width:
        b.configure(width=width)
    return b

def ghost_btn(parent, text, cmd, size=10, **kw):
    """淡色框線按鈕"""
    return tk.Button(parent, text=text, command=cmd,
                     bg=C["white"], fg=C["text_lt"],
                     font=(F, size), relief="flat", cursor="hand2",
                     highlightbackground=C["border"], highlightthickness=1,
                     padx=8, pady=4, **kw)

def danger_btn(parent, text, cmd, size=10, **kw):
    return tk.Button(parent, text=text, command=cmd,
                     bg=C["white"], fg=C["danger"],
                     font=(F, size), relief="flat", cursor="hand2",
                     highlightbackground=C["danger"], highlightthickness=1,
                     padx=8, pady=4, **kw)

def card(parent, title=""):
    """白色卡片，回傳 (outer_frame, inner_frame)"""
    outer = tk.Frame(parent, bg=C["bg"])
    if title:
        lbl(outer, title, size=11, bold=True).pack(anchor="w", pady=(8, 3))
    inner = tk.Frame(outer, bg=C["white"],
                     highlightbackground=C["border"], highlightthickness=1)
    inner.pack(fill="both", expand=True)
    return outer, inner

def sep(parent, orient="h", color=None):
    color = color or C["border"]
    if orient == "h":
        return tk.Frame(parent, bg=color, height=1)
    else:
        return tk.Frame(parent, bg=color, width=1)

def combo(parent, var, values, width=12, state="readonly", **kw):
    cb = ttk.Combobox(parent, textvariable=var, values=values,
                      width=width, state=state, **kw)
    return cb

# ── 樣式化 Treeview ───────────────────────────────────────────────────────────
def styled_tree(parent, cols, headings, widths, height=15, show="headings"):
    style = ttk.Style()
    style.configure("Custom.Treeview",
                    font=(F, 10), rowheight=30,
                    background=C["white"], fieldbackground=C["white"])
    style.configure("Custom.Treeview.Heading",
                    font=(F, 10, "bold"), background=C["header_bg"])
    style.map("Custom.Treeview", background=[("selected", C["primary"])])

    frame = tk.Frame(parent, bg=C["white"])
    tree = ttk.Treeview(frame, columns=cols, show=show,
                        height=height, style="Custom.Treeview")
    for col, heading, width in zip(cols, headings, widths):
        tree.heading(col, text=heading)
        tree.column(col, width=width, anchor="center")

    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")
    return frame, tree

# ── 對話框 ────────────────────────────────────────────────────────────────────
def ask_string(parent, title, prompt, default=""):
    """彈出輸入框，回傳字串或 None"""
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.grab_set()
    dlg.resizable(False, False)
    result = [None]
    tk.Label(dlg, text=prompt, font=FM, padx=20, pady=12).pack()
    var = tk.StringVar(value=default)
    e = tk.Entry(dlg, textvariable=var, font=FL, width=24)
    e.pack(padx=20, pady=(0, 8))
    e.focus_set(); e.select_range(0, "end")
    def ok(ev=None):
        result[0] = var.get().strip()
        dlg.destroy()
    e.bind("<Return>", ok)
    row = tk.Frame(dlg); row.pack(pady=10)
    btn(row, "確定", ok, width=8).pack(side="left", padx=5)
    ghost_btn(row, "取消", dlg.destroy).pack(side="left")
    parent.wait_window(dlg)
    return result[0] if result[0] else None

def ask_choice(parent, title, prompt, choices):
    """彈出單選框，回傳選擇或 None"""
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.grab_set()
    dlg.resizable(False, False)
    result = [None]
    tk.Label(dlg, text=prompt, font=FM, padx=20, pady=12).pack()
    var = tk.StringVar(value=choices[0] if choices else "")
    for c in choices:
        tk.Radiobutton(dlg, text=c, variable=var, value=c,
                       font=FM).pack(anchor="w", padx=24)
    def ok():
        result[0] = var.get()
        dlg.destroy()
    row = tk.Frame(dlg); row.pack(pady=10)
    btn(row, "確定", ok, width=8).pack(side="left", padx=5)
    ghost_btn(row, "取消", dlg.destroy).pack(side="left")
    parent.wait_window(dlg)
    return result[0]

def ask_yesno(parent, title, message):
    """自訂確認框，回傳 True/False"""
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.grab_set()
    dlg.resizable(False, False)
    result = [False]
    tk.Label(dlg, text=message, font=FM, padx=24, pady=16,
             wraplength=300, justify="left").pack()
    row = tk.Frame(dlg); row.pack(pady=10)
    def yes():
        result[0] = True; dlg.destroy()
    btn(row, "確定", yes, width=8).pack(side="left", padx=5)
    ghost_btn(row, "取消", dlg.destroy).pack(side="left")
    parent.wait_window(dlg)
    return result[0]

# ── 日曆選擇器 ────────────────────────────────────────────────────────────────
class DatePicker(tk.Toplevel):
    """點擊按鈕後彈出日曆，選擇後回填到 StringVar"""
    def __init__(self, parent, date_var):
        super().__init__(parent)
        self.date_var = date_var
        self.title("選擇日期")
        self.grab_set()
        self.resizable(False, False)
        # 解析現有值
        try:
            d = datetime.strptime(date_var.get(), "%Y/%m/%d")
            self.year, self.month = d.year, d.month
        except Exception:
            now = datetime.now()
            self.year, self.month = now.year, now.month
        self._build()

    def _build(self):
        # 標題列（年月 + 前後箭頭）
        hdr = tk.Frame(self, bg=C["white"], padx=10, pady=8)
        hdr.pack(fill="x")
        tk.Button(hdr, text="‹", command=self._prev, font=FL,
                  relief="flat", cursor="hand2", bg=C["white"]).pack(side="left")
        self.title_lbl = tk.Label(hdr, font=FLB, bg=C["white"], width=12)
        self.title_lbl.pack(side="left", expand=True)
        tk.Button(hdr, text="›", command=self._next, font=FL,
                  relief="flat", cursor="hand2", bg=C["white"]).pack(side="right")
        # 星期標題
        dow_frame = tk.Frame(self, bg=C["white"])
        dow_frame.pack(fill="x", padx=8)
        for d in ["日","一","二","三","四","五","六"]:
            tk.Label(dow_frame, text=d, font=FS, bg=C["white"],
                     fg=C["text_lt"], width=4).pack(side="left")
        # 日期格
        self.cal_frame = tk.Frame(self, bg=C["white"], padx=8, pady=4)
        self.cal_frame.pack()
        self._render()

    def _render(self):
        for w in self.cal_frame.winfo_children():
            w.destroy()
        self.title_lbl.config(text=f"{self.year} 年 {self.month} 月")
        cal = calendar.monthcalendar(self.year, self.month)
        today = datetime.now()
        try:
            sel = datetime.strptime(self.date_var.get(), "%Y/%m/%d")
        except Exception:
            sel = None
        for week in cal:
            row = tk.Frame(self.cal_frame, bg=C["white"])
            row.pack()
            for day in week:
                if day == 0:
                    tk.Label(row, text="", width=4, bg=C["white"]).pack(side="left")
                else:
                    is_sel = sel and sel.year == self.year and sel.month == self.month and sel.day == day
                    is_today = today.year == self.year and today.month == self.month and today.day == day
                    bg = C["primary"] if is_sel else (C["ok_bg"] if is_today else C["white"])
                    fg = "white" if is_sel else (C["ok_fg"] if is_today else C["text"])
                    tk.Button(row, text=str(day), width=3,
                              bg=bg, fg=fg, relief="flat", cursor="hand2",
                              font=FS,
                              command=lambda d=day: self._pick(d)).pack(side="left", padx=1, pady=1)
        tk.Frame(self, bg=C["white"], height=6).pack()

    def _prev(self):
        self.month -= 1
        if self.month < 1:
            self.month = 12; self.year -= 1
        self._render()

    def _next(self):
        self.month += 1
        if self.month > 12:
            self.month = 1; self.year += 1
        self._render()

    def _pick(self, day):
        self.date_var.set(f"{self.year}/{self.month:02d}/{day:02d}")
        self.destroy()

def date_btn(parent, var, bg=None):
    """日期選擇按鈕，點擊後彈出 DatePicker"""
    bg = bg or C["white"]
    def open_picker():
        DatePicker(parent, var)
    b = tk.Button(parent, textvariable=var, command=open_picker,
                  font=FM, bg=bg, fg=C["text"], relief="flat", cursor="hand2",
                  highlightbackground=C["border"], highlightthickness=1,
                  padx=8, pady=4)
    return b

def bind_mousewheel(canvas):
    """綁定滑鼠滾輪到 Canvas（Windows/Mac 相容）"""
    def _on_win(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    def _on_linux_up(event):
        canvas.yview_scroll(-1, "units")
    def _on_linux_dn(event):
        canvas.yview_scroll(1, "units")
    canvas.bind("<MouseWheel>", _on_win)          # Windows
    canvas.bind("<Button-4>",   _on_linux_up)     # Linux scroll up
    canvas.bind("<Button-5>",   _on_linux_dn)     # Linux scroll down
    # 讓子 widget 也能觸發滾輪
    def _bind_children(widget):
        widget.bind("<MouseWheel>", _on_win)
        widget.bind("<Button-4>",   _on_linux_up)
        widget.bind("<Button-5>",   _on_linux_dn)
        for child in widget.winfo_children():
            _bind_children(child)
    canvas.bind("<Enter>", lambda e: _bind_children(canvas))
