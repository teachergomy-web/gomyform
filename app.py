"""
app.py — 高米補習班兒美成績管理系統
入口點，建立主視窗與側欄導覽
"""
import tkinter as tk
from tkinter import ttk
import data_manager as dm
import ui_utils as U
import pages
import sys
import os
import ctypes

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        try:
            myappid = 'gomy.grade.system.v2'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        self.title("高米補習班 兒美成績管理系統")

        self.data = dm.load()

        # 套用深色/淺色模式
        dark = self.data.get("settings", {}).get("dark_mode", False)
        U.apply_theme(dark)

        self._set_icon()
        self.geometry("1200x780")
        self.minsize(1000, 650)
        self.configure(bg=U.C["bg"])
        self._build()
        self._show("overview")

    def _set_icon(self):
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        icon_p = os.path.join(base, 'gomyscore.ico')
        if os.path.exists(icon_p):
            self.iconbitmap(default=icon_p)

    def _build(self):
        self.sidebar = tk.Frame(self, bg=U.C["sidebar"], width=178)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = tk.Frame(self.sidebar, bg=U.C["sidebar"])
        brand.pack(fill="x", pady=(16, 12), padx=14)
        tk.Label(brand, text="高米補習班", font=(U.F, 9),
                 bg=U.C["sidebar"], fg="#8BBAD4").pack(anchor="w")
        tk.Label(brand, text="兒美成績系統", font=(U.F, 13, "bold"),
                 bg=U.C["sidebar"], fg="white").pack(anchor="w")

        tk.Frame(self.sidebar, bg="#2A6090", height=1).pack(fill="x", padx=12)

        self._nav_btns = {}
        nav_items = [
            ("📝  兒美成績輸入",  "overview"),
            ("🏆  升級考成績輸入", "leveltest"),
            ("👥  班級／學生",    "classes"),
            ("📚  用書管理",      "books"),
            ("📄  輸出成績單",    "export"),
            ("⚙️  系統設定",      "settings"),
        ]
        for label, key in nav_items:
            btn = tk.Button(
                self.sidebar, text=label, bg=U.C["sidebar"], fg="#B8D4EC",
                font=(U.F, 11), relief="flat", cursor="hand2",
                anchor="w", padx=16, pady=11,
                activebackground=U.C["sidebar_sel"], activeforeground="white",
                command=lambda k=key: self._show(k)
            )
            btn.pack(fill="x")
            self._nav_btns[key] = btn

        self.content = tk.Frame(self, bg=U.C["bg"])
        self.content.pack(side="right", fill="both", expand=True)

        self.page_map = {
            "overview":  pages.OverviewPage(self.content, self),
            "classes":   pages.ClassesPage(self.content, self),
            "books":     pages.BooksPage(self.content, self),
            "monthly":   pages.MonthlyInputPage(self.content, self),
            "leveltest": pages.LevelTestPage(self.content, self),
            "export":    pages.ExportPage(self.content, self),
            "settings":  pages.SettingsPage(self.content, self),
        }

    def _show(self, key):
        for k, btn in self._nav_btns.items():
            if k == key:
                btn.configure(bg=U.C["sidebar_sel"], fg="white",
                              font=(U.F, 11, "bold"))
            else:
                btn.configure(bg=U.C["sidebar"], fg="#B8D4EC",
                              font=(U.F, 11))
        for page in self.page_map.values():
            page.pack_forget()
        self.page_map[key].pack(fill="both", expand=True)
        if hasattr(self.page_map[key], "refresh"):
            self.page_map[key].refresh()

    def save(self):
        dm.save(self.data)


if __name__ == "__main__":
    mutex_name = "GomyGrade_Unique_Instance_Mutex"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    if ctypes.windll.kernel32.GetLastError() == 183:
        hwnd = ctypes.windll.user32.FindWindowW(None, "高米補習班 兒美成績管理系統")
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 9)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
        sys.exit()
    app = App()
    app.mainloop()
