import tkinter as tk
from tkinter import ttk
from ttkstyletool import apply_latte_theme, Tooltip, RoundedButton, RoundedEntry, RoundedFrame, RoundedLabel   # 前面做好的模块

class NavApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("导航栏示例")
        self.geometry("700x450")

        # 1. 应用主题
        style = ttk.Style(self)
        apply_latte_theme(style)

        self.paned = ttk.PanedWindow(self, orient="horizontal")
        self.paned.pack(fill="both", expand=True)

        self.nav = ttk.Frame(self.paned, width=160)
        self.paned.add(self.nav, weight=0)  # 0 = 不拉伸

        self.content = ttk.Frame(self.paned)
        self.paned.add(self.content, weight=1)  # 1 = 占剩余空间

        # 5. Notebook 作为“页面栈”
        self.book = ttk.Notebook(self.content)
        self.book.pack(fill="both", expand=True)

        # 6. 创建若干功能页
        self.pages = {
            "首页": self.make_page("首页内容"),
            "设置": self.make_page("设置内容"),
            "帮助": self.make_page("帮助内容"),
        }

        # 7. 把页添加到 Notebook
        for name, page in self.pages.items():
            self.book.add(page, text=name)   # text 仅用于调试，实际隐藏

        # 8. 左侧导航按钮
        self.build_nav_buttons()

    # ---------- 生成单个页面 ----------
    def make_page(self, text):
        frm = ttk.Frame(self.book)
        ttk.Label(frm, text=text).pack(pady=40)
        return frm

    # ---------- 生成导航按钮 ----------
    def build_nav_buttons(self):
        # 让按钮独占整列
        for idx, name in enumerate(self.pages):
            btn = RoundedButton(
                self.nav,
                text=name,
                command=lambda n=name: self.book.select(self.pages[n])
            )
            btn.pack(fill="x", padx=5, pady=5)
            Tooltip(btn, f"切换到 {name}")

if __name__ == "__main__":
    app = NavApp()
    app.mainloop()