# -*- coding: utf-8 -*-
"""
Catppuccin-Latte 主题 + 全局 Tooltip 演示
运行环境：Windows 11（已安装 ABBvoice CNSG 字体）
"""

import tkinter as tk
from tkinter import ttk

# --------------------------------------------------
# 1. 颜色表（Catppuccin-Latte）
# --------------------------------------------------
CATPPUCCIN = {
    "latte": {
        # 基础
        "base":     "#EFF1F5",
        "mantle":   "#E6E9EF",
        "crust":    "#DCE0E8",
        # 文字
        "text":     "black",
        "subtext0": "#6C6F85",
        "subtext1": "#5C5F77",
        # 强调
        "rosewater":"#DC8A78",
        "flamingo": "#DD7878",
        "pink":     "#EA76CB",
        "mauve":    "#8839EF",
        "red":      "#D20F39",
        "maroon":   "#E64553",
        "peach":    "#FE640B",
        "yellow":   "#DF8E1D",
        "green":    "#40A02B",
        "teal":     "#179299",
        "sky":      "#04A5E5",
        "sapphire": "#209FB5",
        "blue":     "#1E66F5",
        "lavender": "#7287FD",
        # 覆盖层
        "overlay0": "#9CA0B0",
        "overlay1": "#8C8FA1",
        "overlay2": "#7C7F93",
        "surface0": "#CCD0DA",
        "surface1": "#BCC0CC",
        "surface2": "#ACB0BE",
    }
}

# --------------------------------------------------
# 2. 全局 Tooltip（延迟创建 + 单例复用）
# --------------------------------------------------
class Tooltip:
    """跨控件共享的悬浮提示，延迟 350 ms 显示，鼠标离开时隐藏"""
    DELAY = 350                    # 延迟毫秒
    OFFSET = (15, 15)              # 相对鼠标偏移
    _tip = None                    # 类级单例
    _after_id = None               # after() 任务 id

    def __init__(self, widget, text, font=("ABBvoice CNSG", 12)):
        self.widget = widget
        self.text = text
        self.font = font
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)
        widget.bind("<Motion>", self._move)

    # ---------- private ----------
    def _schedule(self, _=None):
        """延迟显示"""
        self._hide()
        self._after_id = self.widget.after(self.DELAY, self._show)

    def _show(self):
        """真正创建窗口"""
        if Tooltip._tip:                # 已存在则仅更新文字
            Tooltip._tip.children["!label"].config(text=self.text)
            return
        x, y = self.widget.winfo_pointerxy()
        x += self.OFFSET[0]
        y += self.OFFSET[1]

        top = Tooltip._tip = tk.Toplevel(self.widget)
        top.wm_overrideredirect(True)   # 无边框
        top.wm_geometry(f"+{x}+{y}")
        top.attributes("-topmost", True)
        top.attributes("-alpha", 0.85)

        ttk.Label(top,
                  text=self.text,
                  font=self.font,
                  foreground="#EFF1F5",
                  background="#8839EF",
                  padding=6).pack()

    def _hide(self, _=None):
        """隐藏并取消延迟任务"""
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None
        if Tooltip._tip:
            Tooltip._tip.destroy()
            Tooltip._tip = None

    def _move(self, ev):
        """跟随鼠标"""
        if Tooltip._tip and Tooltip._tip.winfo_exists():
            Tooltip._tip.wm_geometry(f"+{ev.x_root + self.OFFSET[0]}+{ev.y_root + self.OFFSET[1]}")

# --------------------------------------------------
# 3. 创建并应用主题
# --------------------------------------------------
def apply_latte_theme(style, font_name="ABBvoice CNSG"):
    """一次性生成并切换到 latte 主题"""
    col = CATPPUCCIN["latte"]

    # 常用字体变量
    F  = (font_name, 10)
    FB = (font_name, 10, "bold")
    FL = (font_name, 12, "bold")

    style.theme_create(
        "latte",
        settings={
            # 默认全局
            ".": {"configure": {"background": col["base"],
                                "foreground": col["text"]}},
            # Frame
            "TFrame": {"configure": {"background": col["base"]}},
            # LabelFrame
            "TLabelframe": {"configure": {"background": col["base"],
                                          "relief": tk.FLAT,
                                          "bordercolor": col["overlay1"]}},
            "TLabelframe.Label": {"configure": {"background": col["base"],
                                                "foreground": col["text"],
                                                "font": FL}},
            # Label
            "TLabel": {"configure": {"background": col["base"],
                                     "foreground": col["text"],
                                     "font": F}},
            # Notebook
            "TNotebook": {"configure": {"background": col["base"],
                                        "relief": tk.FLAT}},
            "TNotebook.Tab": {"configure": {"background": col["surface0"],
                                            "foreground": col["overlay2"],
                                            "padding": (10, 5),
                                            "font": FB,
                                            "relief": tk.FLAT},
                              "map": {"background": [("selected", col["lavender"]),
                                                     ("active", col["peach"]),
                                                     ("disabled", col["mantle"])],
                                      "foreground": [("selected", col["base"]),
                                                     ("active", col["base"]),
                                                     ("disabled", col["surface0"])]}},
            # Button
            "TButton": {"configure": {"background": col["base"],
                                      "foreground": col["base"],
                                      "font": FB,
                                      "relief": tk.FLAT},
                        "map": {"background": [("active", col["peach"]),
                                               ("disabled", col["mantle"])],
                                "foreground": [("active", col["base"]),
                                               ("disabled", col["surface0"])]}},
            # Entry
            "TEntry": {"configure": {"fieldbackground": col["crust"],
                                     "foreground": col["blue"],
                                     "selectbackground": col["peach"],
                                     "selectforeground": col["base"],
                                     "relief": tk.FLAT},
                       "map": {"fieldbackground": [("disabled", col["mantle"])],
                               "foreground": [("disabled", col["surface0"])]}},
            # Combobox
            "TCombobox": {"configure": {"fieldbackground": col["crust"],
                                        "background": col["crust"],
                                        "foreground": col["blue"],
                                        "arrowcolor": col["lavender"],
                                        "arrowsize": 25,
                                        "relief": tk.FLAT},
                          "map": {"fieldbackground": [("readonly", col["mantle"]),
                                                      ("disabled", col["mantle"])],
                                  "background": [("readonly", col["mantle"]),
                                                 ("active", col["peach"]),
                                                 ("disabled", col["mantle"])],
                                  "foreground": [("readonly", col["text"]),
                                                 ("disabled", col["surface0"])],
                                  "arrowcolor": [("readonly", col["lavender"]),
                                                 ("disabled", col["surface0"])]}},
            # Treeview
            "Treeview": {"configure": {"background": col["mantle"],
                                       "foreground": col["text"],
                                       "fieldbackground": col["mantle"],
                                       "font": F,
                                       "relief": tk.FLAT},
                         "map": {"background": [("selected", col["lavender"])],
                                 "foreground": [("selected", col["base"])]}},
            # Heading
            "Heading": {"configure": {"background": col["surface0"],
                                      "foreground": col["text"],
                                      "font": FB,
                                      "relief": tk.FLAT}},
            # Progressbar
            "Horizontal.TProgressbar": {"configure": {"background": col["green"],
                                                      "troughcolor": col["crust"]}},
            "Vertical.TProgressbar": {"configure": {"background": col["green"],
                                                    "troughcolor": col["crust"]}},
            # Scale
            "Horizontal.TScale": {"configure": {"background": col["teal"],
                                                "troughcolor": col["crust"],
                                                "sliderlength": 15,
                                                "sliderrelief": tk.FLAT}},
            "Vertical.TScale": {"configure": {"background": col["teal"],
                                              "troughcolor": col["crust"],
                                              "sliderlength": 15,
                                              "sliderrelief": tk.FLAT}},
            # Scrollbar
            "Horizontal.TScrollbar": {"configure": {"background": col["teal"],
                                                    "troughcolor": col["crust"],
                                                    "sliderrelief": tk.FLAT}},
            "Vertical.TScrollbar": {"configure": {"background": col["teal"],
                                                  "troughcolor": col["crust"],
                                                  "sliderrelief": tk.FLAT}},
            # Separator
            "Horizontal.TSeparator": {"configure": {"background": col["overlay0"]}},
            "Vertical.TSeparator": {"configure": {"background": col["overlay0"]}},
        })
    style.theme_use("latte")

    # 下拉列表与菜单的统一样式
    opts = {
        "*TCombobox*Listbox.font": F,
        "*TCombobox*Listbox.background": col["mantle"],
        "*TCombobox*Listbox.foreground": col["text"],
        "*TCombobox*Listbox.selectBackground": col["peach"],
        "*TCombobox*Listbox.selectForeground": col["base"],
        "*Menu.font": F,
        "*Menu.background": col["mantle"],
        "*Menu.foreground": col["text"],
        "*Menu.activeBackground": col["peach"],
        "*Menu.activeForeground": col["base"],
        "*Menu.tearOff": 0,
    }
    for k, v in opts.items():
        style.master.option_add(k, v)

def center_window(win):
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

# --------------------------------------------------
# 4. Demo 入口
# --------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    root.title("Catppuccin-Latte 主题演示")
    root.geometry("500x400")

    center_window(root)
    root.deiconify()
    style = ttk.Style()
    apply_latte_theme(style)          # 应用主题

    # 主框架
    frm = ttk.Frame(root, padding=20)
    frm.pack(fill="both", expand=True)

    # 标签 + Tooltip
    lbl = ttk.Label(frm, text="把鼠标悬停在我身上")
    lbl.pack(pady=10)
    Tooltip(lbl, "这是 Label 的提示")

    # 按钮
    btn = ttk.Button(frm, text="点我试试", command=lambda: print("Hello!"))
    btn.pack(pady=10)
    Tooltip(btn, "这是 Button 的提示")

    # Entry
    ent = ttk.Entry(frm, font=("ABBvoice CNSG", 12))
    ent.pack(pady=10)
    Tooltip(ent, "在此输入文字")

    # Notebook
    nb = ttk.Notebook(frm)
    nb.pack(fill="both", expand=True)
    for tab_name in ("Tab A", "Tab B"):
        tab = ttk.Frame(nb)
        nb.add(tab, text=tab_name)
        ttk.Label(tab, text=f"{tab_name} 内容区").pack(pady=20)

    # 树形视图
    tree = ttk.Treeview(frm, columns=("col1", "col2"), show="headings", height=4)
    tree.heading("col1", text="列 1")
    tree.heading("col2", text="列 2")
    for i in range(3):
        tree.insert("", "end", values=(f"数据 {i+1}-A", f"数据 {i+1}-B"))
    tree.pack(pady=10)
    Tooltip(tree, "Treeview 的提示")

    root.mainloop()