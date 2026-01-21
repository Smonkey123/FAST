# -*- coding: utf-8 -*-
"""
Catppuccin-Latte 主题 + 圆角 Label / Button / Entry / Frame + Tooltip
运行环境：Windows 11（已安装 ABBvoice CNSG 字体）
"""

import tkinter as tk
from tkinter import ttk

# --------------------------------------------------
# 1. 颜色表（Catppuccin-Latte）
# --------------------------------------------------
CATPPUCCIN = {
    "latte": {
        "base":     "#EFF1F5",
        "mantle":   "#E6E9EF",
        "crust":    "#DCE0E8",
        "text":     "#4C4F69",
        "subtext0": "#6C6F85",
        "subtext1": "#5C5F77",
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
        "overlay0": "#9CA0B0",
        "overlay1": "#8C8FA1",
        "overlay2": "#7C7F93",
        "surface0": "#CCD0DA",
        "surface1": "#BCC0CC",
        "surface2": "#ACB0BE",
    }
}

# --------------------------------------------------
# 2. 全局 Tooltip（保持原样）
# --------------------------------------------------
class Tooltip:
    """跨控件共享的悬浮提示，延迟 350 ms 显示，鼠标离开时隐藏"""
    DELAY = 350
    OFFSET = (15, 15)
    _tip = None
    _after_id = None

    def __init__(self, widget, text, font=("ABBvoice CNSG", 12)):
        self.widget = widget
        self.text = text
        self.font = font
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)
        widget.bind("<Motion>", self._move)

    def _schedule(self, _=None):
        self._hide()
        self._after_id = self.widget.after(self.DELAY, self._show)

    def _show(self):
        if Tooltip._tip:
            Tooltip._tip.children["!label"].config(text=self.text)
            return
        x, y = self.widget.winfo_pointerxy()
        x += self.OFFSET[0]
        y += self.OFFSET[1]

        top = Tooltip._tip = tk.Toplevel(self.widget)
        top.wm_overrideredirect(True)
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
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None
        if Tooltip._tip:
            Tooltip._tip.destroy()
            Tooltip._tip = None

    def _move(self, ev):
        if Tooltip._tip and Tooltip._tip.winfo_exists():
            Tooltip._tip.wm_geometry(f"+{ev.x_root + self.OFFSET[0]}+{ev.y_root + self.OFFSET[1]}")

# --------------------------------------------------
# 3. 圆角控件实现（与前面相同）
# --------------------------------------------------
def _round_rect(canvas, x1, y1, x2, y2, r=10, **kw):
    points = (
        x1 + r, y1,
        x2 - r, y1,
        x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2,
        x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r,
        x1, y1 + r, x1, y1
    )
    return canvas.create_polygon(points, smooth=True, **kw)

class RoundedLabel(ttk.Frame):
    def __init__(self, parent, text="", bg=None, fg=None,
                 font=None, radius=100, **kw):
        super().__init__(parent, style="Rounded.TFrame")
        self.configure(padding=0)
        col = CATPPUCCIN["latte"]
        self.bg = bg or col["surface0"]
        self.fg = fg or col["text"]
        self.font = font or ("ABBvoice CNSG", 12)
        self.radius = radius

        self.cnv = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.cnv.pack(fill="both", expand=True, padx=0, pady=0)
        self.text_id = None
        self.set(text)
        self.bind("<Configure>", self._redraw)

    def set(self, text):
        self.text = text
        self._redraw()

    def _redraw(self, *_):
        self.cnv.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return
        _round_rect(self.cnv, 0, 0, w, h, self.radius, fill=self.bg, outline="")
        self.cnv.create_text(w // 2, h // 2, text=self.text,
                             fill=self.fg, font=self.font)

class RoundedButton(ttk.Frame):
    def __init__(self, parent, text="", command=None,
                 bg=None, fg=None, active=None, font=None,
                 radius=100, **kw):
        super().__init__(parent, style="Rounded.TFrame")
        self.configure(padding=0)
        col = CATPPUCCIN["latte"]
        self.bg_normal = bg or col["sapphire"]
        self.bg_active = active or col["peach"]
        self.fg = fg or col["base"]
        self.font = font or ("ABBvoice CNSG", 12, "bold")
        self.radius = radius
        self.command = command

        self.cnv = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.cnv.pack(fill="both", expand=True, padx=0, pady=0)
        self.bind("<Configure>", self._redraw)
        self.cnv.bind("<Button-1>", lambda e: self.command and self.command())
        self.bind("<Enter>", lambda e: self._redraw(True))
        self.bind("<Leave>", lambda e: self._redraw(False))
        self.text = text

    def _redraw(self, hover=False):
        bg = self.bg_active if hover else self.bg_normal
        self.cnv.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return
        _round_rect(self.cnv, 0, 0, w, h, self.radius, fill=bg, outline="")
        self.cnv.create_text(w // 2, h // 2, text=self.text,
                             fill=self.fg, font=self.font)

class RoundedEntry(ttk.Frame):
    def __init__(self, parent, width=20, font=None, radius=10, **kw):
        super().__init__(parent, style="Rounded.TFrame")
        self.configure(padding=0)
        col = CATPPUCCIN["latte"]
        self.bg = col["crust"]
        self.fg = col["blue"]
        self.radius = radius
        self.font = font or ("ABBvoice CNSG", 12)

        self.cnv = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.cnv.pack(fill="both", expand=True, padx=0, pady=0)
        self.entry = tk.Entry(self.cnv, bd=0, highlightthickness=0,
                              bg=self.bg, fg=self.fg, font=self.font,
                              insertbackground=self.fg, **kw)
        self.bind("<Configure>", self._redraw)

    def _redraw(self, *_):
        self.cnv.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return
        _round_rect(self.cnv, 0, 0, w, h, self.radius, fill=self.bg, outline="")
        self.cnv.create_window(w//2, h//2, window=self.entry,
                               width=w-2*self.radius-4, height=h-4)

    def get(self):
        return self.entry.get()

class RoundedFrame(ttk.Frame):
    def __init__(self, parent, bg=None, radius=10, **kw):
        super().__init__(parent, style="Rounded.TFrame")
        self.configure(padding=0)
        col = CATPPUCCIN["latte"]
        self.bg = bg or col["surface0"]
        self.radius = radius

        self.cnv = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        self.cnv.pack(fill="both", expand=True, padx=0, pady=0)
        self.inner = ttk.Frame(self.cnv)
        self.bind("<Configure>", self._redraw)

    def _redraw(self, *_):
        self.cnv.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return
        _round_rect(self.cnv, 0, 0, w, h, self.radius, fill=self.bg, outline="")
        self.cnv.create_window(w//2, h//2, window=self.inner,
                               width=w-2*self.radius-4, height=h-2*self.radius-4)

# --------------------------------------------------
# 4. 主题
# --------------------------------------------------
def apply_latte_theme(style, font_name="ABBvoice CNSG"):
    col = CATPPUCCIN["latte"]
    F  = (font_name, 12)
    FB = (font_name, 12, "bold")
    FL = (font_name, 14, "bold")

    style.theme_create(
        "latte",
        parent="clam",
        settings={
            ".": {"configure": {"background": col["base"],
                                "foreground": col["text"]}},
            "TFrame": {"configure": {"background": col["base"]}},
            "TLabelframe": {"configure": {"background": col["base"],
                                          "relief": tk.FLAT}},
            "TLabelframe.Label": {"configure": {"background": col["base"],
                                                "foreground": col["text"],
                                                "font": FL}},
            "TLabel": {"configure": {"background": col["base"],
                                     "foreground": col["text"],
                                     "font": F,
                                     "relief": tk.FLAT}},
            "TNotebook": {"configure": {"background": col["base"],
                                        "relief": tk.FLAT}},
            "TNotebook.Tab": {"configure": {"background": col["surface0"],
                                            "foreground": col["overlay2"],
                                            "padding": (10, 5),
                                            "font": FB},
                              "map": {"background": [("selected", col["lavender"]),
                                                     ("active", col["peach"])],
                                      "foreground": [("selected", col["base"]),
                                                     ("active", col["base"])]}},
            # Button
            "TButton": {"configure": {"background": col["sapphire"],
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
            "Treeview": {"configure": {"background": col["mantle"],
                                       "foreground": col["text"],
                                       "fieldbackground": col["mantle"],
                                       "font": F},
                         "map": {"background": [("selected", col["lavender"])],
                                 "foreground": [("selected", col["base"])]}},
            "Heading": {"configure": {"background": col["surface0"],
                                      "foreground": col["text"],
                                      "font": FB}},
            # Progressbar
            "Horizontal.TProgressbar": {"configure": {"background": col["green"],
                                                      "troughcolor": col["crust"],
                                                      "borderwidth":0}},
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
    style.configure("Rounded.TFrame", background="")

# --------------------------------------------------
# 5. Demo
# --------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Catppuccin-Latte 圆角控件演示")
    root.geometry("500x480")

    style = ttk.Style()
    apply_latte_theme(style)

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
        tree.insert("", "end", values=(f"数据 {i + 1}-A", f"数据 {i + 1}-B"))
    tree.pack(pady=10)
    Tooltip(tree, "Treeview 的提示")

    # 圆角 Label
    lbl = RoundedLabel(frm, text="圆角 Label")
    lbl.pack(pady=10)
    Tooltip(lbl, "圆角 Label 的提示")

    # 圆角 Button
    btn = RoundedButton(frm, text="圆角 Button",
                        command=lambda: print("Hello!"))
    btn.pack(pady=10)
    Tooltip(btn, "圆角 Button 的提示")

    # 圆角 Entry
    ent = RoundedEntry(frm, width=25)
    ent.pack(pady=10)
    Tooltip(ent, "圆角 Entry 的提示")

    # 圆角 Frame
    rf = RoundedFrame(frm, radius=15, bg="#CCD0DA")
    rf.pack(fill="both", expand=True, pady=10, padx=5)
    ttk.Label(rf.inner, text="这是一个圆角 Frame 内部", font=("ABBvoice CNSG", 10)).pack(pady=15)

    # Notebook
    nb = ttk.Notebook(frm)
    nb.pack(fill="both", expand=True, pady=10)
    for tab_name in ("Tab A", "Tab B"):
        tab = ttk.Frame(nb)
        nb.add(tab, text=tab_name)
        ttk.Label(tab, text=f"{tab_name} 内容区").pack(pady=20)
    print(style.theme_use())
    root.mainloop()