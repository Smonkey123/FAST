"""
custom_dialogs.py
在 Tkinter 中替代系统对话框，支持自定义字体、样式、图标。
Author: <xiaoiqng gao>
Date  : 2025-08-04
"""

import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont
from PIL import Image, ImageTk
from typing import List

class CustomDialog:
    ICONS = {"warning": "⚠️", "question": "❓"}

    # ----------------------------------------------------------
    @staticmethod
    def showwarning(title: str,
                    message: str,
                    font_family: str = "ABBvoice CNSG",
                    font_size: int = 10,
                    parent=None):
        lines = break_by_width(message, width=28)   # 关键：按字符宽度 10 切行
        top = _build_top(title, lines, "warning", font_family, font_size, parent)
        tk.Button(top, text="确定", command=top.destroy, font=(font_family, font_size)).pack(pady=10)
        top.wait_window()

    @staticmethod
    def askyesno(title: str,
                 message: str,
                 font_family: str = "ABBvoice CNSG",
                 font_size: int = 10) -> bool:
        lines = break_by_width(message, width=28)
        return _ask_generic(title, lines, "yes_no", font_family, font_size)

    @staticmethod
    def askquestion(title: str,
                    message: str,
                    font_family: str = "ABBvoice CNSG",
                    font_size: int = 10) -> str:
        lines = break_by_width(message, width=28)
        return _ask_generic(title, lines, "question", font_family, font_size)


# ----------------------------------------------------------------------
# 工具：按“字符宽度”切行
def char_width(c: str) -> int:
    """中文字符算 2，其余算 1"""
    return 2 if '\u4e00' <= c <= '\u9fff' else 1

def break_by_width(text: str, width: int = 10) -> List[str]:
    """返回按指定宽度切分后的多行"""
    lines = []
    current_line = ""
    current_width = 0

    for ch in text:
        w = char_width(ch)
        # 如果放一个字符就超宽，则把当前行先存起来
        if current_width + w > width and current_line:
            lines.append(current_line)
            current_line, current_width = ch, w
        else:
            current_line += ch
            current_width += w

    if current_line or not lines:  # 最后一行
        lines.append(current_line)
    return lines


# ----------------------------------------------------------------------
# 内部：创建窗口并返回，尚未放按钮
def _build_top(title, lines, icon_key, font_family, font_size, parent=None):
    top = tk.Toplevel(parent)
    top.withdraw()
    top.title(title)
    top.configure(bg="#EFF1F5")
    top.resizable(False, False)
    top.transient(parent)
    top.grab_set()

    f = tkfont.Font(family=font_family, size=font_size)

    # 图标
    ttk.Label(top, text=CustomDialog.ICONS[icon_key], font=f).pack(pady=8)

    # 消息（已手动换行）
    msg_text = "\n".join(lines)
    msg_lbl = ttk.Label(top, text=msg_text, font=f, justify="left")
    msg_lbl.pack(padx=10, pady=5)

    # 动态计算高度
    line_h = f.metrics("linespace")
    needed_h = line_h * len(lines) + 20  # 20 上下留白
    win_w = 220
    win_h = max(120, needed_h + 110)     # 110 给图标+按钮
    win_h = min(win_h, 450)
    top.geometry(f"{win_w}x{win_h}")

    center_window(top)
    top.deiconify()
    top.focus_set()
    return top


# ----------------------------------------------------------------------
def _ask_generic(title, lines, mode, font_family, font_size):
    result = None
    def _set(val):
        nonlocal result
        result = val
        top.destroy()

    top = _build_top(title, lines,
                     "question" if mode == "question" else "warning",
                     font_family, font_size)

    btn_frame = ttk.Frame(top)
    btn_frame.pack(pady=8)


    if mode == "yes_no":
        tk.Button(btn_frame, text="是", width=6,
                   command=lambda: _set(True), font=(font_family, font_size)).pack(side="left", padx=10)
        tk.Button(btn_frame, text="否", width=6,
                   command=lambda: _set(False), font=(font_family, font_size)).pack(side="left", padx=10)
    else:  # question
        tk.Button(btn_frame, text="是", width=6,
                   command=lambda: _set("yes"), font=(font_family, font_size)).pack(side="left", padx=10)
        tk.Button(btn_frame, text="否", width=6,
                   command=lambda: _set("no"), font=(font_family, font_size)).pack(side="left", padx=10)

    top.wait_window()
    return result


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

class TvTooltip:
    """Treeview 行悬停提示"""
    DELAY = 350
    OFFSET = (15, 15)
    _tip = None
    _after_id = None

    def __init__(self, tree: ttk.Treeview, row2text: dict):
        self.tree = tree
        self.row2text = row2text
        # 只绑在 Treeview 自己身上
        tree.bind('<Motion>', self._schedule)
        tree.bind('<Leave>', self._hide)

    # ---------- 内部 ----------
    def _schedule(self, ev: tk.Event):
        self._hide()
        row = self.tree.identify_row(ev.y)          # 当前行 iid
        if row in self.row2text:                    # 这行有提示
            self._after_id = self.tree.after(
                self.DELAY, lambda: self._show(ev, row))

    def _show(self, ev: tk.Event, row):
        if self._tip:                               # 已存在则更新
            self._tip.children['!label'].config(text=self.row2text[row])
            return
        x, y = ev.x_root + self.OFFSET[0], ev.y_root + self.OFFSET[1]
        top = self._tip = tk.Toplevel(self.tree)
        top.wm_overrideredirect(True)
        top.geometry(f'+{x}+{y}')
        top.attributes('-topmost', True)
        top.attributes('-alpha', 0.85)
        ttk.Label(top,
                  text=self.row2text[row],
                  font=('ABBvoice CNSG', 12),
                  foreground='#EFF1F5',
                  background='#8839EF',
                  padding=6).pack()

    def _hide(self, _=None):
        if self._after_id:
            self.tree.after_cancel(self._after_id)
            self._after_id = None
        if self._tip:
            self._tip.destroy()
            self._tip = None

def center_window(win):
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

def image_label(frame, img, width, height, keep_ratio=True):
    """输入图片信息，及尺寸，返回界面组件"""
    if isinstance(img, str):
        _img = Image.open(img)
    else:
        _img = img
    lbl_image = tk.Label(frame, width=width, height=height)
    tk_img = tkimg_resized(_img, width, height, keep_ratio)
    lbl_image.image = tk_img
    lbl_image.config(image=tk_img)
    return lbl_image

def tkimg_resized(img, w_box, h_box, keep_ratio=True):
    """对图片进行按比例缩放处理"""
    w, h = img.size
    if keep_ratio:
        if w > h:
            width = w_box
            height = int(h_box * (1.0 * h / w))
        if h >= w:
            height = h_box
            width = int(w_box * (1.0 * w / h))
    else:
        width = w_box
        height = h_box
    img1 = img.resize((width, height), Image.ANTIALIAS)
    tkimg = ImageTk.PhotoImage(img1)
    return tkimg

# ------------------------------------------------------------------
# 演示（运行此文件时会执行）
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.geometry("500x400")
    center_window(root)
    root.deiconify()

    # 主框架
    frm = tk.Frame(root)
    frm.pack(fill="both", expand=True)

    # 标签 + Tooltip
    lbl = tk.Label(frm, text="把鼠标悬停在我身上")
    lbl.pack(pady=10)
    Tooltip(lbl, "这是 Label 的提示")

    # 按钮
    btn = tk.Button(frm, text="点我试试", command=lambda: print("Hello!"))
    btn.pack(pady=10)
    Tooltip(btn, "这是 Button 的提示")

    # Entry
    ent = tk.Entry(frm, font=("ABBvoice CNSG", 12))
    ent.pack(pady=10)
    Tooltip(ent, "在此输入文字")

    # Notebook
    nb = ttk.Notebook(frm)
    nb.pack(fill="both", expand=True)
    for tab_name in ("Tab A", "Tab B"):
        tab = tk.Frame(nb)
        nb.add(tab, text=tab_name)
        tk.Label(tab, text=f"{tab_name} 内容区").pack(pady=20)

    # 树形视图
    tree = ttk.Treeview(frm, columns=("col1", "col2"), show="headings", height=4)
    tree.heading("col1", text="列 1")
    tree.heading("col2", text="列 2")
    for i in range(3):
        tree.insert("", "end", values=(f"数据 {i+1}-A", f"数据 {i+1}-B"))
    tree.pack(pady=10)
    Tooltip(tree, "Treeview 的提示")

    # 1. 警告框
    CustomDialog.showwarning("文件未保存", "请先保存safdskfhjhkjahsfkdjsahkajskdhajshdjkasasahdkjhsakjdhkjasdhjhsakjhdjkahskjdhjshdkjahdkjh撒旦撒打算的撒打啊哒哒哒哒哒哒四大撒地方撒发斯蒂芬后再退出！")

    # 2. 是/否
    if CustomDialog.askyesno("确认退出？", "真的阿达撒范德萨发萨芬萨芬萨法撒范德萨发5456456456456要退出吗？"):
        print("用户选择了「是」")
    else:
        print("用户选择了「否」")

    # 3. askquestion 风格
    ans = CustomDialog.askquestion("继续？", "是否继续操sadsadsadsadsa阿德撒旦撒旦撒旦撒作？")
    print("askquestion 返回:", ans)

    root.mainloop()