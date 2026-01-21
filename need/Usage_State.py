import tkinter as tk
from tkinter import ttk
import sqlite3
from tkinter.ttk import Treeview, Style
import shutil, os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import warnings
import pandas as pd
warnings.simplefilter(action='ignore', category=FutureWarning)
from need.custom_dialogs import CustomDialog, center_window, Tooltip, image_label
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
prop = fm.FontProperties(fname=r"C:\Windows\Fonts\ABBvoice_CNSG_Rg.ttf")
plt.rcParams['font.family'] = prop.get_name()
plt.rcParams['axes.unicode_minus'] = False
from need.export_html import export_html

def main(parent, w_ratio, h_ratio):
    global image_export_report
    image_export_report = tk.PhotoImage(file="ico\\export.png")

    # ---------- 1. 准备本地目录 ----------
    local_ylim_dir = r'C:\Temp\FAST_log\ylim'
    local_ylim_file = os.path.join(local_ylim_dir, 'ylim.txt')
    os.makedirs(local_ylim_dir, exist_ok=True)

    # ---------- 2. 拷贝远程文件 ----------
    remote = r'J:\Engineering\ShareFolder\new_ABB_Production_Tools\Pd\ylim\ylim.txt'
    if os.path.isfile(remote):
        try:
            shutil.copy2(remote, local_ylim_file)
        except Exception:
            pass  # 拷贝失败也不管，后面用默认值

    # ---------- 3. 默认值表（与原代码保持一致） ----------
    default_ylim = {
        'ax2': 120,
        'ax3': 200,
        'ax4': 100,
        'ax5': 100,
        'ax6': 500,
        'ax7': 500,
        'ax8': 200,
        'ax9': 200,
        'ax10': 300
    }

    # ---------- 4. 读取本地 ylim.txt ----------
    ylim_dict = default_ylim.copy()
    if os.path.isfile(local_ylim_file):
        try:
            with open(local_ylim_file, 'r', encoding='utf-8') as f:
                lines = [ln.strip() for ln in f if ln.strip()]
            # 只认前 9 行
            for idx, ln in enumerate(lines[:9], 2):
                key = f'ax{idx}'
                if ':' in ln:
                    _, val = ln.split(':', 1)
                    val = val.strip()
                    if val:  # 冒号后有值
                        try:
                            ylim_dict[key] = float(val)
                        except ValueError:
                            pass  # 转换失败就用默认值
        except Exception:
            pass  # 任何异常都忽略，保持默认值


    global image_total_user
    image_total_user = tk.PhotoImage(file="ico\\usage_total_user_amount.png")
    global image_active_user
    image_active_user = tk.PhotoImage(file="ico\\usage_active_user_amount.png")
    global image_deactive_user
    image_deactive_user = tk.PhotoImage(file="ico\\usage_deactive_user_amount.png")
    global image_user_ratio
    image_user_ratio = tk.PhotoImage(file="ico\\usage_ratio_month.png")
    global image_designlist_amount
    image_designlist_amount = tk.PhotoImage(file="ico\\usage_designlist_amount.png")
    global image_predesign
    image_predesign = tk.PhotoImage(file="ico\\usage_predesign.png")
    global image_schdesign
    image_schdesign = tk.PhotoImage(file="ico\\usage_schdesign.png")
    global image_projectdesign
    image_projectdesign = tk.PhotoImage(file="ico\\usage_projectdesign.png")
    global image_project_amount
    image_project_amount = tk.PhotoImage(file="ico\\usage_project_amount_month.png")
    global image_error_amount
    image_error_amount = tk.PhotoImage(file="ico\\usage_error_amount.png")
    global image_top3_item
    image_top3_item = tk.PhotoImage(file="ico\\usage_top3_project.png")
    global image_top1_reason
    image_top1_reason = tk.PhotoImage(file="ico\\usage_top1_reason.png")

    global canvas
    canvas = tk.Canvas(parent, width=int(1750 * w_ratio), height=int(640 * h_ratio), bg="#C9DBE9")
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    canvas.update()
    canvas.bind("<MouseWheel>", on_mousewheel)

    scrollbar_v = tk.Scrollbar(master=parent)
    scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
    scrollbar_v.config(command=canvas.yview)
    canvas.config(yscrollcommand=scrollbar_v.set)

    global content
    content = tk.Frame(canvas)
    # content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    canvas.create_window(0, 1, width=int(1750 * w_ratio), anchor=tk.NW, window=content)

    f1 = tk.Frame(content, bg="#c9dbe9", bd=0)
    tk.Label(f1, text="欢迎查看用户使用情况", bg="#c9dbe9", fg="black", height=int(1 * h_ratio), font=("ABBvoice CNSG", int(20 * h_ratio), "bold")).pack(fill=tk.X)
    f1.pack(fill=tk.X)

    now = datetime.now()
    months = [(now.month + i - 1) % 12 + 1 for i in range(1, 13)]

    # *****************************获取数据*********************************
    global data
    data = get_monthly_data()
    # print(data)

    # *****************************用户概述*********************************
    f2 = tk.LabelFrame(content, text='用户概述', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f2_l = tk.Frame(f2, bg="#eaf1f6", bd=0)
    f2_r = tk.Frame(f2, bg="#eaf1f6", bd=0)

    f2_l_1 = tk.Frame(f2_l, bg="#eaf1f6", bd=0)
    f2_l_2 = tk.Frame(f2_l, bg="#eaf1f6", bd=0)
    f2_l_3 = tk.Frame(f2_l, bg="#eaf1f6", bd=0)
    f2_l_4 = tk.Frame(f2_l, bg="#eaf1f6", bd=0)
    f2_l_1.pack(side='top', fill=tk.X, expand=True)
    f2_l_2.pack(side='top', fill=tk.X, expand=True)
    f2_l_3.pack(side='top', fill=tk.X, expand=True)
    f2_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f2_l_1, text='用户总数：', image=image_total_user, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global user_amount
    user_amount = tk.Label(f2_l_1, bg="#eaf1f6", text=data['total_users'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    user_amount.pack(side=tk.LEFT, fill=tk.X)
    export_button = tk.Button(f2_l_1, image=image_export_report, command=export_report, bg="#eaf1f6", compound='left')
    export_button.pack(side='right', fill=tk.X)
    Tooltip(export_button, '导出报告')

    tk.Label(f2_l_2, text='活跃用户：', image=image_active_user, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global active_user_amount
    active_user_amount = tk.Label(f2_l_2, bg="#eaf1f6", text=data['active_users'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    active_user_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f2_l_3, text='非活用户：', image=image_deactive_user, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global deactive_user_amount
    deactive_user_amount = tk.Label(f2_l_3, bg="#eaf1f6", text=data['inactive_users'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    deactive_user_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f2_l_4, text='月使用率：', image=image_user_ratio, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global user_ratio
    user_ratio_value = data['usage_rate'][11]
    usage_percentage = f'{user_ratio_value:.2%}'
    user_ratio = tk.Label(f2_l_4, bg="#eaf1f6", text=usage_percentage, fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    user_ratio.pack(side=tk.LEFT, fill=tk.X)

    f2_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f2_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f2.pack(fill=tk.X)
    Tooltip(f2_r, "历史数据")
    Tooltip(f2_l, "当月数据")

    fig1, ax1 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig1.patch.set_facecolor('#eaf1f6')
    ax1.set_facecolor("#eaf1f6")

    global user_ratio_list
    user_ratio_list = [round(i*100, 2) for i in data['usage_rate']]

    bars = ax1.bar(months, user_ratio_list, color="#4f81bd", width=0.6)
    ax1.set_xlabel("月份", fontsize=10)
    ax1.set_ylabel("月使用率 (%)", fontsize=10)
    ax1.set_ylim(0, 100)
    ax1.set_xticks(months)

    ax1.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, user_ratio_list):
        ax1.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas1 = FigureCanvasTkAgg(fig1, master=f2_r)
    canvas1.draw()
    canvas1.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************设计传递表*********************************
    f3 = tk.LabelFrame(content, text='设计传递表', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f3_l = tk.Frame(f3, bg="#eaf1f6", bd=0)
    f3_r = tk.Frame(f3, bg="#eaf1f6", bd=0)

    f3_l_1 = tk.Frame(f3_l, bg="#eaf1f6", bd=0)
    f3_l_2 = tk.Frame(f3_l, bg="#eaf1f6", bd=0)
    f3_l_3 = tk.Frame(f3_l, bg="#eaf1f6", bd=0)
    f3_l_4 = tk.Frame(f3_l, bg="#eaf1f6", bd=0)
    f3_l_1.pack(side='top', fill=tk.X, expand=True)
    f3_l_2.pack(side='top', fill=tk.X, expand=True)
    f3_l_3.pack(side='top', fill=tk.X, expand=True)
    f3_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f3_l_1, text='表单总数：', image=image_designlist_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global designlist_amount
    designlist_amount = tk.Label(f3_l_1, bg="#eaf1f6", text=data['total_forms'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    designlist_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f3_l_2, text='提前设计：', image=image_predesign, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global pre_design_amount
    pre_design_amount = tk.Label(f3_l_2, bg="#eaf1f6", text=data['pre_design_forms'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    pre_design_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f3_l_3, text='图纸设计：', image=image_schdesign, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global sch_design_amount
    sch_design_amount = tk.Label(f3_l_3, bg="#eaf1f6", text=data['drawing_design_forms'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    sch_design_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f3_l_4, text='工程设计：', image=image_projectdesign, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global pro_design_amount
    pro_design_amount = tk.Label(f3_l_4, bg="#eaf1f6", text=data['eng_design_forms'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    pro_design_amount.pack(side=tk.LEFT, fill=tk.X)

    f3_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f3_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f3.pack(fill=tk.X)
    Tooltip(f3_r, "历史数据")
    Tooltip(f3_l, "当月数据")

    fig2, ax2 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig2.patch.set_facecolor('#eaf1f6')
    ax2.set_facecolor("#eaf1f6")
    global designlist_amount_every_month
    designlist_amount_every_month = data['total_forms']

    bars = ax2.bar(months, designlist_amount_every_month, color="#4f81bd", width=0.6)
    ax2.set_xlabel("月份", fontsize=10)
    ax2.set_ylabel("表单数量", fontsize=10)
    ax2.set_ylim(0, ylim_dict['ax2'])
    ax2.set_xticks(months)
    ax2.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, designlist_amount_every_month):
        ax2.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas2 = FigureCanvasTkAgg(fig2, master=f3_r)
    canvas2.draw()
    canvas2.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************端子检查*********************************
    f4 = tk.LabelFrame(content, text='端子检查', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f4_l = tk.Frame(f4, bg="#eaf1f6", bd=0)
    f4_r = tk.Frame(f4, bg="#eaf1f6", bd=0)

    f4_l_1 = tk.Frame(f4_l, bg="#eaf1f6", bd=0)
    f4_l_2 = tk.Frame(f4_l, bg="#eaf1f6", bd=0)
    f4_l_3 = tk.Frame(f4_l, bg="#eaf1f6", bd=0)
    f4_l_4 = tk.Frame(f4_l, bg="#eaf1f6", bd=0)
    f4_l_1.pack(side='top', fill=tk.X, expand=True)
    f4_l_2.pack(side='top', fill=tk.X, expand=True)
    f4_l_3.pack(side='top', fill=tk.X, expand=True)
    f4_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f4_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global terminal_check_project_amount
    terminal_check_project_amount = tk.Label(f4_l_1, bg="#eaf1f6", text=data['terminal_projects'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    terminal_check_project_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f4_l_2, text='错误条数：', image=image_error_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global terminal_check_perproject_error
    terminal_check_perproject_error = tk.Label(f4_l_2, bg="#eaf1f6", text=str(data['terminal_errors'][11]) + '/项目', fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    terminal_check_perproject_error.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f4_l_3, text='Top3项目：', image=image_top3_item, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global terminal_check_top_3_project_number
    top3_text = '\n'.join(eval(data['terminal_top3_projects'][11]))
    terminal_check_top_3_project_number = tk.Label(f4_l_3, bg="#eaf1f6", text=top3_text, fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    terminal_check_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f4_l_4, text='Top1原因：', image=image_top1_reason, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global terminal_check_top_1_reason
    terminal_check_top_1_reason = tk.Label(f4_l_4, bg="#eaf1f6", text=data['terminal_top1_reason'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    terminal_check_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f4_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f4_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f4.pack(fill=tk.X)
    Tooltip(f4_r, "历史数据")
    Tooltip(f4_l, "当月数据")

    fig3, ax3 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig3.patch.set_facecolor('#eaf1f6')
    ax3.set_facecolor("#eaf1f6")
    global terminal_check_perproject_error_every_month
    terminal_check_perproject_error_every_month = data['terminal_errors']

    bars = ax3.bar(months, terminal_check_perproject_error_every_month, color="#4f81bd", width=0.6)
    ax3.set_xlabel("月份", fontsize=10)
    ax3.set_ylabel("平均每项目错误数", fontsize=10)
    ax3.set_ylim(0, ylim_dict['ax3'])
    ax3.set_xticks(months)
    ax3.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, terminal_check_perproject_error_every_month):
        ax3.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas3 = FigureCanvasTkAgg(fig3, master=f4_r)
    canvas3.draw()
    canvas3.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************开孔检查*********************************
    f5 = tk.LabelFrame(content, text='开孔检查', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f5_l = tk.Frame(f5, bg="#eaf1f6", bd=0)
    f5_r = tk.Frame(f5, bg="#eaf1f6", bd=0)

    f5_l_1 = tk.Frame(f5_l, bg="#eaf1f6", bd=0)
    f5_l_2 = tk.Frame(f5_l, bg="#eaf1f6", bd=0)
    f5_l_3 = tk.Frame(f5_l, bg="#eaf1f6", bd=0)
    f5_l_4 = tk.Frame(f5_l, bg="#eaf1f6", bd=0)
    f5_l_1.pack(side='top', fill=tk.X, expand=True)
    f5_l_2.pack(side='top', fill=tk.X, expand=True)
    f5_l_3.pack(side='top', fill=tk.X, expand=True)
    f5_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f5_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global lvd_layout_check_project_amount
    lvd_layout_check_project_amount = tk.Label(f5_l_1, bg="#eaf1f6", text=data['hole_projects'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    lvd_layout_check_project_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f5_l_2, text='错误条数：', image=image_error_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global lvd_layout_check_perproject_error
    lvd_layout_check_perproject_error = tk.Label(f5_l_2, bg="#eaf1f6", text=str(data['hole_errors'][11]) + '/项目', fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    lvd_layout_check_perproject_error.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f5_l_3, text='Top3项目：', image=image_top3_item, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global lvd_layout_check_top_3_project_number
    top3_text = '\n'.join(eval(data['hole_top3_projects'][11]))
    lvd_layout_check_top_3_project_number = tk.Label(f5_l_3, bg="#eaf1f6", text=top3_text, fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    lvd_layout_check_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)

    # tk.Label(f5_l_4, text='Top1原因：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global terminal_check_top_1_reason
    # terminal_check_top_1_reason = tk.Label(f5_l_4, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # terminal_check_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f5_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f5_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f5.pack(fill=tk.X)
    Tooltip(f5_r, "历史数据")
    Tooltip(f5_l, "当月数据")

    fig4, ax4 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig4.patch.set_facecolor('#eaf1f6')
    ax4.set_facecolor("#eaf1f6")
    global lvd_layout_check_perproject_error_every_month
    lvd_layout_check_perproject_error_every_month = data['hole_errors']

    bars = ax4.bar(months, lvd_layout_check_perproject_error_every_month, color="#4f81bd", width=0.6)
    ax4.set_xlabel("月份", fontsize=10)
    ax4.set_ylabel("平均每项目错误数", fontsize=10)
    ax4.set_ylim(0, ylim_dict['ax4'])
    ax4.set_xticks(months)
    ax4.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, lvd_layout_check_perproject_error_every_month):
        ax4.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas4 = FigureCanvasTkAgg(fig4, master=f5_r)
    canvas4.draw()
    canvas4.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************尺寸检查*********************************
    f6 = tk.LabelFrame(content, text='尺寸检查', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f6_l = tk.Frame(f6, bg="#eaf1f6", bd=0)
    f6_r = tk.Frame(f6, bg="#eaf1f6", bd=0)

    f6_l_1 = tk.Frame(f6_l, bg="#eaf1f6", bd=0)
    f6_l_2 = tk.Frame(f6_l, bg="#eaf1f6", bd=0)
    f6_l_3 = tk.Frame(f6_l, bg="#eaf1f6", bd=0)
    f6_l_4 = tk.Frame(f6_l, bg="#eaf1f6", bd=0)
    f6_l_1.pack(side='top', fill=tk.X, expand=True)
    f6_l_2.pack(side='top', fill=tk.X, expand=True)
    f6_l_3.pack(side='top', fill=tk.X, expand=True)
    f6_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f6_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global panel_size_check_project_amount
    panel_size_check_project_amount = tk.Label(f6_l_1, bg="#eaf1f6", text=data['size_projects'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    panel_size_check_project_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f6_l_2, text='错误条数：', image=image_error_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global panel_size_check_perproject_error
    panel_size_check_perproject_error = tk.Label(f6_l_2, bg="#eaf1f6", text=str(data['size_errors'][11]) + '/项目', fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    panel_size_check_perproject_error.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f6_l_3, text='Top3项目：', image=image_top3_item, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global panel_size_check_top_3_project_number
    top3_text = '\n'.join(eval(data['size_top3_projects'][11]))
    panel_size_check_top_3_project_number = tk.Label(f6_l_3, bg="#eaf1f6", text=top3_text, fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    panel_size_check_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)

    # tk.Label(f6_l_4, text='Top1原因：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global terminal_check_top_1_reason
    # terminal_check_top_1_reason = tk.Label(f6_l_4, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # terminal_check_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f6_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f6_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f6.pack(fill=tk.X)
    Tooltip(f6_r, "历史数据")
    Tooltip(f6_l, "当月数据")

    fig5, ax5 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig5.patch.set_facecolor('#eaf1f6')
    ax5.set_facecolor("#eaf1f6")
    global panel_size_check_perproject_error_every_month
    panel_size_check_perproject_error_every_month = data['size_errors']

    bars = ax5.bar(months, panel_size_check_perproject_error_every_month, color="#4f81bd", width=0.6)
    ax5.set_xlabel("月份", fontsize=10)
    ax5.set_ylabel("平均每项目错误数", fontsize=10)
    ax5.set_ylim(0, ylim_dict['ax5'])
    ax5.set_xticks(months)
    ax5.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, panel_size_check_perproject_error_every_month):
        ax5.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas5 = FigureCanvasTkAgg(fig5, master=f6_r)
    canvas5.draw()
    canvas5.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************线号检查*********************************
    f7 = tk.LabelFrame(content, text='线号检查', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f7_l = tk.Frame(f7, bg="#eaf1f6", bd=0)
    f7_r = tk.Frame(f7, bg="#eaf1f6", bd=0)

    f7_l_1 = tk.Frame(f7_l, bg="#eaf1f6", bd=0)
    f7_l_2 = tk.Frame(f7_l, bg="#eaf1f6", bd=0)
    f7_l_3 = tk.Frame(f7_l, bg="#eaf1f6", bd=0)
    f7_l_4 = tk.Frame(f7_l, bg="#eaf1f6", bd=0)
    f7_l_1.pack(side='top', fill=tk.X, expand=True)
    f7_l_2.pack(side='top', fill=tk.X, expand=True)
    f7_l_3.pack(side='top', fill=tk.X, expand=True)
    f7_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f7_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global wiring_check_project_amount
    wiring_check_project_amount = tk.Label(f7_l_1, bg="#eaf1f6", text=data['wire_projects'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    wiring_check_project_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f7_l_2, text='错误条数：', image=image_error_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global wiring_check_perproject_error
    wiring_check_perproject_error = tk.Label(f7_l_2, bg="#eaf1f6", text=str(data['wire_errors'][11]) + '/项目', fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    wiring_check_perproject_error.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f7_l_3, text='Top3项目：', image=image_top3_item, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global wiring_check_top_3_project_number
    top3_text = '\n'.join(eval(data['wire_top3_projects'][11]))
    wiring_check_top_3_project_number = tk.Label(f7_l_3, bg="#eaf1f6", text=top3_text, fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    wiring_check_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f7_l_4, text='Top1原因：', image=image_top1_reason, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global wiring_check_top_1_reason
    wiring_check_top_1_reason = tk.Label(f7_l_4, bg="#eaf1f6", text=data['wire_top1_reason'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    wiring_check_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f7_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f7_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f7.pack(fill=tk.X)
    Tooltip(f7_r, "历史数据")
    Tooltip(f7_l, "当月数据")

    fig6, ax6 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig6.patch.set_facecolor('#eaf1f6')
    ax6.set_facecolor("#eaf1f6")
    global wiring_check_perproject_error_every_month
    wiring_check_perproject_error_every_month = data['wire_errors']

    bars = ax6.bar(months, wiring_check_perproject_error_every_month, color="#4f81bd", width=0.6)
    ax6.set_xlabel("月份", fontsize=10)
    ax6.set_ylabel("平均每项目错误数", fontsize=10)
    ax6.set_ylim(0, ylim_dict['ax6'])
    ax6.set_xticks(months)
    ax6.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, wiring_check_perproject_error_every_month):
        ax6.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas6 = FigureCanvasTkAgg(fig6, master=f7_r)
    canvas6.draw()
    canvas6.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************BOM检查*********************************
    f8 = tk.LabelFrame(content, text='BOM检查', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f8_l = tk.Frame(f8, bg="#eaf1f6", bd=0)
    f8_r = tk.Frame(f8, bg="#eaf1f6", bd=0)

    f8_l_1 = tk.Frame(f8_l, bg="#eaf1f6", bd=0)
    f8_l_2 = tk.Frame(f8_l, bg="#eaf1f6", bd=0)
    f8_l_3 = tk.Frame(f8_l, bg="#eaf1f6", bd=0)
    f8_l_4 = tk.Frame(f8_l, bg="#eaf1f6", bd=0)
    f8_l_1.pack(side='top', fill=tk.X, expand=True)
    f8_l_2.pack(side='top', fill=tk.X, expand=True)
    f8_l_3.pack(side='top', fill=tk.X, expand=True)
    f8_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f8_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global bom_check_project_amount
    bom_check_project_amount = tk.Label(f8_l_1, bg="#eaf1f6", text=data['bom_projects'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    bom_check_project_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f8_l_2, text='错误条数：', image=image_error_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global bom_check_perproject_error
    bom_check_perproject_error = tk.Label(f8_l_2, bg="#eaf1f6", text=str(data['bom_errors'][11])+'/项目', fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    bom_check_perproject_error.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f8_l_3, text='Top3项目：', image=image_top3_item, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global bom_check_top_3_project_number
    top3_text = '\n'.join(eval(data['bom_top3_projects'][11]))
    bom_check_top_3_project_number = tk.Label(f8_l_3, bg="#eaf1f6", text=top3_text, fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    bom_check_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f8_l_4, text='Top1原因：', image=image_top1_reason, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global bom_check_top_1_reason

    def split_string(s):
        if len(s) > 22:
            return "\n".join([s[i:i + 22] for i in range(0, len(s), 22)])
        return s

    bom_check_top_1_reason = tk.Label(f8_l_4, bg="#eaf1f6", text=split_string(data['bom_top1_reason'][11]), fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    bom_check_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f8_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f8_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f8.pack(fill=tk.X)
    Tooltip(f8_r, "历史数据")
    Tooltip(f8_l, "当月数据")

    fig7, ax7 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig7.patch.set_facecolor('#eaf1f6')
    ax7.set_facecolor("#eaf1f6")
    global bom_check_perproject_error_every_month
    bom_check_perproject_error_every_month = data['bom_errors']

    bars = ax7.bar(months, bom_check_perproject_error_every_month, color="#4f81bd", width=0.6)
    ax7.set_xlabel("月份", fontsize=10)
    ax7.set_ylabel("平均每项目错误数", fontsize=10)
    ax7.set_ylim(0, ylim_dict['ax7'])
    ax7.set_xticks(months)
    ax7.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, bom_check_perproject_error_every_month):
        ax7.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas7 = FigureCanvasTkAgg(fig7, master=f8_r)
    canvas7.draw()
    canvas7.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************BOM导入SAP*********************************
    f9 = tk.LabelFrame(content, text='BOM导入SAP', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f9_l = tk.Frame(f9, bg="#eaf1f6", bd=0)
    f9_r = tk.Frame(f9, bg="#eaf1f6", bd=0)

    f9_l_1 = tk.Frame(f9_l, bg="#eaf1f6", bd=0)
    f9_l_2 = tk.Frame(f9_l, bg="#eaf1f6", bd=0)
    f9_l_3 = tk.Frame(f9_l, bg="#eaf1f6", bd=0)
    f9_l_4 = tk.Frame(f9_l, bg="#eaf1f6", bd=0)
    f9_l_1.pack(side='top', fill=tk.X, expand=True)
    f9_l_2.pack(side='top', fill=tk.X, expand=True)
    f9_l_3.pack(side='top', fill=tk.X, expand=True)
    f9_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f9_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global bom_import_project_amount
    bom_import_project_amount = tk.Label(f9_l_1, bg="#eaf1f6", text=data['bom_import_sap'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    bom_import_project_amount.pack(side=tk.LEFT, fill=tk.X)

    # tk.Label(f9_l_2, text='错误条数：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global bom_import_perproject_error
    # bom_import_perproject_error = tk.Label(f9_l_2, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # bom_import_perproject_error.pack(side=tk.LEFT, fill=tk.X)
    #
    # tk.Label(f9_l_3, text='Top3项目：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global bom_import_top_3_project_number
    # bom_import_top_3_project_number = tk.Label(f9_l_3, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # bom_import_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)
    #
    # tk.Label(f9_l_4, text='Top1原因：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global bom_import_top_1_reason
    # bom_import_top_1_reason = tk.Label(f9_l_4, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # bom_import_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f9_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f9_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f9.pack(fill=tk.X)
    Tooltip(f9_r, "历史数据")
    Tooltip(f9_l, "当月数据")

    fig8, ax8 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig8.patch.set_facecolor('#eaf1f6')
    ax8.set_facecolor("#eaf1f6")
    global bom_import_project_amount_every_month
    bom_import_project_amount_every_month = data['bom_import_sap']

    bars = ax8.bar(months, bom_import_project_amount_every_month, color="#4f81bd", width=0.6)
    ax8.set_xlabel("月份", fontsize=10)
    ax8.set_ylabel("导入项目数", fontsize=10)
    ax8.set_ylim(0, ylim_dict['ax8'])
    ax8.set_xticks(months)
    ax8.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, bom_import_project_amount_every_month):
        ax8.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas8 = FigureCanvasTkAgg(fig8, master=f9_r)
    canvas8.draw()
    canvas8.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************SAP中BOM对比*********************************
    f10 = tk.LabelFrame(content, text='SAP中BOM对比', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f10_l = tk.Frame(f10, bg="#eaf1f6", bd=0)
    f10_r = tk.Frame(f10, bg="#eaf1f6", bd=0)

    f10_l_1 = tk.Frame(f10_l, bg="#eaf1f6", bd=0)
    f10_l_2 = tk.Frame(f10_l, bg="#eaf1f6", bd=0)
    f10_l_3 = tk.Frame(f10_l, bg="#eaf1f6", bd=0)
    f10_l_4 = tk.Frame(f10_l, bg="#eaf1f6", bd=0)
    f10_l_1.pack(side='top', fill=tk.X, expand=True)
    f10_l_2.pack(side='top', fill=tk.X, expand=True)
    f10_l_3.pack(side='top', fill=tk.X, expand=True)
    f10_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f10_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global bom_compare_project_amount
    bom_compare_project_amount = tk.Label(f10_l_1, bg="#eaf1f6", text=data['bom_compare_sap'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    bom_compare_project_amount.pack(side=tk.LEFT, fill=tk.X)

    # tk.Label(f10_l_2, text='错误条数：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global bom_compare_perproject_error
    # bom_compare_perproject_error = tk.Label(f10_l_2, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # bom_compare_perproject_error.pack(side=tk.LEFT, fill=tk.X)
    #
    # tk.Label(f10_l_3, text='Top3项目：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global bom_compare_top_3_project_number
    # bom_compare_top_3_project_number = tk.Label(f10_l_3, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # bom_compare_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)
    #
    # tk.Label(f10_l_4, text='Top1原因：', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    # global bom_compare_top_1_reason
    # bom_compare_top_1_reason = tk.Label(f10_l_4, bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    # bom_compare_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f10_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f10_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f10.pack(fill=tk.X)
    Tooltip(f10_r, "历史数据")
    Tooltip(f10_l, "当月数据")

    fig9, ax9 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig9.patch.set_facecolor('#eaf1f6')
    ax9.set_facecolor("#eaf1f6")
    global bom_compare_project_amount_every_month
    bom_compare_project_amount_every_month = data['bom_compare_sap']

    bars = ax9.bar(months, bom_compare_project_amount_every_month, color="#4f81bd", width=0.6)
    ax9.set_xlabel("月份", fontsize=10)
    ax9.set_ylabel("对比项目数", fontsize=10)
    ax9.set_ylim(0, ylim_dict['ax9'])
    ax9.set_xticks(months)
    ax9.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, bom_compare_project_amount_every_month):
        ax9.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas9 = FigureCanvasTkAgg(fig9, master=f10_r)
    canvas9.draw()
    canvas9.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # *****************************EPLAN中BOM对比*********************************
    f11 = tk.LabelFrame(content, text='EPLAN中BOM对比', font=("ABBvoice CNSG", int(18 * h_ratio)), bg="#eaf1f6", bd=2)
    f11_l = tk.Frame(f11, bg="#eaf1f6", bd=0)
    f11_r = tk.Frame(f11, bg="#eaf1f6", bd=0)

    f11_l_1 = tk.Frame(f11_l, bg="#eaf1f6", bd=0)
    f11_l_2 = tk.Frame(f11_l, bg="#eaf1f6", bd=0)
    f11_l_3 = tk.Frame(f11_l, bg="#eaf1f6", bd=0)
    f11_l_4 = tk.Frame(f11_l, bg="#eaf1f6", bd=0)
    f11_l_1.pack(side='top', fill=tk.X, expand=True)
    f11_l_2.pack(side='top', fill=tk.X, expand=True)
    f11_l_3.pack(side='top', fill=tk.X, expand=True)
    f11_l_4.pack(side='top', fill=tk.X, expand=True)

    tk.Label(f11_l_1, text='月项目数：', image=image_project_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global eplan_bom_compare_project_amount
    eplan_bom_compare_project_amount = tk.Label(f11_l_1, bg="#eaf1f6", text=data['bom_compare_eplan'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    eplan_bom_compare_project_amount.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f11_l_2, text='错误条数：', image=image_error_amount, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global eplan_bom_compare_perproject_error
    eplan_bom_compare_perproject_error = tk.Label(f11_l_2, bg="#eaf1f6", text=str(data['bom_compare_eplan_err'][11])+'/项目', fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    eplan_bom_compare_perproject_error.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f11_l_3, text='Top3项目：', image=image_top3_item, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global eplan_bom_compare_top_3_project_number
    top3_text = '\n'.join(eval(data['bom_compare_eplan_top3'][11]))
    eplan_bom_compare_top_3_project_number = tk.Label(f11_l_3, bg="#eaf1f6", text=top3_text, fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    eplan_bom_compare_top_3_project_number.pack(side=tk.LEFT, fill=tk.X)

    tk.Label(f11_l_4, text='Top1原因：', image=image_top1_reason, compound='left', bg="#eaf1f6", fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left').pack(side=tk.LEFT, fill=tk.X)
    global eplan_bom_compare_top_1_reason
    eplan_bom_compare_top_1_reason = tk.Label(f11_l_4, bg="#eaf1f6", text=data['bom_compare_eplan_top1'][11], fg="black", font=("ABBvoice CNSG", int(18 * h_ratio)), justify='left')
    eplan_bom_compare_top_1_reason.pack(side=tk.LEFT, fill=tk.X)

    f11_r.pack(side='left', fill=tk.Y, padx=(50, 20), pady=5)
    f11_l.pack(side='left', fill=tk.Y, padx=5, pady=5)
    f11.pack(fill=tk.X)
    Tooltip(f11_r, "历史数据")
    Tooltip(f11_l, "当月数据")

    fig10, ax10 = plt.subplots(figsize=(6, 4.2), dpi=100)
    fig10.patch.set_facecolor('#eaf1f6')
    ax10.set_facecolor("#eaf1f6")
    global eplan_bom_compare_perproject_error_every_month
    eplan_bom_compare_perproject_error_every_month = data['bom_compare_eplan_err']

    bars = ax10.bar(months, eplan_bom_compare_perproject_error_every_month, color="#4f81bd", width=0.6)
    ax10.set_xlabel("月份", fontsize=10)
    ax10.set_ylabel("平均每项目错误数", fontsize=10)
    ax10.set_ylim(0, ylim_dict['ax10'])
    ax10.set_xticks(months)
    ax10.set_xticklabels([f"{m}月" for m in months])

    for bar, v in zip(bars, eplan_bom_compare_perproject_error_every_month):
        ax10.text(bar.get_x() + bar.get_width() / 2, v + 1,
                 f'{v}', ha='center', va='bottom', fontsize=8)

    # 把图嵌进 tk
    canvas10 = FigureCanvasTkAgg(fig10, master=f11_r)
    canvas10.draw()
    canvas10.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    canvas.update_idletasks()
    # content.update_idletasks()
    canvas.config(scrollregion=canvas.bbox('all'))

    # 把 10 张图收进 dict，key 必须和模板里一致
    global fig_map
    fig_map = {
        'fig_user_ratio': fig1,
        'fig_designlist': fig2,
        'fig_terminal': fig3,
        'fig_hole': fig4,
        'fig_size': fig5,
        'fig_wire': fig6,
        'fig_bom': fig7,
        'fig_bom_import': fig8,
        'fig_bom_cmp_sap': fig9,
        'fig_bom_cmp_eplan': fig10
    }

def export_report():
    export_html(data, fig_map)

def get_monthly_data():
    # 生成最近 12 个月 yyyy-mm 列表
    today = datetime.today()
    months = [
        (today.replace(day=1) - relativedelta(months=i)).strftime('%Y-%m')
        for i in range(12)
    ][::-1]  # 升序

    REMOTE_DB_DIR = r'J:\Engineering\ShareFolder\new_ABB_Production_Tools\Pl2\db'  # 远端目录
    REMOTE_DB = os.path.join(REMOTE_DB_DIR, 'monthly_stats_v2.db')

    with sqlite3.connect(REMOTE_DB) as conn:
        # 参数占位符
        placeholders = ','.join(['?' for _ in months])
        sql = f'''
            SELECT * 
            FROM monthly_stats
            WHERE year_month IN ({placeholders})
            ORDER BY year_month ASC
        '''
        df = pd.read_sql(sql, conn, params=months)

    return df


def on_mousewheel(event):
    global canvas
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

