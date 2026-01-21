from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import Treeview
import need.MVE_Pre_Configure as Pre_Configure
import need.CheckList_Export as CheckList_Export
import need.DesignList_Export as DesignList_Export

import need.Usage_State as Usage_State
import need.Comments_Transfer as Comments_Transfer
import need.Terminal_Check as Terminal_Check
import need.Open_Hole as Open_Hole_Check
import need.Panel_Size as Panel_Size_Check
import need.Wiring_Check as Wiring_Check

import need.BOM_Check as BOM_Check
import need.EBOM_to_SAP as EBOM_Import
import need.SAP_to_EBOM as EBOM_Export

import need.Project_Information_Manage as Information_Manage
import need.SAP_Item_View as Item_Info_View

import need.MVE_Panel_Size as Panel_Size_Check2
import need.MVE_CTPT as Param_CTPT_Check

import os, glob, shutil
import sys
import ctypes
from tkinter import Toplevel
import wmi
import winreg
import time
from cryptography.fernet import Fernet
import base64
from PIL import Image, ImageTk
import sqlite3
import hashlib
import csv
import datetime
import logging
import webbrowser
import win32com.client
import re
import psutil
from openpyxl import Workbook

from need.custom_dialogs import CustomDialog, center_window, Tooltip, TvTooltip, image_label

class App:
    def __init__(self):
        self.close_existing_instance()

        global root_window
        self.root = tk.Tk()  # 实例化
        self.root.withdraw()  # 隐藏主窗口，直到许可证窗口被关闭
        root_window = self.root

        # self.root.attributes('-fullscreen', True)
        self.root.configure(bg='white', relief='flat')
        # self.root.tk.call('tk', 'scaling', '-displayof', '.', '1.0')

        global abb_logo_image
        abb_logo_image = tk.PhotoImage(file="ico\\ABB_Logo_Screen_RGB_33px_@1x.png")
        global success_image
        success_image = tk.PhotoImage(file="ico\\abb_check-mark-circle-2_24.png")
        global timeout_image
        timeout_image = tk.PhotoImage(file="ico\\abb_error-circle-2_24.png")
        global wait_image
        wait_image = tk.PhotoImage(file="ico\\abb_hour-glass_24.png")
        global serial_number_image
        serial_number_image = tk.PhotoImage(file="ico\\login_sn.png")
        global progress_image
        progress_image = tk.PhotoImage(file="ico\\login_certificate.png")
        global authorized_person_image
        authorized_person_image = tk.PhotoImage(file="ico\\login_user.png")
        global expiration_date_iamge
        expiration_date_iamge = tk.PhotoImage(file="ico\\login_valid.png")
        global copy_serial_number_image
        copy_serial_number_image = tk.PhotoImage(file="ico\\login_copy.png")
        global send_serial_number_image
        send_serial_number_image = tk.PhotoImage(file="ico\\login_send.png")
        global retry_image
        retry_image = tk.PhotoImage(file="ico\\login_retry.png")
        global login_image
        login_image = tk.PhotoImage(file="ico\\login.png")
        global exit_image
        exit_image = tk.PhotoImage(file="ico\\login_exit.png")
        global login_error_image
        login_error_image = tk.PhotoImage(file="ico\\login_error.png")
        global login_right_image
        login_right_image = tk.PhotoImage(file="ico\\login_right.png")


        global gui_width
        gui_width = 1350
        gui_width = gui_width / 1920.0 * self.root.winfo_screenwidth()

        global gui_height
        gui_height = 900
        gui_height = gui_height / 1080.0 * self.root.winfo_screenheight()
        # print(gui_width, gui_height)

        global w_ratio    # 根据屏幕分辨率计算出的软件界面比例系数，一般情况下是固定值0.625,0.833333334，无论分辨率如何变化
        global h_ratio

        w_ratio = gui_width / self.root.winfo_screenwidth() * 1.0
        h_ratio = gui_height / self.root.winfo_screenheight() * 1.0
        # print(w_ratio, h_ratio)

        w_ratio = 0.7
        h_ratio = 0.84
        global PCR
        PCR = self.root.winfo_screenwidth() / 1920.0

        self.root.geometry("%dx%d" % (gui_width, gui_height))  # 窗体尺寸

        center_window(self.root)  # 将窗体移动到屏幕中央

        self.root.iconbitmap("ico\\logo_new.ico")  # 窗体图标
        self.root.call('tk', 'scaling', 96 / 72.0)  # 设置tkinter的缩放因子，使dpi固定为96
        self.root.title("二次设计辅助工具FAST_V2.2_20260116")
        self.root.resizable(True, True)  # 设置窗体不可改变大小
        self.body()
        # self.previous_f1_height = None  # 添加一个实例变量保存上次f1的高度
        # self.root.bind('<Configure>', self.update_row_height)

        # self.root.protocol('WM_DELETE_WINDOW', self.exit_program)

        # 0. 先只建本地缓存目录
        user = os.getlogin()
        self.local_dir = rf'C:\Users\{user}\temp\FAST_cache'
        os.makedirs(self.local_dir, exist_ok=True)

        self.local_log_path = os.path.join(self.local_dir, 'log.txt')
        self.local_checklog_path = os.path.join(self.local_dir, 'checklog.xlsx')

        # 1. 先让全局指向本地（文件还不一定存在，后面会拉）
        global log_file_path, checklog_file_path
        log_file_path = self.local_log_path
        checklog_file_path = self.local_checklog_path

        # 2. 注册退出回调
        self.root.protocol('WM_DELETE_WINDOW', self._on_exit)

        self.root.attributes("-topmost", True)  # 窗口始终在最前端

        global license_window
        license_window = tk.Toplevel(self.root, bg="#eaf1f6")
        license_window.title('验证FAST许可证')
        license_window.iconbitmap("ico\\license.ico")
        license_window.geometry('500x300+%d+%d' % ((self.root.winfo_screenwidth() - 500) / 2, (self.root.winfo_screenheight() - 300) / 2))
        license_window.resizable(False, False)
        license_window.overrideredirect(True)    # 取消标题栏
        # license_window.attributes("-topmost", True)  # 窗口始终在最前端
        # license_window.protocol('WM_DELETE_WINDOW', lambda: self.on_license_window_close(license_window))  # 控制窗口关闭按钮的行为

        license_f1 = tk.Frame(license_window, bg="#c9dbe9")
        # license_f12 = tk.Frame(license_window, bg="whitesmoke", height=1)
        license_f2 = tk.Frame(license_window, bg="#eaf1f6")
        license_f3 = tk.Frame(license_window, bg="#eaf1f6")
        license_f4 = tk.Frame(license_window, bg="#eaf1f6")
        license_f5 = tk.Frame(license_window, bg="#eaf1f6")
        license_f56 = tk.Frame(license_window, bg="whitesmoke", height=3)
        license_f6 = tk.Frame(license_window, bg="#eaf1f6")
        license_f1.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        # license_f12.pack(side=tk.TOP, fill=tk.X, expand=True)
        license_f2.pack(side=tk.TOP, fill=tk.X, expand=True)
        license_f3.pack(side=tk.TOP, fill=tk.X, expand=True)
        license_f4.pack(side=tk.TOP, fill=tk.X, expand=True)
        license_f5.pack(side=tk.TOP, fill=tk.X, expand=True)
        license_f56.pack(side=tk.TOP, fill=tk.X, expand=True)
        license_f6.pack(side=tk.TOP, fill=tk.X, expand=True, pady=(0,10))

        license_logo_label = tk.Label(license_f1, image=abb_logo_image, bg="#c9dbe9")
        license_logo_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=20)
        license_title_label = tk.Label(license_f1, text='FAST 登录', bg="#c9dbe9", font=("ABBvoice CNSG", int(20 * h_ratio), "bold"))
        license_title_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=70)

        w = wmi.WMI()
        for CS in w.Win32_ComputerSystem():
            device_name = CS.Caption
        # print(device_name)

        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\SQMClient", access=winreg.KEY_READ)
        device_id = winreg.QueryValueEx(key, 'MachineId')[0].replace('{', '').replace('}', '')
        # print(device_id)
        winreg.CloseKey(key)

        global user_account_email
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\Classes\\Local Settings\\Software\\Microsoft\\MSIPC\\aip-addin", access=winreg.KEY_READ)
        user_account_email = winreg.QueryValueEx(reg_key, 'RMSUser')[0]
        # print(user_account_email)
        winreg.CloseKey(reg_key)

        if not os.path.exists('J:/Engineering/ShareFolder/new_ABB_Production_Tools'):
            tk.messagebox.showwarning('提示', '未连接内网，请确保J盘是可访问的，程序将退出...')
            sys.exit()
        else:
            fast_key_file = 'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pf/shfhsak#$hf^sa6saf@#324lczcxzcxz%Ff32gf4321vsha@#!hf.txt'

            if not os.path.exists(fast_key_file):
                tk.messagebox.showwarning('提示', '序列号加密密钥文件不存在')
            else:
                with open(fast_key_file, 'rb') as file:
                    fast_key = file.read()
            global fast_cipher
            fast_cipher = Fernet(fast_key)    # 加密器

            db_key_file = 'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pa/ahsdkjhADAKJ7ad&sadjasg%ashdgjsa%#$FGfhgfhgf.txt'

            if not os.path.exists(db_key_file):
                tk.messagebox.showwarning('提示', '许可证数据库加密密钥文件不存在')
            else:
                with open(db_key_file, 'rb') as file:
                    db_key = file.read()
            global db_cipher
            db_cipher = Fernet(db_key)  # 创建加密器

            license_key_file = 'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pa/809df8h09fdhrgrejlkrejlkgjreljtljre^%cwr#Zhjxcshfsfsfs.txt'

            if not os.path.exists(license_key_file):
                tk.messagebox.showwarning('提示', '许可证文件加密密钥文件不存在')
            else:
                with open(license_key_file, 'rb') as file:
                    license_key = file.read()
            global license_cipher
            license_cipher = Fernet(license_key)  # 创建加密器

            global MAP_FILE
            MAP_FILE = 'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pa/path.csv'

            # tk.Frame(license_f12, height=int(1 * h_ratio), bg="whitesmoke").pack(side=tk.TOP, fill=tk.X)  # 水平分割线
            # tk.Frame(license_f56, height=int(1 * h_ratio), bg="whitesmoke").pack(side=tk.TOP, fill=tk.X)  # 水平分割线

            global fast_serial_number
            fast_serial_number = self.fast_encrypt_data(str(device_name)+str(device_id), fast_cipher, 'f_')
            license_serial_number_label = tk.Label(license_f2, text=' 序列号：', image=serial_number_image, bg="#eaf1f6", compound=tk.LEFT, font=("ABBvoice CNSG", int(12 * h_ratio)))
            license_serial_number_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=(15, 5))

            # license_serial_number_label2 = tk.Label(license_f2, text=fast_serial_number, bg="#eaf1f6", wraplength=450, justify=tk.LEFT)
            license_serial_number_text = tk.Text(license_f2, bg="#eaf1f6", relief='flat', height=1, width=20, font=("ABBvoice CNSG", int(12 * h_ratio)))
            license_serial_number_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            license_serial_number_text.insert(tk.INSERT, fast_serial_number)
            license_serial_number_text.config(state='disabled')
            # license_serial_number_text.config(background="white")
            Tooltip(license_serial_number_text, "电脑硬件序列号")

            license_serial_number_button = tk.Button(license_f2, image=send_serial_number_image, bg="#eaf1f6", cursor='hand2', command=lambda: self.send_serial_number(license_serial_number_text))
            license_serial_number_button.pack(side=tk.RIGHT, fill=tk.BOTH, padx=20)
            # license_serial_number_button.bind('<Enter>', lambda e, name='1':on_enter_for_login_tooltip(e, name))
            # license_serial_number_button.bind('<Leave>', on_level_for_login_tooltip)
            Tooltip(license_serial_number_button, "权限申请/延期")

            license_progress_label = tk.Label(license_f3, text=' 许可证：', image=progress_image, bg="#eaf1f6", compound=tk.LEFT, font=("ABBvoice CNSG", int(12 * h_ratio)))
            license_progress_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=(15, 5))

            progressbar = ttk.Progressbar(license_f3, length=330, mode='determinate')
            progressbar.pack(side=tk.LEFT)

            # progress_result = tk.Label(license_f3, bg="#eaf1f6", image=wait_image, highlightbackground="white", highlightcolor="white", highlightthickness=4)
            # progress_result.pack(side=tk.LEFT, fill=tk.BOTH, padx=15)

            result_label = tk.Label(license_f3, fg='green', bg="#eaf1f6")
            result_label.pack(side=tk.RIGHT, fill=tk.BOTH, padx=20)
            Tooltip(result_label, "许可证查询结果")

            global authorized_person
            authorized_person_label = tk.Label(license_f4, text=' 用户名：', image=authorized_person_image, bg="#eaf1f6", compound=tk.LEFT, font=("ABBvoice CNSG", int(12 * h_ratio)))
            authorized_person_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=(15, 5))
            authorized_person_label2 = tk.Label(license_f4, bg="#eaf1f6", font=("ABBvoice CNSG", int(12 * h_ratio)))
            authorized_person_label2.pack(side=tk.LEFT, fill=tk.BOTH)
            Tooltip(authorized_person_label2, "授权用户邮箱")

            account_valid_label = tk.Label(license_f4, bg="#eaf1f6")
            account_valid_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=20)

            global login_button
            login_button = tk.Button(license_f6, image=login_image, text='进入', compound=tk.LEFT, bg="#eaf1f6", font=("ABBvoice CNSG", int(12 * h_ratio)), cursor='hand2', state='disabled', command=self.enter_fast)
            login_button.pack(side=tk.LEFT, fill=tk.BOTH, padx=(100,0))
            Tooltip(login_button, "登录FAST")

            global retry_button
            retry_button = tk.Button(license_f6, image=retry_image, text='重试', compound=tk.LEFT, bg="#eaf1f6", font=("ABBvoice CNSG", int(12 * h_ratio)), cursor='hand2', state='disabled',
                                     command=lambda: self.re_validation(progressbar, result_label, authorized_person_label2, expiration_date_label2, start_time, date_valid_label, account_valid_label))
            retry_button.pack(side=tk.LEFT, fill=tk.BOTH, padx=80)
            Tooltip(retry_button, "刷新许可证")

            global exit_button
            exit_button = tk.Button(license_f6, image=exit_image, text='退出', font=("ABBvoice CNSG", int(12 * h_ratio)), compound=tk.LEFT, bg="#eaf1f6", cursor='hand2', state='disabled', command=self.exit_program)
            exit_button.pack(side=tk.LEFT, fill=tk.BOTH)
            Tooltip(exit_button, "退出FAST")

            # global entry_or_restart_button
            # entry_or_restart_button = tk.Button(license_f4, bg="#eaf1f6", cursor='hand2')
            # entry_or_restart_button.pack_forget()

            global expiration_date
            expiration_date_label = tk.Label(license_f5, text=' 有效期：', image=expiration_date_iamge, bg="#eaf1f6", compound=tk.LEFT, font=("ABBvoice CNSG", int(12 * h_ratio)))
            expiration_date_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=(15, 5))
            expiration_date_label2 = tk.Label(license_f5, bg="#eaf1f6", font=("ABBvoice CNSG", int(12 * h_ratio)))
            expiration_date_label2.pack(side=tk.LEFT, fill=tk.BOTH)
            Tooltip(expiration_date_label2, "授权有效期")

            date_valid_label = tk.Label(license_f5, bg="#eaf1f6")
            date_valid_label.pack(side=tk.LEFT, fill=tk.BOTH, padx=20)

            # 记录开始验证的时间
            start_time = time.time()
            self.root.after(1000, self.check_license_validation, progressbar, result_label, authorized_person_label2, expiration_date_label2, start_time, date_valid_label, account_valid_label)

            # 让许可证窗口获得焦点，直到它被关闭
            # license_window.focus_force()
            # license_window.grab_set()  # 确保所有事件都发送到许可证窗口

            # self.root.wm_attributes('-disabled', True)    # 禁用主窗口，直到许可证窗口被关闭


    # def on_close(self):
    #     logging.info("Logout from the FAST application")  # 记录日志
    #     self.root.destroy()  # 关闭程序


    def close_existing_instance(self):
        """检查是否有其他实例在运行，如果是，则关闭它"""
        current_pid = os.getpid()
        # current_process_name = psutil.Process(current_pid).name()
        # print(current_pid, current_process_name)

        # 遍历所有正在运行的进程
        for proc in psutil.process_iter(['pid', 'name']):
            # print(proc.info['pid'], proc.info['name'])
            try:
                # 如果找到相同名称的进程，且PID不同，则认为是已有的实例
                if proc.info['name'] == 'FAST_V2.1.exe' and proc.info['pid'] != current_pid:
                    # 终止已有实例
                    # print(f"发现已有实例运行，PID: {proc.info['pid']}，正在关闭...")
                    os.system(f"taskkill /F /PID {proc.info['pid']}")
                    # print(f"已关闭实例 PID: {proc.info['pid']}")
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # 设置进程的DPI感知状态


    def body(self):
        f0 = tk.Frame(self.root, bg="white", relief='flat', borderwidth=0, highlightthickness=2)
        # self.tooltip = None
        # self.setup_ui(f0)
        global home_img
        home_img = tk.PhotoImage(file="ico\\fun_exit.png")
        global help_img
        help_img = tk.PhotoImage(file="ico\\manual.png")
        global feedback_img
        feedback_img = tk.PhotoImage(file="ico\\feedback.png")

        home_label = tk.Label(f0, image=home_img, bg='white')
        home_label.pack(side=tk.LEFT, padx=2, pady=2)
        Tooltip(home_label, "返回功能菜单")
        home_label.bind('<Button-1>', self.back_mainmenu)

        help_label = tk.Label(f0, image=help_img, bg='white')
        help_label.pack(side=tk.LEFT, padx=2, pady=2)
        Tooltip(help_label, "查看说明书")
        help_label.bind('<Button-1>', self.about_help)

        feedback_label = tk.Label(f0, image=feedback_img, bg='white')
        feedback_label.pack(side=tk.LEFT, padx=2, pady=2)
        Tooltip(feedback_label, "软件使用反馈")
        feedback_label.bind('<Button-1>', self.about_feedback)

        Tooltip(f0, "快捷功能栏")
        f0.pack(fill=tk.X)

        global funlist
        # funlist = tk.PhotoImage(file="ico\\abb_folder-open_24.png")
        funlist = tk.PhotoImage(file="ico\\nav_main.png")
        # global funlist_detail
        # funlist_detail = tk.PhotoImage(file="ico\\abb_object_16.png")
        global nav_mve
        nav_mve = tk.PhotoImage(file="ico\\nav_mve.png")
        global nav_designlist
        nav_designlist = tk.PhotoImage(file="ico\\nav_designlist.png")
        global nav_comments
        nav_comments = tk.PhotoImage(file="ico\\nav_comments.png")
        global nav_terminal
        nav_terminal = tk.PhotoImage(file="ico\\nav_terminal.png")
        global nav_lvd
        nav_lvd = tk.PhotoImage(file="ico\\nav_lvd.png")
        global nav_dimension
        nav_dimension = tk.PhotoImage(file="ico\\nav_dimension.png")
        global nav_wiring
        nav_wiring = tk.PhotoImage(file="ico\\nav_wiring.png")
        global nav_bom_check
        nav_bom_check = tk.PhotoImage(file="ico\\nav_bom_check.png")
        global nav_bom_import
        nav_bom_import = tk.PhotoImage(file="ico\\nav_bom_import.png")
        global nav_bom_export
        nav_bom_export = tk.PhotoImage(file="ico\\nav_bom_export.png")
        global nav_project_info
        nav_project_info = tk.PhotoImage(file="ico\\nav_project_info.png")
        global nav_sap_info
        nav_sap_info = tk.PhotoImage(file="ico\\nav_sap_info.png")
        global nav_user_info
        nav_user_info = tk.PhotoImage(file="ico\\nav_user_info.png")

        global f1
        f1 = tk.Frame(self.root, bg="#eaf1f6", bd=0, relief='flat')
        f1.pack(side=tk.LEFT, fill=tk.Y, expand=False, anchor='w')
        global f2
        f2 = tk.Frame(self.root, bg="#eaf1f6", bd=0, relief='flat')
        f2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, anchor='w')

        # 以下内容是Tkinter 8.6.9的BUG解决代码
        def fixed_map(option):  # Fix for setting text colour for Tkinter 8.6.9  , From: https://core.tcl.tk/tk/info/509cafafae
            return [elm
                    for elm in style.map('Treeview', query_opt=option)
                    if elm[:2] != ('!disabled', '!selected')]

        style = ttk.Style()

        style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))
        # style.configure("Custom.Treeview", font=("ABBvoice CNSG", int(13 * h_ratio)))
        # style.configure("Treeview.Heading", font=("ABBvoice CNSG", int(13 * h_ratio)))
        style.configure('nav.Treeview', rowheight=int(70 / (PCR + 1)), font=("ABBvoice CNSG", int(13 * h_ratio)))
        style.configure('nav.Treeview.Heading', font=("ABBvoice CNSG", int(13 * h_ratio)), background="#EFF1F5")

        global app_tree    # 功能导航页
        app_tree = Treeview(f1, show='tree', selectmode='browse', style='nav.Treeview', height=16)
        # Tooltip(app_tree, "功能菜单栏")
        # app_tree.heading('#0', text='导航', anchor='w')
        app_tree.column('#0', width=int(300*w_ratio), anchor='center')

        app_tree.insert('', END, 'Fun1', text=' 配置', image=funlist)
        app_tree.insert('', END, 'Fun2', text=' 设计', image=funlist)
        app_tree.insert('', END, 'Fun3', text=' 物料', image=funlist)
        app_tree.insert('', END, 'Fun4', text=' 信息', image=funlist)
        app_tree.insert('Fun1', END, text=' MVE预配置', image=nav_mve)
        #app_tree.insert('Fun1', END, text='CheckList导出')
        app_tree.insert('Fun1', END, text=' DesignList导出', image=nav_designlist)

        app_tree.insert('Fun2', END, text=' 端子检查', image=nav_terminal)
        app_tree.insert('Fun2', END, text=' 开孔检查', image=nav_lvd)
        app_tree.insert('Fun2', END, text=' 尺寸检查', image=nav_dimension)
        app_tree.insert('Fun2', END, text=' 线号检查', image=nav_wiring)

        app_tree.insert('Fun3', END, text=' BOM检查', image=nav_bom_check)
        app_tree.insert('Fun3', END, text=' BOM导入SAP', image=nav_bom_import)
        app_tree.insert('Fun3', END, text=' SAP中BOM对比', image=nav_sap_info)
        app_tree.insert('Fun3', END, text=' EPLAN中BOM对比', image=nav_bom_export)

        app_tree.insert('Fun4', END, text=' 项目信息管理', image=nav_project_info)
        app_tree.insert('Fun4', END, text=' 图纸意见传递', image=nav_comments)
        app_tree.insert('Fun4', END, text=' 使用情况统计', image=nav_user_info)

        app_tree.bind('<Double-1>', self.open_all_menu)
        app_tree.pack(side=tk.LEFT, fill=tk.Y, expand=True, anchor='w', padx=0)

        app_tree.item('Fun1', open=True)
        app_tree.item('Fun2', open=True)
        app_tree.item('Fun3', open=True)
        app_tree.item('Fun4', open=True)

        tips = {
            'I001': ".mve文件导入MVE软件前的快速读取和配置",
            "I002": "配置和导出项目设计需求表",
            'I003': "基于报表和图纸对柜内端子进行检查",
            'I004': "基于报表和图纸对门板开孔进行检查",
            'I005': "基于报表对柜体尺寸进行检查",
            'I006': "基于报表对原理接线进行检查",
            'I007': "基于报表和图纸对BOM进行检查",
            'I008': "基于报表实现快速导EBOM",
            "I009": "基于SAP对断路器选配、BOM清单进行查阅、对比",
            "I00A": "基于EPLAN和SAP进行配置和物料对比",
            "I00B": "项目基础数据库",
            "I00C": "DD与PM/PE之间传递图纸意见",
            "I00D": "用户使用情况汇总展示",
        }
        TvTooltip(app_tree, tips)

        global app_frame    # 功能内容页
        app_frame = tk.Canvas(f2, width=gui_width, bg="#c9dbe9", height=gui_height)
        app_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, anchor='w', padx=2)

        tk.Label(app_frame, text='欢迎使用二次设计辅助工具', bg="#c9dbe9", fg="black", height=int(2 * h_ratio), font=("ABBvoice CNSG", int(30 * h_ratio), "bold")).pack(side=tk.TOP, expand=True, fill=tk.BOTH)

        # 0. 壁纸缓存（只拉一次）
        self._welcome_pic = None  # 第一次 body() 时真正下载

        pic_path = self._ensure_welcome_pic()
        im = image_label(app_frame, pic_path, 1169, 780, False)
        Tooltip(im, re.search(r'pic1_(.+?)(?:\.[^.]+)?$', os.path.basename(pic_path)).group(1))

        im.configure(bg="#c9dbe9")
        im.pack(side=tk.TOP, expand=True, fill=tk.BOTH)

    # def update_row_height(self, event):
    #     f1_height = f1.winfo_height()
    #     if f1_height != self.previous_f1_height:
    #         self.previous_f1_height = f1_height
    #         row_height = int(55 * (f1_height/gui_height))
    #         style = ttk.Style()
    #         style.configure('Custom.Treeview', rowheight=row_height)

    def exit_program(self):
        self._on_exit()

    def _on_exit(self):
        if not tk.messagebox.askyesno("提示", "确定要退出程序吗？"):
            return
        # flush 日志
        for h in logging.root.handlers:
            h.flush()

        # 如果 authorized_person 还没拿到，说明没登录，直接退
        if not authorized_person:
            self.root.quit()
            return

        else:
            self.remote_log_dir = f'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl2/t/{authorized_person}'
            self.remote_checklog_dir = f'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl2/d/{authorized_person}'

        logging.info("Logout from the FAST application")
        # 回写
        shutil.copy2(self.local_log_path, os.path.join(self.remote_log_dir, 'log.txt'))
        shutil.copy2(self.local_checklog_path, os.path.join(self.remote_checklog_dir, 'checklog.xlsx'))
        self.root.quit()

    def fast_encrypt_data(self, data, data_cipher, prefix):
        encrypted = data_cipher.encrypt(data.encode())
        encoded = base64.urlsafe_b64encode(encrypted).decode().rstrip('=')
        return f'{prefix}{encoded}'

    def fast_decrypt_data(self, data, data_cipher, prefix):
        # 移除列名前缀prefix，如果有的话
        if data.startswith(prefix):
            data = data[2:]

        # Base64 字符串中可能删除了等号，需要添加回原来的填充字符，因为 base64 编码的输出长度应该是 4 的倍数
        padding_needed = 4 - (len(data) % 4)
        if padding_needed:
            data += "=" * padding_needed

        # base64 解码以获取原始的加密字节串
        encrypted_data = base64.urlsafe_b64decode(data.encode())

        # 使用 data_cipher 对象解密，它应该有和加密时一样的密钥
        decrypted_data = data_cipher.decrypt(encrypted_data)

        # 将字节解码回字符串
        decrypted_string = decrypted_data.decode()
        return decrypted_string

    def fast_generate_decrypted_data(self, data, data_cipher, prefix):
        decrypted_data = [self.fast_decrypt_data(col, data_cipher, prefix) for col in data]
        return decrypted_data

    def copy_serial_number(self, license_serial_number_text, root):
        root.clipboard_clear()
        root.clipboard_append(license_serial_number_text.get("1.0", "end-1c"))
        root.update()
        tk.messagebox.showwarning("提示", "软件序列号已复制")
        license_window.update()

    def send_serial_number(self, license_serial_number_text):
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        receiver = 'xiaoqing.gao@cn.abb.com'
        mail.To = receiver

        mail.Subject = 'FAST用户权限申请/延期'
        content = 'Dear ' + re.split(r'[-.@]', receiver)[0].capitalize() + ':\n      请帮忙创建/延期FAST用户权限，我的电脑序列号是：' + license_serial_number_text.get("1.0", "end-1c")
        mail.Body = content
        mail.Importance = 2  # 设置重要性为高
        mail.Send()
        tk.messagebox.showwarning("提示", "权限申请/延期邮件发送成功，请等待邮件通知")

    def enter_fast(self):
        self.root.attributes("-topmost", False)  # 取消窗口始终在最前端
        license_window.destroy()
        root_window.deiconify()

        # log_folder = os.path.join('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl2/t/', '%s' % authorized_person)
        # global checklog_file_path
        # checklog_folder = os.path.join('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl2/d/', '%s' % authorized_person)
        # if not os.path.exists(log_folder):
        #     os.mkdir(log_folder)
        # if not os.path.exists(checklog_folder):
        #     os.mkdir(checklog_folder)
        #
        # # 获取当前时间
        # current_time = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        # # log_file_path = os.path.join(log_folder, 'log_{}.txt'.format(current_time))
        # log_file_path = os.path.join(log_folder, 'log.txt')
        #
        # logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(message)s')
        # logging.info("Login to the FAST application")  # 记录日志

        # 此时 authorized_person 已确定
        self.remote_log_dir = f'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl2/t/{authorized_person}'
        self.remote_checklog_dir = f'J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl2/d/{authorized_person}'

        # 创建远程目录（如果不存在）
        os.makedirs(self.remote_log_dir, exist_ok=True)
        os.makedirs(self.remote_checklog_dir, exist_ok=True)

        # 拉取远程文件到本地
        for remote, local in ((os.path.join(self.remote_log_dir, 'log.txt'), self.local_log_path),
                              (os.path.join(self.remote_checklog_dir, 'checklog.xlsx'), self.local_checklog_path)):
            if os.path.exists(remote):
                shutil.copy2(remote, local)
            elif not os.path.exists(local):
                if remote.endswith('.txt'):
                    open(local, 'w', encoding='utf-8').close()
                else:
                    workbook = Workbook()
                    worksheet = workbook.active

                    # 定义表头
                    headers = ["项目号", "检查功能", "问题条数", "检查时间"]
                    worksheet.append(headers)

                    # 保存表格
                    workbook.save(local)


        # 重新指向本地（保险）
        global log_file_path, checklog_file_path
        log_file_path = self.local_log_path
        checklog_file_path = self.local_checklog_path

        # 初始化 logging（指向本地）
        logging.basicConfig(filename=log_file_path, level=logging.INFO,
                            format='%(asctime)s - %(message)s')
        logging.info("Login to the FAST application")

        # checklog_file_path = os.path.join(checklog_folder, 'checklog.xlsx')
        # if not os.path.exists(checklog_file_path):
        #     # 创建一个新的 Excel 表格
        #     workbook = Workbook()
        #     worksheet = workbook.active
        #
        #     # 定义表头
        #     headers = ["项目号", "检查功能", "问题条数", "检查时间"]
        #     worksheet.append(headers)
        #
        #     # 保存表格
        #     workbook.save(checklog_file_path)

    def re_validation(self, progressbar, result_label, authorized_person_label, expiration_date_label, start_time, date_valid_label, account_valid_label):
        start_time = time.time()
        root_window.after(1000, self.check_license_validation, progressbar, result_label, authorized_person_label, expiration_date_label, start_time, date_valid_label, account_valid_label)

    def check_license_validation(self, progressbar, result_label, authorized_person_label, expiration_date_label, start_time, date_valid_label, account_valid_label):
        global fast_serial_number
        global authorized_person
        global expiration_date

        elapsed_time = time.time() - start_time
        # 检查验证是否成功
        validation_result = self.query_and_validation(fast_serial_number)
        if validation_result:
            # result_label.configure(text="成功", fg="green")
            result_label.configure(image=login_right_image)
            authorized_person_label.configure(text=authorized_person)
            expiration_date_label.configure(text=expiration_date)
            progressbar.stop()
            progressbar["value"] = 100  # 将进度条设置为100%
            # progress_result.configure(image=success_image)

            expiration_date_obj = datetime.datetime.strptime(expiration_date, '%Y-%m-%d')
            current_date = datetime.datetime.now().date()

            if current_date > expiration_date_obj.date():
                exceed_date_flag = 1
            else:
                exceed_date_flag = 0

            global user_account_email
            account_consistency_flag = 0
            if user_account_email.lower() == authorized_person:
                account_consistency_flag = 1

            if not exceed_date_flag and account_consistency_flag:
                # entry_or_restart_button.configure(image=login_image, command=self.enter_fast)
                # entry_or_restart_button.pack(side=tk.RIGHT, fill=tk.BOTH, padx=20)
                login_button.configure(state='normal')
                retry_button.configure(state='normal')
                exit_button.configure(state='normal')
                # account_valid_label.configure(text='√', fg='green')
                account_valid_label.configure(image=login_right_image)
                # date_valid_label.configure(text='√', fg='green')
                date_valid_label.configure(image=login_right_image)
            elif not exceed_date_flag and not account_consistency_flag:
                tk.messagebox.showwarning("提示", "电脑登录用户%s，与注册用户名不一致，请联系管理员注册新用户" % user_account_email)
                login_button.configure(state='disabled')
                retry_button.configure(state='normal')
                exit_button.configure(state='normal')
                # account_valid_label.configure(text='×', fg='red')
                account_valid_label.configure(image=login_error_image)
                # date_valid_label.configure(text='√', fg='green')
                date_valid_label.configure(image=login_right_image)
            elif exceed_date_flag and not account_consistency_flag:
                tk.messagebox.showwarning("提示", "1.电脑登录用户%s，与注册用户名不一致，请联系管理员注册新用户\n\n2.软件许可证已过期，请联系管理员更新授权有效期" % user_account_email)
                login_button.configure(state='disabled')
                retry_button.configure(state='normal')
                exit_button.configure(state='normal')
                # account_valid_label.configure(text='×', fg='red')
                account_valid_label.configure(image=login_error_image)
                # date_valid_label.configure(text='×', fg='red')
                date_valid_label.configure(image=login_error_image)
            else:
                tk.messagebox.showwarning("提示", "软件许可证已过期，请联系管理员更新授权有效期")
                # entry_or_restart_button.configure(image=retry_image, command=lambda: self.re_validation(progressbar, result_label, authorized_person_label, expiration_date_label, start_time, date_valid_label, account_valid_label))
                # entry_or_restart_button.pack(side=tk.RIGHT, fill=tk.BOTH, padx=20)
                login_button.configure(state='disabled')
                retry_button.configure(state='normal')
                exit_button.configure(state='normal')
                # account_valid_label.configure(text='√', fg='green')
                account_valid_label.configure(image=login_right_image)
                # date_valid_label.configure(text='×', fg='red')
                date_valid_label.configure(image=login_error_image)
        else:
            authorized_person = ''
            expiration_date = ''

            # 如果超过60秒仍未成功，则显示验证超时
            if elapsed_time >= 30:
                # result_label.configure(text="超时", fg="red")
                result_label.configure(image=login_error_image)
                # authorized_person_label.configure(text="无许可证")
                # expiration_date_label.configure(text="无许可证")
                authorized_person_label.configure(text=authorized_person)
                expiration_date_label.configure(text=expiration_date)
                progressbar.stop()
                # progress_result.configure(image=timeout_image)
                # entry_or_restart_button.configure(image=retry_image, command=lambda:self.re_validation(progressbar, result_label, authorized_person_label, expiration_date_label, start_time, date_valid_label, account_valid_label))
                # entry_or_restart_button.pack(side=tk.RIGHT, fill=tk.BOTH, padx=20)
                login_button.configure(state='disabled')
                retry_button.configure(state='normal')
                exit_button.configure(state='normal')
                # account_valid_label.configure(text='')
                account_valid_label.configure(image='')
                # date_valid_label.configure(text='')
                date_valid_label.configure(image='')
            else:
                # 更新进度条的完成百分比
                # result_label.configure(text='')
                result_label.configure(image='')
                progress_percentage = (elapsed_time / 30) * 100
                progressbar["value"] = progress_percentage
                # progress_result.configure(image=wait_image)
                # authorized_person_label.configure(text="----------")
                # expiration_date_label.configure(text="----------")
                authorized_person_label.configure(text=authorized_person)
                expiration_date_label.configure(text=expiration_date)
                # 继续定时检查
                self.root.after(1000, self.check_license_validation, progressbar, result_label, authorized_person_label, expiration_date_label, start_time, date_valid_label, account_valid_label)
                # entry_or_restart_button.pack_forget()
                login_button.configure(state='disabled')
                retry_button.configure(state='normal')
                exit_button.configure(state='normal')
                # account_valid_label.configure(text='')
                account_valid_label.configure(image='')
                # date_valid_label.configure(text='')
                date_valid_label.configure(image='')

    def query_and_validation(self, data):
        global authorized_person
        global expiration_date

        if not os.path.exists('J:/Engineering/ShareFolder/new_ABB_Production_Tools'):
            tk.messagebox.showwarning('提示', '请连接内网')
        else:
            if is_folder_hidden('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl/d') or os.path.exists('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl/d'):
                if not os.path.exists('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl/d/l.db'):
                    tk.messagebox.showwarning('提示', 'Licese数据库不存在')
                else:
                    with sqlite3.connect('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl/d/l.db') as conn:
                        cursor = conn.cursor()
                        try:
                            cursor.execute('SELECT * FROM customer_encrypted')
                            result = cursor.fetchall()

                            data_exist_flag = 0
                            for i in range(0, len(result)):
                                temp_list = self.fast_generate_decrypted_data(result[i], db_cipher, 'c_')

                                if self.fast_decrypt_data(temp_list[4], fast_cipher, 'f_') == self.fast_decrypt_data(data, fast_cipher, 'f_'):
                                    authorized_person = temp_list[2]
                                    expiration_date = temp_list[5]
                                    # print(expiration_date)

                                    if len(os.listdir('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl/u/')) == 0:
                                        tk.messagebox.showwarning('提示', 'License文件夹下的所有文件夹及下层许可证文件丢失')
                                    else:
                                        for j in os.listdir('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl/u/'):
                                            path_map = load_path_map(MAP_FILE)
                                            original_path = find_original_path(j, path_map)
                                            if self.fast_decrypt_data(self.fast_decrypt_data(original_path, license_cipher, 'x_'), db_cipher, 'c_') == authorized_person:
                                                for k in os.listdir(os.path.join('J:/Engineering/ShareFolder/new_ABB_Production_Tools/Pl/u/', j)):
                                                    path_map = load_path_map(MAP_FILE)
                                                    original_path = find_original_path(k, path_map)
                                                    # print(original_path)
                                                    # print(self.fast_decrypt_data(original_path, license_cipher, 'x_'))
                                                    # print(self.fast_decrypt_data(self.fast_decrypt_data(original_path, license_cipher, 'x_'), db_cipher, 'c_'))
                                                    # print((self.fast_decrypt_data(self.fast_decrypt_data(original_path, license_cipher, 'x_'), db_cipher, 'c_')).split(expiration_date))
                                                    # print(self.fast_decrypt_data((self.fast_decrypt_data(self.fast_decrypt_data(original_path, license_cipher, 'x_'), db_cipher, 'c_')).split(expiration_date)[0], fast_cipher, 'f_'))
                                                    if self.fast_decrypt_data((self.fast_decrypt_data(self.fast_decrypt_data(original_path, license_cipher, 'x_'), db_cipher, 'c_')).split(expiration_date)[0], fast_cipher, 'f_') == self.fast_decrypt_data(data, fast_cipher, 'f_'):
                                                        data_exist_flag = 1
                                                        return True

                            if not data_exist_flag:
                                authorized_person = ''
                                expiration_date = ''
                                return False


                        except sqlite3.Error as e:
                            tk.messagebox.showwarning('提示', '发生错误：%s' % e)

    # def on_license_window_close(self, license_window):
    #     # 用户尝试关闭许可证窗口时的行为
    #     tk.messagebox.showwarning("提示", "请先验证许可证")
    #     # 保持许可证窗口，不进行关闭
    #     license_window.focus_force()  # 让许可证窗口获得焦点

    def open_all_menu(self, event):
        global checklog_file_path

        tree = app_tree
        column_t = tree.identify_column(event.x)  # 点击的列列号,#0(不显示),#1,#2
        row_t = tree.identify_row(event.y)  # 点击的行行号,I001,I002,I003

        # print(row_t, column_t, type(row_t), type(column_t))     #         #0 <class 'str'> <class 'str'>
                                                                  #   Fun1 #0 <class 'str'> <class 'str'>
                                                                  #   I003 #2 <class 'str'> <class 'str'>
        # print(tree.item('Fun1')['open'])    # {'text': 'Fun1相关功能', 'image': '', 'values': '', 'open': 0, 'tags': ''}
        for widget in app_frame.winfo_children():
            widget.destroy()    # 删除功能页的元素，重新创建

        if row_t == '' and column_t == '#0':    # 点击列标题，进行各行项目全局展开/合并
            # 当单元全开时，点击就全关；非全开，点击就全开
            if tree.item('Fun1')['open'] == 1 and tree.item('Fun2')['open'] == 1 and tree.item('Fun3')['open'] == 1 and tree.item('Fun4')['open'] == 1:
                tree.item('Fun1', open=not tree.item('Fun1')['open'])
                tree.item('Fun2', open=not tree.item('Fun2')['open'])
                tree.item('Fun3', open=not tree.item('Fun3')['open'])
                tree.item('Fun4', open=not tree.item('Fun4')['open'])
            else:
                tree.item('Fun1', open=True)
                tree.item('Fun2', open=True)
                tree.item('Fun3', open=True)
                tree.item('Fun4', open=True)
            tk.Label(app_frame, text='欢迎使用二次设计辅助工具', bg="white", fg="black", height=int(2*h_ratio), font=("ABBvoice CNSG", int(30 * h_ratio), "bold")).pack(side=tk.TOP, expand=True, fill=tk.BOTH)

            pic_path = self._ensure_welcome_pic()
            im = image_label(app_frame, pic_path, 1169, 780, False)
            Tooltip(im, re.search(r'pic1_(.+?)(?:\.[^.]+)?$', os.path.basename(pic_path)).group(1))

            im.configure(bg="#c9dbe9")
            im.pack(side=tk.TOP, expand=True, fill=tk.BOTH)

        else:
            global f1_visiable
            if row_t == 'I001' and column_t == '#0':  # MVE预配置功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: MVE Pre-configure")  # 记录日志
                Pre_Configure.main(app_frame, self.root, w_ratio, h_ratio)

            # elif row_t == 'I002' and column_t == '#0':  # CheckList导出功能
            #     if f1_visiable:
            #         f1.pack_forget()
            #         f1_visiable = False
            #
            #     CheckList_Export.main(app_frame, w_ratio, h_ratio)

            elif row_t == 'I002' and column_t == '#0':  # DesignList导出功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: DesignList Export")
                DesignList_Export.main(app_frame, w_ratio, h_ratio)

            elif row_t == 'I003' and column_t == '#0':  # 端子检查功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: Terminal Check")

                Terminal_Check.main(app_frame, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I004' and column_t == '#0':    # 开孔检查功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: LVD Check")

                Open_Hole_Check.main(app_frame, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I005' and column_t == '#0':    # 尺寸检查功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: Panel Size Check")

                Panel_Size_Check.main(app_frame, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I006' and column_t == '#0':    # 线号检查功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: Wiring Check")

                Wiring_Check.main(app_frame, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I007' and column_t == '#0':    # P/EBOM检查功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: BOM Check")

                BOM_Check.main(app_frame, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I008' and column_t == '#0':    # EBOM导入功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: BOM Import")
                EBOM_Import.main(app_frame, self.root, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I009' and column_t == '#0':    # 柜型查询功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: Item Information View")

                Item_Info_View.main(app_frame, self.root, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I00A' and column_t == '#0':    # EBOM导出功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: BOM Export")
                EBOM_Export.main(app_frame, w_ratio, h_ratio, checklog_file_path)

            elif row_t == 'I00B' and column_t == '#0':    # 信息管理功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: Project Information Manage")
                Information_Manage.main(app_frame, w_ratio, h_ratio)

            elif row_t == 'I00C' and column_t == '#0':  # 意见传递功能
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: Comments Transfer")
                Comments_Transfer.main(app_frame, w_ratio, h_ratio)

            elif row_t == 'I00D' and column_t == '#0':
                if f1_visiable:
                    f1.pack_forget()
                    f1_visiable = False
                logging.info("Access the subfunction: Usage State")
                Usage_State.main(app_frame, w_ratio, h_ratio)

            else:
                tk.Label(app_frame, text='欢迎你，工程师', bg="#c9dbe9", fg="black", height=int(2*h_ratio), font=("ABBvoice CNSG", int(30 * h_ratio), "bold")).pack(side=tk.TOP, expand=True, fill=tk.BOTH)

                pic_path = self._ensure_welcome_pic()
                im = image_label(app_frame, pic_path, 1169, 780, False)
                Tooltip(im, re.search(r'pic1_(.+?)(?:\.[^.]+)?$', os.path.basename(pic_path)).group(1))

                im.configure(bg="#c9dbe9")
                im.pack(side=tk.TOP, expand=True, fill=tk.BOTH)

    # def setup_ui(self, f0):
    #     self.labels = []
    #     self.tooltips_text = {
    #         'home_label': '返回主菜单',
    #         'help_label': '帮助',
    #         'feedback_label': '反馈',
    #     }
    #
    #     global home_img
    #     home_img = tk.PhotoImage(file="ico\\fun_exit.png")
    #     global help_img
    #     help_img = tk.PhotoImage(file="ico\\manual.png")
    #     global feedback_img
    #     feedback_img = tk.PhotoImage(file="ico\\feedback.png")
    #
    #     label_text = {
    #         'home_label': home_img,
    #         'help_label': help_img,
    #         'feedback_label': feedback_img,
    #     }
    #
    #     label_actions = {
    #         'home_label': self.back_mainmenu,
    #         'help_label': self.about_help,
    #         'feedback_label': self.about_feedback,
    #     }
    #     for label_name, img in label_text.items():
    #         label = tk.Label(f0, image=img, bg='white')
    #         label.image = img  # 保持对图像的引用
    #         label.pack(side=tk.LEFT, padx=2, pady=2)
    #         label.bind('<Enter>', lambda e, name=label_name: self.on_enter(e, name))
    #         label.bind('<Leave>', self.on_leave)
    #
    #         if label_name in label_actions:
    #             label.bind('<Button-1>', label_actions[label_name])
    #
    #         self.labels.append((f0, label))

    # def create_tooltip(self, widget, text):
    #     if self.tooltip:
    #         self.tooltip.destroy()
    #
    #     tooltip = Toplevel(widget)
    #     tooltip.wm_overrideredirect(True)
    #     tooltip.wm_geometry("+%d+%d" % (widget.winfo_rootx() + 20, widget.winfo_rooty() + 20))
    #
    #     label = Label(tooltip, text=text, bg="lightyellow", fg="black", relief="solid", bd=1)
    #     label.pack()
    #
    #     self.tooltip = tooltip
    #
    # def on_enter(self, event, label_name):
    #     event.widget['bg'] = 'lightblue'
    #     text = self.tooltips_text.get(label_name, "No tooltip available")
    #     self.create_tooltip(event.widget, text)
    #
    # def on_leave(self, event):
    #     event.widget['bg'] = 'white'
    #     if self.tooltip:
    #         self.tooltip.destroy()
    #         self.tooltip = None

    def back_mainmenu(self, event):
        # start = time.time()
        # root_window.state("normal")
        logging.info("Back to main menu")
        global f1_visiable
        if not f1_visiable:
            for widget in app_frame.winfo_children():
                widget.destroy()  # 删除功能页的元素，重新创建
            f2.pack_forget()
            f1.pack(side=tk.LEFT, fill=tk.Y)

            # app_tree.configure(style='Custom.Treeview')

            f2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, anchor='w')
            f1_visiable = True

            tree = app_tree
            tree.item('Fun1', open=True)
            tree.item('Fun2', open=True)
            tree.item('Fun3', open=True)
            tree.item('Fun4', open=True)

            tk.Label(app_frame, text='欢迎使用二次设计辅助工具', bg="#c9dbe9", fg="black", height=int(2 * h_ratio), font=("ABBvoice CNSG", int(30 * h_ratio), "bold")).pack(side=tk.TOP, expand=True, fill=tk.BOTH)

            pic_path = self._ensure_welcome_pic()
            im = image_label(app_frame, pic_path, 1169, 780, False)
            Tooltip(im, re.search(r'pic1_(.+?)(?:\.[^.]+)?$', os.path.basename(pic_path)).group(1))

            im.configure(bg="#c9dbe9")
            im.pack(side=tk.TOP, expand=True, fill=tk.BOTH)

        # end = time.time()
        # print(str(end-start))

    def _ensure_welcome_pic(self) -> str:
        """
        保证返回一张本地存在的 welcome 壁纸；
        若缓存路径丢失会重新拉取。
        """
        if self._welcome_pic and os.path.isfile(self._welcome_pic):
            return self._welcome_pic

        # 缓存为空或文件被删 → 重新拉
        self._welcome_pic = cache_remote_pic(r'J:\Engineering\ShareFolder\new_ABB_Production_Tools\ico')
        return self._welcome_pic

    def about_help(self, event):
        os.startfile(os.path.abspath('J:\\Engineering\\ShareFolder\\new_ABB_Production_Tools\\Pd\\document\\二次设计辅助工具使用说明书V2.2.pdf'))

    def about_feedback(self, event):
        webbrowser.open("https://abb.sharepoint.com/:x:/r/sites/CNDMXSWGProjectPMSE/_layouts/15/Doc.aspx?sourcedoc=%7B4CF78210-B528-443A-A14A-4A5450AF965C%7D&file=FAST%E4%BD%BF%E7%94%A8%E5%8F%8D%E9%A6%88%E6%94%B6%E9%9B%86%E8%A1%A8.xlsx&action=default&mobileredirect=true")

def is_folder_hidden(fpath):
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(fpath)  # attrs值为18表示该文件夹具有以下属性组合：只读 (1)、隐藏 (2) 和 子文件夹 (16)
        # print(attrs)
        if attrs != -1 and attrs & 2 == 2:  # 对于18（二进制为10010）与2（二进制为00010）进行按位与运算，结果为2（二进制为00010）
            return True
    except OSError:
        pass
    return False


def load_path_map(map_file):
    path_map = {}
    if os.path.exists(map_file):
        with open(map_file, 'r') as file:
            reader = csv.reader(file)
            for short_path, long_path in reader:
                path_map[short_path] = long_path
    return path_map


def find_original_path(short_path, path_map):
    # 根据短路径名查找原始的长路径名
    return path_map.get(short_path)

def cache_remote_pic(remote_dir: str, pattern: str = '*pic1_*', suffixes=('.jpg', '.png')) -> str:
    # 1. 找图
    for ext in suffixes:
        files = glob.glob(os.path.join(remote_dir, pattern + ext))
        if files:
            src = files[0]
            break
    else:
        raise FileNotFoundError(f'目录下找不到文件名包含 {pattern} 的图片')

    # 2. 准备本地目录
    local_dir = rf'C:\Users\{os.getlogin()}\temp'
    os.makedirs(local_dir, exist_ok=True)

    # 3. 复制
    dst = os.path.join(local_dir, os.path.basename(src))
    shutil.copy2(src, dst)
    return dst


if __name__ == "__main__":
    try:
        import pyi_splash

        pyi_splash.close()
    except ImportError:
        pass

    global f1_visiable
    f1_visiable = True
    app = App()
    app.root.mainloop()



