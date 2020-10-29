#!/usr/bin/python
# coding:utf-8

"""
@author: NookVoice
@software: PyCharm
@file: table_column_to_excel.py
@time: 2020/10/28 20:56
"""
import logging
import os
import time
import tkinter as tk
from pprint import pprint
from tkinter import ttk, filedialog
from tkinter import simpledialog

import openpyxl as openpyxl


class DIALOG_FILE(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.title("警告!")
        self.geometry('200x100')
        tk.Label(self, text="文件已存在是否覆盖?").grid(row=1)
        tk.Button(self, text="是", command=self.yes).grid(row=2, column=0)
        tk.Button(self, text="否", command=self.no).grid(row=2, column=1)

    def yes(self):
        self.is_cover = True
        self.destroy()

    def no(self):
        self.is_cover = False
        self.destroy()


class DB_INFO():
    def __init__(self, db_type, db_host, db_port, db_name, db_user, db_passwd):
        self.db_type = db_type
        self.db_host = db_host
        self.db_port = db_port
        self.db_name = db_name
        self.db_user = db_user
        self.db_passwd = db_passwd


class Excel_INFO():
    def __init__(self, file_path, file_name, file_type):
        self.file_path = file_path
        self.file_name = file_name
        self.file_type = file_type
        self.file_name_full = file_path + '/' + file_name + file_type


class BASE_DESK(tk.Tk):

    def __init__(self):
        super().__init__()

    def __init__(self, title):
        super().__init__()
        self.title(title)

        # 初始化变量
        self.init_variable()

        # 初始化数据库类型下拉框
        self.init_cmb()

        # 初始化输入框
        self.init_entry()

    def init_variable(self):
        self.db_types = ["PostgreSQL", "MySQL", "Oracle"]
        self.db_type = ''
        self.db_host = tk.StringVar()
        self.db_port = tk.IntVar()
        self.db_database = tk.StringVar()
        self.db_user = tk.StringVar()
        self.db_passwd = tk.StringVar()

        self.file_path = tk.StringVar()
        self.file_name = tk.StringVar()
        self.file_type = ".xlsx"

    def init_cmb(self):
        # 数据库类型

        tk.Label(self, text="数据库类型: ").grid(row=0, column=0)
        self.cmb = ttk.Combobox(self, values=self.db_types)
        self.cmb.grid(row=0, column=1)
        self.cmb.current(0)
        # 选择后触发操作
        self.cmb.bind("<<ComboboxSelected>>", self.chose_db_type)

    def init_entry(self):
        # 数据库host
        tk.Label(self, text="数据库地址: ").grid(row=1, column=0)
        self.et_host = tk.Entry(self, textvariable=self.db_host)
        self.et_host.grid(row=1, column=1)
        # 输入后触发操作
        self.et_host.bind("<<KeyRelease>>", func=self.update_file_name)

        # 数据库port
        tk.Label(self, text="数据库端口: ").grid(row=2, column=0)
        self.et_port = tk.Entry(self, textvariable=self.db_port)
        self.et_port.grid(row=2, column=1)
        # 输入后触发操作
        self.et_port.bind("<<KeyRelease>>", func=self.update_file_name)

        # 数据库名称
        tk.Label(self, text="数据库名称: ").grid(row=3, column=0)
        self.et_database = tk.Entry(self, textvariable=self.db_database)
        self.et_database.grid(row=3, column=1)
        # 输入后触发操作
        self.et_database.bind("<<KeyRelease>>", func=self.update_file_name)

        # 数据库用户
        tk.Label(self, text="数据库用户: ").grid(row=4, column=0)
        self.et_user = tk.Entry(self, textvariable=self.db_user)
        self.et_user.grid(row=4, column=1)
        # 输入后触发操作
        self.et_user.bind("<<KeyRelease>>", func=self.update_file_name)

        # 数据库密码
        tk.Label(self, text="数据库密码: ").grid(row=5, column=0)
        self.et_passwd = tk.Entry(self, textvariable=self.db_passwd, show='*')
        self.et_passwd.grid(row=5, column=1)
        # 输入后触发操作
        self.et_passwd.bind("<<KeyRelease>>", func=self.update_file_name)

        # 导出文件夹
        tk.Label(self, text="导出文件夹: ").grid(row=6, column=0)
        self.et_file_path = tk.Entry(self, textvariable=self.file_path)
        self.et_file_path.grid(row=6, column=1)
        tk.Button(self, text="浏览", command=self.select_path).grid(row=6, column=2)

        # 导出文件名

        tk.Label(self, text="导出文件名: ").grid(row=7, column=0)
        self.et_file_path = tk.Entry(self, textvariable=self.file_name)
        self.et_file_path.grid(row=7, column=1)
        tk.Label(self, text=self.file_type).grid(row=7, column=2)

        # 提交与退出
        tk.Button(self, text="提交", command=self.submit).grid(row=8, column=0)
        tk.Button(self, text="退出", command=self.quit).grid(row=8, column=1)

    def chose_db_type(self, root):
        func_map = {
            "PostgreSQL": self.default_postgres,
            "MySQL": self.default_mysql,
            "Oracle": self.default_oracle
        }

        func = func_map.get(self.cmb['values'][self.cmb.current()])
        func()

    def default_postgres(self):
        self.db_type = "postgres"
        self.db_host.set("localhost")
        self.db_port.set(5432)
        self.db_database.set("postgres")
        self.db_user.set("postgres")
        self.db_passwd.set("")
        self.file_path.set(os.path.join(os.path.expanduser("~"), "Desktop"))
        self.file_name.set("表结构-" + self.db_host.get() + "-" + str(self.db_port.get()) + "-"
                           + self.db_database.get() + "-" + self.db_user.get()
                           + time.strftime('%Y%m%d', time.localtime(time.time())))

    def default_mysql(self):
        self.db_type = "mysql"
        self.db_host.set("localhost")
        self.db_port.set(3306)
        self.db_database.set("mysql")
        self.db_user.set("root")
        self.db_passwd.set("")
        self.file_path.set(os.path.join(os.path.expanduser("~"), "Desktop"))
        self.file_name.set("表结构-" + self.db_host.get() + "-" + str(self.db_port.get()) + "-"
                           + self.db_database.get() + "-" + self.db_user.get()
                           + time.strftime('%Y%m%d', time.localtime(time.time())))

    def default_oracle(self):
        self.db_type = "oracle"
        self.db_host.set("localhost")
        self.db_port.set(1521)
        self.db_database.set("orcl")
        self.db_user.set("sysdba")
        self.db_passwd.set("")
        self.file_path.set(os.path.join(os.path.expanduser("~"), "Desktop"))
        self.file_name.set("表结构-" + self.db_host.get() + "-" + str(self.db_port.get()) + "-"
                           + self.db_database.get() + "-" + self.db_user.get()
                           + time.strftime('%Y%m%d', time.localtime(time.time())))

    def update_file_name(self, root):
        self.file_name.set("表结构-" + self.db_host.get() + "-" + str(self.db_port.get()) + "-"
                           + self.db_database.get() + "-" + self.db_user.get()
                           + time.strftime('%Y%m%d', time.localtime(time.time())))

    def select_path(self):
        self.file_path.set(filedialog.askdirectory())

    def touch_file(self, file_info: Excel_INFO):
        if os.access(file_info.file_name_full, os.F_OK):
            dialog_file = DIALOG_FILE()
            self.wait_window(dialog_file)
            if not dialog_file.is_cover:
                return False

        if file_info.file_type == '.xlsx':
            target_file = openpyxl.Workbook()
            target_file.save(file_info.file_name_full)
            return True

    def submit(self):
        db_info = DB_INFO(db_type=self.db_type, db_host=self.db_host, db_port=self.db_port, db_name=self.db_database
                          , db_user=self.db_user, db_passwd=self.db_passwd)
        excel_info = Excel_INFO(self.file_path.get(), self.file_name.get(), self.file_type)

        # 检查并创建文件
        self.touch_file(excel_info)

        export_to_file(db_info,excel_info)


def get_db_cur(db_info):
    pass


def export_to_file(db_info, excel_info):
    # 获取数据库连接
    cur = get_db_cur(db_info)
    pass

def main():
    base_desk = BASE_DESK("表字段导出Excel工具")

    base_desk.mainloop()


if __name__ == "__main__":
    # 日志等级
    logging.basicConfig(level=logging.INFO)
    main()
