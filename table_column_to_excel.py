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
from tkinter import ttk, filedialog


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
        self.db_type = ''
        self.db_host = tk.StringVar()
        self.db_port = tk.IntVar()
        self.db_database = tk.StringVar()
        self.db_user = tk.StringVar()
        self.db_passwd = tk.StringVar()

        self.file_path = tk.StringVar()
        self.file_name = tk.StringVar()

    def init_cmb(self):
        # 数据库类型
        values = ["PostgreSQL", "MySQL", "Oracle"]
        tk.Label(self, text="数据库类型: ").grid(row=0, column=0)
        self.cmb = ttk.Combobox(self, values=values)
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
        self.et_passwd = tk.Entry(self, textvariable=self.db_passwd,show='*')
        self.et_passwd.grid(row=5, column=1)
        # 输入后触发操作
        self.et_passwd.bind("<<KeyRelease>>", func=self.update_file_name)

        # 导出文件夹
        tk.Label(self, text="导出文件夹: ").grid(row=6, column=0)
        self.et_file_path = tk.Entry(self, textvariable=self.file_path)
        self.et_file_path.grid(row=6, column=1)
        tk.Button(self,text="浏览",command= self.select_path).grid(row=6, column=2)

        # 导出文件名
        file_type = ".xlsx"
        tk.Label(self, text="导出文件名: ").grid(row=7, column=0)
        self.et_file_path = tk.Entry(self, textvariable=self.file_name)
        self.et_file_path.grid(row=7, column=1)
        tk.Label(self, text=file_type).grid(row=7, column=2)

        # 提交与退出
        tk.Button(self, text="提交",command= self.submit).grid(row=8, column=0)
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


    def submit(self):
        print("提交")


def main():
    base_desk = BASE_DESK("表字段导出Excel工具")

    base_desk.mainloop()


if __name__ == "__main__":
    # 日志等级
    logging.basicConfig(level=logging.INFO)
    main()
