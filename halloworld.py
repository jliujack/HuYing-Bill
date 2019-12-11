#!/usr/bin/env python
# -*- coding: utf-8 -*-
# author:刘杰

import tkinter as tk  # 使用Tkinter前需要先导入
from tkinter import *
from tkinter import scrolledtext
import tkinter.filedialog as fd
import tkinter.messagebox
import handleBill.bill as Bill
import handleBill.assist as assist

window = tk.Tk()
window.title('胡颖账单处理')
window.geometry('900x600')  # 这里的乘是小x
window.resizable(width=True, height=True)

frameLeft = tk.Frame(window, height = 20,width = 300)
frameLeft.grid(row=0, column=0, sticky=W+E)

frameRight = tk.LabelFrame(window,text="Info Box", padx=5, pady=5)
frameRight.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky=E+W+N+S)

window.columnconfigure(0, weight=1)
window.rowconfigure(1, weight=1)
frameRight.columnconfigure(0, weight=1)
frameRight.rowconfigure(0, weight=1)

excelPath = tk.StringVar(value='./业务工作时间_2019-08-06.xlsx')
targetPath = tk.StringVar(value='./target')
templatePath = tk.StringVar(value='./temp')
defaultPath = tk.StringVar(value='./config/default.doc')

def selectExcelPath():
  path_ = fd.askopenfilename()
  print('excelPath:',path_)
  excelPath.set(path_)

def selectTargetPath():
  path_ = fd.askdirectory()
  targetPath.set(path_)
    
def selectTemplatePath():
  path_ = fd.askdirectory()
  templatePath.set(path_)

def selectDefaultPath():
  path_ = fd.askopenfilename()
  defaultPath.set(path_)

tk.Label(frameLeft,text = "账单路径:").grid(row = 0, column = 0)
tk.Entry(frameLeft, textvariable = excelPath).grid(row = 0, column = 1)
tk.Button(frameLeft, text = "路径选择", command = selectExcelPath).grid(row = 0, column = 2)

tk.Label(frameLeft,text = "生成路径:").grid(row = 1, column = 0)
tk.Entry(frameLeft, textvariable = targetPath).grid(row = 1, column = 1)
tk.Button(frameLeft, text = "路径选择", command = selectTargetPath).grid(row = 1, column = 2)

tk.Label(frameLeft,text = "模板路径:").grid(row = 2, column = 0)
tk.Entry(frameLeft, textvariable = templatePath).grid(row = 2, column = 1)
tk.Button(frameLeft, text = "路径选择", command = selectTemplatePath).grid(row = 2, column = 2)

tk.Label(frameLeft,text = "默认路径:").grid(row = 3, column = 0)
tk.Entry(frameLeft, textvariable = defaultPath).grid(row = 3, column = 1)
tk.Button(frameLeft, text = "路径选择", command = selectDefaultPath).grid(row = 3, column = 2)

msg = scrolledtext.ScrolledText(frameRight, width=40, height=10, bg='green')
msg.grid(row=0, column=0,  sticky=E+W+N+S)
def printLog(log = ''):
  msg.insert(END, log + '\n')
#msg = tk.Message(frameRight,textvariable = userLog,bg='')
#msg.grid(row=0, column=0,  sticky=NW)
#msg.grid_propagate(0)

# 定义一个函数功能（内容自己自由编写），供点击Button按键时调用，调用命令参数command=函数名
def handle_bill():
  billObj = Bill.Bill(frameLeft, excelPath, targetPath, templatePath, defaultPath, printLog)
  billObj.displayexcelPath()
  try:
    billObj.handleBill()
  except Exception as e:
    tkinter.messagebox.showerror(title='error', message=(e))
    printLog(e)
    
b = tk.Button(frameLeft, text='开始处理', font=('Arial', 12), width=10, height=1, command=handle_bill).grid(row = 5, columnspan = 3)
 
# 主窗口循环显示
window.mainloop()