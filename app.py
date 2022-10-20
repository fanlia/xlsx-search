#!/usr/bin/python3

import os
import io
from xlsx2csv import Xlsx2csv
from tkinter import *
from tkinter.filedialog import *

win = Tk()

win.title('xlsx 查找小工具')
win.geometry('300x150')

def format_lines(lines):
    if len(lines) == 0:
        return '没有找到数据'

    result = ''

    for (dirpath, filename, i, line) in lines:
        result += dirpath + ' ' + filename + ' ' + str(i) + ' ' + line

    return result

def search_xlsx(text, xlsxpath):
    f = io.StringIO()
    Xlsx2csv(xlsxpath, outputencoding="utf-8").convert(f)
    f.seek(0)

    results = []

    i = 1
    for line in f:
        if line.find(text) == -1:
            continue
        results.append((i, line))

    return results

def search_dir(text, directory):
    lines = []
    files = os.walk(directory)
    for file in files:
        dirpath, _, filenames = file
        for filename in filenames:
            if not filename.endswith('.xlsx'):
                continue
            file_lines = search_xlsx(text, os.path.join(dirpath, filename))
            for (i, line) in file_lines:
                lines.append((dirpath, filename, i, line))

    return format_lines(lines)

    # return text + ' ' + directory

def search():
    text = keyword.get()
    if len(text) == 0:
        return

    directory = askdirectory(title='请选择包含xlsx的文件夹')
    if len(directory) == 0:
        return

    result = search_dir(text, directory)
    label.config(text=result)

keyword = Entry(win)
keyword.grid(row=0)
btn = Button(win, text='搜索', command=search).grid(row=0, column=1)
label = Label(win, text='请输入关键字', justify='left')
label.grid(row=1, column=0)

win.mainloop()
