import tkinter as tk
import time
import sys
from tkinter import filedialog
from pathlib import Path
import pytesseract
from PIL import Image
import re
import numpy as np
import pandas as pd
import os

curPath = os.path.abspath(os.path.dirname(__file__))
filePath = os.path.split(curPath)[0]
sys.path.append(filePath)
sys.path.extend([filePath + '\\' + i for i in os.listdir(filePath)])


def open_file():
    global dir_name
    dir_name = filedialog.askdirectory(title='选择文件夹', initialdir=r'’D:\a’')
    return 0


def show(c):
    text_fm.insert('end', c)
    text_fm.insert('end', "\n")


def ocr():
    global dir_name
    Data = ['文件名', '姓名', '检测日期', '检测结果']
    t1 = time.time()
    for i in Path(dir_name).iterdir():  # 获取文件夹下的所有文件。
        i.as_posix()  # as_posix() 把Path型转为字符串。
        result = pytesseract.image_to_string(Image.open(i), lang='chi_sim')
        show(i)
        result = result.split()
        for o in result:
            if o == ' ':
                result.remove(o)
        result = str(result)
        name = re.compile('姓名:')
        name_loc = name.search(result).regs[0]
        date = re.compile('\d{4}-\d{2}-\d{2}')
        date_loc = date.search(result).regs[0]
        res = re.compile('性')
        res_loc = res.search(result).regs[0]
        data = [i.stem, result[name_loc[1] + 4:name_loc[1] + 7], result[date_loc[0]:date_loc[1]],
                result[res_loc[0] - 1:res_loc[1]]]
        Data = np.vstack((Data, data))
        root.update()
    t = time.time() - t1
    pd_data = pd.DataFrame(Data)
    fileName ='识别结果Output.xlsx'
    writer = pd.ExcelWriter(fileName)
    pd_data.to_excel(writer, 'sheet1', startrow=0, startcol=0)
    writer.save()
    # print('识别结果已保存到“识别结果Output.xlsx”，请打开查看！')
    # print('耗时', t, 's')
    fn = '识别结果已保存到' + str(fileName) + '，请打开查看！'
    tt = '耗时', str(t), 's'
    show(fn)
    show(tt)


if __name__ == '__main__':
    root = tk.Tk()
    root.iconbitmap("hhulogo.ico")
    root.title("OCR_for_HHU")
    root.geometry('600x400')
    root.resizable(0, 0)
    text_fm = tk.Text(root,  width=90, height=27)
    text_fm.place(x=0, y=40)
    open_button = tk.Button(root, width=10, text="选择文件夹", font=('宋体', 15, 'bold'), relief="groove", command=open_file)
    open_button.place(x=50, y=0,)
    ocr_button = tk.Button(root, width=10, text="开始识别", font=('宋体', 15, 'bold'), relief="groove", command=ocr)
    ocr_button.place(x=200, y=0,)
    exit_button = tk.Button(root, width=10, text="退出程序", font=('宋体', 15, 'bold'), relief="groove", command=root.destroy)
    exit_button.place(x=350, y=0,)
    root.update()
    root.mainloop()
