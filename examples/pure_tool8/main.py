#!/usr/bin/env python3
# coding: utf-8
# @carl9527


import time
from loguru import logger
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

from excels_processor import Productivity
from utils import isfile


# create the root window
root = tk.Tk()
root.title('工具八POC金流(PCMS)檔案匯入器')
root.resizable(False, False)
root.geometry('300x150')

btn_txt = '匯入金流(PCMS)檔案'

file_path = tk.StringVar()


def select_file():
    filetypes = (
        ('PDF files', '*.pdf'),
    )

    filename = fd.askopenfilename(
        title='選擇金流(PCMS)檔案',
        initialdir='/',
        filetypes=filetypes)

    showinfo(
        title='選擇的檔案',
        message=f'選擇的檔案路徑：{filename}'
    )

    file_path.set(filename)

    # after close infomation woindow
    progress()

def update_button_txt(message: str=''):
    if (not message) or (len(message) <= 0):
        return

    open_button['text'] = f'{message}'
    root.update()

def update_btn_state(state: bool=True):
    open_button['state'] = 'enabled' if state == True else 'disabled'

def progress():
    update_btn_state(False)
    filename = file_path.get() 

    if isfile(filename) == False:
        showinfo(
            title='選擇的檔案',
            message=f'檔案路徑不存在：{filename}'
        )
        update_btn_state(True)
        return

    logger.info(f'Import and Create poc_tool8 with file {filename} started.')
    update_button_txt(f'{btn_txt} 10%')

    prod = Productivity()

    # check point 1
    prod.set_cash_flow_file(cash_flow=filename)
    update_button_txt(f'{btn_txt} 15%')

    # check point 2
    bsuccess = prod.check_jre()
    if bsuccess == False:
        update_button_txt(f'JAVA 檢測失敗 ...')
        time.sleep(3)
        update_button_txt(f'{btn_txt}')
        update_btn_state(True)
        return

    update_button_txt(f'{btn_txt} 30%')

    # check point 3
    bsuccess, data = prod.analysis_pdf()
    if bsuccess == False:
        update_button_txt(f'取出PCMS資料失敗 ...')
        time.sleep(3)
        update_button_txt(f'{btn_txt}')
        update_btn_state(True)
        return

    update_button_txt(f'{btn_txt} 45%')

    # check point 4
    bsuccess, sdata = prod.strict_pdf(data)
    if bsuccess == False:
        update_button_txt(f'精煉PCMS資料失敗 ...')
        time.sleep(3)
        update_button_txt(f'{btn_txt}')
        update_btn_state(True)
        return

    update_button_txt(f'{btn_txt} 87%')

    # check point 5
    bsuccess = prod.export_data(sdata)
    if bsuccess == False:
        update_button_txt(f'最終匯入PCMS資料失敗 ...')
        time.sleep(3)
        update_button_txt(f'{btn_txt}')
        update_btn_state(True)
        return

    logger.info(f'Import and Create poc_tool8 with file {filename} completed.')
    update_button_txt(f'{btn_txt} 99%')
    
    time.sleep(3)
    update_button_txt(f'{btn_txt}')
    update_btn_state(True)

# open button
open_button = ttk.Button(
    root,
    text='匯入金流(PCMS)檔案 0%',
    command=select_file
)

open_button.pack(expand=True)

# run the application
root.mainloop()
