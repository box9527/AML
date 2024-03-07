#!/usr/bin/env python3
# coding: utf-8
# @voneyay


import time
from loguru import logger
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

from excels_processor import Productivity
from utils.toolkit import isfile


class GUIApp:
    def __init__(self, root): 
        # create the root window
        self.root = root
        self.root.title('工具八POC金流(PCMS)檔案匯入器')
        self.root.resizable(False, False)
        self.root.geometry('300x150')
        self.btn_txt = '匯入金流(PCMS)檔案'

        self.file_path = tk.StringVar()
        self.prod = Productivity()

        # open button
        self.open_button = ttk.Button(
            self.root,
            text='匯入金流(PCMS)檔案 0%',
            command=self.select_file
        )
        
        self.open_button.pack(expand=True)
        
    def run_gui(self):
        self.root.mainloop()

    def select_file(self):
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

        self.file_path.set(filename)

        # after close infomation woindow
        self.progress()

    def update_button_txt(self, message: str=''):
        if (not message) or (len(message) <= 0):
            return

        self.open_button['text'] = f'{message}'
        self.root.update()

    def update_btn_state(self, state: bool=True):
        self.open_button['state'] = 'enabled' if state == True else 'disabled'

    def progress(self):
        self.update_btn_state(False)
        filename = self.file_path.get()

        if isfile(filename) == False:
            showinfo(
                title='選擇的檔案',
                message=f'檔案路徑不存在：{filename}'
            )
            self.update_btn_state(True)
            return

        logger.info(f'Import and Create poc_tool8 with file {filename} started.')
        self.update_button_txt(f'{self.btn_txt} 10%')

        # check point 1
        self.prod.set_cash_flow_file(cash_flow=filename)
        self.update_button_txt(f'{self.btn_txt} 15%')

        # check point 2
        bsuccess = self.prod.check_jre()
        if bsuccess == False:
            self.update_button_txt(f'JAVA 檢測失敗 ...')
            time.sleep(3)
            self.update_button_txt(f'{self.btn_txt}')
            self.update_btn_state(True)
            return

        self.update_button_txt(f'{self.btn_txt} 30%')

        # check point 3
        bsuccess, data = self.prod.analysis_pdf()
        if bsuccess == False:
            self.update_button_txt(f'取出PCMS資料失敗 ...')
            time.sleep(3)
            self.update_button_txt(f'{self.btn_txt}')
            self.update_btn_state(True)
            return

        self.update_button_txt(f'{self.btn_txt} 45%')

        # check point 4
        bsuccess, sdata = self.prod.strict_pdf(data)
        if bsuccess == False:
            self.update_button_txt(f'精煉PCMS資料失敗 ...')
            time.sleep(3)
            self.update_button_txt(f'{self.btn_txt}')
            self.update_btn_state(True)
            return

        self.update_button_txt(f'{self.btn_txt} 87%')

        # check point 5
        bsuccess = self.prod.export_data(sdata)
        if bsuccess == False:
            self.update_button_txt(f'最終匯入PCMS資料失敗 ...')
            time.sleep(3)
            self.update_button_txt(f'{self.btn_txt}')
            self.update_btn_state(True)
            return

        logger.info(f'Import and Create poc_tool8 with file {filename} completed.')
        self.update_button_txt(f'{self.btn_txt} 99%')

        time.sleep(3)
        self.update_button_txt(f'{self.btn_txt}')
        self.update_btn_state(True)

