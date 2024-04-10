#!/usr/bin/env python3
# coding: utf-8
# @voneyay


import time
import signal
from loguru import logger
import logging
import random
import tkinter as tk
import pandas as pd
import numpy as np
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import ttk, VERTICAL, HORIZONTAL
from tkinter import *

from excels_processor import Productivity
from utils.toolkit import isfile, ispdf, isexcel

from utils.simple_file_ui import SimpleFileUi
from utils.console_ui import ConsoleUi
from utils.simple_progress_ui import SimpleProgressUi


class GUIApp:
    def __init__(self, root):
        # create the root window
        self.root = root
        self.root.title('工具八POC金流(PCMS)檔案匯入器')
        self.root.resizable(False, False)
        self.btn_txt = '匯入金流(PCMS)檔案'

        self.root.geometry('800x600')
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        self.file_path = tk.StringVar()
        self.prod = Productivity()
        
        # Create the panes and frames
        vertical_pane = ttk.PanedWindow(self.root, orient=VERTICAL)
        vertical_pane.grid(row=0, column=0, sticky="nsew")
        horizontal_pane = ttk.PanedWindow(vertical_pane, orient=HORIZONTAL)
        vertical_pane.add(horizontal_pane, weight=1)

        ## open button
        form_frame = ttk.Labelframe(horizontal_pane, text="File selected", width=300, height=150)
        form_frame.columnconfigure(1, weight=1)
        form_frame.pack(fill=BOTH,expand=True)
        horizontal_pane.add(form_frame, weight=1)

        # progress_frame
        progress_frame = ttk.Labelframe(horizontal_pane, text="Progress", width=500, height=150)
        progress_frame.pack(fill=BOTH,expand=True)
        horizontal_pane.add(progress_frame, weight=2)

        # console_frame
        console_frame = ttk.Labelframe(vertical_pane, text="Console", width=800, height=450)
        console_frame.columnconfigure(0, weight=1)
        console_frame.rowconfigure(0, weight=1)
        console_frame.pack(fill=BOTH,expand=True)
        vertical_pane.add(console_frame, weight=2)

        self.form_ctrl = SimpleFileUi(form_frame, self.select_file)
        self.progress_ctrl = SimpleProgressUi(progress_frame)
        self.console_ctrl = ConsoleUi(console_frame)

        self.root.protocol('WM_DELETE_WINDOW', self.quit)
        self.root.bind('<Control-q>', self.quit)
        signal.signal(signal.SIGINT, self.quit)

    def quit(self, *args):
        self.root.destroy()

    def run_gui(self):
        self.root.mainloop()

    def select_file(self):
        filetypes = (
            ('PDF files', '*.pdf'),
            ('Excel files', '*.xlsx'),
        )

        filename = fd.askopenfilename(
            title='選擇金流(PCMS)檔案',
            initialdir='/',
            filetypes=filetypes)

        if False:
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

        self.form_ctrl.submit_btn()['text'] = f'{message}'
        self.root.update()

    def update_btn_state(self, state: bool=True):
        self.form_ctrl.submit_btn()['state'] = 'enabled' if state == True else 'disabled'

    def checkpoint_progress(self, init_step: int=5, delta: int=1, start_interval: int=1000):
        self.progress_ctrl.stop()
        self.progress_ctrl.start(start_interval)
        self.progress_ctrl.update_progress_value(init_step)
        self.progress_ctrl.progress(delta)

    def conv_progress_step_val(self, up_bound: int=0) -> float:
        #return round(random.uniform(float(self.progress_ctrl.get_progress_value()), float(up_bound)), 2)
        return round(
                random.uniform(
                    round(float(self.progress_ctrl.get_progress_value()), 2), 
                    round(float(up_bound), 2)
                ), 2)

    def progress(self):
        self.update_btn_state(False)
        filename = self.file_path.get()

        if isfile(filename) == False:
            if len(filename) > 0:
                showinfo(
                    title='選擇的檔案',
                    message=f'檔案路徑不存在：{filename}'
                )
            self.update_btn_state(True)
            return

        sdata = None
        logger.info(f'Import and Create poc_tool8 with file {filename} started.')
        self.checkpoint_progress(init_step=self.conv_progress_step_val(5))

        if ispdf(filename) == True:
            # check point 1
            self.prod.set_cash_flow_file(cash_flow=filename)
            self.checkpoint_progress(init_step=self.conv_progress_step_val(15))

            # check point 2
            bsuccess = self.prod.check_jre()
            if bsuccess == False:
                self.progress_ctrl.set_error_value(f'JAVA detecting process fail ...')
                self.update_btn_state(True)
                return

            self.checkpoint_progress(init_step=self.conv_progress_step_val(28))

            # check point 3
            bsuccess, data = self.prod.analysis_pdf()
            if bsuccess == False:
                self.progress_ctrl.set_error_value(f'PCMS rawdata extraction fail ...')
                self.update_btn_state(True)
                return

            self.checkpoint_progress(init_step=self.conv_progress_step_val(43))

            # check point 4
            bsuccess, sdata = self.prod.strict_pdf(data)
            if bsuccess == False:
                self.progress_ctrl.set_error_value(f'PCMS rawdata refined fail ...')
                self.update_btn_state(True)
                return

        if (sdata is None) and (isexcel(filename) == True):
            bsuccess, sdata = self.prod.strict_excel(filename)

        self.checkpoint_progress(init_step=self.conv_progress_step_val(87))

        # check point 5
        bsuccess = self.prod.export_data(sdata)
        if bsuccess == False:
            self.progress_ctrl.set_error_value(f'PCMS import fail ...')
            self.update_btn_state(True)
            return

        logger.info(f'Import and Create poc_tool8 with file {filename} completed.')
        self.update_button_txt(f'{self.btn_txt}')
        self.update_btn_state(True)

        self.checkpoint_progress(init_step=self.conv_progress_step_val(96))
