#!/usr/bin/env python3
# coding: utf-8
# @voneyay


import os
import glob
import time
import tkinter as tk
from loguru import logger
import shutil
from utils.toolkit import (
    isfile,
    resource_path,
)
from gui import GUIApp


def run():
    start_time = time.time()
    root = tk.Tk()
    gui_app = GUIApp(root)
    gui_app.run_gui()
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    logger.debug(f"程式執行時間: {elapsed_time} 秒")


def check_extra_excels(excel_file: str=''):
    if isfile(excel_file) == False:
        shutil.copy(resource_path(f'extra_excels/{excel_file}'), f'{excel_file}')

if __name__ == "__main__":
    check_extra_excels('警示戶.xlsx')
    check_extra_excels('虛擬帳戶.xlsx')

    run()

