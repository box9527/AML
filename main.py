#!/usr/bin/env python3
# coding: utf-8
# @voneyay


import os
import glob
import time
import tkinter as tk
from loguru import logger
from gui import GUIApp


def run():
    start_time = time.time()
    root = tk.Tk()
    gui_app = GUIApp(root)
    gui_app.run_gui()
    source_folder = gui_app.source_folder
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    logger.debug(f"程式執行時間: {elapsed_time} 秒")


if __name__ == "__main__":
    run()

