import os
import glob
import time
import tkinter as tk
from gui import GUIApp
from txtrank_v2 import TextRankSummarization

if __name__ == "__main__":
    start_time = time.time()
    root = tk.Tk()
    gui_app = GUIApp(root)
    gui_app.run_gui()
    source_folder = gui_app.source_folder

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"程式執行時間: {elapsed_time} 秒")

