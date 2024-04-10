#!/usr/bin/env python3
# coding: utf-8
# @carl9527


import signal
from tkinter import ttk, VERTICAL, HORIZONTAL

import os
import sys
import pathlib
import logging
import tkinter as tk
from tkinter import *

ROOT_PATH = pathlib.Path(__file__).parent.parent.parent.resolve()
sys.path.append(ROOT_PATH)
sys.path.append('/Users/carl/VSCode/workspace/nlp_aml/')
sys.path.append('/Users/carl/VSCode/workspace/nlp_aml/utils/')
print(sys.path)

from utils.sample_form_ui import FormUi
from utils.console_ui import ConsoleUi
from utils.sample_third_ui import ThirdUi
from utils.clock import Clock


class App:

    def __init__(self, root):
        self.root = root
        #self.root.geometry('800x500')
        self.root.geometry('800x600')

        root.title('Logging Handler')
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        # Create the panes and frames
        vertical_pane = ttk.PanedWindow(self.root, orient=VERTICAL)
        vertical_pane.grid(row=0, column=0, sticky="nsew")
        horizontal_pane = ttk.PanedWindow(vertical_pane, orient=HORIZONTAL)
        vertical_pane.add(horizontal_pane, weight=1)

        # panes
        # form_frame
        form_frame = ttk.Labelframe(horizontal_pane, text="MyForm", width=300, height=200)
        form_frame.columnconfigure(1, weight=1)
        form_frame.pack(fill=BOTH,expand=True)
        horizontal_pane.add(form_frame, weight=1)

        # console_frame
        third_frame = ttk.Labelframe(horizontal_pane, text="Third Frame", width=500, height=200)
        third_frame.pack(fill=BOTH,expand=True)
        horizontal_pane.add(third_frame, weight=2)

        # third_frame
        console_frame = ttk.Labelframe(vertical_pane, text="Console", width=800, height=400)
        console_frame.columnconfigure(0, weight=1)
        console_frame.rowconfigure(0, weight=1)
        console_frame.pack(fill=BOTH,expand=True)
        vertical_pane.add(console_frame, weight=1)

        # Initialize all frames
        self.form = FormUi(form_frame)
        self.console = ConsoleUi(console_frame)
        self.third = ThirdUi(third_frame)
        self.clock = Clock()
        self.clock.start()
        self.root.protocol('WM_DELETE_WINDOW', self.quit)
        self.root.bind('<Control-q>', self.quit)
        signal.signal(signal.SIGINT, self.quit)

    def quit(self, *args):
        self.clock.stop()
        self.root.destroy()

def main():
    logging.basicConfig(level=logging.DEBUG)
    root = tk.Tk()
    app = App(root)
    app.root.mainloop()


if __name__ == '__main__':
    main()
