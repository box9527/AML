#!/usr/bin/env python3
# coding: utf-8
# @carl9527


import time
from tkinter import ttk, N, S, E, W
import tkinter as tk
import threading


def go_ahead():
    time.sleep(10)

class SimpleProgressUi:
    def __init__(self, frame):
        self.frame = frame

        # progressbar
        self.pb = ttk.Progressbar(
            self.frame,
            orient='horizontal',
            mode='determinate',
            length=200,
            takefocus=True,
            maximum=100
        )
        self.pb.place(relx=.5, rely=.5, anchor="c")
        self.pb['value'] = 0

        self.value_label = ttk.Label(self.frame, text=self.update_progress_label())
        self.value_label.place(relx=.5, rely=.6, anchor="c")

        self._start_process = False
        
        global submit_thread
        submit_thread = threading.Thread(target=go_ahead)
        submit_thread.daemon = True
        submit_thread.start()

        self.frame.after(10, self.progress)

    def update_progress_label(self):
        return f"Current Progress: {self.pb['value']}%"

    def update_progress_value(self, val: int=-1):
        if val >= 0: self.pb['value'] = val

    def progress(self, delta: int=1):
        if (self._start_process == True):
            if (self.pb['value'] + delta) >= 100:
                self.update_progress_value(100)
                self.set_error_value(f"Current Progress")
            else:
                self.update_progress_value(self.pb['value'] + delta)

            self.value_label['text'] = self.update_progress_label()
            self.frame.update()

        if self.pb['value'] >= 100:
            self._start_process = False
            return

        self.frame.after(10, self.progress)

    def get_progress_value(self):
        return round(self.pb['value'], 2)

    def start(self, interval: int=50):
        self.pb.start(interval)
        self._start_process = True

    def stop(self):
        self._start_process = False
        self.update_progress_value(0)
        self.value_label['text'] = self.update_progress_label()
        self.pb.stop()

    def set_error_value(self, text: str='Unknown Status'):
        self._start_process = False
        curr_val = self.pb['value']
        self.value_label['text'] = f"{text}: {curr_val}%"
        self.pb.stop()
        self.update_progress_value(curr_val)
        self.value_label['text'] = f"{text}: {curr_val}%"
