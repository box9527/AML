#!/usr/bin/env python3
# coding: utf-8
# @carl9527


from tkinter import ttk


class SimpleFileUi:
    def __init__(self, frame, func):
        self.frame = frame

        # open button
        self.open_button = ttk.Button(
            self.frame,
            text='匯入金流(PCMS)檔案',
            command=func
        )

        self.open_button.pack(expand=True)

    def submit_btn(self):
        return self.open_button
