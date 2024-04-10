#!/usr/bin/env python3
# coding: utf-8
# @carl9527


from tkinter import ttk, N, S, E, W


class ThirdUi:
    def __init__(self, frame):
        self.frame = frame
        ttk.Label(self.frame, text='This is just an example of a third frame').grid(column=0, row=1, sticky=W)
        ttk.Label(self.frame, text='With another line here!').grid(column=0, row=4, sticky=W)
