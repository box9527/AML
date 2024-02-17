#!/usr/bin/env python3
# coding: utf-8
# @carl9527


import os
import pathlib


ROOT_PATH = pathlib.Path(__file__).parent.parent.resolve()
CONST_PATH = os.path.join(ROOT_PATH, 'consts')
DAEX_PATH = os.path.join(ROOT_PATH, 'DataExtraction')

STOPWORDS = os.path.join(CONST_PATH, 'stop_wordsv2.txt')
