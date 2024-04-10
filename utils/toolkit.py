#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import os, sys
from loguru import logger
from functools import lru_cache


@lru_cache()
def try_or(func, default=None, expected_exc=(Exception,)):
    try:
        return func()
    except expected_exc:
        return default

@lru_cache()
def stylize_df(s):
    return "font-weight: normal; text-align: center; vertical-align: middle;"

@lru_cache()
def resource_path(relative_path):
    base_path = os.environ.get("_MEIPASS2", os.path.abspath("."))

    try:
        base_path = sys._MEIPASS
    except Exception:
        logger.warning(f"Fallback to default resource path: {base_path}")
        pass

    return os.path.join(base_path, relative_path)

def ispython(path):
    bfile = False
    try:
        # file name without extension
        fname = os.path.splitext( os.path.basename(path) )[0]
        if fname.lower() == 'python':
            bfile = True
    except:
        pass

    return bfile

def isfile(path):
    bfile = False
    try:
        if os.path.isfile(path):
            bfile = True
    except:
        pass

    return bfile

def isexcel(path):
    bexcel = False
    try:
        fname, fext = os.path.splitext(path)
        if (isfile(path) == True) and (fext.lower() == '.xlsx'):
            bexcel = True
    except:
        pass

    return bexcel

def ispdf(path):
    bpdf = False
    try:
        fname, fext = os.path.splitext(path)
        if (isfile(path) == True) and (fext.lower() == '.pdf'):
            bpdf = True
    except:
        pass

    return bpdf

def isdir(path):
    bdir = False
    try:
        if os.path.isdir(path):
            bdir = True
    except:
        pass

    return bdir

def extract_user_info(cash_flow: str='') -> list:
    # cash_flow is a pdf file
    '''
    Performance is bad, just record it but not using.
    '''
    from pdfminer.high_level import extract_text
    text = extract_text(cash_flow)
    texts = text.split('\n')
    for i, t in enumerate(texts):
        if t.startswith('查詢迄日：'):
            break

    return texts
