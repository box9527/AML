#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import time
import os, sys
from loguru import logger
from functools import lru_cache
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import win32com.client
import win32com.client as win32
import numpy as np


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
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.environ.get("_MEIPASS2", os.path.abspath("."))

    return os.path.join(base_path, relative_path)

def isfile(path):
    bfile = False
    try:
        if os.path.isfile(path):
            bfile = True
    except:
        pass

    return bfile

def add_sheets_and_fill_data_to_xlsm(xlsm_file, sheet_data_dict, sheet_strcol_dict):
    # 使用 pywin32 開啟 Excel 應用程式
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    excel_app.Visible = False

    # 開啟 xlsm 文件
    workbook = excel_app.Workbooks.Open(xlsm_file)

    try:
        # 逐一新增 sheet 到最前頭，並填入對應的 DataFrame 資料
        for sheet_name, dataframe in sheet_data_dict.items():
            new_sheet = workbook.Sheets.Add(Before=workbook.Sheets(1))
            new_sheet.Name = sheet_name

            # headers
            headers = dataframe.columns.values.tolist()
            new_sheet.Range(new_sheet.Cells(1, 1), new_sheet.Cells(1, 1 + len(dataframe.columns) - 1)).Value = headers

            # 將整個 DataFrame 寫入 Excel 表格
            start_row = 2  # 從第二行開始填入
            end_row = start_row + len(dataframe) - 1
            start_col = 1

            rowvals = dataframe.values.tolist()
            mrange = new_sheet.Range(new_sheet.Cells(start_row, start_col), new_sheet.Cells(end_row, start_col + len(dataframe.columns) - 1))

            strcols = []
            if sheet_name in list(sheet_strcol_dict.keys()):
                strcols = sheet_strcol_dict[sheet_name]

            for cidx in strcols:
                if cidx > (len(dataframe.columns) - 1):
                    continue
                new_sheet.Range(new_sheet.Cells(start_row, cidx), new_sheet.Cells(end_row, cidx)).NumberFormat = '@'

            mrange.Value = rowvals

    finally:
        # 儲存並關閉 Excel 文件
        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel_app.Quit()

    logger.debug(f"Successfully created {len(sheet_data_dict)} new worksheets in '{xlsm_file}' and filled in the data.")

def add_sheet_fill_and_merge_to_xlsm(xlsm_file, new_sheet_name, dataframe, start_col, end_col):
    # 使用 pywin32 開啟 Excel 應用程式
    excel_app = win32com.client.DispatchEx("Excel.Application")
    excel_app.Visible = False

    # 開啟 xlsm 文件
    workbook = excel_app.Workbooks.Open(xlsm_file)

    try:
        # 在 xlsm 檔案中最前頭新增一個 sheet
        new_sheet = workbook.Sheets.Add(Before=workbook.Sheets(1))
        new_sheet.Name = new_sheet_name

        # 將 DataFrame 的資料寫入到新 sheet
        for i, col in enumerate(dataframe.columns, 1):
            new_sheet.Cells(1, i).Value = col
            for j, val in enumerate(dataframe[col], 2):
                new_sheet.Cells(j, i).Value = val

        # 在新 sheet 中根據指定的 column_index 範圍進行垂直合併
        merge_vertical_cells(new_sheet, start_col, end_col)

        # 儲存並關閉 Excel 文件
        workbook.Save()
    finally:
        workbook.Close(SaveChanges=True)
        excel_app.Quit()

    logger.debug(f"DataFrame successfully added to new sheet '{new_sheet_name}' in '{xlsm_file}'.")

def merge_vertical_cells(sheet, start_col, end_col):
    for col_index in range(start_col, end_col + 1):
        current_value = None
        start_row = None

        for row_index in range(2, sheet.UsedRange.Rows.Count + 1):
            cell_value = sheet.Cells(row_index, col_index).Value
            if cell_value == current_value:
                if start_row is None:
                    start_row = row_index - 1
            else:
                if start_row is not None:
                    end_row = row_index - 1
                    sheet.Range(sheet.Cells(start_row, col_index), sheet.Cells(end_row, col_index)).Merge()
                    start_row = None

            current_value = cell_value


