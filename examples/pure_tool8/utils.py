#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import platform
import os, sys
from loguru import logger


def try_or(func, default=None, expected_exc=(Exception,)):
    try:
        return func()
    except expected_exc:
        return default

def stylize_df(s):
    return "font-weight: normal; text-align: center; vertical-align: middle;"

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

def isdir(path):
    bdir = False
    try:
        if os.path.isdir(path):
            bdir = True
    except:
        pass

    return bdir

def add_sheets_and_fill_data_to_xlsm(xlsm_file, sheet_data_dict, sheet_strcol_dict):
    if platform.system() != 'Windows':
        logger.debug(f"Not a Windows system: {platform.system()}")
        logger.debug(f"Cannot import win32com library. Just skip for test.")
        return

    import win32com.client
    import win32com.client as win32

    # 使用 pywin32 開啟 Excel 應用程式
    excel_app = win32.DispatchEx('Excel.Application')
    excel_app.Visible = False

    # 開啟 xlsm 文件
    workbook = excel_app.Workbooks.Open(xlsm_file)

    try:
        # 逐一新增 sheet 到最前頭，並填入對應的 DataFrame 資料
        for sheet_name, dataframe in sheet_data_dict.items():
            new_sheet = workbook.Sheets.Add(Before=workbook.Sheets(1))
            new_sheet.Name = sheet_name

            hd_start_row = 1
            hd_start_col = 1

            if sheet_name == "1原始資料":
                hd_start_row = 8

            da_start_row = hd_start_row + 1
            da_end_row = da_start_row + len(dataframe) - 1
            da_start_col = 1

            # headers
            headers = dataframe.columns.values.tolist()
            new_sheet.Range(
                    new_sheet.Cells(hd_start_row, hd_start_col), 
                    new_sheet.Cells(hd_start_row, hd_start_col + len(dataframe.columns) - 1)).Value = headers

            # 將整個 DataFrame 寫入 Excel 表格
            rowvals = dataframe.values.tolist()
            mrange = new_sheet.Range(
                    new_sheet.Cells(da_start_row, da_start_col), 
                    new_sheet.Cells(da_end_row, da_start_col + len(dataframe.columns) - 1))

            strcols = []
            if sheet_name in list(sheet_strcol_dict.keys()):
                strcols = sheet_strcol_dict[sheet_name]

            for cidx in strcols:
                if cidx > (len(dataframe.columns) - 1):
                    continue
                new_sheet.Range(new_sheet.Cells(da_start_row, cidx), new_sheet.Cells(da_end_row, cidx)).NumberFormat = '@'

            mrange.Value = rowvals
    except Exception as e:
        logger.debug('Unexpected issue: {e}')

    finally:
        # 儲存並關閉 Excel 文件
        workbook.Save()
        workbook.Close(SaveChanges=True)
        excel_app.Quit()

    logger.debug(f"Successfully created {len(sheet_data_dict)} new worksheets in '{xlsm_file}' and filled in the data.")

