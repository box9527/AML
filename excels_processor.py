#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import time
import os, sys, platform
import os.path as path
from pathlib import Path
from loguru import logger
import copy
import re
import pandas as pd
import numpy as np
import shutil
import tarfile
from openpyxl import Workbook, load_workbook
import tabula
from utils.toolkit import (
        try_or,
        stylize_df,
        ispython,
        isfile,
        isdir,
        resource_path,
)
from utils.user_info import (
        get_user_info,
)
from utils.windows import (
        add_sheets_and_fill_data_to_xlsm,
)
from consts import STOPWORDS


class Productivity:
    def __init__(self):
        self.cash_flow = ''
        self.dwl_acc = ''
        self.vr_acc = ''

        self.tool8_tmpl = f'poc_tool8.xlsm'
        self.tool8_jre = 'jre-8u211-windows-x64.tar.gz'

        # intermediate file
        self._cp_strict_rawdata = "rawdata-strict-tool8.xlsx"

        # result file
        self.cp_combined_result = f'工具8_異常態樣分析摘要.xlsm'

        # 這裡的 "支出" 與 "存入" 用 "Out" 以及 "In"取代，目的是跟工具七的欄位一致
        self.rawdata_cols = {
            '交易日期':[],'帳務日期':[],'交易代號':[],'交易時間':[],'交易分行':[],'交易櫃員':[],
            '摘要':[],'Out':[],'In':[],'餘額':[],
            '轉出入帳號':[],'合作機構/會員編號':[],'金資序號':[],'票號':[],'備註':[],'註記':[]}

        self.strict_cols = copy.deepcopy(self.rawdata_cols)

        java_home = os.getenv('JAVA_HOME', '')

        pass

    def set_cash_flow_file(self, cash_flow: str='') -> bool:
        if cash_flow and (isfile(cash_flow) == True):
            self.cash_flow = cash_flow
            return True

        return False

    def check_jre(self) -> bool:
        bsuccess, _ = self._check_jre()
        if bsuccess == False:
            logger.warning(f'Check JAVA environment failed.')

        return bsuccess

    def analysis_pdf(self) -> (bool, pd.DataFrame):
        bsuccess, rawdata = self._rawdata_from_cash_flow()
        if bsuccess == False:
            logger.error(f'Extract rawdata from PCMS file failed.')

        return bsuccess, rawdata

    def strict_pdf(self, rawdata: pd.DataFrame=None) -> (bool, pd.DataFrame):
        bsuccess, strict_rawdata = self._strict_from_rawdata(rawdata)
        if bsuccess == False:
            logger.error(f'Strict rawdata failed.')

        return bsuccess, strict_rawdata

    def export_data(self, rawdata: pd.DataFrame=None) -> bool:
        bsuccess = self._combine_product(data=rawdata,)
        if bsuccess == False:
            logger.error(f'Export to poc_tool8 failed.')

        return bsuccess

    def _check_jre(self) -> (bool, str):
        bsuccess = False
        jh = ''

        try:
            jf = resource_path(self.tool8_jre)
            p = Path(jf)
            jh = str(p.parent.absolute())

            java_home = os.environ.get('JAVA_HOME', '')
            if isdir(java_home) == False:
                logger.debug(f"Can not find JAVA_HOME, use dynamically installation.")
                jtarget = (jh + '\\jre1.8.0_211\\') if platform.system() == 'Windows' else (jh + '/jre1.8.0_211/')

                if isdir(jtarget) == False:
                    logger.debug(f"Start dynamically tarball extraction ...")
                    tf = tarfile.open(jf)
                    tf.extractall(jh)
                    tf.close()
                    logger.debug(f"Dynamically tarball extraction is Done.")
                else:
                    logger.info(f"Dynamic installation path exists, skip.")

                os.environ['JAVA_HOME'] = jtarget

                extra_path = (os.pathsep + (jtarget + 'bin\\')) if platform.system() == 'Windows' else (os.pathsep + (jtarget + 'bin/'))
                os.environ['PATH'] += extra_path

                logger.debug(f"JAVA_HOME: {os.environ.get('JAVA_HOME', '')}")
                logger.debug(f"PATH: {os.environ.get('PATH', '')}")
                jh = os.environ.get('JAVA_HOME', '')
            bsuccess = True
        except Exception as e:
            logger.debug(f"Dynamically check JAVA_HOME failed.")
            logger.error(f'Failed message: {e}')
            pass

        return bsuccess, jh

    def _rawdata_from_cash_flow(self) -> (bool, pd.DataFrame):
        bsuccess = False
        exportx = None
        try:
            export = pd.DataFrame.from_dict( self.rawdata_cols )
            dfs = tabula.read_pdf(self.cash_flow, area=[120, 5, 800, 1200], pages="all")
            for df in dfs:
                df = df.replace(np.nan, '', regex=True)
                headers = df.columns.values.tolist()
                wrong_idx = [i for i, item in enumerate(headers) if re.search('^Unnamed:', item)]
                (rows, cols) = df.shape

                for row in range(rows):
                    if row <= 0: continue

                    refine1 = []
                    values = df.iloc[row].tolist()

                    next_skip = False
                    for idx, item in enumerate(values):
                        if (idx in wrong_idx) and (idx <= 12):
                            if len(str(item).strip()) <= 0: continue
                            if (len(refine1) > 0) and (len(refine1[-1]) <= 0): del refine1[-1]

                        if idx > 12:
                            if next_skip == True:
                                next_skip = False
                                continue
                            if idx in wrong_idx:
                                next_skip = True
                        refine1.append(str(item).strip())

                    refine2 = []
                    for idx, item in enumerate(refine1):
                        iitem = item
                        try:
                            fitem = float(item)
                            if (idx == 2) or (idx == 9):
                                iitem = str(int(fitem))
                        except Exception as e:
                            pass

                        if idx == 4:
                            (col4_val, col5_val) = iitem.split(' ')
                            refine2.append(col4_val)
                            refine2.append(col5_val)
                            continue

                        refine2.append(iitem)
                    export.loc[len(export.index)] = refine2

            exportx = export.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)

            bsuccess = True
        except Exception as e:
            logger.debug('Extract rawdata from PDF failed.')
            logger.error(f'Failed message: {e}')
            pass

        return bsuccess, exportx

    def _strict_from_rawdata(self, rawdata: pd.DataFrame=None) -> (bool, pd.DataFrame):
        bsuccess = False
        exportx = None

        def rawdata_formatting(row):
            deal_date = try_or(lambda:f"{str(row['交易日期'].year).zfill(4)}/{str(row['交易日期'].month).zfill(2)}/{str(row['交易日期'].day).zfill(2)}",default=f"{row['交易日期']}")
            acc_date = try_or(lambda:f"{str(row['帳務日期'].year).zfill(4)}/{str(row['帳務日期'].month).zfill(2)}/{str(row['帳務日期'].day).zfill(2)}",default=f"{row['帳務日期']}")
            deal_code = try_or(lambda:f"{str(row['交易代號']).zfill(5)}",default=f"{row['交易代號']}")
            deal_time = try_or(lambda:f"{str(row['交易時間'].hour).zfill(2)}:{str(row['交易時間'].minute).zfill(2)}:{str(row['交易時間'].second).zfill(2)}",default=f"{row['交易時間']}")
            deal_branch = try_or(lambda:f"{str(row['交易分行']).zfill(4)}",default=f"{row['交易分行']}")
            deal_teller = try_or(lambda:f"{str(row['交易櫃員']).zfill(5)}",default=f"{row['交易櫃員']}")
            summary = try_or(lambda:f"{row['摘要']}".strip(),default=f"{row['摘要']}")
            m_out = try_or(lambda:f"{row['Out']:,.2f}",default=f"{row['Out']}")
            m_in = try_or(lambda:f"{row['In']:,.2f}",default=f"{row['In']}")
            m_balance = try_or(lambda:f"{row['餘額']:,.2f}",default=f"{row['餘額']}")
            tr_acc = try_or(lambda:f"{row['轉出入帳號']}".strip(),default=f"{row['轉出入帳號']}")
            tr_infra = try_or(lambda:f"{row['合作機構/會員編號']}".strip(),default=f"{row['合作機構/會員編號']}")

            m_number = try_or(lambda:f"{row['金資序號']}".strip(),default=f"{row['金資序號']}")
            t_number = try_or(lambda:f"{row['票號']}".strip(),default=f"{row['票號']}")
            comment = try_or(lambda:f"{row['備註']}".strip(),default=f"{row['備註']}")
            tmp = try_or(lambda:f"{row['註記']}".strip(),default=f"{row['註記']}")

            return deal_date, acc_date, deal_code, deal_time, deal_branch, deal_teller, \
                    summary, m_out, m_in, m_balance, tr_acc, tr_infra, m_number, t_number, comment, tmp

        try:
            if rawdata is None: return (bsuccess, exportx)

            export = pd.DataFrame( self.strict_cols )
            export['交易日期'], export['帳務日期'], export['交易代號'],\
            export['交易時間'], export['交易分行'], export['交易櫃員'],\
            export['摘要'], export['Out'], export['In'],\
            export['餘額'], export['轉出入帳號'], export['合作機構/會員編號'],\
            export['金資序號'], export['票號'],\
            export['備註'], \
            export['註記'] \
            = zip(*rawdata.apply(rawdata_formatting, axis=1))

            export = export[((export['Out'] == export['Out']) ) |
            ((export['In'] == export['In']))]

            exportx = export.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)

            bsuccess = True
        except Exception as e:
            logger.debug('Strict rawdata failed.')
            logger.error(f'Failed message: {e}')
            pass

        if (exportx is not None):# and (platform.system() != 'Windows'):
            exportx = export.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)
            with pd.ExcelWriter(self._cp_strict_rawdata) as writer:
                exportx.to_excel(writer, sheet_name="Rawdata", startrow=8, index=False)

            df_sheet_name = pd.read_excel(self._cp_strict_rawdata, sheet_name='Rawdata')
            exportx = df_sheet_name.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)
            user_info = get_user_info(self.cash_flow)
            exportx.loc[3, 'Unnamed: 0'] = '客戶：'
            exportx.loc[3, 'Unnamed: 1'] = try_or(lambda:f"{str(user_info['客戶'])}",default=f"{user_info['客戶']}")
            exportx.loc[3, 'Unnamed: 3'] = '產品別：'
            exportx.loc[3, 'Unnamed: 4'] = user_info['產品別']
            exportx.loc[3, 'Unnamed: 6'] = '查詢起日：'
            exportx.loc[3, 'Unnamed: 7'] = user_info['查詢起日']
            exportx.loc[4, 'Unnamed: 0'] = '帳號：'
            exportx.loc[4, 'Unnamed: 1'] = try_or(lambda:f"{str(user_info['帳號'])}",default=f"{user_info['帳號']}")
            exportx.loc[4, 'Unnamed: 3'] = '幣別：'
            exportx.loc[4, 'Unnamed: 4'] = user_info['幣別']
            exportx.loc[4, 'Unnamed: 6'] = '查詢迄日：'
            exportx.loc[4, 'Unnamed: 7'] = user_info['查詢迄日']
            exportx.loc[5, 'Unnamed: 0'] = ''
            exportx.loc[6, 'Unnamed: 0'] = user_info['交易內容']
            with pd.ExcelWriter(self._cp_strict_rawdata) as writer:
                exportx.to_excel(writer, sheet_name="Rawdata", index=False, header=False)

            df_sheet_name = pd.read_excel(self._cp_strict_rawdata, sheet_name='Rawdata')
            exportx = df_sheet_name.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)

        return bsuccess, exportx

    def _combine_product(self, data: pd.DataFrame=None) -> bool:
        bsuccess = False
        data.index = np.arange(0, len(data))
        
        gen_dfs = dict()
        plugin_ui = os.getenv('PLUGIN_UI', False)
        if (plugin_ui == True) or (plugin_ui == "True"):
            gen_dfs = self._gen_plugin_data(rawdata=data)
        
        try:
            shutil.copy(resource_path(self.tool8_tmpl), self.cp_combined_result)

            real_result = f"{os.path.dirname(sys.executable)}\\{self.cp_combined_result}"
            if ispython(sys.executable) == True:
                real_result = resource_path(self.cp_combined_result)

            # 定義 sheet 名稱及對應的 DataFrame
            sheet_data_dict = {
                #"Sheet1": pd.DataFrame({'Column1': [1, 2, 3], 'Column2': ['A', 'B', 'C']}),
                #"Sheet2": pd.DataFrame({'Column1': [4, 5, 6], 'Column2': ['D', 'E', 'F']}),
                # 可以根據需要繼續添加
                "1原始資料": data,
            }

            sheet_strcol_dict = {
                # 定義有哪幾欄的資料是"必須要強制為字串格式的"
                "1原始資料": [1, 2, 3, 5, 6, 11, 12, 13, 14, 15, 16],
            }
            
            for tab_name, tab_ctx in gen_dfs.items():
                if (tab_name in list(sheet_data_dict.keys())) or \
                    (tab_name in list(sheet_strcol_dict.keys())):
                        logger.warning(f"Generate Tool8 plugin related data with Same Sheet Names.")
                        logger.warning(f"Skip conflict and leave to human check.")
                        continue

                sheet_data_dict[tab_name] = copy.deepcopy(tab_ctx)
                sheet_strcol_dict[tab_name] = []

            add_sheets_and_fill_data_to_xlsm(real_result, sheet_data_dict, sheet_strcol_dict)
            time.sleep(5)
            bsuccess = True

            # 移除中間檔案
            if isfile(self._cp_strict_rawdata):
                try:
                    os.remove(self._cp_strict_rawdata)
                except Exception as e:
                    logger.debug('Remove intermediate files failed.')
                    logger.error(f'Failed message: {e}')
                    pass

        except Exception as e:
            logger.error(f'Export PCMS data to poc_tool8 xlsm failed.')
            logger.error(f'Error: {e}')
            bsuccess = False

        return bsuccess

    def _gen_plugin_data(self, rawdata: pd.DataFrame=None) -> dict:
        # poc tool8 plugin features
        gen_dfs = dict()
        if rawdata is None: return gen_dfs

        logger.debug(f"開始產生 Tool8 Plugin 相關資料 ...")
        start_time = time.time()
        
        # 關鍵字分析
        logger.debug(f"開始產生關鍵字分析 ...")
        from txt_processors.txtrank_v2 import TextRankSummarization
        trs = TextRankSummarization()
        trs.set_stop_words_file(stop_words_file=STOPWORDS)
        sorted_df = trs.run_processing(rawdata)
        gen_dfs["關鍵字分析"] = copy.deepcopy(sorted_df)
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        logger.debug(f"Tool8 Plugin 產生時間: {elapsed_time} 秒")
        return copy.deepcopy(gen_dfs)
