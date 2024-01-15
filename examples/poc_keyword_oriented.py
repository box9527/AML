#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import os.path as path
from loguru import logger
import copy
import re
import pandas as pd
import numpy as np
from txtrank_summary import TextRankSummarization
from utils import try_or, stylize_df, isfile
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import tabula


class Productivity:
    def __init__(self):
        self.source = 'source-tool7.pdf'
        self.tool7 = 'source-tool7.xlsm'

        # checkpoints
        self.cp_rawdata = 'rawdata-tool8.xlsx'
        self.cp_strict_rawdata = 'rawdata-strict-tool8.xlsx'
        self.cp_intermediate = 'middle-tool8.xlsx' 
        self.cp_result = 'result-tool8.xlsx'

        # variables
        self.tool7_rawdata_sheet = '原始資料'
        self.tool7_atm_sheet = 'Report_MachineManage'
        # 這裡的 "支出" 與 "存入" 用 "Out" 以及 "In"取代，目的是跟工具七的欄位一致
        self.rawdata_cols = {
            '交易日期':[],'帳務日期':[],'交易代號':[],'交易時間':[],'交易分行':[],'交易櫃員':[],
            '摘要':[],'Out':[],'In':[],'餘額':[],
            '轉出入帳號':[],'合作機構':[],'金資序號':[],'票號':[],'備註':[],'註記':[]}

        self.strict_cols = {
            '交易日期':[],'帳務日期':[],'交易代號':[],'交易時間':[],'交易分行':[],'交易櫃員':[],
            '摘要':[],'支出':[],'存入':[],'餘額':[],
            '轉出入帳號':[],'合作機構':[],'備註':[]}

        self.mid_cols = {
            '關鍵字':[],'備註':[],'摘要':[],'交易日期':[],'交易時間':[],
            '交易分行':[],'交易櫃員':[],'ATM機台據點':[],'支出/Out':[],'存入/In':[]}
        pass

    def output(self, from_pdf: str='', from_tool7: str='') -> bool:
        bsuccess = False
        if len(from_pdf) > 0:
            self.source = from_pdf

        if len(from_tool7) > 0:
            self.tool7 = from_tool7

        if (isfile(self.source) == False) or (isfile(self.tool7) == False):
            return bsuccess

        # checkpoint 1
        bsuccess, rawdata = self._rawdata_from_tool7()
        if bsuccess == False: return bsuccess

        bsuccess, atmdata = self._atm_from_tool7()
        if bsuccess == False: return bsuccess

        # checkpoint 2
        bsuccess, strict_rawdata = self._strict_from_rawdata(rawdata)
        if bsuccess == False: return bsuccess

        # checkpoint 3
        bsuccess, sorted_middata = self._intermediate_product(strict_rawdata, atmdata)
        if bsuccess == False: return bsuccess

        bsuccess, final_data = self._final_product(sorted_middata)
        if bsuccess == False: return bsuccess

        bsuccess = True
        logger.info('Keywords procedure is success.')
        return bsuccess

    def _final_product(self, sorted_data: pd.DataFrame=None) -> (bool, pd.DataFrame):
        bsuccess = False
        exportx = None
        try:
            record = ''
            col_values = list()
            headers = sorted_data.columns.values.tolist()
            col_idx = 0
            col_name = sorted_data.columns[col_idx] # actually, it's "關鍵字"
            old_values = sorted_data[col_name].tolist()
            for val in old_values:
                new_val = val
                if record != val:
                    record = val
                else:
                    new_val = ''

                col_values.append(new_val)

            export = sorted_data.copy(deep=True)
            export.drop(col_name, axis = 1, inplace = True)
            export.insert(col_idx, col_name, col_values)

            exportx = export.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)
            (
                pd.DataFrame([exportx.to_dict('list')])
                    .apply(pd.Series.explode)
                    .pivot_table(index=headers, sort=False)
                    .style.applymap_index(stylize_df)
                    .to_excel(self.cp_result, startrow=-1)
            )
            bsuccess = True
        except:
            logger.debug('Create final_product failed.')
            pass

        return bsuccess, exportx

    def _intermediate_product(self, rawdata: pd.DataFrame=None, atmdata: pd.DataFrame=None) -> (bool, pd.DataFrame):
        bsuccess = False
        exportx = None
        # Text4Rank
        metadata = TextRankSummarization()
        similar_max = float(1)
        similar_min = float(0)
        similar_threshold = float(0.44) # 相似閥值拉在 0.44
        tfidf_vectorizer = TfidfVectorizer(analyzer="char")
        ksorts = dict()
        def builder(row, atmdata, kcounts):
            comment = try_or(lambda:f"{row['備註']}".strip(),default=f"{row['備註']}")
            kitems = metadata.keywords(comment, count=1)
            keyword = try_or(lambda:f"{kitems[0]['word'].strip().lower()}",default='')

            summary = try_or(lambda:f"{row['摘要']}".strip(),default=f"{row['摘要']}")
            deal_date = try_or(lambda:f"{row['交易日期']}".strip(),default=f"{row['交易日期']}")
            deal_time = try_or(lambda:f"{row['交易時間']}".strip(),default=f"{row['交易時間']}")
            deal_code = try_or(lambda:f"{row['交易代號']}".strip(),default=f"{row['交易代號']}")
            deal_branch = try_or(lambda:f"{row['交易分行']}".strip(),default=f"{row['交易分行']}")
            deal_teller = try_or(lambda: f"{row['交易櫃員']}".strip(),default=f"{row['交易櫃員']}")
            atm_addr = try_or(lambda:f"{atmdata[atmdata['代碼(記事本)'] == deal_teller]['地址-區域'].to_list()[0]}",default=deal_teller)
            m_out = try_or(lambda:f"{row['支出']}".strip(),default=f"{row['支出']}")
            m_in = try_or(lambda:f"{row['存入']}".strip(),default=f"{row['存入']}")

            sort_idx = f'{deal_date}={deal_time}={deal_branch}={deal_teller}={comment}'
            if len(keyword) > 0:
                if keyword not in list(kcounts.keys()):
                    kcounts[keyword] = list()
                kcounts[keyword].append(sort_idx)

            return keyword, comment, summary, deal_date, deal_time,\
                    deal_branch, deal_teller, atm_addr, m_out, m_in

        try:
            export = pd.DataFrame.from_dict( self.mid_cols )
            export['關鍵字'], export['備註'], export['摘要'], export['交易日期'],\
            export['交易時間'], export['交易分行'],\
            export['交易櫃員'], export['ATM機台據點'], export['支出/Out'], export['存入/In'] \
            = zip(*rawdata.apply(lambda x: builder(x, atmdata, ksorts), axis=1))

            ksorted_keys = sorted(ksorts, key=lambda k: len(ksorts[k]), reverse=True)

            exportx = pd.DataFrame.from_dict( self.mid_cols )
            for kw in ksorted_keys:
                kidxs = ksorts[kw]
                for kidx in kidxs:
                    deal_date, deal_time, deal_branch, deal_teller, comment = '', '', '', '', ''
                    [deal_date, deal_time, deal_branch, deal_teller, comment] = kidx.split('=')

                    temp = export[(export['交易日期'] == deal_date) & \
                                (export['交易時間'] == deal_time) & \
                                (export['交易分行'] == deal_branch) & \
                                (export['交易櫃員'] == deal_teller) & \
                                (export['備註'] == comment)]

                    exportx.loc[len(exportx.index)] = temp.iloc[0].tolist()

            exportx = exportx.replace(np.nan, '', regex=True)
            with pd.ExcelWriter(self.cp_intermediate) as writer:
                exportx.to_excel(writer, sheet_name="Keywords", index=False)
            bsuccess = True
        except:
            logger.debug('Create intermediate_product failed.')
            pass

        return bsuccess, exportx

    def _atm_from_tool7(self) -> (bool, pd.DataFrame):
        bsuccess = False
        exportx = None
        def atm_formatting(row):
            id_code = try_or(lambda:f"{row['ID1']:08d}",default=f"{row['ID1']}")
            deal_code = try_or(lambda:f"{row['剖析']:04d}",default=f"{row['剖析']}")
            deal_teller = try_or(lambda:f"{row['代碼(記事本)']:05d}",default=f"{row['代碼(記事本)']}")
            return id_code, deal_code, deal_teller

        try:
            # 這個是從 工具七 裡讀出 "Report_MachineManage" 頁
            export = pd.read_excel(self.tool7, sheet_name=self.tool7_atm_sheet, skiprows=-1)
            export.columns = export.columns.str.split('\\n').str[0]
            export['ID1'], export['剖析'], export['代碼(記事本)'] = zip(*export.apply(atm_formatting, axis=1))
            exportx = export.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)
            bsuccess = True
        except:
            logger.debug('Extract ATM information from tool7 failed.')
            pass

        return bsuccess, exportx

    def _strict_from_rawdata(self, rawdata: pd.DataFrame=None) -> (bool, pd.DataFrame):
        bsuccess = False
        exportx = None
        def rawdata_formatting(row):
            deal_date = try_or(lambda:f"{row['交易日期'].year:04d}/{row['交易日期'].month:02d}/{row['交易日期'].day:02d}",default=f"{row['交易日期']}")
            acc_date = try_or(lambda:f"{row['帳務日期'].year:04d}/{row['帳務日期'].month:02d}/{row['帳務日期'].day:02d}",default=f"{row['帳務日期']}")
            deal_code = try_or(lambda:f"{row['交易代號']:04d}",default=f"{row['交易代號']}")
            deal_time = try_or(lambda:f"{row['交易時間'].hour:02d}:{row['交易時間'].minute:02d}:{row['交易時間'].second:02d}",default=f"{row['交易時間']}")
            deal_branch = try_or(lambda:f"{row['交易分行']:03d}",default=f"{row['交易分行']}")
            deal_teller = try_or(lambda:f"{row['交易櫃員']:05d}",default=f"{row['交易櫃員']}")
            summary = try_or(lambda:f"{row['摘要']}".strip(),default=f"{row['摘要']}")
            m_out = try_or(lambda:f"{row['Out']:,.2f}",default=f"{row['Out']}")
            m_in = try_or(lambda:f"{row['In']:,.2f}",default=f"{row['In']}")
            m_balance = try_or(lambda:f"{row['餘額']:,.2f}",default=f"{row['餘額']}")
            tr_acc = try_or(lambda:f"{row['轉出入帳號']}".strip(),default=f"{row['轉出入帳號']}")
            tr_infra = try_or(lambda:f"{row['合作機構']}".strip(),default=f"{row['合作機構']}")
            comment = f"{row['備註']}".strip()

            return deal_date, acc_date, deal_code, deal_time, deal_branch, deal_teller, \
                    summary, m_out, m_in, m_balance, tr_acc, tr_infra, comment

        try:
            if rawdata is None:
                return bsuccess, exportx

            export = pd.DataFrame( self.strict_cols )
            export['交易日期'], export['帳務日期'], export['交易代號'],\
            export['交易時間'], export['交易分行'], export['交易櫃員'],\
            export['摘要'], export['支出'], export['存入'],\
            export['餘額'], export['轉出入帳號'], export['合作機構'],\
            export['備註'] \
            = zip(*rawdata.apply(rawdata_formatting, axis=1))

            export = export[(((export['支出'] == export['支出']) & (export['支出'] != '')) |
            ((export['存入'] == export['存入']) & (export['存入'] != ''))) &
            (export['備註'] != '')]

            exportx = export.copy(deep=True)
            exportx = exportx.replace(np.nan, '', regex=True)
            with pd.ExcelWriter(self.cp_strict_rawdata) as writer:
                exportx.to_excel(writer, sheet_name="Rawdata", index=False)
            bsuccess = True
        except:
            logger.debug('Strict rawdata failed.')
            pass

        return bsuccess, exportx

    def _rawdata_from_tool7(self) -> (bool, pd.DataFrame):
        bsuccess = False
        exportx = None
        try:
            export = pd.DataFrame.from_dict( self.rawdata_cols )
            dfs = tabula.read_pdf(self.source, area=[120, 5, 800, 1200], pages="all")
            for df in dfs:
                df = df.replace(np.nan, '', regex=True)

                headers = df.columns.values.tolist()
                wrong_idx = [i for i, item in enumerate(headers) if re.search('^Unnamed:', item)]
                (rows, cols) = df.shape

                for row in range(rows):
                    if row <= 0:
                        continue

                    refine1 = []
                    values = df.iloc[row].tolist()

                    next_skip = False
                    for idx, item in enumerate(values):
                        if (idx in wrong_idx) and (idx <= 12):
                            if len(str(item).strip()) <= 0:
                                continue
                            if (len(refine1) > 0) and (len(refine1[-1]) <= 0):
                                del refine1[-1]

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
                        except:
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
            with pd.ExcelWriter(self.cp_rawdata) as writer:
                exportx.to_excel(writer, sheet_name="Rawdata", index=False)
            bsuccess = True
        except:
            logger.debug('Extract rawdata from PDF failed.')
            pass

        return bsuccess, exportx


if __name__ == '__main__':
    source = 'source-tool7.pdf'
    tool7 = 'source-tool7.xlsm'
    product = Productivity()
    bsuccess = product.output(source, tool7)
    if bsuccess == False:
        logger.critical('Keywords procedure is failed.')

