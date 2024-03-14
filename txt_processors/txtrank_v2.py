#!/usr/bin/env python3
# coding: utf-8
# @voneyay


import re
import os
import sys
import pathlib
import pandas as pd
import importlib
from typing import Any
from collections import Counter
from loguru import logger
from textrank4zh import TextRank4Keyword


class TextRankSummarization():
    def __init__(self):
        try:
            importlib.reload(sys)
            self._stop_words_file = None
        except Exception as e:
            logger.critical(f"TextRankSummarization initial fail: {e}")
        pass

    def set_stop_words_file(self, stop_words_file: Any):
        if stop_words_file is not None:
            self._stop_words_file = stop_words_file

    def extract_keywords(self, content: str = None, count: int = 20, word_min_len: int = 2, topK: int = 1) -> Any:
        try:
            tr4w = TextRank4Keyword(stop_words_file=self._stop_words_file)
            tr4w.analyze(text=content, lower=False, window=3)  # 2)
            keywords_list = tr4w.get_keywords(count, word_min_len=word_min_len)

            # 如果未找到指定長度的關鍵字，再嘗試較小的長度
            if not keywords_list and word_min_len == 3:
                tr4w.analyze(text=content, lower=False, window=3)  # 2)
                keywords_list = tr4w.get_keywords(count, word_min_len=2)

            if keywords_list:
                return re.sub(r"{'word': '(.+?)'}", r'\1', keywords_list[0]['word'])
        except:
            pass

        return content

    def keywords_2(self, content: str = None, count: int = 20, topK: int = 1):
        return self.extract_keywords(content, count, word_min_len=2, topK=topK)

    def keywords_3(self, content: str = None, count: int = 20, topK: int = 1):
        return self.extract_keywords(content, count, word_min_len=3, topK=topK)

    def get_max_keyword(self, row, remark_dict_2, remark_dict_3):
        count_2 = remark_dict_2.get(row['2字關鍵字'], 0)
        count_3 = remark_dict_3.get(row['3字關鍵字'], 0)
        return row['2字關鍵字'] if count_2 >= count_3 else row['3字關鍵字']

    def process_data(self, data: pd.DataFrame=None):
        df = data.fillna(0)
        df = df.iloc[6:].reset_index(drop=True) 
        df.columns = df.iloc[0] 
        df.iloc[0] = df.columns 

        result_df = pd.DataFrame(columns=['備註', '2字關鍵字', '3字關鍵字'])

        df["備註"] = df["備註"].astype(str)
        df = df[df["備註"].str.contains(r'[\u4e00-\u9fa5a-zA-Z]')]

        df['2字關鍵字'] = df['備註'].apply(lambda x: self.keywords_2(x, topK=1))
        df['3字關鍵字'] = df['備註'].apply(lambda x: self.keywords_3(x, topK=1))

        result_df = pd.concat([df['備註'], df['2字關鍵字'], df['3字關鍵字']], axis=1)

        count_2 = Counter(result_df['2字關鍵字'].explode())
        count_3 = Counter(result_df['3字關鍵字'].explode())
        remark_dict_2 = dict(count_2)
        remark_dict_3 = dict(count_3)

        df['關鍵字'] = df.apply(self.get_max_keyword, args=(remark_dict_2, remark_dict_3), axis=1)
        count_max = Counter(df['關鍵字'].explode())
        remark_dict_max = dict(count_max)

        # Remove dummy row
        if df and (len(df) > 0): df = df.iloc[1:, :]

        return df, remark_dict_max

    def run_processing(self, data: pd.DataFrame=None):
        processed_df, remark_dict_max = self.process_data(data)

        selected_columns = ['關鍵字', '備註', '摘要', '交易日期', '交易時間', '交易分行', '交易櫃員', 'Out', 'In']
        selected_columns_df = processed_df[selected_columns]

        sorted_keys = sorted(remark_dict_max, key=remark_dict_max.get, reverse=True)
        selected_columns_df = selected_columns_df.copy()
        selected_columns_df['關鍵字'] = pd.Categorical(selected_columns_df['關鍵字'], categories=sorted_keys, ordered=True)
        sorted_df = selected_columns_df.sort_values(by='關鍵字')
   
        return sorted_df
