import re, os
import pandas as pd
from collections import Counter
from textrank4zh import TextRank4Keyword
import sys
import os,sys

def resource_path(relative_path):
    try:
        # 如果是打包後的可執行文件
        base_path = sys._MEIPASS
    except Exception:
        # 如果是直接執行的腳本
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
#STOPWORDS = 'stop_wordsv2.txt'
STOPWORDS = resource_path('stop_wordsv2.txt')

class TextRankSummarization():
    def __init__(self):
        try:
            import importlib
            importlib.reload(sys)
        except:
            pass

    def extract_keywords(self, content: str = None, count: int = 20, word_min_len: int = 2, topK: int = 1):
        if not content:
            return []

        tr4w = TextRank4Keyword(stop_words_file=STOPWORDS)
        tr4w.analyze(text=content, lower=False, window=3)  # 2)
        keywords_list = tr4w.get_keywords(count, word_min_len=word_min_len)

        # 如果未找到指定長度的關鍵字，再嘗試較小的長度
        if not keywords_list and word_min_len == 3:
            tr4w.analyze(text=content, lower=False, window=3)  # 2)
            keywords_list = tr4w.get_keywords(count, word_min_len=2)

        if keywords_list:
            return re.sub(r"{'word': '(.+?)'}", r'\1', keywords_list[0]['word'])
        else:
            return content

    def keywords_2(self, content: str = None, count: int = 20, topK: int = 1):
        return self.extract_keywords(content, count, word_min_len=2, topK=topK)

    def keywords_3(self, content: str = None, count: int = 20, topK: int = 1):
        return self.extract_keywords(content, count, word_min_len=3, topK=topK)

    def get_max_keyword(self, row, remark_dict_2, remark_dict_3):
        count_2 = remark_dict_2.get(row['2字關鍵字'], 0)
        count_3 = remark_dict_3.get(row['3字關鍵字'], 0)
        return row['2字關鍵字'] if count_2 >= count_3 else row['3字關鍵字']
