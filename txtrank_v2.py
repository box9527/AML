import re
import pandas as pd
import sys
from collections import Counter
from textrank4zh import TextRank4Keyword, TextRank4Sentence

STOPWORDS = 'stop_wordsv2.txt'

class TextRankSummarization():
    _instance = None
    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

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

    # 定義一個函數，用來比較 '2字關鍵字' 和 '3字關鍵字' 的次數，選擇次數較大的關鍵字
    def get_max_keyword(self, row):
        count_2 = remark_dict_2.get(row['2字關鍵字'], 0)
        count_3 = remark_dict_3.get(row['3字關鍵字'], 0)
        return row['2字關鍵字'] if count_2 >= count_3 else row['3字關鍵字']

'''讀取資料＋開始分類'''
# 讀取資料
df = pd.read_csv('許多備註outputv2.csv').iloc[:, :15].fillna(0)
result_df = pd.DataFrame(columns=['備註', '2字關鍵字', '3字關鍵字'])
columns = ["交易日期", "帳務日期", "交易代號", "交易時間", "交易分行 交易櫃員", "摘要", "支出", "存入", "餘額", "轉出入帳號", "合作機構會員編號", "金資序號", "票號","備註", "註記"]

df.columns = columns
df["備註"] = df["備註"].astype(str)
df = df[~df["備註"].str.contains("0|0000")]

df['2字關鍵字'] = df['備註'].apply(lambda x: TextRankSummarization().keywords_2(x, topK=1))
df['3字關鍵字'] = df['備註'].apply(lambda x: TextRankSummarization().keywords_3(x, topK=1))
result_df = pd.concat([df['備註'],df['2字關鍵字'],df['3字關鍵字']], axis=1)

# 使用 Counter 計算每個 2 字和 3 字關鍵字的出現次數
count_2 = Counter(result_df['2字關鍵字'].explode())
count_3 = Counter(result_df['3字關鍵字'].explode())

# 轉換成字典格式
remark_dict_2 = dict(count_2)
remark_dict_3 = dict(count_3)

# 新增一個欄位 '最佳選擇'
df['最佳選擇'] = df.apply(TextRankSummarization().get_max_keyword, axis=1)
count_max = Counter(df['最佳選擇'].explode())
remark_dict_max = dict(count_max)

#print("最佳選擇",remark_dict_max)
#print("最佳選擇總共可分為",len(remark_dict_max),"類分群")
#print("2字關鍵字",remark_dict_2) 
#print("2字關鍵字總共可分為",len(remark_dict_2),"類分群")
#print("3字關鍵字",remark_dict_3)
#print("3字關鍵字總共可分為",len(remark_dict_3),"類分群")
