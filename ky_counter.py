import re
import pandas as pd
from collections import Counter
from textrank4zh import TextRank4Keyword

STOPWORDS = 'stop_wordsv2.txt'

class TextRankSummarization:
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

    def extract_keywords(self, content, word_min_len, count=20):
        if not content:
            return []

        tr4w = TextRank4Keyword(stop_words_file=STOPWORDS)
        tr4w.analyze(text=content, lower=False, window=3)
        keywords_list = tr4w.get_keywords(count, word_min_len=word_min_len)

        if not keywords_list and word_min_len == 3:
            tr4w.analyze(text=content, lower=False, window=3)
            keywords_list = tr4w.get_keywords(count, word_min_len=2)

        if keywords_list:
            return re.sub(r"{'word': '(.+?)'}", r'\1', keywords_list[0]['word'])
        else:
            return content

# 讀取資料
df = pd.read_csv('許多備註outputv2.csv').iloc[:, :15].fillna(0)
result_df = pd.DataFrame(columns=['備註', '2字關鍵字', '3字關鍵字'])

columns = ["交易日期", "帳務日期", "交易代號", "交易時間", "交易分行 交易櫃員", "摘要", "支出", "存入", "餘額", "轉出入帳號", "合作機構會員編號", "金資序號", "票號","備註", "註記"]
df.columns = columns

df["備註"] = df["備註"].astype(str)
df = df[~df["備註"].str.contains("0|0000")]

# 定義函數來提取關鍵字(使用 txtrank)
def get_keywords_txtrank(content, word_min_len, topK=1):
    return TextRankSummarization().extract_keywords(content, word_min_len, count=topK)

df['2字關鍵字'] = df['備註'].apply(lambda x: get_keywords_txtrank(x, word_min_len=2))
df['3字關鍵字'] = df['備註'].apply(lambda x: get_keywords_txtrank(x, word_min_len=3))
result_df = pd.concat([df['備註'], df['2字關鍵字'], df['3字關鍵字']], axis=1)

# 使用 Counter 計算每個 2 字和 3 字關鍵字的出現次數
count_2 = Counter(result_df['2字關鍵字'].explode())
count_3 = Counter(result_df['3字關鍵字'].explode())

# 轉換成字典格式
remark_dict_2 = dict(count_2)
remark_dict_3 = dict(count_3)

dfv2 = df[['2字關鍵字', '3字關鍵字']]
dfv2.to_csv('dfv4.csv', index=False)
