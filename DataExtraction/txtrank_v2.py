import re
import os
import sys
import pandas as pd
from collections import Counter
from textrank4zh import TextRank4Keyword
from extract_text import PDFProcessor

def resource_path(relative_path):
    try:
        # 如果是打包後的可執行文件
        base_path = sys._MEIPASS
    except Exception:
        # 如果是直接執行的腳本
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

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

    def process_data(self, all_df, output_csv, pdf_path):
        # 這裡你需要根據實際情況調整相依性
        # trs = TextRankSummarization()
        df = all_df.fillna(0)
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

        df['最佳選擇'] = df.apply(self.get_max_keyword, args=(remark_dict_2, remark_dict_3), axis=1)
        count_max = Counter(df['最佳選擇'].explode())
        remark_dict_max = dict(count_max)

        return df, remark_dict_max

    def run_processing(self, source_folder):
        pdf_processor = PDFProcessor(source_folder)
        processed_pdf_paths = pdf_processor.process_folder()

        if processed_pdf_paths:
            for pdf_path in processed_pdf_paths:
                print(f"現在正在分類: {pdf_path} ")
                output_file = f"{os.path.splitext(pdf_path)[0]}.csv"

                all_df = pdf_processor.process_pdf(pdf_path)

                processed_df, remark_dict_max = self.process_data(all_df, output_file, pdf_path)
                print("最佳選擇", remark_dict_max)
                print("最佳選擇總共可分為", len(remark_dict_max), "類分群")

                pdf_processor.all_df = processed_df
