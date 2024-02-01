import os
import pandas as pd
import tabula
from tkinter import filedialog

class PDFProcessor:
    def __init__(self, source_folder=None):
        self.source_folder = source_folder
        self.columns_to_check = ['支出', '存入', '餘額', '備註']
        self.all_df = None  # 初始化 all_df 為 None

    def process_folder(self):
        if self.source_folder:
            pdf_files = [f for f in os.listdir(self.source_folder) if f.endswith(".pdf")]

            if not pdf_files:
                print("資料夾中沒有 PDF 檔案。")
                return

        processed_pdf_paths = []  # 新增一個空列表，用於存放處理完的每篇 PDF 檔案的路徑
        for pdf_file in pdf_files:
            print("pdf檔案名稱:",pdf_file)
            pdf_path = os.path.join(self.source_folder, pdf_file)

            # 在每次迭代中初始化 temp_dfs
            temp_dfs = [self.process_pdf(pdf_path)]
            # 將處理完的 PDF 檔案的路徑加入列表
            processed_pdf_paths.append(pdf_path)

        return processed_pdf_paths


    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.pdf_folder = folder_path
            self.process_folder()
            return self.all_df  # 返回 all_df

    def process_pdf(self, pdf_path):
        temp_dfs = []
        df = tabula.read_pdf(pdf_path, area=[120, 5, 800, 1200], pages="all")
        all_df = pd.DataFrame()

        for page_number, page_df in enumerate(df, start=1):
            if not page_df.empty:
                page_df = self.process_page(page_df)
                temp_dfs.append(page_df)

        all_df = pd.concat(temp_dfs, ignore_index=True)
        all_df = self.drop_unnamed_columns(all_df)
        
        return all_df.copy()# 使用 copy 創建新的 DataFrame

    def process_page(self, page_df):
        page_df = page_df.iloc[1:]

        # 處理 "交易分行 交易櫃員" 列
        split_columns = page_df.iloc[:, 4].str.split(expand=True)
        page_df.insert(4, "交易分行", split_columns[0])
        page_df.insert(5, "交易櫃員", split_columns[1])
        page_df = page_df.drop("交易分行 交易櫃員", axis=1)
        page_df = page_df.rename(columns={"帳務日期 交易代號": "帳務日期", "Unnamed: 0": "交易代號"})

        # 處理Unnamed欄位的移動操作
        unnamed_columns = page_df.filter(like='Unnamed:').columns

        # 檢查指定的列是否全為空值
        empty_columns = page_df[self.columns_to_check].columns[page_df[self.columns_to_check].isnull().all()]

        # 檢查Unnamed欄位是否有值
        unnamed_columns_with_values = page_df[unnamed_columns].columns[page_df[unnamed_columns].notnull().any()]
        if len(empty_columns) == len(unnamed_columns_with_values):
            for target_col, col in zip(empty_columns, unnamed_columns_with_values):
                page_df[target_col] = page_df[col]
        else:
            for col in unnamed_columns_with_values:
                col_index = page_df.columns.get_loc(col)
                page_df.iloc[:, col_index - 1] = page_df.iloc[:, col_index]

        page_df = page_df.rename(columns={"支出": "Out", "存入": "In"})

        return page_df

    def drop_unnamed_columns(self, df):
        # 刪除 Unnamed 欄位'
        df = df.drop(df.columns[df.columns.str.contains('Unnamed:')], axis=1)
        return df
