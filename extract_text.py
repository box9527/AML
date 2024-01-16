import os
import pandas as pd
import tabula

class PDFProcessor:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.columns_to_check = ['支出', '存入', '餘額', '備註']
        self.temp_dfs = []

    def process_pdf(self):
        df = tabula.read_pdf(self.pdf_path, area=[120, 5, 800, 1200], pages="all")
        all_df = pd.DataFrame()

        for page_number, page_df in enumerate(df, start=1):
            if not page_df.empty:
                page_df = self.process_page(page_df)
                self.temp_dfs.append(page_df)

        all_df = pd.concat(self.temp_dfs, ignore_index=True)
        all_df = self.drop_unnamed_columns(all_df)
        
        return all_df
        #output_csv = f"{os.path.splitext(self.pdf_path)[0]}.csv"
        #all_df.to_csv(output_csv, index=False)

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

if __name__ == "__main__":
    source_file = "客戶備註很多.pdf"
    pdf_processor = PDFProcessor(source_file)
    pdf_processor.process_pdf()
