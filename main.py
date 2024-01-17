import os
import glob
import time
import pandas as pd
from extract_text import PDFProcessor
from txtrank_v2 import TextRankSummarization
from collections import Counter

def process_data(input_df, output_csv):
    df = input_df.fillna(0)
    result_df = pd.DataFrame(columns=['備註', '2字關鍵字', '3字關鍵字'])

    df["備註"] = df["備註"].astype(str)
    df = df[df["備註"].str.contains(r'[\u4e00-\u9fa5a-zA-Z]')]

    df['2字關鍵字'] = df['備註'].apply(lambda x: trs.keywords_2(x, topK=1))
    df['3字關鍵字'] = df['備註'].apply(lambda x: trs.keywords_3(x, topK=1))
    result_df = pd.concat([df['備註'], df['2字關鍵字'], df['3字關鍵字']], axis=1)

    count_2 = Counter(result_df['2字關鍵字'].explode())
    count_3 = Counter(result_df['3字關鍵字'].explode())
    remark_dict_2 = dict(count_2)
    remark_dict_3 = dict(count_3)

    df['最佳選擇'] = df.apply(trs.get_max_keyword, args=(remark_dict_2, remark_dict_3), axis=1)
    count_max = Counter(df['最佳選擇'].explode())
    remark_dict_max = dict(count_max)
    
    # 保存 DataFrame 到 CSV 檔案
    #df.to_csv(output_csv, index=False)
    return df, remark_dict_max

if __name__ == "__main__":
    start_time = time.time()
    
    #source_file = "交易備註例1.pdf"
    pdf_folder = "./docs"
    source_file =glob.glob(f"{pdf_folder}/*.pdf")[0]
    pdf_processor = PDFProcessor(source_file)
    all_df = pdf_processor.process_pdf()
    trs = TextRankSummarization()
    output_file = f"{os.path.splitext(source_file)[0]}_processed.csv"
    processed_df, remark_dict_max = process_data(all_df, output_file)
    print("最佳選擇", remark_dict_max)
    print("最佳選擇總共可分為",len(remark_dict_max),"類分群")

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"程式執行時間: {elapsed_time} 秒")
