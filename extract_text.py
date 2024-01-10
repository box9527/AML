import tabula
import pandas as pd

tabula.environment_info()
# 設定欄位名稱
columns = ["交易日期", "帳務日期", "交易代號", "交易時間", "交易分行 交易櫃員", "摘要", "支出", "存入", "餘額", "轉出入帳號", "合作機構會員編號", "金資序號", "票號", "備註", "註記"]

df = tabula.read_pdf("客戶備註很多.pdf", area=[120, 5, 800, 1200], pages="all")
# 初始化合併的 DataFrame
all_df = pd.DataFrame()

# 逐頁處理
for page_df in df:
    # 如果有表格，轉換為 DataFrame
    if not page_df.empty:
        # 從第二行開始合併
        page_df = page_df.iloc[1:]
        # 將 DataFrame 合併到結果中
        all_df = pd.concat([all_df, page_df], ignore_index=True)

# 遍歷整個 DataFrame
for index, row in all_df.iterrows():
    # 檢查第 15 或第 16 欄位是否為空值
    if pd.notna(row[15]):
        all_df.iloc[index, 13] = row[15]
    elif len(row) > 16 and pd.notna(row[16]):
        all_df.iloc[index, 13] = row[16]
    else:
        pass

# 前15個欄位
result_df = all_df.iloc[:, :15]
# 將 DataFrame 寫入 CSV
result_df.to_csv("客戶備註很多完整版.csv", index=False, header = columns)
