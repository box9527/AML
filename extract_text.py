import tabula
import pandas as pd

tabula.environment_info()
# 設定欄位名稱
columns = ["交易日期", "帳務日期", "交易代號", "交易時間", "交易分行 交易櫃員", "摘要", "支出", "存入", "餘額", "轉出入帳號", "合作機構會員編號", "金資序號", "票號", "備註", "註記",""]
# 逐頁處理
df = tabula.read_pdf("pdfone.pdf", area=[120, 5, 800, 1200], pages="all")
# 初始化合併的 DataFrame
result_df = pd.DataFrame()

# 逐頁處理
for page_df in df:
    # 如果有表格，轉換為 DataFrame
    if not page_df.empty:
        # 從第二行開始合併
        page_df = page_df.iloc[1:]

        # 將 DataFrame 合併到結果中
        result_df = pd.concat([result_df, page_df], ignore_index=True)

# 將 DataFrame 寫入 CSV
result_df.to_csv("output.csv", index=False, header = columns)
