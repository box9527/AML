import pandas as pd
import numpy as np

# 讀取資料，空值補0
df = pd.read_csv("許多備註output.csv", decimal=".").fillna(0)# 指定小數點符號為句號

#將 "支出、存入" 列中的逗號刪除並轉換為浮點數
df["支出"] = df["支出"].str.replace(",", "").astype(float)
df["存入"] = df["存入"].str.replace(",", "").astype(float)

# 定義類別與相對應的關鍵字
categories = {
    "放貸/會錢": ["日會","會錢","貸款","利息","借款","還款","結清","本金"],
    "虛擬貨幣": ["英屬維京群島商","現代財富科技有","凱基商業銀行受","幣","USTD","binance","BINGX"],
    "地下匯兌/代儲": ["兌換","換幣","草","RMB","rmb","人民幣","USD","幫幫寶","支付寶"],
    "遊戲/代儲": ["遊戲","代儲","星城","王者","天堂","天2","鑽","儲值","TIKTOK","抖音","微信"],
    "電子菸": ["電子菸","彈","煙","菸","油","IQOS","iqos"],
    "網拍/代購": ["鞋","包","貨款","香奈兒","香水","衣服","手錶","皮帶","代購","Gucci","LV","Prada","精品"],
    "股票/代操": ["股票","股款","配股","配息","認股","交割"],
    "台彩": ["台彩","刮","發財金","營收","兌獎金額","銷售佣金","銷售額度"],
    "運彩/博弈": ["運彩","牌","分析費","世足","贏","發財金","投資","紅利","重注","娛樂","賽事","MITRADE"]
}

# 創建一個空的 DataFrame 來存儲結果
result_df = pd.DataFrame(columns=["交易活動","關鍵字", "支出", "支出次數", "存入", "存入次數"])

# 使用迴圈處理每個關鍵字
for category,keywords in categories.items():
    for keyword in keywords:

        # 選擇符合關鍵字條件的資料行
        selected_data = df[df["備註"].str.contains(keyword, na=False)]
    
        # 計算支出和存入的總和
        total_expense = selected_data["支出"].sum()
        total_income = selected_data["存入"].sum()
        
        # 初始化支出和存入的次數
        expense_occurrences = 0        
        income_occurrences = 0
        # 計算出現的次數
        expense_occurrences += selected_data["支出"].count() if total_expense > 0 else 0
        income_occurrences += selected_data["存入"].count() if total_income > 0 else 0
    
        # 將結果追加到結果 DataFrame
        result_df = result_df.append({
            "交易活動": category,
            "關鍵字": keyword,
            "支出": total_expense,
            "支出次數": expense_occurrences,
            "存入": total_income,
            "存入次數": income_occurrences
        }, ignore_index=True)

# 顯示最終的結果
print(result_df)
result_df.to_csv("result_df.csv", index=False)
