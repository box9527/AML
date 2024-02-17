#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


# Import the required Module
import tabula
from loguru import logger
import copy
import pandas as pd 
from collections import OrderedDict
from txtrank_summary import TextRankSummarization
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np


'''
Text4Rank
'''
g_metadata = TextRankSummarization()

'''
這個函式用來快速遞迴 dataframe, 
會比 pd.DataFrame.iteritems 快，尤其在巨大的 dataframe 情況下會越加明顯
'''
def cleaner(s, counts, kwords, kcounts, scounts):
    '''
    避開header的部分以及備註欄位是Nan的
    '''
    if (s['備註'] == 'nan') or (s['備註'] != s['備註']):
        return None, None, None

    idx1 = str(s['交易日期']).replace('/', '').strip() if len(str(s['交易日期'])) > 0 else ''
    idx2 = str(s['帳務日期 交易代號']).replace('/', '').strip() if len(str(s['帳務日期 交易代號'])) > 0 else ''
    idx4 = str(int(s['Unnamed: 0'])).strip() if len(str(int(s['Unnamed: 0']))) > 0 else '' # 交易代號，特殊處理
    idx3 = str(s['交易時間']).replace(':', '').strip() if len(str(s['交易時間'])) > 0 else ''
    idx5 = str(s['摘要']).strip() if len(str(s['摘要'])) > 0 else ''
    idx = f'{idx1}-{idx2}-{idx3}-{idx4}'
    comment = s['備註'].strip()

    '''
    利用Text4Rank來取得TF-IDF 權重最高的關鍵字
    '''
    kitems = g_metadata.keywords(comment, count=1) # count=1, 這裡只取權重最高的關鍵字使用，回傳一個list

    klist = '' # keyword list, 用來處理第一個sheet 中的關鍵字
    for ki in kitems:
        if ('word' not in list(ki.keys()) ) or (len(ki['word'].strip()) <= 0):
            continue
        if len(klist) > 0:
            klist += '\n'
        kw = ki['word'].strip().lower()
        klist += kw

        if kw not in list(kcounts.keys()):
            kcounts[kw] = list()
        kcounts[kw].append(f'{comment}')

        if kw not in list(scounts.keys()):
            scounts[f'{kw}'] = list()
        scounts[f'{kw}'].append(f'{idx5}')

    if comment not in list(counts.keys()):
        counts[comment] = list()
    counts[comment].append(idx)

    if comment not in list(kwords.keys()):
        kwords[comment] = ''
    kwords[comment] = klist

    return idx, comment, klist

source_pdf = f'../docs/許多備註.pdf'
middle_csv = f'middle_csv.csv'
sink_excel = f'sink_excel.xlsx'

# Step1, Read a PDF File
dtls = {'d_idx': [], 'd_comment': [], 'd_keywords': []}
result_df = pd.DataFrame(copy.deepcopy(dtls)) # initial variable
df_list = tabula.read_pdf(source_pdf, pages='all') # [0] , 0 is the first page

# Step2, Convert PDF into CSV， for debugging
# 這裡我的做法會少最後一頁或最後幾行
tabula.convert_into(source_pdf, middle_csv, output_format="csv", pages='all')

# Step3, Concat，將所有的page 的df 合起來
count_dict = dict() # 以備註為key, 關聯到的交易資訊為value, 用來測試依照關聯到的交易筆數多寡來排序
kword_dict = dict() # 以備註為key, 關聯到的TF-IDF 截取的關鍵字"字串(A\nB\nC...)"為value。但因為只取count=1, 所以只有一個
kcount_dict = dict() # 以TF-IDF 截取的關鍵字為key, 擷取來源備註為value
scount_dict = dict() # 以TF-IDF 截取的關鍵字為key, 擷取來源的交易資訊為value
for idx, df_raw in enumerate(df_list):
    df = pd.DataFrame(copy.deepcopy(dtls))
    df['d_idx'], df['d_comment'], df['d_keywords'] = zip(*df_raw.apply(lambda x: cleaner(x, count_dict, kword_dict, kcount_dict, scount_dict), axis=1))
    df.dropna(subset=['d_comment'], inplace=True)
    result_df = pd.concat([result_df, df], ignore_index=True)

# Step4, 測試排序
logger.debug(f'%%%%%%%%%%%%%%%%%%%%%%')
res1 = '\n'.join(sorted(count_dict, key=lambda key: len(count_dict[key]), reverse=True))
logger.debug(f"Sorted keys by value list : {res1}")
logger.debug(f'%%%%%%%%%%%%%%%%%%%%%%')
res2 = OrderedDict(sorted(count_dict.items(), key = lambda x : len(x[1]), reverse=True)).keys()
logger.debug(f"Sorted keys by value list : {res2}")
logger.debug(f'%%%%%%%%%%%%%%%%%%%%%%')
res3 = [k for k, v in sorted(count_dict.items(), key=lambda item: len(item[1]), reverse=True)]
logger.debug(f"Sorted keys by value list : {res3}")
logger.debug(f'%%%%%%%%%%%%%%%%%%%%%%')
res4 = [k for _, k in sorted(
    zip(map(len, count_dict.values()), count_dict.keys()), reverse=True)]
logger.debug(f"Sorted keys by value list: {res4}")
logger.debug(f'%%%%%%%%%%%%%%%%%%%%%%')
res5 = [k for _, k in sorted(
    zip(map(len, kcount_dict.values()), kcount_dict.keys()), reverse=True)]
logger.debug(f"Sorted keys by value list: {res5}")
logger.debug('%%%%%%%%%%%%%%%%%%%%%%')

# Step5-1, 組合第一個sheet 的內容
dtls = {'備註': [], '關鍵字': [], '交易日期-帳務日期-交易時間-交易代號': []}
for k in res4:
    dtls['備註'].append(k)
    dtls['關鍵字'].append(kword_dict[k])
    dtls['交易日期-帳務日期-交易時間-交易代號'].append('\n'.join(count_dict[k]))

comment_df = pd.DataFrame(copy.deepcopy(dtls))

# Step5-2, 組合第二個sheet 的內容
# 初始化
dtls = {'關鍵字': [], '備註': [], '計數': [], '相似關鍵字': [], '相似備註': [], '摘要': []}
for k in res5:
    dtls['關鍵字'].append(k)
    clist = kcount_dict[k]
    dtls['備註'].append('\n'.join(clist))
    dtls['計數'].append(str(len(clist)))

    dtls['相似關鍵字'].append('')
    dtls['相似備註'].append('')

    slist = scount_dict[k]
    dtls['摘要'].append('\n'.join(slist))

# 取得相似關鍵字
similar_max = float(1)
similar_min = float(0)
similar_threshold = float(0.44) # 相似閥值拉在 0.44
kw_list = dtls['關鍵字']
tfidf_vectorizer = TfidfVectorizer(analyzer="char")

for idx, kw in enumerate(kw_list):
    sparse_matrix = tfidf_vectorizer.fit_transform([kw]+kw_list)
    cosine = cosine_similarity(sparse_matrix[0,:],sparse_matrix[1:,:])
    tmppd = pd.DataFrame({'cosine':cosine[0],'strings':kw_list}).sort_values('cosine',ascending=False)

    tmp_cosine = tmppd['cosine'].to_list()
    tmp_strings = tmppd['strings'].to_list()

    similar_ks = '' # 相似關鍵字
    similar_ks_list = []
    similar_cm = '' # 相似關鍵字關聯到的備註==相似備註
    similar_cm_list = []
    for idx_cos, cos in enumerate(tmp_cosine):
        if float(cos) >= similar_max: continue
        if float(cos) <= similar_threshold: continue # 這裡相似閥值拉在 0.44
        if float(cos) > similar_min:
            if len(similar_ks) > 0: similar_ks += '\n'
            if len(similar_cm) > 0: similar_cm += '\n'
            similar_ks_list.append( str(tmp_strings[idx_cos]) + ',' + str(cos) )
            similar_cm_list.append( str( kcount_dict[str(tmp_strings[idx_cos])] ) )

    if (len(similar_ks_list) > 0) and (len(similar_cm_list) > 0):
        similar_ks = '\n'.join(similar_ks_list)
        similar_cm = '\n'.join(similar_cm_list)

    if len(similar_ks) > 0:
        dtls['相似關鍵字'][idx] = similar_ks 
        dtls['相似備註'][idx] = similar_cm
        
kwords_df = pd.DataFrame(copy.deepcopy(dtls))

# Step6, Convert to excel
with pd.ExcelWriter(sink_excel) as writer:
    comment_df.to_excel(writer, sheet_name="Comments", index=False)
    kwords_df.to_excel(writer, sheet_name="Keywords", index=False)

exit()
