#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


from loguru import logger


def extract_user_info(cash_flow: str='') -> list:
    # cash_flow is a pdf file
    '''
    Performance is bad, just record it but not using.
    '''
    from pdfminer.high_level import extract_text
    text = extract_text(cash_flow)
    texts = text.split('\n')
    for i, t in enumerate(texts):
        if t.startswith('查詢迄日：'):
            break

    return texts

def get_user_info(cash_flow: str='') -> dict:
    # cash_flow is a pdf file
    '''
    Good performance by comparing with pdfminer.high_level
    '''
    import fitz # install using: pip install PyMuPDF
    info = {"客戶": "", "產品別": "", "查詢起日": "", "帳號": "",
            "幣別": "", "查詢迄日": "", "交易內容": ""}
    # 2, 11, 9, 12, 8, 10, 16
    try:
        with fitz.open(cash_flow) as doc:
            texts = list()
            for page in doc:
                text = page.get_text()
                if text and (len(str(text).strip()) > 0):
                    texts.append(text)

                '''
                We only need the user information on the first page
                '''
                break

            if len(texts) > 0:
                cinfo = texts[0].split('\n')
                info["客戶"] = str(cinfo[2]).strip()
                info["產品別"] = str(cinfo[11].split('產品別：')[1]).strip()
                info["查詢起日"] = str(cinfo[9].split('查詢起日：')[1]).strip()
                info["帳號"] = str(cinfo[12]).strip()
                info["幣別"] = str(cinfo[8].split('幣別：')[1]).strip()
                info["查詢迄日"] = str(cinfo[10].split('查詢迄日：')[1]).strip()
                info["交易內容"] = str(cinfo[16]).strip()

    except Exception as e:
        logger.warning(f"Cannot retrieve user information.")
        logger.debug(f"{e}")
        pass

    if len(info["客戶"]) <= 0:
        logger.warning(f"Cannot retrieve user information.")

    return info
