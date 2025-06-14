# -*- coding: utf-8 -*-
"""
港股IPO数据批量补全脚本：提取发行价区间、最终定价、配售结构、绿鞋机制、首日开盘/最高/收盘价等字段。
适用于2019年至今所有公司，基于已整理好的公司列表（含代码、名称、上市日期）。
"""
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import os
import yfinance as yf  # 获取首日开盘/最高/收盘价
from datetime import datetime, timedelta

# 读取本地公司清单（需包含：公司名称、股份代号、上市日期）
base_path = r'C:\Users\19525\Downloads'
flist = ['NLR2020_Chi.xlsx', 'NLR2021_Chi.xlsx', 'NLR2022_Chi.xlsx', 'NLR2023_Chi.xlsx', 'NLR2024_Chi.xlsx']
# 使用列表推导式读取所有xlsx文件
dfs = [pd.read_excel(os.path.join(base_path, fl), header=1) for fl in flist]
# 合并所有DataFrame，并重置索引
companies = pd.concat(dfs, ignore_index=True)

# 输出表格初始化
columns = ["公司名称", "股份代号", "上市日期", "发行价区间", "最终定价",
           "初始公配比例", "初始国际配售比例", "甲组中签率", "乙组中签率",
           "公开申购倍数", "国际申购倍数", "是否触发回拨", "是否触发特别回拨",
           "是否使用绿鞋机制", "首日开盘价", "首日最高价", "首日收盘价",
           "招股书URL", "配售公告URL"]
output = pd.DataFrame(columns=columns)

# 工具函数：构造PDF链接
def construct_urls(code, year):
    base = f"https://www1.hkexnews.hk/app/app_{year}_{code.zfill(5)}"
    return {
        "prospectus": f"{base}/prospectus/cwp_chi.pdf",
        "allotment": f"{base}/allotment/cwp_chi.pdf",
    }

# 工具函数：查询股票首日行情
def get_first_day_prices(code, listing_date):
    try:
        stock = yf.Ticker(f"{code}.HK")
        start_date = datetime.strptime(str(listing_date), "%Y-%m-%d")
        df = stock.history(start=start_date, end=start_date + timedelta(days=2))
        first = df.iloc[0]
        return first["Open"], first["High"], first["Close"]
    except:
        return "", "", ""

# 主循环：逐公司处理
for _, row in companies.iterrows():
    marker = str(row.iloc[-2]).strip()  # 检查这一行最后一列的取值是否为"a"或"b"
    if marker == "(b)":
        # 当前行为国际配售行：除集资额外，其他字段均沿用上一行的记录
        name = prev_name
        code = prev_code
        date = prev_date
        sponsor = prev_sponsor
        reporting_accountant = prev_reporting_accountant
        property_valuer = prev_property_valuer
        amount_of_fund_raised = row["集資額\n(HK$)"]  # 国际集资额
        year = prev_year
        urls = prev_urls
        # 抓取首日行情
        open_p, high_p, close_p = get_first_day_prices(code, date)
        record = {
            "公司名称": name,
            "股份代号": code,
            "上市日期": date,
            "发行价区间": "",
            "最终定价": "",
            "初始公配比例": prev_amount_of_fund_raised,
            "初始国际配售": amount_of_fund_raised,
            "甲组中签率": "",
            "乙组中签率": "",
            "公开申购倍数": "",
            "国际申购倍数": "",
            "是否触发回拨": "",
            "是否触发特别回拨": "",
            "是否使用绿鞋机制": "",
            "首日开盘价": open_p,
            "首日最高价": high_p,
            "首日收盘价": close_p,
            "招股书URL": urls["prospectus"],
            "配售公告URL": urls["allotment"]
        }
        output = pd.concat([output, pd.DataFrame([record])], ignore_index=True)
    elif marker == "(a)":
        # 当前行为香港配售行：直接读取当前行的信息，并更新缓存供后续国际配售行使用
        name = row["上市時公司名稱\n(不包括第二十章下的投資工具個案)"]
        code = str(row["股份代號"]).zfill(4)
        date = row["上市日期\n(日日/月月/年年)"]
        sponsor = row["保薦人"]
        reporting_accountant = row["申報會計師"]
        property_valuer = row["物業估值師"]
        amount_of_fund_raised= row["集資額\n(HK$)"]  # 香港集资额
        year = pd.to_datetime(date).year
        urls = construct_urls(code, year)
        prev_name = name
        prev_code = code
        prev_date = date
        prev_sponsor = sponsor
        prev_reporting_accountant = reporting_accountant
        prev_property_valuer = property_valuer
        prev_amount_of_fund_raised = amount_of_fund_raised
        prev_year = year
        prev_urls = urls
    else:
        # 默认情况：若最后一列既不是"a"也不是"b"，则按香港配售处理
        name = row["上市時公司名稱\n(不包括第二十章下的投資工具個案)"]
        code = str(row["股份代號"]).zfill(4)
        date = row["上市日期\n(日日/月月/年年)"]
        sponsor = row["保薦人"]
        reporting_accountant = row["申報會計師"]
        property_valuer = row["物業估值師"]
        amount_of_fund_raised = row["集資額\n(HK$)"]
        prev_name = name
        prev_code = code
        prev_date = date
        prev_sponsor = sponsor
        prev_reporting_accountant = reporting_accountant
        prev_property_valuer = property_valuer
    
    

# 导出为Excel
output.to_excel("HK_IPO_Full_Dataset.xlsx", index=False)
print("✅ 所有数据处理完毕，已保存 HK_IPO_Full_Dataset.xlsx")
