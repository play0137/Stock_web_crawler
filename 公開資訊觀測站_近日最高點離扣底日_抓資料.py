"""
example:
請輸入起訖年月份:202001 202005
"""
import requests
import pandas as pd
from datetime import datetime
import global_vars

stockNo = 2891 # 輸入股號

current_year = datetime.now().year
current_month = datetime.now().month
while True:
    user_input = input("請輸入起訖年月份(e.g., 202001 202005):").split()  
    if len(user_input[0]) == 5:
        split_len = 3
    elif len(user_input[0]) == 6:
        split_len = 4
        
    start_year = int(user_input[0][:split_len])
    end_year = int(user_input[1][:split_len])
    if start_year < 1911:
        start_year += 1911
    if end_year < 1911:
        end_year += 1911
    
    start_month = int(user_input[0][split_len:])
    end_month = int(user_input[1][split_len:])
    
    if (end_year > current_year) or (end_year == current_year and end_month > current_month):
        print("年月份輸入錯誤，請重新輸入")
        continue
    break
    
dates = list()
year = start_year
month = start_month
if start_year != end_year:
    end_month += (end_year-start_year)*12

month = start_month
while month < end_month+1:
    dates.append(f"{year}{month:02}")
    if month == 12:
        year += 1
        if end_month >= 12:
            end_month -= 12
    month += 1
    if month > 12:
        month -= 12
print(dates)

file_name = f"{stockNo}.xlsx"
file_path = global_vars.DIR_PATH + "近日最高點離扣底日/" + file_name
with pd.ExcelWriter(file_path, mode="w", engine="xlsxwriter") as writer:
    for date in dates:
        url = f"https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=html&date={date}01&stockNo={stockNo}"
        data = pd.read_html(requests.get(url).text)[0]
        data.columns = data.columns.droplevel(0)
        sheet_name = str(date)
        data.to_excel(writer, index=False, encoding="BIG5", sheet_name=sheet_name)