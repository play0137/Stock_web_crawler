import pandas as pd
import os
import sys

import global_vars

dir_path = f"{global_vars.DIR_PATH}近日最高點離扣底日/"
files = [file for file in os.listdir(path=dir_path) if not file.startswith('.')]
for i, file in enumerate(files):
    files[i] = files[i].rstrip('.xlsx')
print("可選檔案:")
print(', '.join(files))
while True:
    closeprice_value = int(input("請輸入\n1.收盤價\n2.成交股數\n")) # 輸入 收盤價 或 成交股數
    if closeprice_value == 1:
        closeprice_value = "收盤價"
        break
    elif closeprice_value == 2:
        closeprice_value = "成交股數"
        break
    else:
        print("請重新輸入")

stockNo = int(input("請輸入股號:"))
print()
file_path = f"{global_vars.DIR_PATH}近日最高點離扣底日/{stockNo}.xlsx"
if not os.path.isfile(file_path):
    sys.exit(f"No {stockNo} file")

dfs = pd.read_excel(file_path, sheet_name=None)
col_names = ["日期", closeprice_value]

df = pd.DataFrame(columns = col_names) # create an empty dataframe
for sheet_name in dfs:
    df = df.append(dfs[sheet_name][{"日期", closeprice_value}], ignore_index=True)
if len(df) < 60:
    sys.exit("資料少於60筆")
    
today_price = df.iloc[len(df)-1].to_list()
df = df.sort_values("日期", ascending=False)
df = df.iloc[:60]
season = df.iloc[59].to_list()

l = list()
if closeprice_value =="收盤價":
    l.append("最高點")
    l.append("今日收盤價")
else:
    l.append("最高成交股數")
    l.append("今日成交股數")

print(f"季扣底:\n{season[0]}, {season[1]}")
df = df.sort_values(closeprice_value, ascending=False)
print()
highest = df.iloc[0]
print(f"{l[0]}:\n{highest[0]}, {highest[1]}")
print()
print(f"{l[1]}:\n{today_price[0]}, {today_price[1]}")    