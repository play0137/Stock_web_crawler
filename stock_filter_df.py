"""
Original data:公司股市代號對照表.csv

Conditions:
1.單月營收歷月排名 1高
    from 月營收創新高.xlsx
2.負債比 < 40% 
    季度
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E8%B2%A0%E5%82%B5%E7%B8%BD%E9%A1%8D%E4%BD%94%E7%B8%BD%E8%B3%87%E7%94%A2%E6%AF%94%E6%9C%80%E9%AB%98%40%40%E8%B2%A0%E5%82%B5%E7%B8%BD%E9%A1%8D%40%40%E8%B2%A0%E5%82%B5%E7%B8%BD%E9%A1%8D%E4%BD%94%E7%B8%BD%E8%B3%87%E7%94%A2%E6%AF%94%E6%9C%80%E9%AB%98
3.全體董監持股 + 本國法人持股 > 30%
    全體董監
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E5%85%A8%E9%AB%94%E8%91%A3%E7%9B%A3%E6%8C%81%E8%82%A1%E6%AF%94%E4%BE%8B%28%25%29%40%40%E5%85%A8%E9%AB%94%E8%91%A3%E7%9B%A3%40%40%E6%8C%81%E8%82%A1%E6%AF%94%E4%BE%8B%28%25%29
    本國法人
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E6%9C%AC%E5%9C%8B%E6%B3%95%E4%BA%BA%E6%8C%81%E8%82%A1%E6%AF%94%E4%BE%8B%28%25%29%40%40%E6%9C%AC%E5%9C%8B%E6%B3%95%E4%BA%BA%40%40%E6%8C%81%E8%82%A1%E6%AF%94%E4%BE%8B%28%25%29
4.全體董監質押比 < 10%
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E5%85%A8%E9%AB%94%E8%91%A3%E7%9B%A3%E8%B3%AA%E6%8A%BC%E6%AF%94%E4%BE%8B%28%25%29%40%40%E5%85%A8%E9%AB%94%E8%91%A3%E7%9B%A3%40%40%E8%B3%AA%E6%8A%BC%E6%AF%94%E4%BE%8B%28%25%29
    和全體董監持股是相同的資料，僅排序不同
5.毛利率, 營益率, 稅後淨利率上升(三率三升)
    毛利率歷季排名 1 (第一名)
    營益率歷季排名 1
    淨利率歷季排名 1
6.殖利率 > 1%
    現金殖利率
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E7%8F%BE%E9%87%91%E6%AE%96%E5%88%A9%E7%8E%87+%28%E6%9C%80%E6%96%B0%E5%B9%B4%E5%BA%A6%29%40%40%E7%8F%BE%E9%87%91%E6%AE%96%E5%88%A9%E7%8E%87%40%40%E6%9C%80%E6%96%B0%E5%B9%B4%E5%BA%A6
7.多項排列(均線)
    均線 向上
    
    8.現金流量比率 > 0 or 營業現金流量 > 0(skip)
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E6%99%BA%E6%85%A7%E9%81%B8%E8%82%A1&INDUSTRY_CAT=%E6%9C%88K%E7%B7%9A%E7%AA%81%E7%A0%B4%E5%AD%A3%E7%B7%9A%40%40%E6%9C%88K%E7%B7%9A%E5%90%91%E4%B8%8A%E7%AA%81%E7%A0%B4%E5%9D%87%E5%83%B9%E7%B7%9A%40%40%E5%AD%A3%E7%B7%9A
    9.股票尚未經歷大漲大跌(skip)
"""
"""
-利率
-把資訊附在篩選後的後面欄位
-不當成篩選條件了

-均線
-月季年三欄 放在篩選後股票的後面欄位

篩選後的股票每個一個分頁(in stock_filter檔案裡)
分頁裡放
1.個股月營收
2.經營績效-合併報表 單季
抓下來後讓使用者輸入第幾季
只秀出第幾季的資料就好
個股月營收放上面row
經營績效放下面row

-個股月營收程式從stock_web_crawler.py拉出到新的程式寫
-月營收創新高的月份 抓 營收月份 那一欄 判斷是否為month-1
-Mode改成append, engine改成openpyxl
-excel_formatting(openpyxl and xlsxwriter versions)
-freeze the first row and the first, second columns
"""

import pdb
import time

import pandas as pd

import global_vars
from stock_web_crawler import stock_crawler, delete_header, excel_formatting

# global variables
DEBT_RATIO = 40         # 負債比
STAKEHOLDING = 30       # 持股
PLEDGE_RATIO = 10       # 質押比
GROSS_MARGIN = 1        # 毛利率
OPERATING_MARGIN = 1    # 營益率
NET_PROFIT_MARGIN = 1   # 稅後淨利率
DIVIDEND_YIELD = 1      # 現金殖利率

def main():
    file_path = global_vars.DIR_PATH + "公司股市代號對照表.csv"
    stock_ID = list()
    stock_name = list()
    with open(file_path, 'r', encoding="UTF-8") as file_r:
        file_r.readline() # skip the first row
        for line in file_r:
            line = line.split(",")
            stock_ID.append(line[0])
            stock_name.append(line[1])
    df_combine = pd.DataFrame(list(zip(stock_ID, stock_name)), columns=["代號", "名稱"])
    
    file_path = global_vars.DIR_PATH + "月營收創新高.xlsx"
    df = pd.read_excel(file_path)
    df["代號"] = df["代號"].astype(str)
    df = df[["代號", "單月營收歷月排名"]]
    df_combine = pd.merge(df_combine, df, on=["代號"], how="left")
    
    file_path = global_vars.DIR_PATH + "stock_crawler.xlsx"
    dfs = pd.read_excel(file_path, sheet_name=None) # set sheet_name=None if you want to read all sheets
    sheet_names = ["負債比", "全體董監持股_全體董監質押比", "本國法人持股", "全體董監持股_全體董監質押比", "單季營業毛利創歷季新高", "單季營業毛利創歷季新高", "單季稅後淨利創歷季新高", "殖利率", "均線", "均線", "均線"]
    columns = ["負債總額(%)", "全體董監持股(%)", "本國法人(%)", "全體董監質押(%)", "毛利歷季排名", "營益率歷季排名", "淨利率歷季排名", "現金殖利率", "月均線", "季均線", "年均線"]
    for i, sheet_name in enumerate(sheet_names):
        df = dfs[sheet_name]
        df["代號"] = df["代號"].astype(str)
        df = df[["代號", columns[i]]]
        df_combine = pd.merge(df_combine, df, on=["代號"], how="left")

    writer = pd.ExcelWriter(global_vars.DIR_PATH + "stock_filter.xlsx", engine="xlsxwriter")
    df_combine.to_excel(writer, index=False, encoding="UTF-8", sheet_name="All", freeze_panes=(1,2))
    excel_formatting(writer, df_combine, "All")
    
    input_menu() # user menu
    # filters
    df_combine = df_combine[df_combine["單月營收歷月排名"] == "1高"]
    df_combine = df_combine[df_combine["負債總額(%)"] < DEBT_RATIO]
    df_combine = df_combine[(df_combine["全體董監持股(%)"]+df_combine["本國法人(%)"]) > STAKEHOLDING]
    df_combine = df_combine[df_combine["全體董監質押(%)"] < PLEDGE_RATIO]
    df_combine = df_combine[df_combine["現金殖利率"] > DIVIDEND_YIELD]
    df_combine.to_excel(writer, index=False, encoding="UTF-8", sheet_name="Filtered", freeze_panes=(1,2))
    excel_formatting(writer, df_combine, "Filtered")
    
    # pdb.set_trace()
    # infos of filtered stocks in different sheets
    # stocks_ID = ','.join(df_combine["代號"].values)
    # stock_info(stocks_ID, writer)
    
    writer.save()


def stock_info(stocks_ID, writer):
    stocks_ID = "2330,1305" # 實驗暫時用的
    
    global_vars.initialize_proxy()
    stocks_ID = stocks_ID.split(",")
    stock_dict = stock_ID_mapping()
    # headers = ["月別", "開盤", "收盤", "最高", "最低", "漲跌(元)", "漲跌(%)", "單月營收(億)", "單月月增(%)", "單月年增(%)", "累計營收(億)", "累計年增(%)", "合併單月營收(億)", "合併單月月增(%)", "合併單月年增(%)", "合併累計營收(億)", "合併累計年增(%)"]
    table_ID = "#divDetail"
    for stock_ID in stocks_ID:
        url = f"https://goodinfo.tw/StockInfo/ShowSaleMonChart.asp?STOCK_ID={stock_ID}"
        df = stock_crawler(url, None, table_ID)
       
        # reassign headers
        headers = list()
        for i in range(len(df.columns)):
            headers.append('_'.join(pd.Series(df.columns[i]).drop_duplicates().tolist()))
        df.columns = headers
        
        delete_header(df, headers)
        sheet_name = f"{stock_dict[stock_ID]}"
        df.to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, freeze_panes=(1,2)) # write to different sheets
        excel_formatting(writer, df, sheet_name)
        time.sleep(1)

# 1101,台泥,台灣水泥股份有限公司
def stock_ID_mapping():
    stock_dict = dict()
    with open(global_vars.DIR_PATH + "公司股市代號對照表.csv", "r", encoding="UTF-8") as file_r:
        file_r.readline()
        for line in file_r:
            line = line.split(",")
            stock_ID = line[0]
            stock_name = line[1]
            if stock_ID not in stock_dict:
                stock_dict[stock_ID] = stock_name
            if stock_name not in stock_dict:
                stock_dict[stock_name] = stock_ID
    return stock_dict

def input_menu():
    global DEBT_RATIO, STAKEHOLDING, PLEDGE_RATIO, GROSS_MARGIN, OPERATING_MARGIN, NET_PROFIT_MARGIN, DIVIDEND_YIELD

    menu()
    while True:
        input_num = input()
        if input_num == "1":
            DEBT_RATIO = int(input())
        elif input_num == "2":
            STAKEHOLDING = int(input())
        elif input_num == "3":
            PLEDGE_RATIO = int(input())
        elif input_num == "4":
            DIVIDEND_YIELD = int(input())
        elif input_num == "5":
            DEBT_RATIO = 40
            STAKEHOLDING = 30
            PLEDGE_RATIO = 10
            GROSS_MARGIN = 1
            OPERATING_MARGIN = 1
            NET_PROFIT_MARGIN = 1
            DIVIDEND_YIELD = 1
        elif input_num == "6":
            print(f"負債比 < {DEBT_RATIO}%\n"\
                  f"董監+法人持股 > {STAKEHOLDING}%\n"\
                  f"董監質押比 < {PLEDGE_RATIO}%\n"\
                  f"現金殖利率 > {DIVIDEND_YIELD}%\n")
        elif input_num == "7":
            break
        menu()
 
def menu():
    print("輸入參數:",
          f"1.負債比(< {DEBT_RATIO}%)",
          f"2.董監+法人持股(> {STAKEHOLDING}%)",
          f"3.董監質押比(< {PLEDGE_RATIO}%)",
          f"4.現金殖利率(> {DIVIDEND_YIELD}%)",
          "5.預設",
          "6.目前參數值", 
          "7.離開", sep="\n")
 

if __name__ == "__main__":
    main()