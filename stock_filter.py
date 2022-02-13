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
import sys
import pdb
import time
import pandas as pd
import random
from datetime import datetime

import global_vars
from stock_web_crawler import stock_crawler, delete_header, excel_formatting
from stock_info import stock_ID_name_mapping

# global variables
DEBT_RATIO = 40         # 負債比
STAKEHOLDING = 30       # 持股
PLEDGE_RATIO = 10       # 質押比
GROSS_MARGIN = 20       # 毛利率
OPERATING_MARGIN = 20   # 營益率
NET_PROFIT_MARGIN = 20  # 稅後淨利率
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
    try:
        last_month = datetime.now().month-1
        if last_month <= 0:
            last_month = 12
        df = pd.read_excel(file_path, sheet_name=f"{last_month}月")
    except ValueError as ve:
        print("ValueError:", ve)
        sys.stderr.write("Please excute stock_web_crawler.py first.\n")
        sys.exit(0)
    df["代號"] = df["代號"].astype(str)
    df = df[["代號", "單月營收歷月排名"]]
    df_combine = pd.merge(df_combine, df, on=["代號"], how="left")
    
    file_path = global_vars.DIR_PATH + "stock_crawler.xlsx"
    dfs = pd.read_excel(file_path, sheet_name=None) # set sheet_name=None if you want to read all sheets
    sheet_names = ["負債比", "全體董監持股_全體董監質押比", "本國法人持股", "全體董監持股_全體董監質押比", "單季營業毛利創歷季新高", "單季營業毛利創歷季新高", "單季稅後淨利創歷季新高", "殖利率", "均線", "均線", "均線"]
    columns = ["負債總額(%)", "全體董監持股(%)", "本國法人(%)", "全體董監質押(%)", "毛利率歷季排名", "營益率歷季排名", "淨利率歷季排名", "現金殖利率", "月均線", "季均線", "年均線"]
    for i, sheet_name in enumerate(sheet_names):
        df = dfs[sheet_name]
        df["代號"] = df["代號"].astype(str)
        df = df[["代號", columns[i]]]
        df_combine = pd.merge(df_combine, df, on=["代號"], how="left")

    with pd.ExcelWriter(global_vars.DIR_PATH + "stock_filter.xlsx", engine="xlsxwriter") as writer:
        df_combine.to_excel(writer, index=False, encoding="UTF-8", sheet_name="All", freeze_panes=(1,2))
        excel_formatting(writer, df_combine, "All")
        
        input_menu() # user menu
        # filters
        df_combine = df_combine[df_combine["單月營收歷月排名"] == "1高"]
        df_combine = df_combine[df_combine["負債總額(%)"] <= DEBT_RATIO]
        df_combine = df_combine[(df_combine["全體董監持股(%)"]+df_combine["本國法人(%)"]) >= STAKEHOLDING]
        df_combine = df_combine[df_combine["全體董監質押(%)"] <= PLEDGE_RATIO]
        # df_combine = df_combine[df_combine["毛利率歷季排名"] <= GROSS_MARGIN]
        # df_combine = df_combine[df_combine["營益率歷季排名"] <= OPERATING_MARGIN]
        # df_combine = df_combine[df_combine["淨利率歷季排名"] <= NET_PROFIT_MARGIN]
        df_combine = df_combine[df_combine["現金殖利率"] >= DIVIDEND_YIELD]
        df_combine.to_excel(writer, index=False, encoding="UTF-8", sheet_name="Filtered", freeze_panes=(1,2))
        excel_formatting(writer, df_combine, "Filtered")
        
        # info of filtered stocks in different sheets
        stocks_ID = ','.join(df_combine["代號"].values)
        if not stocks_ID:
            sys.exit("無符合篩選條件的股票, 請重新調整參數")
        stock_dict = stock_ID_name_mapping()
        print("The size of filtered  stocks:", len(stocks_ID.split(',')))
        print("filtered stocks:")
        for stock_ID in stocks_ID.split(','):
            print(f"{stock_dict[stock_ID]}({stock_ID})")
        stock_info(stocks_ID, writer)
    
    # write filter results to another file that contain history data
    file_path = global_vars.DIR_PATH + "stock_filter_history.xlsx"
    sheet_name = f"{last_month}月"
    with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        df_combine.to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, freeze_panes=(1,2))
        excel_formatting(writer, df_combine, sheet_name)
 

# the info of monthly revenue and consollidated financial statements
def stock_info(stocks_ID, writer):
    global_vars.initialize_proxy()
    # stocks_ID = "1210,1215,2069,2221,2241" # 實驗暫時用的
    
    global_vars.initialize_proxy()
    stocks_ID = stocks_ID.split(",")
    stock_dict = stock_ID_name_mapping()
    
    table_ID = "#divDetail"
    for stock_ID in stocks_ID:
        monthly_revenue_url = f"https://goodinfo.tw/StockInfo/ShowSaleMonChart.asp?STOCK_ID={stock_ID}"
        CFS_url = f"https://goodinfo.tw/StockInfo/StockBzPerformance.asp?STOCK_ID={stock_ID}&YEAR_PERIOD=9999&RPT_CAT=M_QUAR"
        url_list = [monthly_revenue_url, CFS_url]
        
        df_list = list()
        for url in url_list:
            df = stock_crawler(url, None, table_ID)
            # reassign headers
            headers = list()
            for i in range(len(df.columns)):
                headers.append('_'.join(pd.Series(df.columns[i]).drop_duplicates().tolist()))
            df.columns = headers
            delete_header(df, headers)
            df_list.append(df)
            time.sleep(random.randint(3,9))
        
        sheet_name = f"{stock_dict[stock_ID]}"
        df_list[0].to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, freeze_panes=(1,1)) # write to different sheets
        df_list[1].to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, startrow=df_list[0].shape[0]+2) # write to different sheets
        excel_formatting(writer, df_list[1], sheet_name)
        excel_formatting(writer, df_list[0], sheet_name)
        time.sleep(random.randint(5,12))
      
def input_menu():
    global DEBT_RATIO, STAKEHOLDING, PLEDGE_RATIO, GROSS_MARGIN, OPERATING_MARGIN, NET_PROFIT_MARGIN, DIVIDEND_YIELD

    menu()
    while True:
        input_num = input()
        if input_num == "0":
            DEBT_RATIO = int(input())
        elif input_num == "1":
            STAKEHOLDING = int(input())
        elif input_num == "2":
            PLEDGE_RATIO = int(input())
        elif input_num == "3":
            DIVIDEND_YIELD = int(input())
        # elif input_num == "4":
            # GROSS_MARGIN = int(input())
        # elif input_num == "5":
            # OPERATING_MARGIN = int(input())
        # elif input_num == "6":
            # NET_PROFIT_MARGIN = int(input())
        elif input_num == "7":
            DEBT_RATIO = 40
            STAKEHOLDING = 30
            PLEDGE_RATIO = 10
            GROSS_MARGIN = 20
            OPERATING_MARGIN = 20
            NET_PROFIT_MARGIN = 20
            DIVIDEND_YIELD = 1
        elif input_num == "8":
            print(f"負債比 < {DEBT_RATIO}%",
                  f"董監+法人持股 > {STAKEHOLDING}%",
                  f"董監質押比 < {PLEDGE_RATIO}%",
                  f"現金殖利率 > {DIVIDEND_YIELD}%\n", sep='\n')
                #   f"毛利率 前{GROSS_MARGIN}名",
                #   f"營益率 前{OPERATING_MARGIN}名",
                #   f"稅後淨利率 前{NET_PROFIT_MARGIN}名\n", sep='\n')
        elif input_num == "9":
            break
        menu()
 
def menu():
    print("輸入參數:",
          f"0.負債比(< {DEBT_RATIO}%)",
          f"1.董監+法人持股(> {STAKEHOLDING}%)",
          f"2.董監質押比(< {PLEDGE_RATIO}%)",
          f"3.現金殖利率(> {DIVIDEND_YIELD}%)",
        #   f"4.毛利率前{GROSS_MARGIN}名",
        #   f"5.營益率前{OPERATING_MARGIN}名",
        #   f"6.稅後淨利率前{NET_PROFIT_MARGIN}名",
          "7.預設",
          "8.目前參數值", 
          "9.繼續", sep="\n")
 
if __name__ == "__main__":
    main()