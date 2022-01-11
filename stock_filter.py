import pdb
import pandas as pd
from datetime import datetime
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
    
    5.現金流量比率 > 0 or 營業現金流量 > 0(skip)
6.毛利率, 營益率, 稅後淨利率上升(三率三升)
    毛利率歷季排名 1 (第一名)
    營益率歷季排名 1
    淨利率歷季排名 1
7.殖利率 > 5%
    現金殖利率
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E7%86%B1%E9%96%80%E6%8E%92%E8%A1%8C&INDUSTRY_CAT=%E7%8F%BE%E9%87%91%E6%AE%96%E5%88%A9%E7%8E%87+%28%E6%9C%80%E6%96%B0%E5%B9%B4%E5%BA%A6%29%40%40%E7%8F%BE%E9%87%91%E6%AE%96%E5%88%A9%E7%8E%87%40%40%E6%9C%80%E6%96%B0%E5%B9%B4%E5%BA%A6
8.多項排列(均線)
    均線 向上
    https://goodinfo.tw/StockInfo/StockList.asp?RPT_TIME=&MARKET_CAT=%E6%99%BA%E6%85%A7%E9%81%B8%E8%82%A1&INDUSTRY_CAT=%E6%9C%88K%E7%B7%9A%E7%AA%81%E7%A0%B4%E5%AD%A3%E7%B7%9A%40%40%E6%9C%88K%E7%B7%9A%E5%90%91%E4%B8%8A%E7%AA%81%E7%A0%B4%E5%9D%87%E5%83%B9%E7%B7%9A%40%40%E5%AD%A3%E7%B7%9A
    9.股票尚未經歷大漲大跌(skip)

Result:
    代號 名稱 各篩選欄位資訊
1.全篩且變數不可調
2.選篩且變數可調    
"""

# global variable
DEBT_RATIO = 40
STAKEHOLDING = 30
PLEDGE_RATIO = 10
GROSS_MARGIN = 1
OPERATING_MARGIN = 1
NET_PROFIT_MARGIN = 1
DIVIDEND_YIELD = 1

def main():
    stock_dict = construct_stock_dict()
    
    file_path = "C:/Users/play0/OneDrive/桌面/stock/月營收創新高.xlsx"
    dfs = pd.read_excel(file_path, sheet_name=None) # set sheet_name=None if you want to read all sheets
    get_stock_data(stock_dict, dfs, "_highest_sale_month", f"{datetime.now().month}月", "單月營收歷月排名")
    
    #  read data from sheets in excel file
    file_path = "C:/Users/play0/OneDrive/桌面/stock/stock_crawler.xlsx"
    dfs = pd.read_excel(file_path, sheet_name=None)
    get_stock_data(stock_dict, dfs, "_debt_ratio", "負債比", "負債總額(%)")
    get_stock_data(stock_dict, dfs, "_supervisors_stakeholding", "全體董監持股_全體董監質押比", "全體董監持股(%)")
    get_stock_data(stock_dict, dfs, "_institutional_stakeholding", "本國法人持股", "本國法人(%)")
    get_stock_data(stock_dict, dfs, "_pledge_ratio", "全體董監持股_全體董監質押比", "全體董監質押(%)")
    get_stock_data(stock_dict, dfs, "_gross_margin", "單季營業毛利創歷季新高", "毛利歷季排名")
    get_stock_data(stock_dict, dfs, "_operating_margin", "單季營業利益創歷季新高", "營益率歷季排名")
    get_stock_data(stock_dict, dfs, "_net_profit_margin", "單季稅後淨利創歷季新高", "淨利率歷季排名")
    get_stock_data(stock_dict, dfs, "_dividend_yield", "殖利率", "現金殖利率")
    
    file_path = "C:/Users/play0/OneDrive/桌面/stock/stock_filter.csv"
    with open(file_path, "w", encoding="UTF-8") as file_w:
        file_w.write("代號, 名稱, 單月營收歷月排名, 負債總額(%), 董監持股(%), 本國法人持股(%), 董監質押(%), 毛利歷季排行, 營益率歷季排行, 淨利率歷季排行, 現金殖利率\n")
        for symbol, data in stock_dict.items():
            if(
                data._highest_sale_month and data._highest_sale_month == "1高" and
                data._debt_ratio and data._debt_ratio < DEBT_RATIO and
                data._supervisors_stakeholding and data._institutional_stakeholding and
                data._supervisors_stakeholding + data._institutional_stakeholding > STAKEHOLDING and
                data._pledge_ratio and data._pledge_ratio < PLEDGE_RATIO and
                data._gross_margin and data._gross_margin == GROSS_MARGIN and
                data._operating_margin and data._operating_margin == OPERATING_MARGIN and
                data._net_profit_margin and data._net_profit_margin == NET_PROFIT_MARGIN and
                data._dividend_yield and data._dividend_yield > DIVIDEND_YIELD
                ):
                    file_w.write(f"{symbol}, {data._name}, {data._highest_sale_month}, {data._debt_ratio}, {data._supervisors_stakeholding}, {data._institutional_stakeholding}, {data._pledge_ratio}, {data._gross_margin}, {data._operating_margin}, {data._net_profit_margin}, {data._dividend_yield}\n")


def construct_stock_dict():
    stock_dict = dict()
    file_path = "C:/Users/play0/OneDrive/桌面/stock/公司股市代號對照表.csv"
    with open(file_path, "r", encoding="UTF-8") as file_r:
        file_r.readline()
        for line in file_r:
            line = line.strip().split(",")
            stock_symbol = int(line[0])
            stock_abbrev_name = line[1]
            stock_full_name = ''.join(line[2:])
            if stock_abbrev_name not in stock_dict:
                stock_dict[stock_symbol] = Stock(stock_symbol, stock_abbrev_name, stock_full_name)
    return stock_dict

def get_stock_data(stock_dict, dfs, attr, sheet, col):
    df = dfs[sheet][["代號", col]]
    for index, row in df.iterrows():
        # 不處理 "2897A", "00713" 這兩種情況
        if isinstance(row["代號"], str) and (not row["代號"].isdigit() or row["代號"].startswith("00")):
            continue
        if int(row["代號"]) in stock_dict:
            setattr(stock_dict[int(row["代號"])], attr, row[col])
 
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
            print(f"負債比 <{DEBT_RATIO}%\n"\
                  f"董監+法人持股 > {STAKEHOLDING}%\n"\
                  f"董監質押比 < {PLEDGE_RATIO}%\n"\
                  f"現金殖利率 > {DIVIDEND_YIELD}%\n")
        elif input_num == "7":
            break
        menu()
 
def menu():
    print("輸入參數:",
          "1.負債比(<40%)",
          "2.董監+法人持股(>30%)",
          "3.董監質押比(<10%)",
          "4.現金殖利率(>5%)",
          "5.預設",
          "6.目前參數值", 
          "7.離開", sep="\n")
               
class Stock():
    def __init__(self, symbol, name=None, full_name=None, highest_sale_month=None, debt_ratio=None, supervisors_stakeholding=None, institutional_stakeholding=None, pledge_ratio=None, gross_margin=None, operating_margin=None, net_profit_margin=None, cash_flow_ratio=None, operating_cash_flow=None, dividend_yield=None):
        self._symbol = symbol
        self._name = name
        self._full_name = full_name
        self._highest_sale_month = highest_sale_month                    # 月營收新高
        self._debt_ratio = debt_ratio                                    # 負債比
        self._supervisors_stakeholding = supervisors_stakeholding        # 董監持股
        self._institutional_stakeholding = institutional_stakeholding    # 法人持股
        self._pledge_ratio = pledge_ratio                                # 質押比
        self._gross_margin = gross_margin                                # 毛利率
        self._operating_margin = operating_margin                        # 營益率
        self._net_profit_margin = net_profit_margin                      # 稅後淨利率
        # self._cash_flow_ratio = cash_flow_ratio                          # 現金流量比率
        # self._operating_cash_flow = operating_cash_flow                  # 營業現金流量
        self._dividend_yield = dividend_yield                            # 現金殖利率

if __name__ == "__main__":
    main()