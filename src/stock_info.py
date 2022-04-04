""" Monthly revenue (月營收) and Consolidated Financial Statements (合併財務報表) """

import random
import re
import sys
import time
import pandas as pd

# self-defined modules
import global_vars
from stock_web_crawler import stock_crawler, delete_header, excel_formatting

def main():
    stocks_ID, report_type = Input()
    stock_info(stocks_ID, report_type)

# Get the information of monthly revenue (月營收) and consolidated financial statements (合併財務報表)
def stock_info(stocks_ID, report_type):
    global_vars.initialize_proxy()
    stock_dict = stock_ID_name_mapping()
    
    table_ID = "#divDetail"
    for stock_ID in stocks_ID:
        monthly_revenue_url = f"https://goodinfo.tw/StockInfo/ShowSaleMonChart.asp?STOCK_ID={stock_ID}"
        CFS_url = f"https://goodinfo.tw/StockInfo/StockBzPerformance.asp?STOCK_ID={stock_ID}&YEAR_PERIOD=9999&RPT_CAT=M_QUAR"
        if report_type == 1:
            url_list = [monthly_revenue_url]
        elif report_type == 2:
            url_list = [CFS_url]
        else:
            url_list = [monthly_revenue_url, CFS_url]
        
        df_list = list()
        for url in url_list:
            df = stock_crawler(url, None, table_ID)
            
            # reassign headers
            headers = list()
            for i in range(len(df.columns)):
                headers.append('_'.join(list(dict.fromkeys(df.columns[i]))))
            df.columns = headers
            
            delete_header(df, headers)
            df_list.append(df)
            time.sleep(random.randint(3, 9))
        
        sheet_name = f"{stock_dict[stock_ID]}"
        # write to different sheets if the stock is different
        with pd.ExcelWriter(global_vars.DIR_PATH + "stock_info.xlsx", engine="xlsxwriter") as writer:
            df_list[0].to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, freeze_panes=(1,1))
            if report_type == 3: # both monthly revenue and CFS
                df_list[1].to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, startrow=df_list[0].shape[0]+2)
                excel_formatting(writer, df_list[1], sheet_name)
            excel_formatting(writer, df_list[0], sheet_name)
            time.sleep(random.randint(5, 12))

def Input():
    # Report types
    report_type = 0
    while int(report_type) not in {1,2,3}:
        print("Input the type of report: (press q to exit)\n" +
              "1. Monthly revenue (個股月營收)\n" + 
              "2. Consolidated financial statements (合併財務報表)\n" + 
              "3. Both (全部)", end='')
        report_type = input()
        if report_type == "q":
            sys.exit(0)
               
    # Stocks
    """ valid input formats """
    # inputs = "台積電 聯電"
    # inputs = "2330 2314"
    # inputs = "台積電, 2314"
    inputs = input("Input stock company names (公司名稱) or symbols (股號):\n")
    delims = r"[\s\t,\.;]+"
    inputs = re.split(delims, inputs)
    
    stocks_ID = list()
    stock_dict = stock_ID_name_mapping()
    for stock in inputs:
        if stock not in stock_dict:
            print("Invalid input!", stock, "is not in the stock ticker symbol table")
            sys.exit(-1)
        if stock.isdigit():
            stocks_ID.append(stock)
        else: # map company name to stock ID
            stocks_ID.append(stock_dict[stock])
    
    return [stocks_ID, int(report_type)]
    
# 1101,台泥,台灣水泥股份有限公司
def stock_ID_name_mapping():
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

if __name__ == "__main__":
    main()