import re
import sys
import pdb

import pandas as pd

from stock_web_crawler import stock_crawler, delete_header, excel_formatting
import global_vars

def main():
    global_vars.initialize_proxy()
    
    """ 個股資訊 """
    # inputs = "台積電 聯電"
    inputs = "2330 2314"
    # inputs = "2330"
    
    # inputs = input("輸入要搜尋的公司名稱或股號:\n(press q to exit)")
    if inputs == "q":
        sys.exit(0)
    
    stocks_ID = list()
    stock_dict = stock_ID_mapping()
    delims = r"[\s\t,\.;]+"
    inputs = re.split(delims, inputs)
    for stock in inputs:
        if stock not in stock_dict:
            print("Invalid input!!", stock, "is not in the stock ticker symbol table")
            sys.exit(-1)
        if stock.isdigit():
            stocks_ID.append(stock)
        else: # map company name to stock ID
            stocks_ID.append(stock_dict[stock]) 
    print("stocks ID:", stocks_ID)
    
    stock_salemon_file = global_vars.DIR_PATH + "個股月營收.xlsx"
    writer = pd.ExcelWriter(stock_salemon_file, engine="xlsxwriter")
    headers = ["月別", "開盤", "收盤", "最高", "最低", "漲跌(元)", "漲跌(%)", "單月營收(億)", "單月月增(%)", "單月年增(%)", "累計營收(億)", "累計年增(%)", "合併單月營收(億)", "合併單月月增(%)", "合併單月年增(%)", "合併累計營收(億)", "合併累計年增(%)"]
    table_ID = "#divDetail"
    for stock_ID in stocks_ID:
        url = f"https://goodinfo.tw/StockInfo/ShowSaleMonChart.asp?STOCK_ID={stock_ID}"
        df = stock_crawler(url, None, table_ID)
        df.columns = headers
        delete_header(df, headers)
        sheet_name = f"{stock_dict[stock_ID]}"
        df.to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, freeze_panes=(1,2)) # write to different sheets
        excel_formatting(writer, df, sheet_name)
    writer.save()


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

if __name__ == "__main__":
    main()