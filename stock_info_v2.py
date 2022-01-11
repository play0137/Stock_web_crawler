import re
import time
import random
import pandas as pd
import global_vars
from stock_web_crawler import stock_crawler, delete_header, excel_formatting

def main():
    stocks_ID = input("input stock symbols:\n")
    writer = pd.ExcelWriter(global_vars.DIR_PATH + "stock_filter_manually.xlsx", engine="xlsxwriter")
    stock_info(stocks_ID, writer)
    writer.save()

# the info of monthly revenue and consollidated financial statements
def stock_info(stocks_ID, writer):
    
    global_vars.initialize_proxy()
    delims = r"[\s\t,\.;]+"
    stocks_ID = re.split(delims, stocks_ID)
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
        # write to different sheets
        df_list[0].to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, freeze_panes=(1,1))
        df_list[1].to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, startrow=df_list[0].shape[0]+2)
        excel_formatting(writer, df_list[1], sheet_name)
        excel_formatting(writer, df_list[0], sheet_name)
        time.sleep(random.randint(5,12))
        
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