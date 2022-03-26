#  build-in libraries
import os
import re
import sys
import pdb
import requests
from datetime import datetime

#  external libraries
from fake_useragent import UserAgent
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select

"""
stock web crawler
crawl the stock information from the Goodinfo! website
crawl the website by the webdriver

1.
GoodInfo -> 股票篩選 -> 智慧選股 -> 月營收 -> 月營收創排名紀錄 -> 單月營收創歷史新高
該月15號之前更新，15號之後不更新

2.
GoodInfo -> 股票篩選 -> 智慧選股 -> 獲利狀況 -> 營業毛利創新高/低 -> 單季營業毛利創歷季新高
                                               營業利益創新高/低 -> 單季營業利益創歷季新高
                                               稅後淨利創新高/低 -> 單季稅後淨利創歷季新高

3.
GoodInfo -> 每月營收 -> 輸入某個 股票代號/名稱 -> 抓取過去各年月的歷史資料(營業收入部分, 單月及累計下的資料)

"""
def main():
    
    """ stock highest sale month data """
    # stock_highest_salemon_file = "C:/Users/play0/OneDrive/桌面/stock_highest_saleMon_data.xlsx"
    url = "https://goodinfo.tw/StockInfo/StockList.asp?MARKET_CAT=%E8%87%AA%E8%A8%82%E7%AF%A9%E9%81%B8&INDUSTRY_CAT=%E6%88%91%E7%9A%84%E6%A2%9D%E4%BB%B6&FILTER_ITEM0=%E5%B9%B4%E5%BA%A6%E2%80%93ROE%28%25%29&FILTER_VAL_S0=30&FILTER_VAL_E0=&FILTER_SHEET=%E5%B9%B4%E7%8D%B2%E5%88%A9%E8%83%BD%E5%8A%9B&WITH_ROTC=F&FILTER_QUERY=%E6%9F%A5++%E8%A9%A2"
    driver = webdriver.Edge("C:/Users/play0/OneDrive/桌面/stock/edgedriver_win64/msedgedriver.exe")
    driver.get(url)

    # hide menu10 and show menu9
    # 隱藏元素可以被定位到(is_displayed()為false)，但無法操作(click, send_keys, clear, .etc)
    # ref:https://www.itread01.com/content/1545648793.html
    js = "document.getElementById('MENU10').style.display='none'"       # hide 自訂篩選
    driver.execute_script(js)
    js = "document.getElementById('MENU9').style.display='block'"       # show 智慧選股
    driver.execute_script(js)
    # print(element.is_displayed())
    
    js = "document.getElementById('MENU9_0_1100').style.display='none'" # hide 智慧選股_交易狀況
    driver.execute_script(js)
    
    js = "document.getElementById('MENU9_11_0').style.display='block'"  # show 月營收
    driver.execute_script(js)
    element = driver.find_element_by_id("MENU9_11_0")

    # show the stock highest sale month data
    element = element.find_elements_by_class_name("sel_opt_black")
    highest_sale_month = element[1]
    select = Select(highest_sale_month).options[1]
    select.click()
    
    # get the stock highest sale month of tables
    html = driver.page_source
    soup = BeautifulSoup(html, 'lxml')
    div = soup.select_one("#tblStockList")
    df = pd.read_html(str(div))[0]
    
    stock_highest_salemon_file = "C:/Users/play0/OneDrive/桌面/stock/月營收創新高.xlsx"
    writer = pd.ExcelWriter(stock_highest_salemon_file)
    # delete headers
    headers = list(df.iloc[0]) # first row of the table is the headers
    row_of_headers = 1
    row_of_block = 19          # a table consists of 19 rows
    df = deleteHeaders(df, headers, row_of_headers, row_of_block)
    
    # write to different months
    today = datetime.now()
    sheet_name = f"{today.month}月"
    df.to_excel(writer, index=False, encoding="BIG5", sheet_name=sheet_name)
    writer.save()
    
    """ 營業毛利創新高/低, 營業利益創新高/低, 稅後淨利創新高/低 """
    js = "document.getElementById('MENU9_11_0').style.display='none'"      # hide 月營收
    driver.execute_script(js)
    js = "document.getElementById('MENU9_12_11000').style.display='block'" # show 獲利狀況
    driver.execute_script(js)
    
    stock_profit_file = "C:/Users/play0/OneDrive/桌面/stock/獲利狀況.xlsx"
    writer = pd.ExcelWriter(stock_profit_file)

    sheet_name = "單季營業毛利創歷季新高"
    element_id = "MENU9_12_11000"
    element_order = 1
    element_options_oder = 1
    table_id = "#divStockList"
    profitSituation(driver, writer, sheet_name, element_id, element_order, element_options_oder, table_id)
    
    sheet_name = "單季營業利益創歷季新高"
    element_id = "MENU9_12_11000"
    element_order = 3
    profitSituation(driver, writer, sheet_name, element_id, element_order, element_options_oder, table_id)
    
    sheet_name = "單季稅後淨利創歷季新高"
    element_id = "MENU9_12_11000"
    element_order = 7
    profitSituation(driver, writer, sheet_name, element_id, element_order, element_options_oder, table_id)
    
    writer.save()
    writer.close()
    
    driver.quit() # quit the driver
    
    
    """ stock sale month data """
    # inputs = "台積電 聯電"
    # inputs = "2330 2314"
    # inputs = "2330"
    print("輸入要搜尋的公司名稱或股號:")
    inputs = input()
    stocks_ID = list()
    
    stock_dict = stockIDMapping()
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
    
    stock_salemon_file = "C:/Users/play0/OneDrive/桌面/stock/個股月營收.xlsx"
    writer = pd.ExcelWriter(stock_salemon_file)
    
    headers = ["月別", "開盤", "收盤", "最高", "最低", "漲跌(元)", "漲跌(%)", "單月營收(億)", "單月月增(%)", "單月年增(%)", "累計營收(億)", "累計年增(%)", "合併單月營收(億)", "合併單月月增(%)", "合併單月年增(%)", "合併累計營收(億)", "合併累計年增(%)"]
    row_of_headers = 3
    row_of_block = 21
    # write to different sheets
    for stock_ID in stocks_ID:
        soup = stockCrawler(stock_ID)
        data = soup.select_one("#divSaleMonChartDetail")
        dfs = pd.read_html(data.prettify())
        df = dfs[1]
        df = deleteHeaders(df, headers, row_of_headers, row_of_block)
        df.to_excel(writer, index=False, encoding="BIG5", sheet_name=f"{stock_dict[stock_ID]}")
    
    writer.save()
    writer.close()


def profitSituation(driver, writer, sheet_name, element_id, element_order, element_options_oder, table_id):
    # click 單季營業毛利創歷季新高 in 營業毛利創新高/低
    element = driver.find_element_by_id(element_id)
    element = element.find_elements_by_class_name("sel_opt_black")
    highest_month = element[element_order]
    select = Select(highest_month).options[element_options_oder]
    select.click()
    
    # get the table of 單季營業毛利創歷季新高
    html = driver.page_source
    soup = BeautifulSoup(html, 'lxml')
    div = soup.select_one(table_id)
    df = pd.read_html(str(div))[0]
    
    # delete headers
    headers = list(df.iloc[0]) # first row of the table is the headers
    row_of_headers = 1
    row_of_block = 19          # a table consists of 19 rows
    df = deleteHeaders(df, headers, row_of_headers, row_of_block)
    df.to_excel(writer, index=False, encoding="BIG5", sheet_name=sheet_name)

def deleteHeaders(df, headers, row_of_headers, row_of_block):
    tmp_file = "C:/Users/play0/OneDrive/桌面/stock/tmp.xlsx"
    df.to_csv(tmp_file, sep=',', index=False, header=False, encoding='UTF-8')
    row_num = len(df.index)
    step = row_num // row_of_block
    skip_rows = list()
    for i in range(step+1):
        for j in range(row_of_headers):
            skip_rows += [j+row_of_block*i]
    df = pd.read_csv(tmp_file, skiprows=skip_rows, names=headers)
    os.remove(tmp_file)
    
    return df

def stockCrawler(stock_ID):
    # website url
    url = f"https://goodinfo.tw/StockInfo/ShowSaleMonChart.asp?STOCK_ID={stock_ID}"
    
    # fake user agent
    fake_ua = UserAgent()
    headers = {'User-Agent': fake_ua.random}
    # headers = {"User-Agent": "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36"}

    # request
    list_req = requests.post(url, headers = headers)
    if list_req.status_code == requests.codes.ok: # status_code:200
        print("Request successful!")
    else:
        print("Request failed!")
        
    # crawl the website
    soup = BeautifulSoup(list_req.content, "lxml")
    soup.encoding = "UTF-8"
    
    return soup

# 1101,台泥,台灣水泥股份有限公司
def stockIDMapping():
    stock_dict = dict()
    with open("C:/Users/play0/OneDrive/桌面/stock/公司股市代號對照表.csv", "r", encoding="UTF-8") as file_r:
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
