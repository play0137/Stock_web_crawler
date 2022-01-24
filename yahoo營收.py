"""
網址:https://tw.stock.yahoo.com/
抓單一選單的表格  
"""
import os
import re
import sys
import pdb
import time
import random
import numpy as np
import pandas as pd
from openpyxl import Workbook
from msedge.selenium_tools import Edge, EdgeOptions
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from stock_info import stock_ID_name_mapping
import global_vars

def main():
    options = EdgeOptions()
    options.use_chromium = True
    options.add_argument("disable-notifications") # disable notifications
    # options.add_argument("headless")              # executing selenium without running the browser
    driver = Edge(global_vars.DIR_PATH + "edgedriver_win64/msedgedriver.exe", options=options)

    inputs = input("輸入要搜尋的公司名稱或股號:\n(press q to exit)\n")
    if inputs == "q":
        sys.exit(0)
    
    stocks_id = list()
    stock_dict = stock_ID_name_mapping()
    delims = r"[\s\t,\.;]+"
    inputs = re.split(delims, inputs)
    for stock in inputs:
        if stock not in stock_dict:
            print("Invalid input!", stock, "is not in the stock ticker symbol table")
            sys.exit(-1)
        if stock.isdigit():
            stocks_id.append(stock)
        else: # map company name to stock ID
            stocks_id.append(stock_dict[stock])
    print("stocks ID:", stocks_id)
    
    for stock_id in stocks_id:
        financial(driver, stock_dict, stock_id)
    
def financial(driver, stock_dict, stock_id):
    action = webdriver.ActionChains(driver)
    url = f'https://tw.stock.yahoo.com/quote/{stock_id}'
    driver.get(url)
    driver.implicitly_wait(10)
    financial_XPath = '//*[@id="main-1-QuoteTabs-Proxy"]/nav/div/div/div[6]'
    element = WebDriverWait(driver, random.randint(10,15)).until(EC.presence_of_element_located((By.XPATH, financial_XPath)))
    action.move_to_element(element).perform() # 移動滑鼠到指定位置
    
    file_path = global_vars.DIR_PATH + f'萬能營收表_{stock_dict[str(stock_id)]}.xlsx'
    # create an excel file if not exist
    if not os.path.isfile(file_path):
        wb = Workbook()
        wb.worksheets[0].title = "營收表"
        wb.save(file_path)
    writer =  pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace")
        
    elements_financial = WebDriverWait(driver, random.randint(10,15)).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="main-1-QuoteTabs-Proxy"]/nav/div/div/div[6]/div/ul/li')))
    table_XPath = '/html/body/div[1]/div/div/div/div/div[5]/div[1]/div[1]/div/div[4]/div/section[2]/div/div'
    sheet_names = ["營收表", "每股盈餘", "損益表", "資產負債表", "現金流量表"]
    column_numbers = [8, 5, 21, 21, 21]
    for i, element_financial in enumerate(elements_financial):
        print(element_financial.text)
        element_financial.click()
        time.sleep(random.randint(4,6))
        write_table_to_excel(driver, table_XPath, column_numbers[i], writer, sheet_names[i])
        
        tmp_element = driver.find_element_by_xpath(financial_XPath)
        action.move_to_element(tmp_element).perform()
    writer.save()
    driver.close()

def write_table_to_excel(driver, table_XPath, column_number, writer, sheet_name):
    data = driver.find_element_by_xpath(table_XPath).text
    while "資料載入" in data:
        print("資料載入")
        data = driver.find_element_by_xpath(table_XPath).text
    data = data.split('\n')
    
    column_names = list()
    i = 0
    while i < column_number:
        if "單月合併" in data[0] or "累計合併" in data[0]:
            data.pop(0)
            continue
        column_names.append(data.pop(0))
        i += 1
    row_num = len(data) // column_number
    reshaped_array = np.reshape(np.array(data), (row_num, column_number))
    df = pd.DataFrame(reshaped_array, columns=column_names)
    df.to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, freeze_panes=(1,2))
    
if __name__ == "__main__":
    main()