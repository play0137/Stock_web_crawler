"""
自結損益(Preliminary Earnings)
"""
import time
import random
from datetime import datetime
import pdb

import pandas as pd
from fake_useragent import UserAgent
from msedge.selenium_tools import Edge, EdgeOptions

from stock_web_crawler import stock_crawler, delete_header, excel_formatting
import global_vars

def main():
    global_vars.initialize_proxy()
    
    # webdriver
    options = EdgeOptions()
    options.use_chromium = True
    options.add_argument (f"--user-agent={UserAgent().random}")
    options.add_argument("--incognito")           # incognito mode  
    options.add_argument("headless")              # executing selenium without running the browser
    options.add_argument("disable-notifications") # disable notifications
    driver = Edge(global_vars.DIR_PATH + "edgedriver_win64/msedgedriver.exe", options=options)
    driver.implicitly_wait(random.randint(1,4))
    
    url = "https://mops.twse.com.tw/mops/web/t138sb01"
    driver.get(url)
    # click the radio and find the input area
    element = driver.find_element_by_xpath('//*[@id="search_bar1"]/div/input[1]').click()
    element = driver.find_element_by_xpath('//*[@id="search_bar1"]/div/input[2]')

    # pdb.set_trace()
    year = int(input("輸入年份:"))
    file_path = global_vars.DIR_PATH + "stock_PS/" + f"stock_PS_{year}.xlsx"
    writer = pd.ExcelWriter(file_path, mode="w", engine="xlsxwriter")
    if year > 1911:
        year -= 1911
    if year == datetime.now().year - 1911:
        months = datetime.now().month
    else:
        months = 13
    for month in range(1, months):
        date = f"{year}{month:02}"
        print(date)
        element.clear()
        element.send_keys(date)
        # click the search button
        driver.find_element_by_xpath('/html/body/center/table/tbody/tr/td/div[4]/table/tbody/tr/td/div/table/tbody/tr/td[3]/div/div[3]/form/table/tbody/tr/td[4]/table/tbody/tr/td[2]/div/div/input').click()
        time.sleep(random.randint(3,6))
        
        sheet_name = date
        table_ID = "#table01"
        dfs = stock_crawler(None, driver.page_source, table_ID, -1)
        start_row = 0
        for df in dfs[:-1]:
            if any(string in str(df.iloc[0,0]) for string in ["個別", "合併"]):
                df.to_excel(writer, index=False, header=False, encoding="UTF-8", sheet_name=sheet_name, startrow=start_row)
                start_row += df.shape[0]
            else:
                delete_header(df, df.columns)
                df.to_excel(writer, index=False, encoding="UTF-8", sheet_name=sheet_name, startrow=start_row)
                excel_formatting(writer, df, sheet_name)
                start_row += df.shape[0]+2
    writer.save()

if __name__ == "__main__":
    main()