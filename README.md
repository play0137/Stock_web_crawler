# Stock_web_crawler
從Goodinfo!股票網站抓取資料並根據不同的參數來篩選  
可調整的參數包含 負債比、董監及法人持股、質押比、現金殖利率、毛利率、營益率、稅後淨利率

# 如何執行  
1. 執行 stock_web_crawler.py 來爬取Goodinfo!的資料  
輸出檔案為 "月營收創新高.xlsx" 和 "stock_crawler.xlsx"  
月營收創新高.xlsx: 抓取上一個月月營收創新高股票的相關資訊  

2. 執行 stock_filter_df.py 來過濾股票，輸出檔案為 "stock_filter.xlsx"  
stock_filter.xlsx 的第一及第二個分頁(分頁名字:ALL及Filtered)顯示所有及過濾後的股票  
每個欄位皆為用來做篩選的參數，參數是可調整的  
基本參數設定  
DEBT_RATIO < 40%        # 負債比  
STAKEHOLDING > 30%      # 持股  
PLEDGE_RATIO < 10%      # 質押比  
DIVIDEND_YIELD > 1%     # 現金殖利率  
毛利率、營益率、稅後淨利率前20名  
GROSS_MARGIN <= 20      # 毛利率  
OPERATING_MARGIN <= 20  # 營益率  
NET_PROFIT_MARGIN <= 20 # 稅後淨利率  

3. 執行 stock_info.py 得到股票的基本資訊
輸入可以是一個股號或名字，或是多個股號或名字  
分隔符號可以是空白鍵或逗號  
例如: "2330" 或 "台積電, 聯電" 或 "2330 聯電"  
輸出檔案為 "個股月營收.xlsx"  

4. 執行 stock_PS.py 得到公司的自結損益(Preliminary Earnings)  
輸入年份後，會抓取一整年份的資料

5. Optional  
global_vars.py 已經有預設一些可使用的proxy了  
執行get_free_proxies.py會從網站抓取可用的proxy並寫入proxies.txt  
執行其他程式時會讀取proxies.txt的資料並使用，在proxy無法使用時替換成其他的proxy
當全部的proxy都不可用時，請重新執行一次get_free_proxies.py來更新proxy

# 公司
公司股市代號對照表.csv: 包含所有在台灣上市的公司股號、簡稱、全名對照表
