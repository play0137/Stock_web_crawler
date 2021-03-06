[English version](https://github.com/play0137/Stock_web_crawler/blob/master/README_EN.md)

# 介紹
從 Goodinfo! 股票網站抓取資料並根據不同的參數來篩選股票  
可調整的參數包含 負債比、董監及法人持股、質押比、現金殖利率、毛利率、營益率、稅後淨利率

# 如何執行
1. 透過以下指令安裝所需要的packages  
<code> pip install -r requirements.txt </code>

2. 執行 *stock_web_crawler.py* 來爬取 Goodinfo! 的資料  
輸出檔案為 *月營收創新高.xlsx* 和 *stock_crawler.xlsx*  
*月營收創新高.xlsx*: 抓取上一個月月營收創新高股票的相關資訊    

3. 執行 *stock_filter.py* 來過濾股票，輸出檔案為 *stock_filter.xlsx*  
*stock_filter.xlsx* 的第一及第二個分頁(分頁名字: ALL 及 Filtered)顯示所有及過濾後的股票  
每個欄位皆為用來做篩選的參數，參數是可調整的  
預設參數設定  
DEBT_RATIO < 40%        # 負債比  
STAKEHOLDING > 30%      # 持股  
PLEDGE_RATIO < 10%      # 質押比  
DIVIDEND_YIELD > 1%     # 現金殖利率  
毛利率、營益率、稅後淨利率前20名  
GROSS_MARGIN <= 20      # 毛利率  
OPERATING_MARGIN <= 20  # 營益率  
NET_PROFIT_MARGIN <= 20 # 稅後淨利率  

4. 執行 *stock_info.py* 得到股票的 月營收 (monthly revenue) 及 合併財務報表 (consolidated financial statements)  
輸入可以是一個股號或公司名稱，或是多個股號或公司名稱  
分隔符號可以是空白鍵或逗號  
例如: "2330" 或 "2330 2331" 或 "台積電, 聯電" 或 "2330 聯電"  
輸出檔案為 *stock_info.xlsx*  

5. Optional  
* 執行 *continuous_highest_sale_month.py* 會將過去N個月連續M個月，月營收創新高的股票標記顏色
* *global_vars.py* 有預設一些可使用的 proxy  
  執行 *get_free_proxies.py* 會從網站抓取可用的 proxy 並寫入 *proxies.txt*  
  執行其他程式時會讀取 *proxies.txt* 的資料並使用，在預設的 proxy 無法使用時替換成其他的 proxy  
  當全部的 proxy 都不可用時，請重新執行一次 *get_free_proxies.py* 來更新 proxy

# 需求
### 瀏覽器
Microsoft Edge 瀏覽器  

### 需安裝套件
* [beautifulsoup4](https://pypi.org/project/beautifulsoup4/)
* [fake_useragent](https://pypi.org/project/fake-useragent/)
* [numpy](https://pypi.org/project/numpy/)
* [openpyxl](https://pypi.org/project/openpyxl/)
* [pandas](https://pypi.org/project/pandas/)
* [selenium](https://pypi.org/project/selenium/)
