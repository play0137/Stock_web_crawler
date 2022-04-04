[中文版](https://github.com/play0137/Stock_web_crawler/blob/master/README.md)  
# Overview
Crawl the information from Goodinfo! stock website, and filter the data according to different arguments  
The arguments are adjustable, including debt_ratio, stakeholding, pledge_ratio and dividend_yield  

# Usage
1. Execute *stock_web_crawler.py* to crawl the information from Goodinfo!, and the output files are *月營收創新高.xlsx* and *stock_crawler.xlsx*.  
*月營收創新高.xlsx*: stock highest revenue in last month

2. Execute *stock_filter.py* to filter the data, and the output file is *stock_filter.xlsx*.  
The first and the second sheets (sheet names:All and Filtered) of *stock_filter.xlsx* shows the information of unfiltered and filtered sotcks.
The rest of the sheets show the basic information of each filtered stock.  
The columns of each sheet are those arguments mentioned above, and you can change them.    
The default settings are:  
DEBT_RATIO < 40%  
STAKEHOLDING > 30%  
PLEDGE_RATIO < 10%  
DIVIDEND_YIELD > 1%  

3. Execute *stock_info.py* to get the monthly revenue and consolidated financial statements, the output file is *stock_info.xlsx*  
The input can be a single stock symbol or name, or stock symbols or names, e.g., "2330" or "台積電, 聯電" or "2330 聯電"
The seperated symbols can be space, tab or comma

4. Optional
* Execute *continuous_highest_sale_month.py* to mark the stocks that have highest monthly revenue for the past N months
* There are some default proxies in *global_vars.py*.  
  Exexute *get_free_proxies.py* to get the free proxies and write to *proxies.txt*
  Proxies in *proxies.txt* will be used when the default proxies are all invalid.

# Prerequisites
### Browser
A Microsoft Edge browser 

### Dependencies
* [beautifulsoup4](https://pypi.org/project/beautifulsoup4/)
* [fake_useragent](https://pypi.org/project/fake-useragent/)
* [numpy](https://pypi.org/project/numpy/)
* [msedge-selenium-tools](https://pypi.org/project/msedge-selenium-tools/)
* [openpyxl](https://pypi.org/project/openpyxl/)
* [pandas](https://pypi.org/project/pandas/)
* [selenium](https://pypi.org/project/selenium/)
