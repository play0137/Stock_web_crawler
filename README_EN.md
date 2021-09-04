# Stock_web_crawler
Crawl the information from Goodinfo! stock website, and filter the data according to different arguments  
The arguments are adjustable, including debt_ratio, stakeholding, pledge_ratio, gross_margin, operating_margin, net_profit_margin and dividend_yield  

# How to run
1. Execute stock_web_crawler.py to crawl the information from Goodinfo!, and the the output files are 月營收創新高.xlsx and stock_crawler.xlsx  
月營收創新高.xlsx: stock highest revenue in different months  
2. Execute stock_filter_df.py to filter the data, and the output file is stock_filter.xlsx  
The first and the second sheets (sheet names:All and Filtered) of stock_filter.xlsx shows the infotmation of unfiltered and filtered sotcks
The rest of the sheets show the basic information of each filtered stock, including monthly revenue and consolidated financial statements
You can change the arguments mentioned above  
The default settings are:  
DEBT_RATIO < 40%       # 負債比  
STAKEHOLDING > 30%     # 持股  
PLEDGE_RATIO < 10%     # 質押比  
GROSS_MARGIN = 1       # 毛利率  
OPERATING_MARGIN = 1   # 營益率  
NET_PROFIT_MARGIN = 1  # 稅後淨利率  
DIVIDEND_YIELD > 1%    # 現金殖利率  

3. Execute stock_info.py to get the basic information of companies  
The input can be a single stock symbol or name, or stock symbols or names, e.g., 2330, 聯電

# Companies
公司股市代號對照表.csv contains the companies listed in Taiwan, the ones not in this file won't be processed  
