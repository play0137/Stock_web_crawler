"""
股價 > X
AND
成交張數 > Y
AND
成交價位於季均線區間
OR
成交價位於年均線區間
"""

import pandas as pd
import global_vars

def main():
    stock_price = 100  # 股價
    stock_trading_number = 500 # 成交張數
    season_range = 0.1 # 季線區間
    year_range = 0.3   # 年線區間
    
    file_path = global_vars.DIR_PATH + "StockList.csv"
    df = pd.read_csv(file_path)
    for col in df.columns:
        if "線" in col:
            df[col] = df[col].str.replace("↗", "")
            df[col] = df[col].str.replace("↘", "")
            df[col] = df[col].str.replace("→", "")
        if "成交張數" in col:
            df[col] = df[col].str.replace(",", "")
            
    df = df[df["成交"].astype(float) >= stock_price]
    df = df[df["成交張數"].astype(float) >= stock_trading_number]
    df_season = df[df["成交"].astype(float) > df["季均線"].astype(float)*(1-season_range)] # 季線
    df_season = df_season[df_season["成交"].astype(float) < df_season["季均線"].astype(float)*(1+season_range)] # 季線
    df_year = df[df["成交"].astype(float) > df["年均線"].astype(float)*(1-year_range)]    # 年線
    df_year = df_year[df_year["成交"].astype(float) < df_year["年均線"].astype(float)*(1+year_range)] # 年線
    df = pd.concat([df_season, df_year], axis=1) # 季均線 UNION 年均線
    
    with pd.ExcelWriter(global_vars.DIR_PATH + "價格在季年均線區間.xlsx", engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, encoding="UTF-8", freeze_panes=(1,1))

if __name__ == "__main__":
    main()