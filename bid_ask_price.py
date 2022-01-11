import pdb
"""
證交稅:賣的股價*(0.3/100)
交易手續費:買賣皆為 股價*(0.1425/100)  (股價不足 20 元者以 20 元計算)
每賺x塊賣y張的獲利最大
"""
def main():
    original_stock_price = 199
    expected_sell_price = 201
    buy_lot = 1  # 買入張數
    max_profit = 0
    best_sell_pair = list()
    
    # pdb.set_trace()
    
    buy_price = original_stock_price * 1000 * (1+0.1425/100)
    for step_stock_price in range(1, expected_sell_price - original_stock_price + 1): # 每隔幾塊賣出
        for sell_lot in range(buy_lot + 1): # 賣出幾張
            profit = 0
            buy_lot_tmp = buy_lot
            stock_price = original_stock_price+1
            while buy_lot_tmp-sell_lot > 0 and stock_price+step_stock_price <= expected_sell_price:
                sell_price = stock_price * sell_lot * 1000 * (1 - 0.1425/100 - 0.3/100)
                profit = profit + sell_price - buy_price*sell_lot
                buy_lot_tmp = buy_lot_tmp - sell_lot
                stock_price = stock_price + step_stock_price
            if buy_lot_tmp > 0: # 將剩下的全部賣出
                sell_price = stock_price * buy_lot_tmp * 1000 * (1 - 0.1425/100 - 0.3/100)
                profit = profit + sell_price - buy_price*buy_lot_tmp
            if profit > max_profit:
                max_profit = profit
                best_sell_pair = [step_stock_price, sell_lot]
                
    print("買入")
    print(f"股價:{original_stock_price:,}元")
    print(f"張數:{buy_lot}")
    print(f"手續費:{original_stock_price * buy_lot * 1000 * (0.1425/100):.2f}元")
    print(f"總價格:{buy_price*buy_lot:,.2f}元\n")
    
    print("賣出")
    print(f"預期賣出價格:{expected_sell_price}元")
    print(f"每隔{best_sell_pair[0]}元賣出{best_sell_pair[1]}張")
    print(f"最大利益:{max_profit:,.2f}元")
    
if __name__ == "__main__":
    main()