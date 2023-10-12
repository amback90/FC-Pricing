import talib as ta
import time
import pandas as pd
import MetaTrader5 as mt5

def bid_price(symbol):return mt5.symbol_info_tick(symbol).ask
def ask_price(symbol):return mt5.symbol_info_tick(symbol).bid


def symbol_df_updater(symbol_list ,symbol_df , dfs):
    print("ok")
    while True:
        for symbol in symbol_list:
            df = pd.DataFrame(mt5.copy_rates_from_pos(symbol, mt5.TIMEFRAME_M5, 0, 200)) 
            symbol_df[symbol] = df
        
            dfs['bid_price'][symbol]= mt5.symbol_info_tick(symbol).bid
            #print(dfs['bid_price'][symbol])
            dfs['ask_price'][symbol] = mt5.symbol_info_tick(symbol).ask
            #print(dfs['ask_price'][symbol])
            dfs['spred'][symbol] = dfs['bid_price'][symbol] - dfs['ask_price'][symbol]
            #print(dfs['spred'][symbol])
            dfs['rsi'][symbol] = (ta.RSI(symbol_df[symbol]["close"], timeperiod=14)).iloc[-1]
            #print(dfs['rsi'][symbol])
            dfs['moving_average'][symbol] = (ta.MA(symbol_df[symbol]["close"])).iloc[-1]
            a ,b, c = ta.MACD(symbol_df[symbol]["close"], fastperiod=12, slowperiod=26, signalperiod=9)
            dfs['macd'][symbol] , dfs['macdsignal'][symbol] , dfs['macdhist'][symbol] = a.iloc[-1]  ,  b.iloc[-1]   ,c.iloc[-1]
            dfs['parabollic_sar'][symbol] = (ta.SAR(symbol_df[symbol]["high"], symbol_df[symbol]["low"], acceleration=0, maximum=0)).iloc[-1]
        
            