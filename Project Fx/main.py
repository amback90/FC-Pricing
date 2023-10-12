from datetime import *
import time
import pandas as pd
import MetaTrader5 as mt5
import xlwings as xw
from threading import Thread
import ast
import os
current_path = os.getcwd()
path = os.path.join(current_path, "Book1.xlsx")

wb = xw.Book(path)
dashbord = wb.sheets['dashbord']
control = wb.sheets['control']

username = int(dashbord.range("D1").value)
password = dashbord.range("D2").value
server = dashbord.range("D3").value

def sync_excel(path):
    
    wb = xw.Book(path)
    global buy_order_df ,sell_order_df
    control = wb.sheets['control']
    time.sleep(2)
    
    while True:    
        
        for symbol in buy_order_df.index:
            for cell in control.range("A2:A16"):
                if cell.value == symbol:
                    break
            row_index = cell.row
            col_index = cell.column
            if control["B"+str(row_index)+":U"+str(row_index)].expand("right").value==None:
                buy_order_df["order_list"].loc[symbol]=[]
            else:
                buy_order_df["order_list"].loc[symbol] =  [ast.literal_eval(s) for s in (control["B"+str(row_index)+":U"+str(row_index)].expand("right").value)]
                   
            if control.range("W"+str(row_index)).value == None:
                buy_order_df["total_loss"].loc[symbol]=10000
            else:
                buy_order_df["total_loss"].loc[symbol] = control.range("W"+str(row_index)).value
            
            if control.range("X"+str(row_index)).value == None:
                buy_order_df["total_profit"].loc[symbol]=10000
            else:
                buy_order_df["total_profit"].loc[symbol] = control.range("X"+str(row_index)).value
                
            if control.range("Y"+str(row_index)).value == None:
                buy_order_df["sl_value"].loc[symbol]=0
            else:
                buy_order_df["sl_value"].loc[symbol] = control.range("Y"+str(row_index)).value
             
            if control.range("Z"+str(row_index)).value == None:
                buy_order_df["tp_value"].loc[symbol]=0
            else:
                buy_order_df["tp_value"].loc[symbol] = control.range("Z"+str(row_index)).value
        
        for symbol in sell_order_df.index:
            for cell in control.range("A19:A32"):
                if cell.value == symbol:
                    break
            row_index = cell.row
            col_index = cell.column
            if control["B"+str(row_index)+":U"+str(row_index)].expand("right").value==None:
                sell_order_df["order_list"].loc[symbol]=[]
            else:
                sell_order_df["order_list"].loc[symbol] =  [ast.literal_eval(s) for s in (control["B"+str(row_index)+":U"+str(row_index)].expand("right").value)]
  
            if control.range("W"+str(row_index)).value == None:
                sell_order_df["total_loss"].loc[symbol]=10000
            else:
                sell_order_df["total_loss"].loc[symbol] = control.range("W"+str(row_index)).value
            
            if control.range("X"+str(row_index)).value == None:
                sell_order_df["total_profit"].loc[symbol]=10000
            else:
                sell_order_df["total_profit"].loc[symbol] = control.range("X"+str(row_index)).value
                
            if control.range("Y"+str(row_index)).value == None:
                sell_order_df["sl_value"].loc[symbol]=0
            else:
                sell_order_df["sl_value"].loc[symbol] = control.range("Y"+str(row_index)).value
             
            if control.range("Z"+str(row_index)).value == None:
                sell_order_df["tp_value"].loc[symbol]=0
            else:
                sell_order_df["tp_value"].loc[symbol] = control.range("Z"+str(row_index)).value
             
        time.sleep(1)
        
        
def point(symbol):
    return mt5.symbol_info(symbol).point

def bid_price(symbol):
    return mt5.symbol_info_tick(symbol).ask

def ask_price(symbol):
    return mt5.symbol_info_tick(symbol).bid

def total_buy_positions(symbol):
    tbp = pd.DataFrame(mt5.positions_get(symbol=symbol))
    if (len(tbp))==0:
        return 0
    else:
        return len(tbp[tbp[5]==0])
def total_sell_positions(symbol):
    tsp = pd.DataFrame(mt5.positions_get(symbol=symbol))
    if (len(tsp))==0:
        return 0
    else:
        return len(tsp[tsp[5]==1])
# symbol="GBPJPY"
# total_sell_positions("GBPJPY")    
def buy_positions(symbol):
    bpdf = pd.DataFrame(mt5.positions_get(symbol=symbol))
    return (bpdf[bpdf[5]==0]).reset_index(drop=True)
def sell_positions(symbol):
    spdf = pd.DataFrame(mt5.positions_get(symbol=symbol))
    return (spdf[spdf[5]==1]).reset_index(drop=True)


def current_buy_profit(symbol):
    positions = pd.DataFrame(mt5.positions_get(symbol=symbol))
    return (positions[positions[5]==0][15]).sum()
def current_sell_profit(symbol):
    positions = pd.DataFrame(mt5.positions_get(symbol=symbol))
    return (positions[positions[5]==0][15]).sum()

def bid_ask_updater():
    while True:
        for symbol in buy_symbols:
            buy_order_df["bid_price"].loc[symbol] = bid_price(symbol)
            buy_order_df["ask_price"].loc[symbol] = ask_price(symbol)
        for symbol in sell_symbols:
            sell_order_df["bid_price"].loc[symbol] = bid_price(symbol)
            sell_order_df["ask_price"].loc[symbol] = ask_price(symbol)
        time.sleep(0.2)

def order_start(symbol, lot, type, price, sl, tp, comment):
    request = {
        "action": mt5.TRADE_ACTION_DEAL,
        "symbol": symbol,
        "volume": lot,
        "type": type,
        "price": price,
        "sl": sl,
        "tp": tp,
        "deviation": 2,
        "magic": 234000,
        "comment": comment,
        "type_time": mt5.ORDER_TIME_GTC,
        "type_filling": mt5.ORDER_FILLING_IOC,
    }
    result = mt5.order_send(request)
    print(result)
    return result


def order_close(symbol, lot, type, position, price, comment):
    request = {
        "action": mt5.TRADE_ACTION_DEAL,
        "symbol": symbol,
        "volume": lot,
        "type": type,
        "position": position,
        "price": price,
        "deviation": 2,
        "magic": 234000,
        "comment": comment,
        "type_time": mt5.ORDER_TIME_GTC,
        "type_filling": mt5.ORDER_FILLING_IOC,
    }
    result = mt5.order_send(request)
    print(result)
    return result

def current_buy_order_ticket(symbol):
    positions = pd.DataFrame(mt5.positions_get(symbol=symbol))
    positions = positions[positions[5]==0]
    df = pd.DataFrame(positions)
    return df[0].tolist()

def current_sell_order_ticket(symbol):
    positions = pd.DataFrame(mt5.positions_get(symbol=symbol))
    positions = positions[positions[5]==1]
    df = pd.DataFrame(positions)
    return df[0].tolist()


#########################################################################
def next_buy_order(symbol , i ):
    next_order_price = (buy_positions(symbol)[10].iloc[-1]) - (point(symbol)*buy_order_df["order_list"].loc[symbol][i][1]) 
    next_lot = buy_order_df["order_list"].loc[symbol][i][0]
    print("next order: ",next_order_price,"     next not: ",next_lot)
    return next_order_price , next_lot
 
def buy_order_function(symbol):
    _total_positions = total_buy_positions(symbol)
    
    if _total_positions>0 and _total_positions<22 :
        next_order_price , next_lot = next_buy_order(symbol , _total_positions-1)
        if bid_price(symbol) < next_order_price:
            result = order_start(symbol, next_lot , mt5.ORDER_TYPE_BUY, bid_price(symbol),  buy_order_df["sl_value"].loc[symbol] , buy_order_df["tp_value"].loc[symbol] , "fist order completed")
            if result.retcode == mt5.TRADE_RETCODE_DONE:
                print("Buy  order send completed")
            else:
                print(f"sell  order send faild , with  retcode={result.retcode}")
                

def all_buy_order_close(symbol):
    if ((total_buy_positions(symbol) > 0) and (  (current_buy_profit(symbol) >  buy_order_df["total_profit"].loc[symbol] ) or ( current_buy_profit(symbol) <  buy_order_df["total_loss"].loc[symbol]) )):
        for ticket in current_buy_order_ticket(symbol):
            volume = current_ticket_volume(ticket)
            try:
                result = order_close(symbol, volume, mt5.ORDER_TYPE_SELL, ticket, ask_price(symbol), "STOCH out")
            except:
                print(result)
            else:
                if result.retcode == mt5.TRADE_RETCODE_DONE:
                    print(
                        f"buy order closed, with  retcode={result.retcode}")
                else:
                    print(
                        f"buy order closed Faild, with  retcode={result.retcode}")
            
                        
def buy_order_sltp(symbol):
  
    if (buy_order_df["last_sl_value"].loc[symbol] != buy_order_df["sl_value"].loc[symbol]) or ((buy_order_df["last_tp_value"].loc[symbol] != buy_order_df["tp_value"].loc[symbol])):
        for i in range(total_buy_positions(symbol)):
            _ticket=int(buy_positions(symbol)[0][i])
      
            request = {
                "action": mt5.TRADE_ACTION_SLTP,
                "position": _ticket,
                "symbol" : symbol,
                "sl": buy_order_df["sl_value"].loc[symbol],
                "tp" :buy_order_df["tp_value"].loc[symbol]
            }
            result_SL = mt5.order_send(request)
            print(result_SL)
        buy_order_df["last_sl_value"].loc[symbol] = buy_order_df["sl_value"].loc[symbol]
        buy_order_df["last_tp_value"].loc[symbol] = buy_order_df["tp_value"].loc[symbol]
    
#####################################################################################################

def next_sell_order(symbol , i):

    next_order_price = (sell_positions(symbol)[10].iloc[-1]) + (point(symbol)*sell_order_df["order_list"].loc[symbol][i][1] ) 
    next_lot = sell_order_df["order_list"].loc[symbol][i][0]
    print("next sell order: ",next_order_price,"     next not: ",next_lot)
    return next_order_price , next_lot

def sell_order_function(symbol):
    _total_positions = total_sell_positions(symbol)
    if _total_positions>0 and _total_positions<22 :
        next_order_price , next_lot = next_sell_order(symbol , _total_positions-1)
        if bid_price(symbol) > next_order_price:
            result = order_start(symbol, next_lot , mt5.ORDER_TYPE_SELL, ask_price(symbol),  sell_order_df["sl_value"].loc[symbol] , sell_order_df["tp_value"].loc[symbol] , "fist order completed")
            if result.retcode == mt5.TRADE_RETCODE_DONE:
                print("Buy  order send completed")
            else:
                print(f"sell  order send faild , with  retcode={result.retcode}")
                

def all_sell_order_close(symbol):
    if ((total_sell_positions(symbol) > 0) and (  (current_sell_profit(symbol) >   sell_order_df["total_profit"].loc[symbol]) or ( current_sell_profit(symbol) <  sell_order_df["total_loss"].loc[symbol])  )):
        for ticket in current_sell_order_ticket(symbol):
            volume = current_ticket_volume(ticket)
            try:
                result = order_close(symbol, volume, mt5.ORDER_TYPE_BUY, ticket, bid_price(symbol), "STOCH out")
            except:
                print(result)
            else:
                if result.retcode == mt5.TRADE_RETCODE_DONE:
                    print(
                        f"sell order closed, with  retcode={result.retcode}")
                else:
                    print(
                        f"sell order closed Faild, with  retcode={result.retcode}")
            print("this")

def sell_order_sltp(symbol):
    if (sell_order_df["last_sl_value"].loc[symbol] != sell_order_df["sl_value"].loc[symbol]) or (sell_order_df["last_tp_value"].loc[symbol] != sell_order_df["tp_value"].loc[symbol]):
        for i in range(total_sell_positions(symbol)):
            _ticket=int(sell_positions(symbol)[0][i])
            
            request = {
                "action": mt5.TRADE_ACTION_SLTP,
                "position": _ticket,
                "symbol" : symbol,
                "sl": float(sell_order_df["sl_value"].loc[symbol]),
                "tp" :float(sell_order_df["tp_value"].loc[symbol])
            }
            result_SL = mt5.order_send(request)
            print(result_SL)
        
        sell_order_df["last_sl_value"].loc[symbol] =   sell_order_df["sl_value"].loc[symbol]
        sell_order_df["last_tp_value"].loc[symbol] =   sell_order_df["tp_value"].loc[symbol]
         
        
#####################################################################################################

def current_ticket_volume(ticket):
    positions = mt5.positions_get(ticket=ticket)
    tp = positions[0]
    return tp[9]

if __name__=="__main__":
    
    if not mt5.initialize():
        print("initialize() failed, error code =", mt5.last_error())
        quit()
    time.sleep(3)
    
    # if mt5.login(username, password=password, server=server ):
    #     print("connected to account #{}".format(username))
    # else:
    #     print("failed to connect at account #{}, error code: {}".format(username, mt5.last_error()))
    # time.sleep(2)

    print("your current balance : " + str(mt5.account_info().balance))
    
    buy_symbols =[item for item in control.range('A2:A16').value if item is not None] 
    sell_symbols =[item for item in control.range('A19:A33').value if item is not None] 

    buy_order_df = pd.DataFrame(index=buy_symbols,columns=["symbol","order_list" , "bid_price", "ask_price" , "total_loss", "total_profit" , "sl_value" , "tp_value","last_sl_value","last_tp_value"])
    buy_order_df["symbol"] = buy_symbols
    sell_order_df = pd.DataFrame(index=sell_symbols,columns=["symbol","order_list" , "bid_price", "ask_price" , "total_loss", "total_profit" , "sl_value" , "tp_value","last_sl_value","last_tp_value"])
    sell_order_df["symbol"] = sell_symbols
    
    Thread(target=sync_excel , args=(path,)).start()
    Thread(target=bid_ask_updater).start()
    time.sleep(4)
    
    while True:
        for symbol in buy_order_df.index:
            print(symbol,end=" : ")
            
            if total_buy_positions(symbol) > 0 :
                buy_order_function(symbol)
                all_buy_order_close(symbol)
                buy_order_sltp(symbol)
            if total_sell_positions(symbol)>0:
                print(total_sell_positions(symbol))
                sell_order_function(symbol)
                all_sell_order_close(symbol)
                sell_order_sltp(symbol)
            else:
                print("not any oreder")
        time.sleep(1) 





