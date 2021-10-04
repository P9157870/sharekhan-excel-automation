import pandas as pd
import xlsxwriter
import datetime
import yfinance as yf

#* reading and seggregating data
excel_file = r"Master Stock Data.xlsx"
df = pd.read_excel(excel_file,engine='openpyxl')
buy_df = df[df['Buy/Sell'] == 'B'].reset_index(drop=True)
sell_df = df[df['Buy/Sell'] == 'S'].reset_index(drop=True)


#* resultant df
resultant_df =  pd.DataFrame(columns = [
    "S.No",
    "Entry Date",       #? Trade Date aka Buy date in Equity Tradelisting sheet
    "Stocks",           #? Scrip Symbol - E.g : AJANTPHARM
    "Transaction",      #? Buy or Sell - here in this column, it's BUY only
    "Buy Quantity",     #? Buying Quantity  - column F  in BUY Category 
    "Entry Price",      #? Rate - Column G in BUY Category
    "Buy Cost",         #? Brokerage + (IGST + SGST + CGST + UTGST + Cess + Stamp + TO Charges + Sebi Fees + STT Amount)
    "Buy Value",        #? Buy Quantity * Entry Price + Buy Cost
    
    "Exit Date",        #? Sell Date in respective Sell Category --> look at immediate sold stock for respective stock
    "Sell Quantity",    #? Sell Quantity (need not be same as entry)
    "Exit Price",       #? Rate - Column G in SELL Category
    "Sell Cost",        #? Brokerage + (IGST + SGST + CGST + UTGST + Cess + Stamp + TO Charges + Sebi Fees + STT Amount)
    "Sell value",       #? Sell Quantity * Exit Price + Sell Cost
    
    "PNL",              #? Profit or Loss : Exit Price - Entry Price
    "Net PNL",          #? Sell Value - Buy Value
    "Percentage",       #? Net PNL / Buy Value
    "Cumulative Profit",#? We have to sum the Net PNL for every past row and current row and make that as CP
])


#* adding initial buy data to resultant_df
resultant_df["Entry Date"] = buy_df["Trade Date"]
resultant_df["Stocks"] = buy_df["Scrip Symbol"]
resultant_df["Transaction"] = "Buy"
resultant_df["Buy Quantity"] = buy_df["Quantity"]
resultant_df["Entry Price"] = buy_df["Rate"]
resultant_df["Buy Cost"] = buy_df["Brokerage"] + buy_df["IGST"] + buy_df["SGST"] + buy_df["CGST"]\
                            + buy_df["UTGST"] + buy_df["Cess"] + buy_df["Stamp"] + buy_df["TO Charges"] + buy_df["Sebi Fees"]\
                            + buy_df["STT Amount"]
resultant_df["Buy Value"] = (resultant_df["Buy Quantity"] * resultant_df["Entry Price"]) + resultant_df["Buy Cost"]



#* function for one buy and one sell
def fill_data_for_single_buy_single_sell(stock, individual_stock_sell_df):
    temp_sell_df    = individual_stock_sell_df.iloc[0]
    exit_date       = temp_sell_df["Trade Date"]
    sell_quantity   = temp_sell_df["Quantity"]
    exit_price      = temp_sell_df["Rate"]
    sell_cost       = temp_sell_df["Brokerage"] + temp_sell_df["IGST"] + temp_sell_df["SGST"] + temp_sell_df["CGST"]\
                        + temp_sell_df["UTGST"] + temp_sell_df["Cess"] + temp_sell_df["Stamp"] + temp_sell_df["TO Charges"]\
                        + temp_sell_df["Sebi Fees"] + temp_sell_df["STT Amount"]
                    
    resultant_df.loc[resultant_df['Stocks'] == stock,
                    [
                    "Exit Date",
                    "Sell Quantity",
                    "Exit Price",
                    "Sell Cost"
                    ]
                ] = [
                        exit_date,
                        sell_quantity,
                        exit_price,
                        sell_cost
                    ]

#* code till single buy and single sell

total_stock_types = list(set(resultant_df["Stocks"].to_list()))

for stock in total_stock_types:
    individual_stock_buy_df = resultant_df[resultant_df["Stocks"] == stock]
    if len(individual_stock_buy_df) == 1:
        individual_stock_sell_df = sell_df[sell_df["Scrip Symbol"] == stock]
        if len(individual_stock_sell_df) == 1:
            if individual_stock_buy_df.iloc[0]["Buy Quantity"] == individual_stock_sell_df.iloc[0]["Quantity"]:
                fill_data_for_single_buy_single_sell(stock,individual_stock_sell_df)
                
                