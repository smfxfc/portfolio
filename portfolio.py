#! python3
# portfolio.py - Pull stock data from an API and update portfolio tracker accordingly
# the script will run once a day after market close in order to capture and save a snapshot of the portfolio's value on any date whilst script was active 

from datetime import datetime

import text_myself
import openpyxl
import yfinance as yf

portfolio = 'portfolio.xlsx'
portfolio_history = 'portfolio_history.xlsx'

wb = openpyxl.load_workbook(portfolio)
destination = openpyxl.load_workbook(portfolio_history)

current = wb['Current']

# copy 'total market value' cell in to a new cell and store it as 'yesterday's portfolio value'. This will be implemented before the cells are updated to reflect the current day's market holdings, so that after current prices are updated I can calculate today's +- by taking subtracting the two.  
current.cell(row=2, column=13).value = current.cell(row=7, column=8).value

# Loop through rows to grab ticker references which will be used to pull stock data from yf
for row_num in range(2, current.max_row): # skipping first row because it's the header. leaving last row because it's the total
    stock_ticker = (current.cell(row=row_num, column=2)).value
    print(stock_ticker)
    stock_ticker_string = str(stock_ticker)
    stock_info = yf.Ticker(stock_ticker_string)
    print(stock_info)
    stock_data = stock_info.history(period = '5d')
    current_quote = (stock_data.tail(1)['Close'].iloc[0])
    yesterday_quote = (stock_data.tail(2)['Close'].iloc[0])
    current.cell(row=row_num, column=7).value = current_quote
    current.cell(row=row_num, column=3).value = (current_quote - yesterday_quote)/yesterday_quote
    current.cell(row=row_num, column=8).value = current.cell(row=row_num, column=7).value * current.cell(row=row_num, column = 4).value
    current.cell(row=row_num, column=9).value = (current.cell(row=row_num, column=8).value - current.cell(row=row_num, column=6).value)
    current.cell(row=row_num, column=10).value = current.cell(row=row_num, column=9).value / current.cell(row=row_num, column=6).value * 100

# calcate today's portfolio +-  
current.cell(row=7, column=8).value = current.cell(row=2, column=8).value + current.cell(row=3, column=8).value + current.cell(row=4, column=8).value + current.cell(row=5, column=8).value + current.cell(row=6, column=8).value
current.cell(row=1, column=13).value = round(float(current.cell(row=7, column=8).value) - float(current.cell(row=2, column=13).value),2)

# create tab in portfolio_history.xlsx file to store the day's portfolio value, naming the tabs as the current date
destination.create_sheet(datetime.today().strftime('%y-%m-%d'))
destination.save(portfolio_history)
export_tab = destination[datetime.today().strftime('%y-%m-%d')]

# get data from portfolio.xlsx in order to copy it to portfolio_history.xlsx
mr = current.max_row
mc = current.max_column

for i in range(1, mr + 1):
    for j in range(1, mc + 1):
        c = current.cell(row = i, column = j)
        export_tab.cell(row = i, column = j).value = c.value

wb.save(portfolio)
destination.save(portfolio_history)


text_myself.textmyself("Today's portfolio movement: $" + str(current.cell(row=1, column=13).value)) # TODO: round the value to 2 decemals