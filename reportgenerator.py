"""
NSE Ticker Statistics Report Generator
"""

import os
import yfinance as yf
import pandas as pd 
import numpy as np
from datetime import datetime 

"""
FUNCTIONS
"""

# last 52 week stock column stats
def past52wkTickerColStat(df, stock_column, statistic):
    """
    @params:
    df: Multi-Indexed Pandas Dataframe for a given list of NSE Tickers (day-level row index)
    stock_column: One of the provided NSE stock price columns (one of: Open, High, Low, Close, Adj Close, Volume)
    statistic: One of 'mean', 'median', 'min', 'max'

    @returns:
    pd.Series object with <TICKER>.NS as entries in row index, values correspond to the desired column statistic
    """
    assert isinstance(df.columns, pd.MultiIndex), 'downloaded pandas DataFrame from yfinance must have multiIndexed columns!'
    assert stock_column in ['Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume'], 'invalid stock_column. must be one of Open, High, Low, Close, Adj Close, Volume'
    assert statistic in ['mean', 'median', 'min', 'max']

    columns = list(filter(lambda x: x[0] == stock_column, df.columns))

    return df[columns].droplevel(axis=1, level=0)\
                    .apply(statistic)\
                    .rename('52w {} (INR)'.format(stock_column))

# Current Stock Price: (last 1d trading average price)
def currentTickerPrice(df):
    """
    @params:
    df: Multi-Indexed Pandas Dataframe for a given list of NSE Tickers (minute-level row index)

    @returns:
    pd.Series object with <TICKER>.NS as entries in row index, values correspond to the last 24 mean 'Close' stock price
    """
    assert isinstance(df.columns, pd.MultiIndex), 'downloaded pandas DataFrame from yfinance must have multiIndexed columns!'

    columns = list(filter(lambda x: x[0] == 'Close', df.columns))

    return df[columns]\
                .ffill()\
                .droplevel(axis=1, level=0)\
                .iloc[-30: ].mean()\
                .rename('Last 1d Avg Close (INR)')
    
"""
MAIN
"""
CD = datetime.today()

print("""
######################################
NSE TICKER STATISTICS REPORT GENERATOR

Date/Time: {}
######################################
""".format(str(CD)))

# all NSE Tickers of interest
ticker_metadata_df = pd.read_csv('TICKERS.csv')

tickers = ticker_metadata_df['SYMBOL'].values
 
yf_ticker_string = ' '.join([x+'.NS' for x in tickers])

# 52w period, 1d interval Ticker statistics
print('Last 52 weeks; 1 day Interval Ticker Data Download')
daily_52w_ticker_stats = yf.download(yf_ticker_string, period='1y')

# 1d period, 1m interval Ticker statistics
print('Last 1 Day; 1 min Interval Ticker Data Download')
minute_1d_ticker_stats = yf.download(yf_ticker_string, interval='1m', period='1d')


print('Compiling Report..')
# computing Ticker-level stats of interest
low_52w = past52wkTickerColStat(daily_52w_ticker_stats, 'Low', 'min')
high_52w = past52wkTickerColStat(daily_52w_ticker_stats, 'High', 'max')
current_1d = currentTickerPrice(minute_1d_ticker_stats)

# constructing ticker_stats_df:
ticker_stats_df = pd.concat([low_52w, high_52w, current_1d], axis=1)\
                        .reset_index()\
                        .rename(columns={'index': 'SYMBOL.NS'})  
ticker_stats_df['SYMBOL'] = ticker_stats_df['SYMBOL.NS'].str.split('.').str.get(0)
ticker_stats_df.drop(columns=['SYMBOL.NS'], inplace=True)

# Output:
report = ticker_metadata_df[['SYMBOL', 'NAME OF COMPANY']]\
                .merge(ticker_stats_df, how='left', left_on='SYMBOL', right_on='SYMBOL')\
                .reset_index(drop=True) 

report_datetime = str(CD.date()) + '-' + ''.join(str(CD.time()).split(':')[:2])

report_path = 'REPORTS'
if not os.path.exists(report_path):
    os.makedirs(report_path)

report_filename = report_path + '\\' + 'NSE_TICKER_STATS_{}.xlsx'.format(report_datetime)

report.to_excel(report_filename, engine='xlsxwriter')

print("""
##############################
NSE REPORT GENERATION COMPLETE
##############################
""")
