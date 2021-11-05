# compiled dataframe is stored in this csv
dataframe_csv_filename ="./data/compiled_dataframe.csv" #this file name does not need to be changed


import openpyxl as xl  # openpyxl is an excel manipulation library
import pandas as pd
import numpy as np
import utils
import pickle
import os
import pdb

df = pd.read_csv(dataframe_csv_filename) # read in the compiled dataframe from the csv


# clean up data
df = df[df.apply(utils.hasNoNumbers, axis = 1 )] # removes tickers that have numbers-- bonds and private funds
df['Market Cap'] = pd.to_numeric(df['Market Cap'])
df['Market Val'] = pd.to_numeric(df['Market Val'].replace(r'^\s*$', 0, regex=True))
df = df[df['Market Val'] != 0]
df['Mkt Val Next'] = df['PX Ratio Next'] * df['Market Val']
df['Mkt Val Next'] = pd.to_numeric(df['Mkt Val Next'].replace(r'^\s*$', 0, regex=True))
df['Mkt Val Next'] = df['Mkt Val Next'].fillna(0)
# df = df[df['Mkt Val Next'] != 0]

df['Market Cap%'] = 100 * df['Market Val'] / (1000000 * df['Market Cap'])

df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%y')
df['Month-Year'] = df['Date'].dt.to_period('M')
df_reportCAP = pd.DataFrame(columns=df.columns)
df_reportVAL = pd.DataFrame(columns=df.columns)

# start work here


#For Market Cap%
#######################################

#creating a list of top Market Cap % owned
df_topCAP = df.groupby(by=['Month-Year','Fund Name'] , as_index = False).apply(lambda df_iterate: df_iterate.nlargest(20,'Market Cap%') ) # for each month-year and fund combination, find the largest 20 in Market Cap % owned

#creating a list that filters the last Market Cap% list by the number of funds that hold each position
df_DateTickerCAP = df_topCAP.groupby(by=['Month-Year','Ticker'], as_index=False) #for each month-year and ticker combination...
df_countsCAP = df_DateTickerCAP.agg(['count']) #... count repeated instances for each column category
df_filteredCAP = df_countsCAP['Fund Name'][df_countsCAP['Fund Name']>=2].dropna() #show which positions have X or more funds holding them, drop those N/A

# my= month-year, t= ticker
for my, t in df_filteredCAP.index:
    df_popularCAP = df_topCAP[(df_topCAP['Month-Year'] == my) & (df_topCAP['Ticker'] == t)]
    df_reportCAP = df_reportCAP.append(df_popularCAP)
    print(df_popularCAP[['Date', 'Ticker', 'Fund Name']])

#final report has these parameters, can add more so long as they are defined in the _raw file or previously defined in ticker_count_across_fundsCAP.py
df_report_nicerCAP = df_reportCAP[['Date', 'Ticker', 'Fund Name', 'Market Cap%']]
df_report_nicerCAP.to_csv('./reports/temp_reportCAP.csv', index=False)



#For Market Val
#######################################


df_topVAL = df.groupby(by=['Month-Year','Fund Name'] , as_index = False).apply(lambda df_iterate: df_iterate.nlargest(10,'Market Val') ) ###

df_DateTickerVAL = df_topVAL.groupby(by=['Month-Year','Ticker'], as_index=False)
df_countsVAL = df_DateTickerVAL.agg(['count'])

df_filteredVAL = df_countsVAL['Fund Name'][df_countsVAL['Fund Name']>=2].dropna() ###


for my, t in df_filteredVAL.index:
    df_popularVAL = df_topVAL[(df_topVAL['Month-Year'] == my) & (df_topVAL['Ticker'] == t)]
    df_reportVAL = df_reportVAL.append(df_popularVAL)
    print(df_popularVAL[['Date', 'Ticker', 'Fund Name']])

df_report_nicerVAL = df_reportVAL[['Date', 'Ticker', 'Fund Name', 'Market Val']]
df_report_nicerVAL.to_csv('./reports/temp_reportVAL.csv', index=False)








