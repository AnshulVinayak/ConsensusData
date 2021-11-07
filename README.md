# ConsensusData
Programs to Run in Sequence:
1. 'get_data_from_bloomberg'
   - Iterates through an individual fund's Excel data dump, and finds the cell-row ranges for each fiscal quarter
   - For each range, indexes relevant fund/position metrics and removes excess data (odd labels, spaces, etc. leftover from Bloomberg)
   - Creates new indices for missing data, using Bloomberg API
   - Prints organized data into 'copypasteFundName_raw.xlsx' file, and stores ranges in 'FundName_ranges.pickle'
2. 'compile_dataframe'
   - Iterates through each individual 'copypasteFundName_raw.xlsx' file by range, and transfers data into an aggregate dataframe
   - Aggregate fund dataframe stored in 'compiled_dataframe.csv'
3. 'ticker_count_across_funds'
   - Throws out bonds, private funds, and other non-equities by filtering tickers containing numeric values
   - Throws out mergers and acquisitions by filtering NaN in 'MarketValueNext'
   - Creates Bloomberg portfolios with two methods, stored in 'temp_reportCAP.csv' and 'temp_reportVAL.csv'
   
Other Programs:
1. 'bl_getdata' - defines functions for Bloomberg API
2. 'utils' - defines functions using Pyxcel library