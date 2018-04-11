# stock-forex-API-XLS
Get stock and forex quotes from API, export them in Excel tables.

Brief description:
This script gets stock (& ETF) and foreign currency exchange ('forex') quotes from two different online API-s.
The relevant data (eg. closing price) is exported and stored in an Ms Excel file, with the spreadsheet table ranges being expanded each time, so other functions or pivot tables can still use the reference (table name) when the file is reopened. 

Motivation:
I needed a quick and easy way to track current prices in ETF investment portfolio, along with it's value in foreign currencies and Hungarian Forint (HUF).
Since I'm new to Python (and coding in general), this turned out to be a fun and useful way to learn about Classes, API-s, the JSON format, the 'requests' module to get online data, and ways to edit Excel worksheets and tables.

The script
  1. Finds the date of the last update in the spreadsheets
      (Cell A2, since the the dates and quotes are stored with the latest at the top.)
  2. If the latest date is earlier than today, inserts the proper number of empty rows under the header row, for new dates and data.
  3. Fills column A with dates from present to the starting date in past, sets cell value types 'datetime.date'.
  4. Gets the API data for the stock worksheet first, in JSON format. 
  5. In columns B--, writes the quote for the date (column A) and ticker (header row) in each cell.
  6. Cells for dates with no online data (weekend days) will get the last available value.
  7. On the forex worksheet, in a similar but not same manner:
      Columns 4-6 (EUR, GBP, HUF) get quotes from another API. The API gives only USD based quotes.
  8. I'm interested in HUF based prices, columns 2-3 (EUR/HUF, GBP/HUF) will be written formulas to calculate those.
  9. After each data update, the ranges for the existing Excel tables ('QuoteTable', 'ForexTable') will be modified.
      That is: closing coordinate of table will be same column, but new "last row" in worksheet.
  10. Worksheets can be saved together to the same filname, or another one based on user input.
  11. At various points the script responds to unexpected input (date mismatch in first column, no ticker in header row, file can't be         overwritten.)
  
  Configuration:
  Certain parameters are in a separate config file for quick access:
  - file name, working directory
  - name of worksheets and tables
  - first cell where the 'dates' data start
  - access keys to the API-s
  - the functions that get the API data.
    
  Requirements:
  1. Script was written in Python 3.6
  2. Uses the following modules:
      - os, time, datetime (probably installed with your Python package)
      - requests, openpyxl: these might have to be installed before use.
  3. You will need an API access key to get online data. These can be requested for free. Copy-paste it in the config file.
      https://www.alphavantage.co/
      https://apilayer.com/
     If you want to use another API service, or data with a different content and JSON structure, you will probably have to write your      own functions for that, or at least edit mine.     
  4. An Excel (.xlsx) file with the following parameters, if you want to use the script 'as-is':
      - two separates with worksheets for the quote and forex data (if you wan't to use both).
      - Column A down from row 2 should be left empty (or filled with proper dates).
      - tables set up in each, starting from A1, with the header row from B1-- showing the tickers (eg, 'CSX5.AS' / currency pairs (eg.         'USDEUR'). The tables should include the header row. 
      - You will have to find out how your stock ticker is stored in the API's database, especially if you want to get the quote on a           specific stock exchange outside the US.
        
   Again, please note that this script is not completely user-friendly / fool-proof, meaning that if you want to customize it for a       different spreadhseet structure, different API, or different stock data type, that will need tinkering around with the code!
      
    
