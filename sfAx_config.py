file_to_load = "sfAx.xlsx"
working_dir = False
'''
import openpyxl
excel_file = openpyxl.load_workbook(file_to_load)
'''

# Names of worksheets.
quote_sheet_name = "Quote"
quote_table_name = "QuoteTable"

# Names of tables in each worksheet.
forex_sheet_name = "Forex"
forex_table_name = "ForexTable"


first_cell_of_date_column = 'A2'

alphavantage_key = ''
apilayer_key = ''

import requests

def get_quote_json(ticker):

    apikey = alphavantage_key
    function = 'TIME_SERIES_DAILY'
    datatype = 'json'
    outputsize = 'compact'

    res = requests.get('https://www.alphavantage.co/query?function=' + function +
                       '&symbol=' + ticker +
                       '&outputsize=' + outputsize +
                       '&datatype=' + datatype +
                       '&apikey=' + apikey)

    res.raise_for_status()
    return res.json()["Time Series (Daily)"]


def get_forex_json(reference_date):

    access_key = apilayer_key

    res = requests.get('http://apilayer.net/api/historical?access_key='
                       + access_key
                       + '&date='
                       + reference_date)

    res.raise_for_status()
    return res.json()["quotes"]
