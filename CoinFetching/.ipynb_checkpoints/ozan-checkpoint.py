import json
import time
import requests
import xlwings as xw

WORKBOOK_PATH = 'yakuphan.xlsx'
SHEET_NAME = 'Sheet1'

# Open Excel workbook and select worksheet
workbook = xw.Book(WORKBOOK_PATH)
worksheet = workbook.sheets[SHEET_NAME]

# Add headers to Excel sheet
worksheet.range('A1').value = 'Timestamp'
worksheet.range('B1').value = 'BTC Rate'
worksheet.range('C1').value = 'ETH Rate'
worksheet.range('D1').value = 'STL Rate'

# API endpoint URL
url = 'http://api.coinlayer.com/api/live?access_key=1263ca355c08daeece63debf97d4bfe5'

row = 2  # Start writing data from row 2

# Loop to update Excel sheet with latest rates
while True:
    # Get data from API
    response = requests.get(url)
    data = json.loads(response.text)

    # Extract relevant data from response
    success = data['success']
    terms = data['terms']
    privacy = data['privacy']
    timestamp = data['timestamp']
    target = data['target']
    rates = data['rates']

    # Extract specific rates
    btc_rate = rates['BTC']
    eth_rate = rates['ETH']
    stl_rate = rates['STK']

    # Write data to Excel sheet
    worksheet.range(f'A{row}').value = timestamp
    worksheet.range(f'B{row}').value = btc_rate
    worksheet.range(f'C{row}').value = eth_rate
    worksheet.range(f'D{row}').value = stl_rate

   

    # Wait for 10 seconds
    print(timestamp, btc_rate, eth_rate, stl_rate)
    time.sleep(30)