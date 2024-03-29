import xlwings as xw
import requests
import datetime
from concurrent.futures import ThreadPoolExecutor

# @xw.sub
def extract_data():
    # Define the URLs for the APIs
    price_url = "https://fapi.binance.com/fapi/v1/premiumIndex"
    last_price_url = "https://fapi.binance.com/fapi/v1/ticker/24hr"
    
    # Create a ThreadPoolExecutor to make the API requests in parallel
    with ThreadPoolExecutor(max_workers=2) as executor:
        # Send requests to the APIs and get the responses in parallel
        price_response = executor.submit(requests.get, price_url)
        last_price_response = executor.submit(requests.get, last_price_url)
        
        # Convert the responses to JSON format
        price_data = price_response.result().json()
        last_price_data = last_price_response.result().json()
    
    # Create a list to store the data
    data = []
    # Loop through the price data and extract the symbol and price
    for item in price_data:
        symbol = item["symbol"]
        price = item["markPrice"]
        funding_rate = item['lastFundingRate']
        timestamp = item['time']
        time = datetime.datetime.fromtimestamp(timestamp / 1000).strftime('%H:%M %d/%m/%Y')
        
        # Loop through the last price data and find the corresponding last price
        for last_price in last_price_data:
            if last_price["symbol"] == symbol:
                last_price = last_price["lastPrice"]
                break
        else:
            last_price = "N/A"
        
        # Append the data to the list
        data.append([symbol, price, funding_rate, last_price,time])
    
    # Write the data to the Excel sheet
    sheet = xw.sheets.active
    sheet.clear_contents()
    sheet.range("A1").value = ["Symbol", "Mark Price", "Funding Rate", "Last Price","Time"]
    sheet.range("A2").value = data
    # Autofit the columns
    sheet.range("A1").expand("table").columns.autofit()

