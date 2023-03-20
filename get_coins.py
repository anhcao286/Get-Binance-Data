import xlwings as xw
import requests

# @xw.sub
def extract_data():
    # Define the URLs for the APIs
    price_url = "https://fapi.binance.com/fapi/v1/ticker/price"
    funding_rate_url = "https://fapi.binance.com/fapi/v1/fundingRate"
    last_price_url = "https://fapi.binance.com/fapi/v1/ticker/24hr"
    
    # Send requests to the APIs and get the responses
    price_response = requests.get(price_url)
    funding_rate_response = requests.get(funding_rate_url)
    last_price_response = requests.get(last_price_url)
    
    # Convert the responses to JSON format
    price_data = price_response.json()
    funding_rate_data = funding_rate_response.json()
    last_price_data = last_price_response.json()
    
    # Create a list to store the data
    data = []
    
    # Loop through the price data and extract the symbol and price
    for item in price_data:
        symbol = item["symbol"]
        price = item["price"]
        
        # Loop through the funding rate data and find the corresponding funding rate
        for rate in funding_rate_data:
            if rate["symbol"] == symbol:
                funding_rate = rate["fundingRate"]
                break
        else:
            funding_rate = "N/A"
        
        # Loop through the last price data and find the corresponding last price
        for last_price in last_price_data:
            if last_price["symbol"] == symbol:
                last_price = last_price["lastPrice"]
                break
        else:
            last_price = "N/A"
        
        # Append the data to the list
        data.append([symbol, price, funding_rate, last_price])
    
    # Write the data to the Excel sheet
    sheet = xw.sheets.active
    sheet.clear_contents()
    sheet.range("A1").value = ["Symbol", "Price", "Funding Rate", "Last Price"]
    sheet.range("A2").value = data
