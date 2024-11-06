import requests
import time
import pandas as pd  # Import pandas for data handling
import random
import xlwings as xw

API_KEY = {'X-API-key': '3CJNDK9B'}  # Step 2
BASE_API_URL = 'http://localhost:9999/v1'

class APIClient:
    def __init__(self, endpoint):
        self.api_url = f'{BASE_API_URL}/{endpoint}'
        self.session = requests.Session()
        self.session.headers.update(API_KEY)

    def fetch_data(self):
        resp = self.session.get(self.api_url)
        if resp.ok:
            return resp.json()
        else:
            resp.raise_for_status()

class Case(APIClient):
    def __init__(self):
        super().__init__('case')  # Calls the base class with 'case' endpoint
        case_data = self.fetch_data()
        self.tick = case_data['tick']
        self.period = case_data['period']
        self.name = case_data['name']
        self.ticks_per_period = case_data['ticks_per_period']
        self.total_periods = case_data['total_periods']
        self.status = case_data['status']

    def get_current_tick(self):
        case_data = self.fetch_data()  # Fetch the latest case data
        self.tick = case_data['tick']  # Update the current tick
        
        # Adjust current_tick based on case_data["period"]
        if case_data["period"] == 1:
            current_tick = self.tick  # Keep current_tick the same
        elif case_data["period"] == 2:
            current_tick = self.tick + 312  # Increment current_tick by 312
        
        return current_tick  # Return the adjusted current_tick

class Securities(APIClient):
    def __init__(self):
        super().__init__('securities')
        self.securities_data = self.fetch_data()

    def collect_valid_cb_zc_data(self, current_tick):
        valid_data = []
        for security in self.securities_data:
            if self.is_valid_ticker(security["ticker"]):
                valid_data.append({
                    "ticker": security["ticker"],
                    "tick": current_tick,  # Include the current tick
                    "bid": security["bid"],
                    "last": security["last"],
                    "ask": security["ask"]
                })
        return valid_data

    def is_valid_ticker(self, ticker):
        # Check if the ticker starts with "CB" or "ZC" and is followed by digits
        return ticker != "CAD"

    def export_to_excel(self, valid_data,instance):
        df = pd.DataFrame(valid_data)
        filename = 'valid_cb_zc_data'+str(instance)+'.xlsx'
        df.to_excel(filename, index=False)  # Save DataFrame to Excel
        print(f'Data exported to {filename}')


class News(APIClient):
    def __init__(self):
        super().__init__('news')
        self.news_data = self.fetch_data()
        #self.collected_news_ids = set()  # Set to track collected news IDs

    def collect_news_data(self, current_tick):
        news_collected = []
        for news in self.news_data:
            # Check if the news_id has already been collected
            if current_tick != 0:
                # print("#####NEWS#####")
                # print(news['period'])
                if int(news['period']) == 1:
                    news_collected.append({
                        "news_id": news['news_id'],
                        "tick": news['tick'],
                        "ticker": news['ticker'],
                        "headline": news['headline'],
                        "body": news['body']
                    })
                else:
                    news_collected.append({
                        "news_id": news['news_id'],
                        "tick": int(news['tick']) + 312,
                        "ticker": news['ticker'],
                        "headline": news['headline'],
                        "body": news['body']
                    })
        return news_collected

    def export_news_to_excel(self, news_data,instance):
        df = pd.DataFrame(news_data)
        filename = 'news_data'+str(instance)+'.xlsx'
        df.to_excel(filename, index=False)  # Save DataFrame to Excel
        print(f'News data exported to {filename}')

def main(instance):
    case_instance = Case()
    securities_instance = Securities()
    news_instance = News()  # Create a  n instance of News

    all_valid_cb_zc_data = []  # List to collect all valid CB and ZC data for export
    all_news_data = []  # List to collect all news data for export
    previous_tick = -1  # Initialize previous tick to an invalid value
    no_update_duration = 0  # Timer for tracking no updates
    update_threshold = 45  # Duration in seconds before export

    while True:  # Continuously monitor for tick updates
        # df = pd.read_excel('RIT - Decision Support - FI Capstone Case - vRelease.xlsx', sheet_name='Base Support Sheet')  # Specify the sheet name
        current_tick = case_instance.get_current_tick()  # Get the adjusted current tick
        wb = xw.Book("RIT - Decision Support - FI Capstone Case - vRelease.xlsx")  # Specify your workbook name
        sheet = wb.sheets["Base Support Sheet"]  # Specify your sheet name

        if current_tick != previous_tick:  # Check if the tick has changed
            # print(df['Unnamed: 7'])
            #print(f'Tick updated to: {current_tick}')  # Print the updated tick
            value = sheet.range("H32").value
            print("Live value in cell H32:", value)
            
            # Fetch updated securities data
            securities_instance.securities_data = securities_instance.fetch_data()
            
            # Collect data for valid CB and ZC tickers
            valid_cb_zc_data = securities_instance.collect_valid_cb_zc_data(current_tick)
            all_valid_cb_zc_data.extend(valid_cb_zc_data)  # Add current tick data to the list

            # Collect news data for the current tick
            news_data = news_instance.collect_news_data(current_tick)
            #print(news_data)
            all_news_data.extend(news_data)
            # Only append unique news and print if new news is collected
            # if news_data:
            #     for news in news_data:
            #         if news not in all_news_data:  # Check for uniqueness
            #             all_news_data.append(news)  # Append unique news data
            #             print(f'New News Appended: Tick: {news["tick"]}, News ID: {news["news_id"]}, Ticker: {news["ticker"]}, Headline: {news["headline"]}')

            # Print the collected data for the current tick
            # for data in valid_cb_zc_data:
            #     print(f'Tick: {data["tick"]}, Ticker: {data["ticker"]}, Bid: {data["bid"]}, Last: {data["last"]}, Ask: {data["ask"]}')
            
            for data in news_data:
                if int(data["tick"]) == current_tick:
                    print(f'Tick: {data["tick"]}, Ticker: {data["ticker"]}, Headline: {data["headline"]}, Body: {data["body"]}')

            previous_tick = current_tick  # Update the previous tick
            no_update_duration = 0  # Reset the no update timer

        else:
            # Increment the no update timer
            no_update_duration += 1
            print(f'No update for {no_update_duration} seconds.')

        # Check if no updates have occurred for the threshold duration
        if no_update_duration >= update_threshold:
            print(f'No updates for {update_threshold} seconds. Exporting data...')
            securities_instance.export_to_excel(all_valid_cb_zc_data,instance)
            news_instance.export_news_to_excel(all_news_data,instance)
            break  # Exit the loop

        time.sleep(1)  # Sleep for a bit before checking again; adjust as needed

def delegate_capacity(max_capacity, num_items):
    # Create a list to hold the allocated capacities
    allocations = [0] * num_items
    remaining_capacity = max_capacity

    for i in range(num_items - 1):
        # Randomly allocate between 0 and the remaining capacity
        allocation = random.randint(0, remaining_capacity)
        allocations[i] = allocation
        remaining_capacity -= allocation

    # Assign the remaining capacity to the last item
    allocations[-1] = remaining_capacity

    return allocations

def trade():
    with requests.Session() as s:
        s.headers.update(API_KEY)   
        capacity = 25000
        num_items = 3  # Number of items to allocate to

        # Allocate capacities
        allocations = delegate_capacity(capacity, num_items)
        tickers = {
            "CB2017": [allocations[0], "BUY"],
            "CB2020": [allocations[1], "BUY"],
            "CB2025": [allocations[2], "SELL"]
        }

        for ticker in tickers:
            quant = tickers[ticker][0]
            print(f"Initial quantity for {ticker}: {quant}")
            
            while quant > 0:
                print(f"Processing {quant} for {ticker}")

                if quant >= 1000:
                    mkt_buy_params = {
                        'ticker': ticker,
                        'type': 'MARKET',
                        'quantity': 1000,
                        'action': tickers[ticker][1]
                    }
                    quant -= 1000
                else:
                    mkt_buy_params = {
                        'ticker': ticker,
                        'type': 'MARKET',
                        'quantity': quant,
                        'action': tickers[ticker][1]
                    }
                    quant -= quant

                # Send the order to the API
                resp = s.post('http://localhost:9999/v1/orders', params=mkt_buy_params)

                if resp.ok:
                    mkt_order = resp.json()
                    order_id = mkt_order['order_id']
                    print(f'The market order for {ticker} was submitted and has ID: {order_id}')
                else:
                    print(f"FAILED to submit order for {ticker}. Response: {resp.text}")

                time.sleep(0.015)
                        
                

            # mkt_buy_params = {'ticker': 'CB2017', 'type': 'MARKET', 'quantity': 1000,
            # 'action': 'BUY'}
            # resp = s.post('http://localhost:9999/v1/orders', params=mkt_buy_params)
            # if resp.ok:
            #     mkt_order = resp.json()
            #     id = mkt_order['order_id']
            #     print('The market buy order was submitted and has ID', id)
            #     time.sleep(0.015)
            # else:
            #     break

def update_portfolio(previous_portfolio, max_holdings):
    # Calculate the total bonds currently held (long)
    total_long = sum(max(0, amount) for amount in previous_portfolio)
    total_short = sum(max(0, -amount) for amount in previous_portfolio)

    # Maximum capacity after selling and buying
    max_remaining_capacity = max_holdings - total_long + total_short

    # Generate new long and short positions randomly
    new_long = random.randint(0, max_remaining_capacity)
    new_short = random.randint(0, total_long + total_short)

    # Determine how many bonds to buy/sell to achieve new positions
    bond_transactions = []
    for amount in previous_portfolio:
        if amount >= 0:  # Long position
            bonds_to_sell = min(amount, new_short)
            bond_transactions.append(-bonds_to_sell)  # Sell (negative value)
            new_short -= bonds_to_sell
        else:  # Short position
            bonds_to_buy = min(-amount, new_long)
            bond_transactions.append(bonds_to_buy)  # Buy (positive value)
            new_long -= bonds_to_buy

    # Add remaining buy/sell to the last bond in the portfolio
    bond_transactions.append(new_long - new_short)

    return bond_transactions

# Example usage
previous_portfolio = [5000, -2000, 10000, -5000]  # Current portfolio with long and short positions
max_holdings = 25000

# Generate new portfolio transactions
new_transactions = update_portfolio(previous_portfolio, max_holdings)

# print(f"Previous Portfolio: {previous_portfolio}")
# print(f"New Transactions: {new_transactions}")
# print(f"Updated Portfolio: {[prev + trans for prev, trans in zip(previous_portfolio, new_transactions)]}")



if __name__ == "__main__":
    # for i in range(3,9):
    #     main(i)
    main(2)
    #trade()



# if __name__ == '__main__':
#     try:
#         # Case information
#         # case_instance = Case()  # Create an instance of Case
#         # case_instance.display_case_info()  # Display the case information

#         # News information
#         # news_instance = News()
#         # news_instance.display_news_info()

#         # asset_instance = Securities()
#         # asset_instance.display_security_info()

#         # get_live_data(None)
#         # print(table)
#         main()
#     except Exception as e:
#         print(f'An error occurred: {e}')


# table = {"CB2017":[],"CB2020":[],"CB2025":[]}

# class News(APIClient):
#     def __init__(self):
#         super().__init__('news')
#         self.news_data = self.fetch_data()
    
#     def display_news_info(self):
#         for news in self.news_data:
#             print(f'The news id is: {news['news_id']}')
#             print(f'The period period is: {news['period']}')
#             print(f'The news is on tick: {news['tick']}')
#             print(f'The news ticker is: {news['ticker']}')
#             print(f'The headline is: {news['headline']}')
#             print(f'The body of the news: {news['body']}')

# def get_live_data(tick):
#     case = Case()
#     security = Securities()
#     if case.tick != tick:
#         print("\n")
#         tick = case.tick
#         print("Time: "+str(tick))
#         if int(tick) == 624:
#             print("POOFAS")
#             print(tick)
#             exit()
#         else:
#             security.security_info()
#             for sec in security.securities_data:
#                 if "CB" in sec["ticker"]:
#                     table[sec["ticker"]].append(sec["last"])
#             time.sleep(1)
#             get_live_data(case.tick)

