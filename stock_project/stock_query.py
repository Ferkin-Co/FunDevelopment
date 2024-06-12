import yfinance as yfin
import pandas as pd
import numpy as np
import traceback



class StockQuery:
    def __init__(self, ticker: str = None, ticker_time: str = None):
        self.stock_tick = ticker
        self.stock_time = ticker_time

    def query_ticker(self):
        ticker_query = yfin.Ticker(self.stock_tick)

        return ticker_query

    def create_stock_dataframe(self):
        try:
            ticker_data = self.query_ticker()
            df = pd.DataFrame(ticker_data.history(period=self.stock_time)).reset_index()
            pd.set_option('display.max_columns', None)

            formatted_df = df[["Date", "Open", "Close"]].copy()
            formatted_df["Date"] = pd.to_datetime(formatted_df["Date"]).dt.date
            formatted_df["Average Growth"] = formatted_df["Close"] - formatted_df["Open"]
            total_mean = np.mean(formatted_df["Average Growth"])
            # formatted_df.at[0, "Total Mean"] = np.mean(formatted_df["Average Growth"])
            formatted_df.rename_axis(f"Company:{self.stock_tick}", inplace=True)

            return print(f"{formatted_df}\nTotal Mean Growth: {total_mean}")
        except Exception as e:
            traceback_info = traceback.format_exc()
            return print(f"Error: {e}\n{traceback_info}")

# try:
#     while True:
#         user_input = input("Enter a stock ticker >> ")
#         if 0 < len(user_input) <= 4 and user_input.isalpha():
#             break
#         else:
#             print("Stock tickers cannot be more than 4 characters or less than 0")
#             continue
# except Exception as e:
#     traceback_info = traceback.format_exc()
#     print(f"Error: {e}\n{traceback_info}")



# stonk = StockQuery(user_input.upper(), "1mo")
# stonk.create_stock_dataframe()
