import yfinance as yfin
import pandas as pd
import numpy as np



class StockQuery:
    def __init__(self, ticker: str = None, ticker_time: str = None):
        self.stock_tick = ticker
        self.stock_time = ticker_time
        self.total_mean_growth = None

    def query_ticker(self):
        try:
            ticker_query = yfin.Ticker(self.stock_tick)
            return ticker_query
        except Exception:
            return None

    def create_stock_dataframe(self):
        try:
            ticker_data = self.query_ticker()
            df = pd.DataFrame(ticker_data.history(period=self.stock_time)).reset_index()

            df[["Open", "Close"]] = df[["Open", "Close"]].round(2)
            formatted_df = df[["Date", "Open", "Close"]].copy()
            formatted_df["Date"] = pd.to_datetime(formatted_df["Date"]).dt.date
            formatted_df["Average_Change"] = (formatted_df["Close"] - formatted_df["Open"]).round(2)

            formatted_df.rename_axis(f"Company:{self.stock_tick}", inplace=True)

            self.get_mean_change(formatted_df)

            return formatted_df
        except Exception as e:
            df = pd.DataFrame()

            return df

    def get_mean_change(self, df):
        total_mean_change = np.mean(df["Average_Change"].values).round(2)
        return total_mean_change

    def get_total_change(self, df):
        total_change = (df["Close"].iloc[-1] - df["Open"].iloc[0]).round(2)
        return total_change

    def get_percent_change(self, df):
        close = df["Close"].iloc[-1]
        open = df["Open"].iloc[0]

        percent_change = ((close - open)/open) * 100
        return percent_change
