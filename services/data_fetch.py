from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
import pandas as pd
import yfinance as yf

class StockDataFetcher:
    def __init__(self, tickers, start_date, end_date=None):
        self.tickers = tickers
        self.start_date = start_date
        self.end_date = end_date or date.today()

    def fetch_single(self, ticker):
        try:
            data = yf.Ticker(ticker).history(start=self.start_date, end=self.end_date)
            close_col = f"{ticker}-Close"
            return data[['Close']].rename(columns={'Close': close_col})
        except Exception as e:
            print(f"Failed to fetch data for {ticker}: {e}")
            return None

    def fetch_all(self):
        dfs = []
        with ThreadPoolExecutor() as executor:
            futures = {executor.submit(self.fetch_single, t): t for t in self.tickers}
            for future in as_completed(futures):
                result = future.result()
                if result is not None:
                    dfs.append(result)

        if not dfs:
            raise ValueError("No valid stock data fetched.")

        return pd.concat(dfs, axis=1).dropna()
