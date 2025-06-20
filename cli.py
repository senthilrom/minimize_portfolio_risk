import argparse
from datetime import datetime
from pathlib import Path
import pandas as pd

from services.data_fetch import StockDataFetcher
from services.portfolio import PortfolioOptimizer

def run_cli():
    parser = argparse.ArgumentParser(description="Run stock portfolio optimization from command line.")
    parser.add_argument("--tickers", nargs="+", required=True, help="List of NSE tickers (e.g., RELIANCE INFY TCS)")
    parser.add_argument("--start-date", required=True, help="Start date in YYYY-MM-DD format")
    parser.add_argument("--output", default="output/Stock-Risk-CLI.xlsx", help="Path to save output Excel file")
    args = parser.parse_args()

    try:
        start_date = datetime.strptime(args.start_date, "%Y-%m-%d").date()
        tickers = [t + ".BO" for t in args.tickers]

        print("📥 Fetching data...")
        fetcher = StockDataFetcher(tickers, start_date)
        df = fetcher.fetch_all()

        print("📊 Optimizing portfolio...")
        optimizer = PortfolioOptimizer(df)
        result = optimizer.optimize_weights()

        print("✅ Optimization Complete")
        print(f"Risk: {result.risk:.4f}")
        print(f"Expected Return: {result.expected_return:.4f}")
        print("\nOptimized Weights:")
        print(result.weights[['weights_rounded']])

        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(out_path) as writer:
            df.index = df.index.tz_localize(None)
            df.to_excel(writer, sheet_name="Stock-Data")
            result.weights.to_excel(writer, sheet_name="Optimized-Weights")

        print(f"📁 Results saved to {out_path}")

    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    run_cli()