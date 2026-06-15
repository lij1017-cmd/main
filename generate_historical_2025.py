import pandas as pd
import warnings
from backtest_updated import BacktesterVol, clean_data, apply_new_stocks_registry, calculate_metrics_dual, export_to_excel_premium

warnings.filterwarnings('ignore')
DATA_FILE = "資料26Q2-1.xlsx"
LAST_DATE = "2025-12-31"
NEW_STOCKS_REGISTRY = {"2026-04-01": ["3481群創", "6446藥華藥", "2368金像電", "2334華邦電", "3037欣興", "2449京元電", "7769鴻勁"]}

def main():
    prices_raw, volumes_raw, code_to_name = clean_data(DATA_FILE)
    prices_filtered, volumes_filtered = apply_new_stocks_registry(prices_raw, volumes_raw, NEW_STOCKS_REGISTRY)
    prices_filtered = prices_filtered.loc[prices_filtered.index <= pd.to_datetime(LAST_DATE)]
    volumes_filtered = volumes_filtered.loc[volumes_filtered.index <= pd.to_datetime(LAST_DATE)]
    bt = BacktesterVol(prices_filtered, volumes_filtered, code_to_name)
    eq, t, t2, d, h = bt.run(use_market_filter=False)
    metrics = calculate_metrics_dual(eq, 30000000, 150000000)
    export_to_excel_premium(eq, t, t2, d, h, metrics, "trendstrategy_results_historical_2019-2025.xlsx")
if __name__ == "__main__":
    main()
