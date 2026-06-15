import pandas as pd
import warnings
from backtest_updated import BacktesterVol, clean_data, apply_new_stocks_registry, calculate_metrics_dual, export_to_excel_premium

warnings.filterwarnings('ignore')
DATA_FILE = "資料26Q2-1.xlsx"
PREV_FILE = "trendstrategy_results_historical_2019-2025.xlsx"
START_DATE_2026 = "2026-01-01"

def get_warm_start_data(prev_file):
    df_eq = pd.read_excel(prev_file, sheet_name='Equity_Curve')
    last_equity = df_eq.iloc[-1]['權益']
    df_daily = pd.read_excel(prev_file, sheet_name='Daily')
    last_date = df_daily.iloc[-1]['日期']
    df_last_day = df_daily[df_daily['日期'] == last_date]
    cash = last_equity - df_last_day['市值'].sum()
    warm_slots = {i: {'code': str(row['股票代號']), 'shares': row['持有股數'], 'entry_price': row['買進成本'], 'max_price': row['追蹤最高價'], 'budget': row['買入總市值'], 'entry_date': row['買進日期']} for i, (_, row) in enumerate(df_last_day.iterrows())}
    return cash, warm_slots

def main():
    try:
        cash, warm_slots = get_warm_start_data(PREV_FILE)
    except:
        return
    prices_raw, volumes_raw, code_to_name = clean_data(DATA_FILE)
    prices_filtered, volumes_filtered = apply_new_stocks_registry(prices_raw, volumes_raw, {})
    prices_filtered = prices_filtered.loc[prices_filtered.index >= pd.to_datetime(START_DATE_2026)]
    volumes_filtered = volumes_filtered.loc[volumes_filtered.index >= pd.to_datetime(START_DATE_2026)]
    bt = BacktesterVol(prices_filtered, volumes_filtered, code_to_name, warm_start_cash=cash, warm_start_slots=warm_slots)
    eq, t, t2, d, h = bt.run(use_market_filter=False, start_date=START_DATE_2026)
    metrics = calculate_metrics_dual(eq, 30000000, 150000000)
    export_to_excel_premium(eq, t, t2, d, h, metrics, f"trendstrategy_results_2026_daily_{pd.to_datetime(eq.iloc[-1]['日期']).strftime('%y%m%d')}.xlsx")
if __name__ == "__main__":
    main()
