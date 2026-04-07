import pandas as pd
import numpy as np
from backtest_v2 import clean_data, BacktesterV2, calculate_metrics

def main():
    # 1. 資料讀取與清洗
    prices, volumes, code_to_name = clean_data('樣本集-1.xlsx')

    # 2. 設定參數與執行回測
    # SMA=303, ROC=14, SL=0.0999, Reb=9
    sma, roc, sl, reb = 303, 14, 0.0999, 9
    bt = BacktesterV2(prices, volumes, code_to_name)
    eq, trades, hold, trades2, daily = bt.run(sma, roc, sl, reb, 'peak', 10)

    # 3. 篩選指定期間 (2019-2025)
    mask = (eq['日期'] >= '2019-01-01') & (eq['日期'] <= '2025-12-31')
    eq_full = eq[mask].copy()

    # 4. 年度績效計算
    annual_results = []
    years = range(2019, 2026)
    for year in years:
        year_mask = (eq_full['日期'].dt.year == year)
        eq_year = eq_full[year_mask].copy()

        if eq_year.empty:
            continue

        # 計算年度報酬率
        # 使用該年度最後一天權益 / 該年度第一天權益 - 1
        # 但如果是從 2019 年開始，且 2019 年只有部分數據（因為指標計算需要緩衝），
        # 則計算該段時間的表現。
        total_ret = (eq_year['權益'].iloc[-1] / eq_year['權益'].iloc[0]) - 1

        # 年度年化報酬率 (CAGR)
        days = (eq_year['日期'].iloc[-1] - eq_year['日期'].iloc[0]).days
        if days > 0:
            years_frac = days / 365.25
            cagr = (1 + total_ret) ** (1 / years_frac) - 1
        else:
            cagr = total_ret

        # 計算年度 MDD (當年度內的最大回撤)
        eq_year['Peak'] = eq_year['權益'].cummax()
        eq_year['Drawdown'] = (eq_year['權益'] - eq_year['Peak']) / eq_year['Peak']
        mdd = eq_year['Drawdown'].min()

        annual_results.append({
            '年份': f"{year}年",
            'CAGR': cagr,
            'MDD': mdd
        })

    df_annual = pd.DataFrame(annual_results)

    # 5. 全期間績效計算 (2019-2025)
    cagr_full, mdd_full, calmar_full, total_ret_full = calculate_metrics(eq_full)

    summary_data = [
        {'項目': '2019-2025 總計', 'CAGR': cagr_full, 'MDD': mdd_full, '總報酬率': total_ret_full}
    ]
    df_summary = pd.DataFrame(summary_data)

    # 6. 輸出到 EXCEL
    output_file = 'equityV-annual.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df_annual.to_excel(writer, sheet_name='年度績效', index=False)
        df_summary.to_excel(writer, sheet_name='全期間總計', index=False)

        # 格式化
        workbook = writer.book
        percent_fmt = workbook.add_format({'num_format': '0.00%'})

        # 年度績效分頁格式化
        worksheet_annual = writer.sheets['年度績效']
        worksheet_annual.set_column('B:C', 15, percent_fmt)

        # 總計分頁格式化
        worksheet_summary = writer.sheets['全期間總計']
        worksheet_summary.set_column('B:D', 15, percent_fmt)

    print(f"結果已存入 {output_file}")

    # 打印結果供確認
    print("\n年度績效:")
    print(df_annual.to_string(index=False))
    print("\n全期間總計:")
    print(df_summary.to_string(index=False))

if __name__ == "__main__":
    main()
