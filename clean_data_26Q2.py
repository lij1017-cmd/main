import pandas as pd
import numpy as np

def clean_and_save_data(input_file, output_file):
    print(f"Cleaning data from {input_file} and saving to {output_file}...")

    # Load Excel file (reading all sheets)
    xls = pd.ExcelFile(input_file)
    sheets = xls.sheet_names

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for sheet in sheets:
            print(f"Processing sheet: {sheet}")
            df = pd.read_excel(input_file, sheet_name=sheet, header=None)

            # Data part starts from row 2, column 1
            data = df.iloc[2:, 1:].astype(float)

            if sheet == '還原收盤價':
                # Prices: ffill for suspensions, then bfill for late listings
                # The user specified: "較晚上市標的：以首次出現的價格往前期填補 ； 中途暫停交易標的：以前一日收盤價填補中間空白區域。"
                # So ffill (forward fill) first for suspensions, then bfill (backward fill) for late listings.
                data = data.ffill(axis=0).bfill(axis=0)
            elif sheet == '成交量':
                # Volumes: fill with 0
                data = data.fillna(0)
            else:
                # Any other sheets if exist
                data = data.ffill(axis=0).bfill(axis=0)

            # Put back into dataframe
            df.iloc[2:, 1:] = data

            # Save to output file
            df.to_excel(writer, sheet_name=sheet, index=False, header=False)

    print(f"Successfully saved cleaned data to {output_file}")

if __name__ == "__main__":
    RAW_DATA = '資料26Q2.xlsx'
    CLEAN_DATA = '資料26Q2-1.xlsx'
    clean_and_save_data(RAW_DATA, CLEAN_DATA)
