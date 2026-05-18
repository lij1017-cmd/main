import pandas as pd

def check_island():
    df = pd.read_csv('opt_best_effort.csv')
    # Target: roc=10, reb=9, atr_p=15, atr_m=4.3, mkt_t=0.42, mkt_s=14

    # Filter neighbors
    neighbors = df[
        (df['roc'] == 10) &
        (df['reb'] == 9) &
        (df['mkt_t'] == 0.42) &
        (df['mkt_s'] == 14)
    ]

    print("Neighbors around ATR_P=15, ATR_M=4.3:")
    print(neighbors[['atr_p', 'atr_m', 'Full_CAGR', 'Full_Calmar', 'Ret_2022']].to_string())

if __name__ == "__main__":
    check_island()
