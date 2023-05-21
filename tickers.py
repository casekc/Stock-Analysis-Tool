import pandas as pd

technology_tickers_csv = r"C:\Users\cummi\Desktop\SmallCapPortfolio\nasdaq_screener_1684221162997.csv"
ds = pd.read_csv(technology_tickers_csv, sep=',', header=0)
technology_tickers = ds['Symbol'].tolist()

