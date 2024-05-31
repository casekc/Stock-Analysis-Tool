import pandas as pd

ds = pd.read_csv(technology_tickers_csv, sep=',', header=0)
technology_tickers = ds['Symbol'].tolist()

