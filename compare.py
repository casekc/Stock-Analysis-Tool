import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook


def comparativeAnalysis():
    net_income_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx'
    historical_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses'

    # Net income analysis

    df = pd.read_excel(net_income_analysis_path, sheet_name=0)
    ticker_list = df['ticker'].tolist()
    AAGR_list = df['AAGR'].tolist()
    AvgPco_list = []
    AvgPhl_list = []

    # Historical analysis
    for i in os.listdir(r'C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses'):
        df = pd.read_excel(os.path.join(historical_analysis_path, i), sheet_name=0)
        global AvgPco
        global AvgPhl
        AvgPco = df['AvgPco'].tolist()
        AvgPhl = df['AvgPhl'].tolist()
        AvgPco_list.append(AvgPco[0])
        AvgPhl_list.append(AvgPhl[0])
    print(ticker_list)
    print(AAGR_list)
    print(AvgPco_list)
    print(AvgPhl_list)

    df_final = pd.DataFrame(
        {'Tickers': ticker_list, 'AAGR': AAGR_list, 'AvgPco': AvgPco_list, 'AvgPhl': AvgPhl_list}
    )

    df_final.to_excel(r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx')
