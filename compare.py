import os
import pandas as pd
import numpy as np
import xlrd
import openpyxl
from openpyxl import load_workbook, Workbook
from pyxlsb import open_workbook as open_xlsb
import xlwings as xw
import pythoncom

def comparativeAnalysis():
    net_income_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\converted_files\net_income_analyses\net_income_analysis.xlsx'
    historical_price_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\converted_files\historical_price_analyses\.xlsx'
    comparative_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx'

    wb = openpyxl.load_workbook(net_income_analysis_path, data_only=True)
    sheet = wb.active

    aagr_values = [cell.value for cell in sheet['I'][1:]]

    tickers = [cell.value for cell in sheet['A'][1:]]

    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active

    new_sheet['A1'] = 'Ticker'
    new_sheet['B1'] = 'AAGR'
    new_sheet['C1'] = 'PPI'
    new_sheet['D1'] = 'VI'

    for i, (ticker, aagr) in enumerate(zip(tickers, aagr_values), start=2):
        new_sheet['A{}'.format(i)] = ticker
        new_sheet['B{}'.format(i)] = aagr




    new_wb.save(comparative_analysis_path)

    historical_price_analysis_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\converted_files\\historical_price_analyses\\'

    data = []
    columns = ['PPI', 'VI']

    for i in os.listdir(historical_price_analysis_path):
        wb = openpyxl.load_workbook(historical_price_analysis_path + i, data_only=True)
        ws = wb.active
        AvgPco = ws['I2']
        AvgPhl = ws['J2']
        data.append([AvgPco.value, AvgPhl.value])

    df = pd.DataFrame(data, columns=columns)

    print(df)

    existing_wb = openpyxl.load_workbook(
        r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx')
    writer = pd.ExcelWriter(r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx',
                            engine='openpyxl')
    writer.book = existing_wb
    writer.sheets = dict((ws.title, ws) for ws in existing_wb.worksheets)
    df.to_excel(writer, index=False, sheet_name='Sheet', startcol=2, startrow=1, header=False)
    writer.save()
