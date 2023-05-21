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

    # Load the net income analysis workbook using openpyxl with data_only option
    wb = openpyxl.load_workbook(net_income_analysis_path, data_only=True)
    sheet = wb.active

    # Get the calculated values of AAGR column
    aagr_values = [cell.value for cell in sheet['I'][1:]]

    # Extract the tickers and AAGR values
    tickers = [cell.value for cell in sheet['A'][1:]]

    # Create a new workbook and sheet using openpyxl
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active

    # Write the extracted data to the new sheet
    new_sheet['A1'] = 'Ticker'
    new_sheet['B1'] = 'AAGR'
    new_sheet['C1'] = 'PPI'
    new_sheet['D1'] = 'VI'

    for i, (ticker, aagr) in enumerate(zip(tickers, aagr_values), start=2):
        new_sheet['A{}'.format(i)] = ticker
        new_sheet['B{}'.format(i)] = aagr




    # Save the new workbook
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

    # Append the DataFrame to the existing sheet in comparative_analysis.xlsx using openpyxl
    existing_wb = openpyxl.load_workbook(
        r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx')
    writer = pd.ExcelWriter(r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx',
                            engine='openpyxl')
    writer.book = existing_wb
    writer.sheets = dict((ws.title, ws) for ws in existing_wb.worksheets)
    df.to_excel(writer, index=False, sheet_name='Sheet', startcol=2, startrow=1, header=False)
    writer.save()

def old_evaluate_formula(cell):
    # Evaluate the formula in the cell and return the calculated value
    calculated_value = cell.parent.parent.calculate_cell(cell.coordinate).value
    return calculated_value

def old_old_old_comparativeAnalysis():
    net_income_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx'
    comparative_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx'

    # Load the net income analysis workbook using openpyxl with data_only option
    wb = openpyxl.load_workbook(net_income_analysis_path, data_only=True)
    sheet = wb.active

    # Get the calculated values of AAGR column
    aagr_values = [cell.value for cell in sheet['I'][1:]]

    # Extract the tickers and AAGR values
    tickers = [cell.value for cell in sheet['A'][1:]]

    # Create a new workbook and sheet using openpyxl
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active

    # Write the extracted data to the new sheet
    new_sheet['A1'] = 'Ticker'
    new_sheet['B1'] = 'AAGR'
    for i, (ticker, aagr) in enumerate(zip(tickers, aagr_values), start=2):
        new_sheet['A{}'.format(i)] = ticker
        new_sheet['B{}'.format(i)] = aagr

    # Save the new workbook
    new_wb.save(comparative_analysis_path)


def old_old_comparativeAnalysis():
    net_income_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx'
    comparative_analysis_path = r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses\comparative_analysis.xlsx'

    # Load the net income analysis workbook
    wb = openpyxl.load_workbook(net_income_analysis_path)
    sheet_name = wb.active

    # Extract the required columns
    ticker_column = [cell.value for cell in sheet_name['A'][1:]]
    aagr_column = [cell.value for cell in sheet_name['I'][1:]]

    # Create a new workbook and sheet
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active

    # Write the extracted columns to the new sheet
    new_sheet['A1'] = 'Tickers'
    new_sheet['B1'] = 'AAGR'
    for i, (ticker, aagr) in enumerate(zip(ticker_column, aagr_column), start=2):
        new_sheet['A{}'.format(i)] = ticker
        new_sheet['B{}'.format(i)] = aagr

    # Save the new workbook
    new_wb.save(comparative_analysis_path)

