import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from extractor import stockInput
import xlrd
import os
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
import zipfile
from openpyxl import load_workbook, Workbook

ticker_list = ['goog', 'pltr']

def analyzeHistoricalPrices(ticker_list):
    xlsx_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\xlsx\\'
    output_directory = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\historical_price_analyses\\'

    for ticker in ticker_list:
        # Load the existing workbook
        wb = load_workbook(xlsx_path + ticker + '.xlsx')
        ws = wb.active

        length = ws.max_row

        # Add the analysis headers
        ws['G1'] = 'Pco'
        ws['H1'] = 'Phl'
        ws['I1'] = 'AvgPco'
        ws['J1'] = 'AvgPhl'

        for i in range(2, length + 1):
            i = str(i)
            ws['G' + i] = '=B' + i + '-D' + i  # B2-D2
            ws['H' + i] = '=E' + i + '-F' + i  # E2-F2

        ws['I2'] = '=AVERAGE(G2:G' + str(length) + ')'
        ws['J2'] = '=AVERAGE(H2:H' + str(length) + ')'

        # Save the modified workbook to the output directory
        wb.save(output_directory + ticker + '.xlsx')
analyzeHistoricalPrices(ticker_list)


def nothing(ticker_list):
    xlsx_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\xlsx\\'

    for ticker in ticker_list:
        wb = None

        try:
            # Load the existing workbook using openpyxl
            wb = load_workbook(xlsx_path + ticker + '.xlsx', read_only=True, data_only=True)
        except KeyError:
            # If sharedStrings.xml is missing, try opening the archive and modifying the case of the filename
            archive = zipfile.ZipFile(xlsx_path + ticker + '.xlsx')
            shared_strings_path = None

            for file in archive.namelist():
                if file.lower() == 'xl/sharedstrings.xml':
                    shared_strings_path = file
                    break

            if shared_strings_path:
                archive.extract(shared_strings_path, xlsx_path)
                shared_strings_path = xlsx_path + shared_strings_path

                # Modify the case of the extracted sharedStrings.xml filename
                renamed_path = shared_strings_path.replace('xl/sharedstrings.xml', 'xl/SharedStrings.xml')
                os.rename(shared_strings_path, renamed_path)

                # Load the workbook with the modified sharedStrings.xml filename
                wb = load_workbook(xlsx_path + ticker + '.xlsx', read_only=True, data_only=True)

        if wb:
            ws = wb.active

            length = ws.max_row

            # Add the analysis headers
            ws['G1'] = 'Pco'
            ws['H1'] = 'Phl'
            ws['I1'] = 'AvgPco'
            ws['J1'] = 'AvgPhl'

            for i in range(2, length + 1):
                i = str(i)
                ws['G' + i] = '=B' + i + '-D' + i  # B2-D2
                ws['H' + i] = '=E' + i + '-F' + i  # E2-F2

            ws['I2'] = '=AVERAGE(G2:G' + str(length) + ')'
            ws['J2'] = '=AVERAGE(H2:H' + str(length) + ')'

            wb.save(r'C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses')

'''
def analyzeHistoricalPrices(ticker_list):
    xlsx_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\xlsx\\'

    for ticker in ticker_list:
        # Load the existing workbook in read-only mode
        wb = load_workbook(xlsx_path + ticker + '.xlsx', read_only=True, data_only=True)
        ws = wb.active

        length = ws.max_row

        # Add the analysis headers
        ws['G1'] = 'Pco'
        ws['H1'] = 'Phl'
        ws['I1'] = 'AvgPco'
        ws['J1'] = 'AvgPhl'

        for i in range(2, length + 1):
            i = str(i)
            ws['G' + i] = '=B' + i + '-D' + i  # B2-D2
            ws['H' + i] = '=E' + i + '-F' + i  # E2-F2

        ws['I2'] = '=AVERAGE(G2:G' + str(length) + ')'
        ws['J2'] = '=AVERAGE(H2:H' + str(length) + ')'

        wb.save()
'''


def analyzeNetIncome(ticker_list):
    # Loads the net income workbook
    wb = load_workbook(filename=r'C:\Users\cummi\Desktop\webscrap\bin\net_income.xlsx')

    # Loads the first sheet of the net income workbook
    ws = wb.active

    # Assigns cells F1:I1 the values of G1, G2, G3, and AGR, which means growth in period 1, 2, 3, and average annual
    # growth rate.
    ws['F1'] = 'G1'
    ws['G1'] = 'G2'
    ws['H1'] = 'G3'
    ws['I1'] = 'AAGR'

    # Iterates over each row and writes the growth rate formulas
    for i in range(2, len(ticker_list) + 2):
        i = str(i)
        ws['F' + i] = '=((D' + i + '- E' + i + ')/E' + i + ')*100'
        ws['G' + i] = '=((C' + i + '- D' + i + ')/D' + i + ')*100'
        ws['H' + i] = '=((B' + i + '- C' + i + ')/C' + i + ')*100'
        ws['I' + i] = '=AVERAGE(F' + i + ':H' + i + ')'

    # Saves the analysis to the net_income_analyses folder
    wb.save(r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx')
