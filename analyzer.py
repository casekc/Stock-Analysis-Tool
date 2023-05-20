from openpyxl import Workbook
from openpyxl import load_workbook
from extractor import stockInput

def analyzeNetIncome(ticker_list):
    # Loads the net income workbook
    wb = load_workbook(filename = r'C:\Users\cummi\Desktop\webscrap\bin\net_income.xlsx')

    # Loads the first sheet of the net income workbook
    ws = wb.active

    # Assigns cells F1:I1 the values of G1, G2, G3, and AGR, which means growth in period 1, 2, 3, and average annual
    # growth rate.
    ws['F1'] = 'G1'
    ws['G1'] = 'G2'
    ws['H1'] = 'G3'
    ws['I1'] = 'AAGR'

    # Iterates over each row and writes the growth rate formulas
    for i in range(2, len(ticker_list)+2):
        i = str(i)
        ws['F' + i] = '=((D' + i + '- E' + i + ')/E' + i + ')*100'
        ws['G' + i] = '=((C' + i + '- D' + i + ')/D' + i + ')*100'
        ws['H' + i] = '=((B' + i + '- C' + i + ')/C' + i + ')*100'
        ws['I' + i] = '=AVERAGE(F' + i + ':H' + i + ')'

    # Saves the analysis to the net_income_analyses folder
    wb.save(r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses\net_income_analysis.xlsx')


def analyzeHistoricalPrices(ticker_list):

    for ticker in ticker_list:
        wb = load_workbook(r'C:\Users\cummi\Desktop\webscrap\bin\csv\' + ' + ticker + '.csv')

        ws = wb.active

        ws['F1'] = 'Pco'
        ws['G1'] = 'Phl'

        for i in range(2, len(ticker_list) + 2):
            i = str(i)
            ws['F' + i] = '=F'
            ws['G' + i] = '=((C' + i + '- D' + i + ')/D' + i + ')*100'
            ws['H' + i] = '=((B' + i + '- C' + i + ')/C' + i + ')*100'
            ws['I' + i] = '=AVERAGE(F' + i + ':H' + i + ')'
