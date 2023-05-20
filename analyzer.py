from openpyxl import Workbook
from openpyxl import load_workbook
from extractor import stockInput


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



ticker_list = stockInput()

def analyzeHistoricalPrices(ticker_list):
    xlsx_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\xlsx\\'
    historical_analysis_path = 'C:\\Users\\cummi\\Desktop\\webscrap\\bin\\historical_price_analyses\\'
    for ticker in ticker_list:
        # Create a new workbook
        wb = load_workbook(xlsx_path + ticker + '.xlsx')
        ws = wb.active

        length = int(ws.max_row)

        ws['G1'] = 'Pco'
        ws['H1'] = 'Phl'
        ws['I1'] = 'AvgPco'
        ws['J1'] = 'AvgPhl'

        for i in range(2, length):
            i = str(i)
            ws['G' + i] = '=B' + i + '-D' + i  # B2-D2
            ws['H' + i] = '=E' + i + '-F' + i  # E2-F2

        ws['I2'] = '=AVERAGE(G2:G' + str(length)
        ws['J2'] = '=AVERAGE(H2:H' + str(length)

        wb.save(historical_analysis_path + ticker + '.xlsx')

analyzeHistoricalPrices(ticker_list)

