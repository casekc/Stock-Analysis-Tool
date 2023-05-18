from openpyxl import Workbook
from openpyxl import load_workbook
#from main import ticker_list

wb = load_workbook(filename = r'C:\Users\cummi\Desktop\webscrap\bin\net_income.xlsx')

ws = wb.active

ws['F1'] = 'G1'


count = 2

for i in range(2, 4):
    i = str(i)
    ws['F' + i] = '=((D' + i + '+ E' + i + ')/E' + i + ')*100'
    ws['F' + i] = '=((D' + i + '+ E' + i + ')/E' + i + ')*100'

wb.save("net_income_new.xlsx")