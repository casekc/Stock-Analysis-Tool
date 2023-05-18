# imports
import os
import urllib.request
from selenium import webdriver
import time
import requests
import urllib3
from bs4 import BeautifulSoup
import pandas as pd
from extractor import stockInput, extractFinancials, extractHistorical
from formatting import format_list, move_hyphen, split_list, extract_net_income
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


ticker_list = stockInput()
print(ticker_list)

financials_list = []

for ticker in ticker_list:
    f = extractFinancials(ticker)
    financials_list.append(f)

net_income_2d_list = extract_net_income(financials_list)

for i in net_income_2d_list:
    del i[0]

data = net_income_2d_list
tickers = ticker_list

df = pd.DataFrame(data, index=tickers, columns=[ '2022', '2021', '2020', '2019'])
df.index.name = 'ticker'

# Get the project directory
project_dir = r'C:\Users\cummi\Desktop\webscrap'

export_dir = os.path.join(project_dir, 'bin')

os.makedirs(export_dir, exist_ok=True)

export_path = os.path.join(export_dir, 'net_income.xlsx')

df.to_excel(export_path)

def setup_chrome_driver():
    # Specify the full path to chromedriver.exe
    driver_path = r'C:\path\to\chromedriver.exe'  # Replace with the actual path to chromedriver.exe

    # Set up the Chrome options
    options = webdriver.ChromeOptions()

    # Specify the default download directory
    prefs = {"download.default_directory": "C:\\Users\\cummi\\Desktop\\webscrap\\bin\\csv"}
    options.add_experimental_option("prefs", prefs)

    # Set up the Service object with the executable path
    service = Service(executable_path=driver_path)

    # Set up the Chrome driver with the Service and options
    driver = webdriver.Chrome(service=service, options=options)

    # Return the driver object
    return driver

# Call the function to set up the Chrome driver
driver = setup_chrome_driver()
for ticker in ticker_list:
    extractHistorical(ticker)






















''' 
output1 = format_list(financials_list)
output2 = move_hyphen(output1)
for sublist in output2:
    index = None  # Initialize the index variable
    if '-' in sublist:
        index = sublist.index('-')
    if index is not None and index + 1 < len(sublist):
        sublist[index] = '-' + sublist[index + 1]
        del sublist[index + 1]

# Modify the code to filter the desired output
desired_output = [output2[0], output2[2]]  # Select the first and third sublist

print(desired_output)

['Total Revenue', '$282,836,000', '$257,637,000', '$182,527,000', '$161,857,000', 'Cost of Revenue', '$126,203,000', '$110,939,000', '$84,732,000', '$71,896,000', 'Gross Profit', '$156,633,000', '$146,698,000', '$97,795,000', '$89,961,000', 'Operating Expenses', '', '', '', '', 'Research and Development', '$39,500,000', '$31,562,000', '$27,573,000', '$26,018,000', 'Sales, General and Admin.', '$42,291,000', '$36,422,000', '$28,998,000', '$29,712,000', 'Non-Recurring Items', '--', '--', '--', '--', 'Other Operating Items', '--', '--', '--', '--', 'Operating Income', '$74,842,000', '$78,714,000', '$41,224,000', '$34,231,000', "Add'l income/expense items", '-$3,514,000', '$12,020,000', '$6,858,000', '$5,394,000', 'Earnings Before Interest and Tax', '$71,328,000', '$90,734,000', '$48,082,000', '$39,625,000', 'Interest Expense', '--', '--', '--', '--', 'Earnings Before Tax', '$71,328,000', '$90,734,000', '$48,082,000', '$39,625,000', 'Income Tax', '$11,356,000', '$14,701,000', '$7,813,000', '$5,282,000', 'Minority Interest', '--', '--', '--', '--', 'Equity Earnings/Loss Unconsolidated Subsidiary', '--', '--', '--', '--', 'Net Income-Cont. Operations', '$59,972,000', '$76,033,000', '$40,269,000', '$34,343,000', 'Net Income', '$59,972,000', '$76,033,000', '$40,269,000', '$34,343,000', 'Net Income Applicable to Common Shareholders', '$59,972,000', '$76,033,000', '$40,269,000', '$34,343,000', 'Current Assets', '', '', '', '', 'Cash and Cash Equivalents', '$21,879,000', '$20,945,000', '$26,465,000', '$18,498,000', 'Short-Term Investments', '$91,883,000', '$118,704,000', '$110,229,000', '$101,177,000', 'Net Receivables', '$40,258,000', '$39,304,000', '$31,384,000', '$27,492,000', 'Inventory', '$2,670,000', '$1,170,000', '$728,000', '$999,000', 'Other Current Assets', '$8,105,000', '$8,020,000', '$5,490,000', '$4,412,000', 'Total Current Assets', '$164,795,000', '$188,143,000', '$174,296,000', '$152,578,000', 'Long-Term Assets', '', '', '', '', 'Long-Term Investments', '$30,492,000', '$29,549,000', '$20,703,000', '$13,078,000', 'Fixed Assets', '$127,049,000', '$110,558,000', '$96,960,000', '$84,587,000', 'Goodwill', '$28,960,000', '$22,956,000', '$21,175,000', '$20,624,000', 'Intangible Assets', '$2,084,000', '$1,417,000', '$1,445,000', '$1,979,000', 'Other Assets', '$6,623,000', '$5,361,000', '$3,953,000', '$2,342,000', 'Deferred Asset Charges', '$5,261,000', '$1,284,000', '$1,084,000', '$721,000', 'Total Assets', '$365,264,000', '$359,268,000', '$319,616,000', '$275,909,000', 'Current Liabilities', '', '', '', '', 'Accounts Payable', '$65,392,000', '$60,966,000', '$54,291,000', '$43,313,000', 'Short-Term Debt / Current Portion of Long-Term Debt', '--', '--', '--', '--', 'Other Current Liabilities', '$3,908,000', '$3,288,000', '$2,543,000', '$1,908,000', 'Total Current Liabilities', '$69,300,000', '$64,254,000', '$56,834,000', '$45,221,000', 'Long-Term Debt', '$14,701,000', '$14,817,000', '$13,932,000', '$4,554,000', 'Other Liabilities', '$24,006,000', '$22,770,000', '$22,264,000', '$22,633,000', 'Deferred Liability Charges', '$1,113,000', '$5,792,000', '$4,042,000', '$2,059,000', 'Misc. Stocks', '--', '--', '--', '--', 'Minority Interest', '--', '--', '--', '--', 'Total Liabilities', '$109,120,000', '$107,633,000', '$97,072,000', '$74,467,000', 'Stock Holders Equity', '', '', '', '', 'Common Stocks', '$68,184,000', '$61,774,000', '$58,510,000', '$50,552,000', 'Capital Surplus', '$195,563,000', '$191,484,000', '$163,401,000', '$152,122,000', 'Retained Earnings', '--', '--', '--', '--', 'Treasury Stock', '--', '--', '--', '--', 'Other Equity', '-$7,603,000', '-$1,623,000', '$633,000', '-$1,232,000', 'Total Equity', '$256,144,000', '$251,635,000', '$222,544,000', '$201,442,000', 'Total Liabilities & Equity', '$365,264,000', '$359,268,000', '$319,616,000', '$275,909,000', 'Net Income', '$59,972,000', '$76,033,000', '$40,269,000', '$34,343,000', 'Cash Flows-Operating Activities', '', '', '', '', 'Depreciation', '$15,928,000', '$12,441,000', '$13,697,000', '$11,781,000', 'Net Income Adjustments', '$17,830,000', '$4,701,000', '$9,331,000', '$7,577,000', 'Changes in Operating Activities', '', '', '', '', 'Accounts Receivable', '-$2,317,000', '-$9,095,000', '-$6,524,000', '-$4,340,000', 'Changes in Inventories', '--', '--', '--', '--', 'Other Operating Activities', '-$5,046,000', '-$1,846,000', '-$1,330,000', '-$621,000', 'Liabilities', '$5,128,000', '$9,418,000', '$9,681,000', '$5,780,000', 'Net Cash Flow-Operating', '$91,495,000', '$91,652,000', '$65,124,000', '$54,520,000', 'Cash Flows-Investing Activities', '', '', '', '', 'Capital Expenditures', '-$31,485,000', '-$24,640,000', '-$22,281,000', '-$23,548,000', 'Investments', '$16,567,000', '-$8,806,000', '-$9,822,000', '-$4,017,000', 'Other Investing Activities', '-$5,380,000', '-$2,077,000', '-$670,000', '-$1,926,000', 'Net Cash Flows-Investing', '-$20,298,000', '-$35,523,000', '-$32,773,000', '-$29,491,000', 'Cash Flows-Financing Activities', '', '', '', '', 'Sale and Purchase of Stock', '-$68,561,000', '-$60,126,000', '-$34,069,000', '-$22,941,000', 'Net Borrowings', '-$1,196,000', '-$1,236,000', '$9,661,000', '-$268,000', 'Other Financing Activities', '--', '--', '--', '--', 'Net Cash Flows-Financing', '-$69,757,000', '-$61,362,000', '-$24,408,000', '-$23,209,000', 'Effect of Exchange Rate', '-$506,000', '-$287,000', '$24,000', '-$23,000', 'Net Cash Flow', '$934,000', '-$5,520,000', '$7,967,000', '$1,797,000', 'Liquidity Ratios', '', '', '', '', 'Current Ratio', '238%', '293%', '307%', '337%', 'Quick Ratio', '234%', '291%', '305%', '335%', 'Cash Ratio', '164%', '217%', '241%', '265%', 'Profitability Ratios', '', '', '', '', 'Gross Margin', '55%', '57%', '54%', '56%', 'Operating Margin', '26%', '31%', '23%', '21%', 'Pre-Tax Margin', '25%', '35%', '26%', '24%', 'Profit Margin', '21%', '30%', '22%', '21%', 'Pre-Tax ROE', '28%', '36%', '22%', '20%', 'After Tax ROE', '23%', '30%', '18%', '17%']
'''




