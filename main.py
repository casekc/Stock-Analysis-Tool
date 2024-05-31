import os
import urllib.request
from selenium import webdriver
import time
import requests
import urllib3
from bs4 import BeautifulSoup
import pandas as pd
from extractor import stockInput, extractFinancials, extractHistorical
from formatting import format_list, move_hyphen, split_list, extract_net_income, renameCSV, convertCSV, fix_excel_files
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from analyzer import analyzeNetIncome, analyzeHistoricalPrices
from compare import comparativeAnalysis
import subprocess
from tickers import technology_tickers

def clearFiles():
    #os.remove()

    for f in os.listdir(csv_dir):
        os.remove(os.path.join(csv_dir, f))

    for f in os.listdir(historical_price_analyses_dir):
        os.remove(os.path.join(historical_price_analyses_dir, f))

    for f in os.listdir(net_income_analyses_dir):
        os.remove(os.path.join(net_income_analyses_dir, f))

    for f in os.listdir(xlsx_dir):
        os.remove(os.path.join(xlsx_dir, f))

    for f in os.listdir(comparative_analyses_dir):
        os.remove(os.path.join(comparative_analyses_dir, f))

    for f in os.listdir(hpa_converted_files_dir):
        os.remove(os.path.join(hpa_converted_files_dir, f))

    for f in os.listdir(nia_converted_files_dir):
        os.remove(os.path.join(nia_converted_files_dir, f))

clear_input = input("Would you like to clear files in the bin (y for yes, n for no): ")
if str(clear_input) == 'y':
    clearFiles()


sector_input = input("Would you like the technology sector small caps instead? (y for yes, n for no): ")
if str(sector_input) == 'y':
    ticker_list = technology_tickers
else:
    ticker_list = stockInput()
    print(ticker_list)

financials_list = []

for ticker in ticker_list:
    time.sleep(5)
    f = extractFinancials(ticker)
    financials_list.append(f)

net_income_2d_list = extract_net_income(financials_list)

for i in net_income_2d_list:
    del i[0]

data = net_income_2d_list
tickers = ticker_list

df = pd.DataFrame(data, index=tickers, columns=['2022', '2021', '2020', '2019'])
df.index.name = 'ticker'

export_dir = os.path.join(project_dir, 'bin')

os.makedirs(export_dir, exist_ok=True)

export_path = os.path.join(export_dir, 'net_income.xlsx')

df.to_excel(export_path)


def setup_chrome_driver():
    options = webdriver.ChromeOptions()

    prefs = {"download.default_directory":}
    options.add_experimental_option("prefs", prefs)

    service = Service(executable_path=driver_path)

    driver = webdriver.Chrome(service=service, options=options)

    return driver


driver = setup_chrome_driver()

for ticker in ticker_list:
    time.sleep(5)
    extractHistorical(ticker)

analyzeNetIncome(ticker_list)

renameCSV(ticker_list)

convertCSV()


analyzeHistoricalPrices(ticker_list)

try:
    subprocess.call(["bash", script_path])
    print("Script executed successfully.")
except subprocess.CalledProcessError as e:
    print(f"Error executing script: {e}")

time.sleep(5)

comparativeAnalysis()
