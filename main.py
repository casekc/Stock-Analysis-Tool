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
from formatting import format_list, move_hyphen, split_list, extract_net_income, renameCSV, convertCSV, fix_excel_files
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from analyzer import analyzeNetIncome, analyzeHistoricalPrices
from compare import comparativeAnalysis
import subprocess


historical_price_analyses_dir = r'C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses'
net_income_analyses_dir = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses'

def clearFiles():
    csv_dir = r'C:\Users\cummi\Desktop\webscrap\bin\csv'
    historical_price_analyses_dir = r'C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses'
    net_income_analyses_dir = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses'
    xlsx_dir = r'C:\Users\cummi\Desktop\webscrap\bin\xlsx'
    comparative_analyses_dir = r'C:\Users\cummi\Desktop\webscrap\bin\comparative_analyses'
    hpa_converted_files_dir = r'C:\Users\cummi\Desktop\webscrap\bin\converted_files\historical_price_analyses'
    nia_converted_files_dir = r'C:\Users\cummi\Desktop\webscrap\bin\converted_files\net_income_analyses'

    os.remove(r'C:\Users\cummi\Desktop\webscrap\bin\net_income.xlsx')

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

clear_input = input("Would you like to clear files in the bin (Y for yes, N for no): ")
if str(clear_input) == 'Y':
    clearFiles()
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
    time.sleep(5)
    extractHistorical(ticker)

analyzeNetIncome(ticker_list)

renameCSV(ticker_list)

convertCSV()


analyzeHistoricalPrices(ticker_list)

hpa_directory = r'C:\Users\cummi\Desktop\webscrap\bin\historical_price_analyses'
nia_directory = r'C:\Users\cummi\Desktop\webscrap\bin\net_income_analyses'


script_path = r"C:\Users\cummi\Desktop\webscrap\bin\bash\convert_files.sh"

try:
    subprocess.call(["bash", script_path])
    print("Script executed successfully.")
except subprocess.CalledProcessError as e:
    print(f"Error executing script: {e}")

time.sleep(5)

comparativeAnalysis()