# imports
import urllib.request
from selenium import webdriver
import time
import requests
import urllib3
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


# Setting Chrome driver options
#options = webdriver.ChromeOptions()
#prefs = {"download.default_directory": "C:\\Users\\cummi\\Desktop\\webscrap\\bin\\csv"}

#options.add_experimental_option("prefs", prefs)

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


def stockInput():
    true = True

    global ticker_list
    ticker_list = []


    while true:
        ticker = input("Enter ticker (enter 0 to stop): ")
        if ticker == '0':
            true = False
            break
        else:
            ticker_list.append(ticker)
    return ticker_list

def extractFinancials(ticker):
    financials_url = 'https://www.nasdaq.com/market-activity/stocks/' + ticker + '/financials'
    driver = webdriver.Chrome()
    driver.get(financials_url)
    resp = driver.page_source
    driver.close()

    soup = BeautifulSoup(resp, 'html.parser')
    numbers_with_class = soup.find_all(class_='financials__cell')
    numbers_list = [number.get_text(strip=True) for number in numbers_with_class]

    return numbers_list

def extractHistorical(ticker):
    historical_url = 'https://www.nasdaq.com/market-activity/stocks/' + ticker + '/historical'
    global driver
    driver= setup_chrome_driver()

    driver.get(historical_url)

    wait = WebDriverWait(driver, 10)

    max_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@class="table-tabs__tab" and @data-value="y10"]')))

    max_button.click()

    action = ActionChains(driver)

    download_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'historical-data__controls-button--download')))
    action.move_to_element(download_button).perform()

    download_button.click()
    time.sleep(5)

    driver.close()








