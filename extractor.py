import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

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
    driver= setup_chrome_driver()
    time.sleep(5)
    driver.get(historical_url)
    time.sleep(2)
    wait = WebDriverWait(driver, 10)
    time.sleep(2)
    max_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@class="table-tabs__tab" and @data-value="y10"]')))

    max_button.click()

    action = ActionChains(driver)

    download_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'historical-data__controls-button--download')))
    action.move_to_element(download_button).perform()

    download_button.click()
    time.sleep(5)

    driver.close()





