from selenium import webdriver
from bs4 import BeautifulSoup


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


numbers_list = extractFinancials('goog')
print(numbers_list)