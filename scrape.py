"""_summary_

    Returns:
        _type_: _description_

    Yields:
        _type_: _description_
    """
import time
import os
import shutil

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import yfinance as yf

from constants import *
from scoring import *
from credentials import *


class Scraper:
    """_summary_
    """

    def __init__(self):
        self.logged_in = False
        options = webdriver.ChromeOptions()
        prefs = {
            'download.default_directory': os.getcwd(),
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': False,
            'safebrowsing.disable_download_protection': True,
            "profile.default_content_settings.popups": 0
        }

        options.add_experimental_option('prefs', prefs)
        options.add_experimental_option(
            "excludeSwitches", ['enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-setuid-sandbox')
        options.add_argument('--start-maximized')
        options.add_argument('--ignore-ssl-errors=yes')
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--headless')
        options.add_argument('--log-path=./chromedriver.log')
        user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.53 Safari/537.36'
        options.add_argument(f'user-agent={user_agent}')
        service = ChromeService(executable_path=CHROMEDRIVER_PATH)
        self.driver = webdriver.Chrome(service=service, options=options)

        self.driver.command_executor._commands["send_command"] = (
            "POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {
            'behavior': 'allow', 'downloadPath': os.getcwd()}}
        command_result = self.driver.execute("send_command", params)

    def __del__(self):
        self.driver.quit()

    def login(self):
        """_summary_
        """
        self.driver.get("https://stockanalysis.com")
        self.driver.set_window_size(2000, 2000)
        login_button = self.driver.find_elements(
            By.XPATH, "//header/div/div[2]/a")[0]
        login_button.click()

        time.sleep(0.5)
        [user, passwd] = self.driver.find_elements(
            By.XPATH, "//main/div/form/input")

        user.send_keys(EMAIL)
        passwd.send_keys(PASSWD)

        form = self.driver.find_elements(By.XPATH, "//main/div/form")
        form[0].submit()
        self.logged_in = True

    def download_price_chart(self, ticker):
        """_summary_

        Args:
            ticker (_type_): _description_
        """
        self.driver.get(
            f"https://stockanalysis.com/stocks/{ticker.lower()}/chart/")
        self.driver.set_window_size(2000, 2000)

        try:
            export_button = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//span[text() = 'Export']/parent::button")))
            export_button.click()

            export_button = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//button[text() = 'Export to Excel']")))
            export_button.click()
        except TimeoutException:
            print(f"Could not download price chart for {ticker}")
            self.driver.save_screenshot('fail.png')

    def download_financial_data(self, ticker):
        """_summary_

        Args:
            ticker (_type_): _description_
        """
        url = f"https://stockanalysis.com/stocks/{ticker.lower()}/financials"
        self.driver.get(url)
        self.driver.set_window_size(2000, 2000)
        try:
            export_dropdown = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//span[text()='Export']/parent::button")))
            export_dropdown.click()
            bulk_export_button = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//main//button[text()='Bulk Export']")))
            bulk_export_button.click()
        except TimeoutException:
            print(f"Could not download financial data for {ticker}")
            self.driver.save_screenshot('fail.png')

    def download_historical_data(self, ticker):
        """_summary_

        Args:
            ticker (_type_): _description_
        """
        folder = os.path.join(os.getcwd(), "data", ticker)
        financial_data_excel = f"{ticker.lower()}-financials.xlsx"
        chart_excel = f"{ticker.lower()}-chart.xlsx"
        dividend_excel = f"{ticker.lower()}-dividends.xlsx"

        if not os.path.exists(folder):
            os.makedirs(folder)

        print(f"--- Downloading historical data for {ticker} ----")

        if not os.path.exists(os.path.join(folder, financial_data_excel)):
            print('Financial Data...', end='')
            self.download_financial_data(ticker)
            print('OK')
        if not os.path.exists(os.path.join(folder, chart_excel)):
            print('Price Chart...', end='')
            self.download_price_chart(ticker)
            print('OK')
        if not os.path.exists(os.path.join(folder, dividend_excel)):
            print('Dividend History...', end='')
            self.download_dividend_history(ticker)
            print('OK')

        while any([f.endswith('crdownload') for f in os.listdir(os.getcwd())]):
            continue

        time.sleep(0.5)

        if os.path.exists(financial_data_excel):
            shutil.move(financial_data_excel, folder)
        if os.path.exists(chart_excel):
            shutil.move(chart_excel, folder)
        if os.path.exists(dividend_excel):
            shutil.move(dividend_excel, folder)

    def download_company_info(self, ticker):
        """_summary_

        Args:
            ticker (_type_): _description_

        Returns:
            _type_: _description_

        Yields:
            _type_: _description_
        """
        url = f"https://stockanalysis.com/stocks/{ticker}/"
        self.driver.get(url)
        overview = self.driver.find_elements(By.XPATH, "//main/div/div")
        company_name_w_ticker = overview[0].text.split('\n')[0]
        price = float(overview[1].text.split(' ')[0].split('\n')[0])
        div_and_yield = self.driver.find_elements(
            By.XPATH, "//main/div[2]/div/table/tbody/tr[8]")
        div = float(div_and_yield[0].text.split('(')[0].split('$')[1].strip())
        yld = float(div_and_yield[0].text.split('(')[1][:-2])

        beta = self.driver.find_element(
            By.XPATH, "//main/div[2]/div/table[2]/tbody/tr[6]")
        beta = float(beta.text.split(' ')[1])

        return {
            NAME: company_name_w_ticker.split('(')[0].strip(),
            TICKER: ticker,
            CURRENT_PRICE: price,
            CURRENT_DPS: div,
            CURRENT_YLD: yld,
            CREDIT_RATING: 'AAA',
            BETA: beta,
        }

    def year_newest_report(self, ticker):
        """_summary_

        Args:
            ticker (_type_): _description_

        Returns:
            _type_: _description_
        """
        url = f"https://stockanalysis.com/stocks/{ticker.lower()}/financials"
        self.driver.get(url)
        self.driver.set_window_size(2000, 2000)
        last_year_available = self.driver.find_element(
            By.XPATH, "//table/thead//th[2]").text
        return last_year_available

    def download_current_data(self, ticker):
        """_summary_

        Args:
            ticker (_type_): _description_

        Returns:
            _type_: _description_
        """
        tkr = yf.Ticker(ticker)
        name = tkr.info['longName']
        dps = tkr.info['dividendRate'] / 4
        price = tkr.history(period="1d").Close[0]
        yld = dps * 4 / price
        beta = tkr.info['beta']
        sector = tkr.info['sector']
        return {
            NAME: name,
            TICKER: ticker,
            CURRENT_PRICE: price,
            CURRENT_DPS: dps,
            CURRENT_YLD: yld,
            CREDIT_RATING: 'AAA',
            BETA: beta,
            SECTOR: sector
        }

    def download_dividend_history(self, ticker):
        """_summary_

        Args:
            ticker (_type_): _description_
        """
        url = f"https://dividendhistory.org/payout/{ticker}/"
        self.driver.get(url)
        self.driver.set_window_size(2000, 2000)

        # Read page by page
        payout_dates = []
        amounts = []

        try:
            nxt = self.driver.find_element(By.XPATH,
                                           "//a[@id='dividend_table_next']")
            while 'disabled' not in nxt.get_attribute('class').split():
                pdd = self.driver.find_elements(By.XPATH,
                                                "//table[@id='dividend_table']/tbody/tr/td[2]")
                payout_dates.extend(map(lambda x: x.text, pdd))

                am = self.driver.find_elements(By.XPATH,
                                               "//table[@id='dividend_table']/tbody/tr/td[3]")
                amounts.extend(map(lambda x: x.text, am))

                nxt.click()

                nxt = self.driver.find_element(By.XPATH,
                                               "//a[@id='dividend_table_next']")

            # Read last page
            pdd = self.driver.find_elements(By.XPATH,
                                            "//table[@id='dividend_table']/tbody/tr/td[2]")
            payout_dates.extend(map(lambda x: x.text, pdd))

            am = self.driver.find_elements(By.XPATH,
                                           "//table[@id='dividend_table']/tbody/tr/td[3]")

            amounts.extend(map(lambda x: x.text, am))

            amounts = list(map(lambda x: x.replace(
                "$", "").replace("*", ""), amounts))

            df = pd.DataFrame(list(zip(payout_dates, amounts)),
                              columns=["Date", "Amount"])
            df["Date"] = df["Date"].astype("datetime64")
            df["Amount"] = df["Amount"].astype("float")

            file_path = os.path.join(
                os.getcwd(), "data", ticker, f"{ticker.lower()}-dividends.xlsx")
            df.to_excel(file_path, index=False)
        except NoSuchElementException:
            print(f'Failed to download dividend history for {ticker}')
            self.driver.save_screenshot('fail.png')
