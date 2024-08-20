from RPA.Browser.Selenium import Selenium
from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
import os
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl


class Robot:
    def __init__(self):
        print("Starting robot")
        self.browser = self._initialize_browser()
        self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        self._setup_worksheet()

    def _initialize_browser(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920x1080")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--start-maximized")
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-web-security')
        options.page_load_strategy = 'eager'

        browser = Selenium(auto_close=False)
        return browser

    def _setup_worksheet(self):
        print("Creating sheet headers")
        self.sheet.append(["Title", "Description", "Date", "Image File", "Phrase Count", "Contains Money Reference"])

    def _open_browser(self):
        print("Opening browser")
        self.browser.open_available_browser(url="https://www.latimes.com/")

    def _search_phrase(self, phrase):
        print(f"Searching for phrase: {phrase}")
        self.browser.click_button_when_visible(locator='//html/body/ps-header/header/div[2]/button')
        self.browser.input_text_when_element_is_visible(locator='//html/body/ps-header/header/div[2]/div[2]/form/label/input', text=phrase)
        self.browser.click_button(locator='//html/body/ps-header/header/div[2]/div[2]/form/button')
        self.browser.wait_until_element_is_visible(locator='//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[1]/div[2]/div/label/select', timeout=10)
        self.browser.select_from_list_by_index('//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[1]/div[2]/div/label/select', '1')
        time.sleep(5)

    def _fetch_results(self):
        print("Fetching results")
        return self.browser.find_elements(locator='//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li')

    def _process_news_item(self, news, phrase, time_period):
        attempts = 0
        while attempts < 10:
            try:
                self.browser.wait_until_element_is_visible(locator='//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li', timeout=10)
                
                _date = self._get_news_date(news)
                if not self._is_within_time_period(_date, time_period):
                    return  # Skip this news item
                
                title, description, count_phrase, check_money = self._extract_news_info(news, phrase)
                image_filename = self._download_image(news, title)
                
                self._save_news_to_sheet(title, description, _date, image_filename, count_phrase, check_money)
                return
            except StaleElementReferenceException:
                print(f"Trying to find element. Number attempt: {attempts}")
                if attempts == 9:
                    raise Exception("Could not find element. Trying again")
                attempts += 1
                time.sleep(1)
            except Exception as e:
                print(f"Error processing news item: {str(e)}")
                return

    def _get_news_date(self, news):
        return news.find_element("css selector", ".promo-timestamp").get_attribute("data-timestamp")

    def _is_within_time_period(self, _date, time_period):
        if int(time_period) == 0 or int(time_period) == 1:
            date_period = datetime.today()
        else:
            date_period = datetime.today() - relativedelta(months=int(time_period) - 1)
        date_period = date_period.replace(day=1)
        timestamp_month = int(date_period.timestamp() * 1000)
        return int(_date) >= timestamp_month

    def _extract_news_info(self, news, phrase):
        title = news.find_element("css selector", ".promo-title a").text
        description = news.find_element("css selector", ".promo-description").text
        
        count_phrases_title = title.lower().count(phrase.lower())
        count_phrases_description = description.lower().count(phrase.lower())
        count_phrase = count_phrases_description + count_phrases_title
        
        check_money = any(keyword in title.lower() or keyword in description.lower() for keyword in ["$", "dollars", "usd"])
        return title, description, count_phrase, check_money

    def _download_image(self, news, title):
        image_element = news.find_element("xpath", './/div[@class="promo-media"]/a/picture/img')
        image_filename = f"./output/{title}.jpg"
        time.sleep(1)
        self.browser.screenshot(image_element, image_filename)
        return image_filename

    def _save_news_to_sheet(self, title, description, _date, image_filename, count_phrase, check_money):
        print("Saving data to sheet")
        self.sheet.append([title, description, _date, image_filename, count_phrase, check_money])

    def _save_workbook(self):
        print("Saving workbook")
        if not os.path.exists("./output"):
            os.makedirs("./output")
        self.workbook.save("./output/results.xlsx")
        print("Workbook saved")

    def start_robot(self, phrase: str, time_period: int):
        self._open_browser()
        self._search_phrase(phrase)
        results = self._fetch_results()
        for news in results:
            self._process_news_item(news, phrase, time_period)
        self._save_workbook()
