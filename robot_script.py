from RPA.Browser.Selenium import Selenium
import requests
from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException
import time
from datetime import datetime
from dateutil.relativedelta import relativedelta
import openpyxl

class Robot:
    def __init__(self):
        print("Starting robot")
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--disable-gpu")
        self.options.add_argument("--window-size=1920x1080")
        self.options.add_argument("--disable-extensions")
        self.options.add_argument("--disable-popup-blocking")
        self.options.add_argument("--start-maximized")
        self.options.add_argument('--no-sandbox')
        self.options.add_argument('--disable-web-security')
        self.options.page_load_strategy = 'eager'

        self.browser = Selenium(auto_close=False)
        self.news_data = list()
        
    def start_robot(self, phrase: str, time_period: int):
        print(f"Phrase: {phrase} | Time Period: {time_period}")
        
        print("Creating worksheet...")
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        
        print("Creating sheet headers")
        sheet.append(["Title", "Description", "Date", "Image File", "Phrase Count", "Contains Money Reference"])
        
        print("Opening browser")
        self.browser.open_available_browser(url="https://www.latimes.com/", options=self.options)
        
        print("Starting search...")
        self.browser.click_button_when_visible(locator='//html/body/ps-header/header/div[2]/button')
        self.browser.input_text_when_element_is_visible(locator='//html/body/ps-header/header/div[2]/div[2]/form/label/input', text=phrase)
        self.browser.click_button(locator='//html/body/ps-header/header/div[2]/div[2]/form/button')
        self.browser.wait_until_element_is_visible(locator='//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[1]/div[2]/div/label/select', timeout=10)
        self.browser.select_from_list_by_index('//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/div[1]/div[2]/div/label/select', '1')
        
        print("Waiting for page to load completely...")
        time.sleep(5)
        
        print("Loading results")
        results = self.browser.find_elements(locator='//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li')
        
        for news in results:
            attempts = 0
            # Page keeps updating sometimes, so tries 10 times to find element to avoid errors for not finding it
            while attempts < 10:
                try:
                    self.browser.wait_until_element_is_visible(locator='//html/body/div[2]/ps-search-results-module/form/div[2]/ps-search-filters/div/main/ul/li', timeout=10)
                    
                    # Number of month to receive news
                    # Example: time_period = 0 or 1: current month | time_period = 2: current and previous month | and so on
                    _date = news.find_element("css selector", ".promo-timestamp").get_attribute("data-timestamp")
                    if int(time_period) == 0 or int(time_period)==1:
                        date_period = datetime.today()
                    else:
                        date_period = datetime.today()-relativedelta(months=int(time_period)-1)
                    date_period = date_period.replace(day=1)
                    timestamp_month = int(date_period.timestamp() * 1000)
                    if int(_date) < timestamp_month:
                        attempts=10
                        continue
                    
                    # Getting title, date, description, picture filename, count of search phrases and check for amount of money
                    title = news.find_element("css selector", ".promo-title a").text
                    description = news.find_element("css selector", ".promo-description").text
                    
                    count_phrases_title = title.lower().count(phrase.lower())
                    count_phrases_description = description.lower().count(phrase.lower())
                    count_phrase = count_phrases_description+count_phrases_title
                    
                    if "$" in title or "dollars" in title.lower() or "usd" in title.lower() or "$" in description or "dollars" in description.lower() or "usd" in description.lower():
                        check_money = True
                    else:
                        check_money=False

                    image_url = news.find_element("xpath", './/div[@class="promo-media"]/a/picture/img').get_attribute('src')
                    
                    attempts = 10
                except StaleElementReferenceException:
                    print(f"Trying to find element. Number attempt: {attempts}")
                    if attempts == 9:
                        raise Exception("Couldnt find element. Trying again")
                    attempts += 1
                    time.sleep(1)
                except Exception as e:
                    print(f"Error to find element: {str(e)}")
                    break
                
                # Getting image and saving on outputs folder
                if image_url:
                    image_response = requests.get(image_url)
                    if image_response.status_code == 200:
                        with open(f"./outputs/{title}.jpg", "wb") as file:
                            file.write(image_response.content)

                print("Data loaded. Adding it to sheets")
                sheet.append([title, description, _date, f"{title}.jpg", count_phrase, check_money])

        print("Creating results.xlsx")
        workbook.save("./outputs/results.xlsx")
        print("Worksheet created")
