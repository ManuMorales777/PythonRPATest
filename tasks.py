import os
import time
import re
import logging
import requests
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys,ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from robocorp.tasks import task
from robocorp import workitems, browser, vault

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class DateCalculator:
    """This class contains all the logic 
    to calculate the time period input to search news between that days."""
    
    @staticmethod
    def calculate(actual_date, months):
        """Calculates the start date by subtracting 
        the given number of months."""
        logging.info("Calculating date from %d months ago.", months)
        start_date = actual_date.replace(day=1)
        for _ in range(months):
            start_date -= timedelta(days=start_date.day)
        return start_date, actual_date


class ExcelCreator:
    """Class to handle the creation and management of an Excel file."""
    
    def __init__(self, filename):
        self.filename = filename
        self.workbook = Workbook()
        self.sheet = self.workbook.active

    def create_headers(self, headers):
        """Create headers for the Excel sheet."""
        logging.info("Creating Excel headers: %s", headers)
        self.sheet.append(headers)

    def add_row(self, data):
        """Add a row of data to the Excel sheet."""
        logging.info("Adding row to Excel: %s", data)
        self.sheet.append(data)

    def save_file(self):
        """Save the Excel file to the specified directory."""
        output_dir = Path(os.environ.get('ROBOT_ARTIFACTS'))
        output_path = output_dir / self.filename
        logging.info("Saving Excel file to %s", output_path)
        self.workbook.save(filename=output_path)


class FoxNewsSearch:
    """Class to handle web scraping and search operations on Fox News."""
    @staticmethod
    def download_image(image_url, filename):
        """Download an image from the specified URL. Review if exists using the status code from Url
        and then download the image to the robot folder"""
        output_dir = Path(os.environ.get('ROBOT_ARTIFACTS'))
        output_path = output_dir / filename
        logging.info("Downloading image from %s to %s", image_url, output_path)
        response = requests.get(image_url, stream=True)
        if response.status_code == 200: 
            with open(output_path, 'wb') as f:
                shutil.copyfileobj(response.raw, f)
            logging.info("Image downloaded successfully.")
        else:
            logging.error("Error downloading image.")

    @staticmethod
    def phrase_counter(text, phrase):
        """Count occurrences of a phrase in the given text by spliting the word and the phrase to compare char by char if 
        the word is in the phrase"""
        logging.debug("Counting occurrences of phrase '%s' in text.", phrase)
        text_words = text.lower().split()
        phrase_words = phrase.lower().split()
        return sum(
            1 for i in range(len(text_words) - len(phrase_words) + 1)
            if text_words[i:i + len(phrase_words)] == phrase_words
        )

    @staticmethod
    def contains_money(text):
        """Check if the text contains any monetary references by using a Regex."""
        logging.debug("Checking for money references in text.")
        money_pattern = r'\$[\d,]+(?:\.\d+)?|\b\d+\s*(?:dollars|USD)\b'
        return bool(re.findall(money_pattern, text))

    def close(self):
        """Close the Selenium WebDriver."""
        logging.info("Closing WebDriver.")
        time.sleep(10)
        self.driver.quit()


@task
def minimal_task():
    logging.info("Starting minimal task.")
    
    # Retrieve work item payload
    item = workitems.inputs.current
    logging.info("Received payload: %s", item.payload)
    
    payload = item.payload
    # Set up of the variables to use, according to setups in robocorp
    date_parameter = payload.get('Month', '0')
    phrase_to_search = payload.get('Phrase', 'Economy')
    category_to_search = payload.get('Category', 'LatinAmerica')
    phrase_category_search = f"{phrase_to_search} in {category_to_search}"
    actual_date = datetime.now()
    if date_parameter < 0:
        logging.error("Invalid date parameter: %d", date_parameter)
        return

    # Calculate the date range
    date_parameter -= 1
    start_date, actual_date = DateCalculator.calculate(actual_date, date_parameter)
    past_month, past_year = start_date.month, start_date.year
    current_month, current_day, current_year = actual_date.month, actual_date.day, actual_date.year

    # Convert int variables to 00 String format
    past_month_formatted = f"0{past_month}" if past_month < 10 else str(past_month)
    current_month_formatted =  f"0{current_month}" if current_month < 10 else str(current_month)
    current_day_formatted =  f"0{current_day}" if current_day < 10 else str(current_day)

    # Init chrome
    logging.info("Configuring browser for automation.")
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--remote-debugging-port=9222")
    service = Service('/usr/bin/chromedriver')
    browser = webdriver.Chrome(service=service, options=chrome_options)
    # Retrieve secrets for authentication
    secrets = vault.get_secret('Rpa_Challenge')
    browser.get(secrets['url'])
    time.sleep(3)

    # Perform the search on the website
    logging.info("Performing search on the website.")
    browser.find_element(By.XPATH, "//div[@class='search-toggle tablet-desktop']/a[@class='js-focus-search']").click()
    text_input = browser.find_element(By.XPATH,"//input[@type='text' and (@aria-label='search foxnews.com' or @placeholder='Search foxnews.com') and @name='q']") 

    #Search the input phrase with his category
    ActionChains(browser)\
        .send_keys_to_element(text_input, phrase_category_search)\
        .perform()
    
    browser.find_element(By.XPATH,"//input[@type='submit' and @aria-label='submit search' and @class='resp_site_submit']").click()
    time.sleep(5)

    #Before clicking in the news searching output, we review if there is no Alert Banner in the page, if exists then close it
    try:
        close_button_element = browser.find_element(By.XPATH, "//div[@class='alert-banner is-breaking ']//a[@class='close']")
        close_button_element.click()
        logging.info("there was an alert message")
    except NoSuchElementException:
        logging.info("there is no alert message")
    
    """ Select Date Range (From and To)
    Clicks on each list and selects depending on dates
    """
    logging.info("Selecting date range for the search.")
    date_min_element = browser.find_element(By.XPATH, "//div[@class='date min']")
    date_min_element.find_element(By.XPATH, ".//div[@class='sub month']").click()
    date_min_element.find_element(By.XPATH, f".//div[@class='sub month']//ul[@name='select']//li[@id='{past_month_formatted}']").click()
    date_min_element.find_element(By.XPATH, ".//div[@class='sub day']").click()
    date_min_element.find_element(By.XPATH, ".//div[@class='sub day']//ul[@name='select']//li[@id='01']").click()
    date_min_element.find_element(By.XPATH, ".//div[@class='sub year']").click()
    date_min_element.find_element(By.XPATH, f".//div[@class='sub year']//ul[@name='select']//li[@id='{start_date.year}']").click()
    time.sleep(1)
    date_max_element = browser.find_element(By.XPATH, "//div[@class='date max']")
    date_max_element.find_element(By.XPATH, ".//div[@class='sub month']").click()
    date_max_element.find_element(By.XPATH, f".//div[@class='sub month']//ul[@name='select']//li[@id='{current_month_formatted}']").click()
    date_max_element.find_element(By.XPATH, ".//div[@class='sub day']").click()
    date_max_element.find_element(By.XPATH, f".//div[@class='sub day']//ul[@name='select']//li[@id='{current_day_formatted}']").click()
    date_max_element.find_element(By.XPATH, ".//div[@class='sub year']").click()
    date_max_element.find_element(By.XPATH, f".//div[@class='sub year']//ul[@name='select']//li[@id='{actual_date.year}']").click()

    # Start searching
    logging.info("Starting search.")
    browser.find_element(By.XPATH,"//div[@class='button']/a[text()='Search']").click()
    time.sleep(3)

    #In this loop we collapse all the articles to retrieve all the info of all news within given days
    logging.info("Entering the loop to load more news.")
    while True:
        try:
            load_more_button = browser.find_element(By.XPATH, "//span[text()='Load More']")
            if not load_more_button.is_displayed():
                break
            load_more_button.click()
            time.sleep(3)
        except Exception as e:
            logging.error(f"An error occurred: {e}")
            break

    logging.info("Finished loading news.")

    """ Get the number of news articles resulting
    """
    news_amount = int(browser.find_element(By.XPATH,"//div[@class='num-found']/span[2]/span").text)
    logging.info("Amount of news articles found: %d", news_amount)

    # Create Excel table
    logging.info("Creating Excel table.")
    excel_creator = ExcelCreator('data.xlsx')
    headers = ['Title', 'Date', 'Description', 'Picture Filename', 'Count of Search Phrases', 'Contains Money']
    excel_creator.create_headers(headers)
    
    # Find all articles using XPath
    articles = browser.find_elements(By.XPATH, "//article[@class='article']")
    
    # Variable to handle images dynamic names
    i=1 
    
    # Loop through each article
    for article in articles:

        # Find elements within the current article using relative XPaths
        title_element = article.find_element(By.XPATH, ".//h2/a")
        image_element = article.find_element(By.XPATH, ".//div[@class='m']//img")
        description_element = article.find_element(By.XPATH, ".//div[@class='info']//div[@class='content']//p[@class='dek']")
        date_element = article.find_element(By.XPATH, ".//div[@class='info']//header[@class='info-header']//div[@class='meta']//span[@class='time']")

        # Extract data from elements
        title = title_element.text
        image_src = image_element.get_attribute("src")
        description = description_element.text
        date = date_element.text

        # Count occurrences of the search phrase in the title and description
        phrase_counter = (
            FoxNewsSearch.phrase_counter(title, phrase_to_search) +
            FoxNewsSearch.phrase_counter(description, phrase_to_search)
        )
        contains_money = FoxNewsSearch.contains_money(title + description)

        # Download the associated image
        FoxNewsSearch.download_image(image_src, f"img_{i}.jpg")

        # Add the extracted data to the Excel sheet
        data = [title, date, description, f"img_{i}.jpg", phrase_counter, contains_money]
        excel_creator.add_row(data)
        
        #Increment the counter of images
        i += 1
        
        logging.debug("Processed article: %s", title)
    
    logging.info("Completed processing all news articles.")
    
    # Save the Excel file
    excel_creator.save_file()
    logging.info("Excel saved succesfully.")
    # Close the page
    page.close()
    logging.info("RPA task completed.")
