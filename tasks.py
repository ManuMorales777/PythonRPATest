import os
from datetime import datetime, timedelta
from openpyxl import Workbook
from robocorp.tasks import task
from robocorp import workitems, browser, vault
from selenium.webdriver.chrome.options import Options
from RPA.Robocorp.WorkItems import WorkItems
import time
import re
import requests
import shutil
from selenium import webdriver
from selenium.common import ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
class DateCalculator:
    def dateCalculate(actualDate, months):
        StartDate = actualDate.replace(day=1) 
        for _ in range(months):
         StartDate -= timedelta(days=StartDate.day)
        return StartDate, actualDate
class ExcelCreator:
    def __init__(self, filename):
        self.filename = filename
        self.workbook = Workbook()
        self.sheet = self.workbook.active

    def create_headers(self, headers):
        self.sheet.append(headers)

    def add_row(self, data):
        self.sheet.append(data)

    def save_file(self):
        output_dir = Path(os.environ.get('ROBOT_ARTIFACTS'))
        #output_dir = "C:/Users/manue/PycharmProjects/RPA_Challenge/output/"
        output_path = os.path.join(output_dir, self.filename)
        self.workbook.save(filename=output_path)
class FoxNewsSearch:
    def __init__(self):
        options = Options()
        options.add_argument("--no-sandbox")  
        options.add_argument("--disable-dev-shm-usage")  
        options.add_argument("--headless") 
        options.add_argument("--remote-debugging-port=9222")

        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 10)
    
    def click(self, xpath):
        time.sleep(2)
        self.driver.find_element(By.XPATH, xpath).click()
    def findbyXPath(self, xpath):
        self.driver.find_element(By.XPATH, xpath)
    def search(self, keyword):
        button_xpath="/html/body/div[3]/header/div[2]/div/div/div[2]/div[1]/a"
        search_input_xpath = "/html/body/div[3]/header/div[4]/div[1]/div/div/form/fieldset/input[1]"
        self.click(button_xpath)
        self.driver.find_element(By.XPATH, search_input_xpath).send_keys(keyword)
        self.driver.find_element(By.XPATH, search_input_xpath).send_keys(Keys.ENTER)

    def download_image(self, image_url, filename):
        output_dir = Path(os.environ.get('ROBOT_ARTIFACTS'))
        output_path = os.path.join(output_dir, filename)
        response = requests.get(image_url, stream=True)
        if response.status_code == 200:
            with open(output_path, 'wb') as f:
                response.raw.decode_content = True
                shutil.copyfileobj(response.raw, f)
            print(f"Image Downloaded '{output_path}'")
        else:
            print("Download error")

    def phrase_counter(self,textGiven, phraseGiven):
        text = textGiven.lower()
        phrase = phraseGiven.lower()
        text_words = text.split()
        phrase_words = phrase.split()
        counter = 0
        for i in range(len(text_words) - len(phrase_words) + 1):
            if text_words[i:i + len(phrase_words)] == phrase_words:
                counter += 1

        return counter

    def contains_money(self,textGiven):
        #Regular Expresion
        money_pattern = r'\$[\d,]+(?:\.\d+)?|\b\d+\s*(?:dollars|USD)\b'
        matches = re.findall(money_pattern, textGiven)
        return bool(matches)
    def close(self):
        time.sleep(10)
        self.driver.quit()
    
    
@task
def minimal_task():
    item = workitems.inputs.current
    print("Received payload:", item.payload)
    payload_actual = item.payload
    dateparameter = payload_actual.get('Month', '0')
    phraseToSearch = payload_actual.get('Phrase', 'Economy in LatinAmerica')
    actualdate = datetime.now()
    #dateparameter = 0
    #phraseToSearch = "Economy in LatinAmerica"
    if dateparameter < 0:
        print("Error.")
        return
    dateparameter = dateparameter - 1
    startDate, actualDate = DateCalculator.dateCalculate(actualdate, dateparameter)
    pastmonth = int(startDate.strftime("%m"))
    pastyear = int(startDate.strftime("%Y"))
    currentmonth = int(actualDate.strftime("%m"))
    currentday = int(actualDate.strftime("%d"))
    currentyear = int(actualDate.strftime("%Y"))
    yearIndex = (currentyear - pastyear) + 1
    print("Starting the automation")
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )
    secrets =  vault.get_secret('Rpa_Challenge')
    page = browser.goto(secrets['url'])
    page.click("//div[@class='header']/div[1]/a")
    
    page.fill("/html/body/div[3]/header/div[4]/div[1]/div/div/form/fieldset/input[1]").send_keys(send_enter = True)   
    time.sleep(3)
    #Month_From
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[1]/button")
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[1]/ul/li["+str(pastmonth)+"]")
    #Day_From
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[2]/button")
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[2]/ul/li["+str(1)+"]")
    #Year_From
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[3]/button")
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[1]/div[3]/ul/li[1]")
    #Month_To
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[1]/button")
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[1]/ul/li["+str(currentmonth)+"]")
    #Day_To
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[2]/button")
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[2]/ul/li["+str(currentday)+"]")
    #Year_From
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[3]/button")
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[2]/div[3]/div[2]/div[3]/ul/li["+str(yearIndex)+"]")
    #SearchButton
    bot.click("/html/body/div[1]/div/div/div[2]/div[1]/div/div[1]/div[2]/div/a")
    time.sleep(3)
    while True:
     try:
         bot.findbyXPath("/html/body/div[1]/div/div/div[2]/div[2]/div/div[3]/div[2]/a")
         bot.click("/html/body/div[1]/div/div/div[2]/div[2]/div/div[3]/div[2]/a")
         time.sleep(1)
     except ElementNotInteractableException:
        break
    print("Exit loading news")
    time.sleep(3)
    print("Start reading all news")
    newsAmount = int(bot.driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div[2]/div[1]/div/div[1]/div[1]/span[1]").text)
    print("Amount of News: " + str(newsAmount))
    print("Creating Excel Table")
    # Create the Excel File
    excel_creator = ExcelCreator('data.xlsx')
    headers = ['Title', 'Date', 'Description', 'Picture Filename', 'Count of Search Phrases', 'True or False']
    excel_creator.create_headers(headers)



    for i in range(1, newsAmount):
     elements = bot.driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div[2]/div[2]/div/div[3]/div[1]/article["+str(i)+"]").text
     #Get Picture Source
     picture = bot.driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div[2]/div[2]/div/div[3]/div[1]/article["+str(i)+"]/div[1]/a/picture/img")
     pictureSRC = picture.get_attribute("src")
     #Get Date Text
     dateText = bot.driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div[2]/div[2]/div/div[3]/div[1]/article["+str(i)+"]/div[2]/header/div/span[2]").text
     #Get Title
     titleText = bot.driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div[2]/div[2]/div/div[3]/div[1]/article["+str(i)+"]/div[2]/header/h2/a").text
     # Get Description
     descriptionText = bot.driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div[2]/div[2]/div/div[3]/div[1]/article["+str(i)+"]/div[2]/div/p/a").text
     #print(elements)
     print(pictureSRC)
     print(dateText)
     print(titleText)
     print(descriptionText)
     TitleCounter = bot.phrase_counter(titleText,phraseToSearch)
     DescriptionCounter = bot.phrase_counter(descriptionText, phraseToSearch)
     phraseCounter = TitleCounter + DescriptionCounter
     print("Phrase Counter: "+str(phraseCounter))
     print("Money in Title: "+str(bot.contains_money(titleText+descriptionText)))
     bot.download_image(pictureSRC,"img_"+str(i)+".jpg")
     # Añadir una fila con los datos de ejemplo
     data = [titleText, dateText, descriptionText, "img_"+str(i)+".jpg", phraseCounter, bot.contains_money(titleText+descriptionText)]
     excel_creator.add_row(data)
     print("--------------------------------------------------------------------------------------------------------------------------------")

    print("End reading all news")
    # Save Excel File
    excel_creator.save_file()
    bot.close()

