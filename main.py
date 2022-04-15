import requests
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from category import *

categories = [fridge, washer_dryer, washing_machine, tv, soundbars, aircond, microwave_oven, cooking_hood, hob_stove]

# wb = Workbook()
# ws = wb.active

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()

# For each appliance category
for category in categories:
    # For each online store for the category
    for shop, webpage_list in category["website"].items():
        # For each search link for the online store
        for webpage in webpage_list:
            driver.get(webpage)
            if shop == 'senq':
                for _ in range(10):
                    driver.execute_script("window.scrollBy(0, document.body.scrollHeight)")
                    time.sleep(1)
                all_items = driver.find_elements(By.CLASS_NAME, 'grid-item')
                for item in all_items:
                    print(item.text)

# 1) get each items element
# 2) Go to next page after all items
# 3) create item class
# 4) for each item, detect for duplicates, if unique, create item object
# 5) Create workbook and worksheet for each category
# 6) Add new row for each unique item
