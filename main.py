import time
import openpyxl.drawing.image
import urllib3
import io
import math
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
try:
    from openpyxl.cell import get_column_letter
except ImportError:
    from openpyxl.utils import get_column_letter
    from openpyxl.utils import column_index_from_string
from category import *


PRODUCT_BRAND_COL = 1
PRODUCT_MODEL_COL = 2
PRODUCT_DETAILS_COL = 3
PRODUCT_IMAGE_COL = 4
PRODUCT_IMAGE_COL_LETTER = 'D'
SENQ_PRICE_COL = 5
HN_PRICE_COL = 6
BHB_PRICE_COL = 7

categories = [fridge, washer_dryer, washing_machine, tv, soundbars, aircond, microwave_oven, cooking_hood, hob_stove]

wb = Workbook()
wb.save('Text.xlsx')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()


def senq_scrape():
    global excel_row_num
    # Scroll down 10 times
    for _ in range(20):
        driver.execute_script("window.scrollBy(0, document.body.scrollHeight)")
        time.sleep(1)

    # Get all products in product page
    all_products = driver.find_elements(By.CLASS_NAME, 'grid-item')
    for product in all_products:
        product_full_name = product.find_element(By.CLASS_NAME, 'MuiTypography-body1').text
        product_brand = product_full_name.split(' ')[0]

        # If product brand is included in category list
        if product_brand in category["brands"]:
            # Product Brand
            ws.cell(row=excel_row_num, column=PRODUCT_BRAND_COL, value=product_brand)

            # Product Model number
            product_model = product_full_name.split(' ')[-1]
            ws.cell(row=excel_row_num, column=PRODUCT_MODEL_COL, value=product_model)

            # Product Details
            product_details = ' '.join(product_full_name.split(' ')[1:-1])
            ws.cell(row=excel_row_num, column=PRODUCT_DETAILS_COL, value=product_details)

            # Product Image
            product_image_src = product.find_element(By.CLASS_NAME, 'img-bg-load').get_attribute('src')
            http = urllib3.PoolManager()
            # Get image from URL
            r = http.request('GET', product_image_src)
            image_file = io.BytesIO(r.data)
            # Insert image into excel cell
            img = openpyxl.drawing.image.Image(image_file)
            img.height = 220
            img.width = 300
            img.anchor = ws.cell(row=excel_row_num, column=PRODUCT_IMAGE_COL).coordinate
            ws.add_image(img)

            # Product Price
            try:
                product_price = product.find_element(By.TAG_NAME, 'strike').text
            except NoSuchElementException:
                product_price = product.find_element(By.CLASS_NAME, 'price_text').text
            chars = ['RM ', ',']
            for char in chars:
                product_price = product_price.replace(char, '')
                try:
                    product_price = float(product_price)
                except ValueError:
                    pass
            ws.cell(row=excel_row_num, column=SENQ_PRICE_COL, value=product_price)

            # Autofit rows
            ws.row_dimensions[excel_row_num].height = 165

            # Next product (row)
            excel_row_num += 1


def hn_scrape():
    global excel_row_num

    # Number of pages based on total product count
    total_products = int(driver.find_element(By.CLASS_NAME, 'pagination-amount').text.split(' ')[0])
    no_of_pages = math.ceil(total_products / 20)

    # While next page is available
    for page in range(no_of_pages):

        # Get all products in product page
        all_products = driver.find_elements(By.CLASS_NAME, 'product-col')
        for product in all_products:
            product_full_name = product.find_element(By.CLASS_NAME, 'product-info').text
            product_brand = product_full_name.split(' ')[0]

            # If product brand is included in category list
            if product_brand in category["brands"]:
                # Product Brand
                ws.cell(row=excel_row_num, column=PRODUCT_BRAND_COL, value=product_brand)

                # Product Model number
                product_model = product_full_name.split(' ')[1]
                ws.cell(row=excel_row_num, column=PRODUCT_MODEL_COL, value=product_model)

                # Product Details
                product_details = ' '.join(product_full_name.split(' ')[2:])
                ws.cell(row=excel_row_num, column=PRODUCT_DETAILS_COL, value=product_details)

                # Product Image
                product_image_src = product.find_element(By.TAG_NAME, 'img').get_attribute('src')
                http = urllib3.PoolManager()
                # Get image from URL
                r = http.request('GET', product_image_src)
                image_file = io.BytesIO(r.data)
                # Insert image into excel cell
                img = openpyxl.drawing.image.Image(image_file)
                img.height = 210
                img.width = 290
                img.anchor = ws.cell(row=excel_row_num, column=PRODUCT_IMAGE_COL).coordinate
                ws.add_image(img)

                # Product Price
                try:
                    product_price = product.find_elements(By.CLASS_NAME, 'list-price')[-1].text
                except IndexError:
                    product_price = product.find_elements(By.CLASS_NAME, 'price-num')[-1].text
                product_price = product_price.replace(',', '')
                try:
                    product_price = float(product_price)
                except ValueError:
                    pass
                ws.cell(row=excel_row_num, column=HN_PRICE_COL, value=product_price)

                # Autofit rows
                ws.row_dimensions[excel_row_num].height = 165

                # Next product (row)
                excel_row_num += 1

        if page < (no_of_pages - 1):
            # Go to next page
            next_page_clicked = False
            while not next_page_clicked:
                try:
                    next_page = driver.find_elements(By.CSS_SELECTOR, '#pagination_contents li')[-1]
                    next_page.click()
                    time.sleep(2)
                    next_page_clicked = True
                # Close pop up if it appears
                except ElementClickInterceptedException:
                    pop_up = driver.find_element(By.XPATH, '/html/body/div[9]/iframe')
                    driver.switch_to.frame(pop_up)
                    driver.find_element(By.XPATH, '//*[@id="icon-close-button-1523407214825"]').click()
                    driver.switch_to.default_content()


# For each appliance category
for category in categories:
    ws = wb.create_sheet(category["name"])
    excel_row_num = 4
    # For each online store for the category
    for shop, webpage_list in category["website"].items():
        # For each search link for the online store
        for webpage in webpage_list:
            driver.get(webpage)
            # For SenQ Website
            if shop == 'senq':
                senq_scrape()
            elif shop == "harvey_norman":
                hn_scrape()

    # Autofit columns
    for row_cells in ws.columns:
        new_row_length = max(len(str(cell.value)) for cell in row_cells)
        new_column_letter = (get_column_letter(row_cells[0].column))
        if new_row_length > 0:
            ws.column_dimensions[new_column_letter].width = new_row_length * 1.23
    # Set column width for image column
    ws.column_dimensions[PRODUCT_IMAGE_COL_LETTER].width = 42.14

    wb.save('Text.xlsx')


# 1) get each items element
# 2) Go to next page after all items
# 3) create item class
# 4) for each item, detect for duplicates, if unique, create item object
# 5) Create workbook and worksheet for each category
# 6) Add new row for each unique item
